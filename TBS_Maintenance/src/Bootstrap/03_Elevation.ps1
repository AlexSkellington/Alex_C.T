# ===================================================================================================
#                                  SECTION: Ensure Admin Execution
# ---------------------------------------------------------------------------------------------------
# Description:
#   Relaunches the script elevated ONLY if it is not already running as Administrator.
#     - Already admin            -> continues normally (no relaunch).
#     - Saved .ps1 / packaged .exe, interactive session -> relaunches itself with RunAs and the
#       same arguments, then exits this (non-elevated) instance.
#     - Inline (irm | iex), worker mode, or a non-interactive session -> left alone.
#   If the UAC prompt is cancelled the script continues without elevation.
#
# Variables:
#   - $IsAdmin   [bool] running elevated (read later by the startup code)
#   - Reads the launch facts from 01_Launch_Detection ($__isSavedPs1Launch, $__isExeLaunch, ...)
# ===================================================================================================

$IsAdmin = $false
try
{
	$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
		[Security.Principal.WindowsBuiltInRole]::Administrator)
}
catch { $IsAdmin = $false }

if (-not $IsAdmin -and [Environment]::UserInteractive -and ($__isSavedPs1Launch -or $__isExeLaunch))
{
	if ($script:ShowConsole) { Write-Host "Requesting elevated launch..." -ForegroundColor DarkYellow }

	# Start-Process joins -ArgumentList with spaces and does not quote, so quote anything with
	# whitespace (script paths such as "...\TBS Scripts\..." would otherwise be split).
	$__quoteArg = {
		param ([string]$Value)
		if ($Value -match '\s' -and $Value -notmatch '^".*"$') { return '"' + $Value + '"' }
		return $Value
	}

	try
	{
		if ($__isSavedPs1Launch)
		{
			$__elevArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (& $__quoteArg $__scriptPath))
			foreach ($__a in @($args)) { $__elevArgs += (& $__quoteArg ([string]$__a)) }

			$__psExe = $null
			try { $__psExe = (Get-Command powershell.exe -ErrorAction SilentlyContinue).Source }
			catch { }
			if ([string]::IsNullOrWhiteSpace($__psExe))
			{
				try { $__psExe = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Source }
				catch { }
			}
			if ([string]::IsNullOrWhiteSpace($__psExe) -or -not (Test-Path -LiteralPath $__psExe)) { $__psExe = "powershell.exe" }

			$null = Start-Process -FilePath $__psExe -ArgumentList $__elevArgs -Verb RunAs -ErrorAction Stop
			exit 0
		}

		if ($__isExeLaunch)
		{
			# -ArgumentList rejects an empty array, so only pass it when there are arguments
			if (@($__launchArgs).Count -gt 0)
			{
				$__exeArgs = @($__launchArgs | ForEach-Object { & $__quoteArg ([string]$_) })
				$null = Start-Process -FilePath $__entryPath -ArgumentList $__exeArgs -Verb RunAs -ErrorAction Stop
			}
			else
			{
				$null = Start-Process -FilePath $__entryPath -Verb RunAs -ErrorAction Stop
			}
			exit 0
		}
	}
	catch
	{
		if ($script:ShowConsole) { Write-Host "Elevation canceled or unavailable. Continuing." -ForegroundColor DarkYellow }
	}
}

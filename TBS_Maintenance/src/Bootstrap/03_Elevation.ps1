# ===================================================================================================
#                               SECTION: Ensure Admin Execution
# ---------------------------------------------------------------------------------------------------
# Description:
#   Relaunches the script elevated ONLY if not already running as Administrator.
#   If already admin -> continues normally (NO relaunch).
#   Preserves script path and original arguments.
#   Safe for PowerShell 5.1 and Scheduled Tasks.
# ===================================================================================================

try
{
	$IsAdmin = ([Security.Principal.WindowsPrincipal] `
		[Security.Principal.WindowsIdentity]::GetCurrent()
	).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
catch
{
	$IsAdmin = $false
}

$scriptPath = $null
$entryPath = $null
$launchArgs = @()
$isSavedPs1Launch = $false
$isExeLaunch = $false

try
{
	$scriptPath = $PSCommandPath
	if ([string]::IsNullOrWhiteSpace($scriptPath))
	{
		try { $scriptPath = $MyInvocation.MyCommand.Path }
		catch { $scriptPath = $null }
	}
}
catch
{
	$scriptPath = $null
}

try
{
	$cmdLineArgs = [Environment]::GetCommandLineArgs()
	if ($cmdLineArgs -and $cmdLineArgs.Count -gt 0)
	{
		$entryPath = $cmdLineArgs[0]
		if ($cmdLineArgs.Count -gt 1)
		{
			$launchArgs = @($cmdLineArgs[1 .. ($cmdLineArgs.Count - 1)])
		}
	}
}
catch
{
	$entryPath = $null
	$launchArgs = @()
}

try
{
	if (-not [string]::IsNullOrWhiteSpace($scriptPath) -and
		(Test-Path -LiteralPath $scriptPath) -and
		([string]::Equals([System.IO.Path]::GetExtension($scriptPath), '.ps1', [System.StringComparison]::OrdinalIgnoreCase)))
	{
		$isSavedPs1Launch = $true
		$script:LaunchSourceType = 'Ps1'
		$script:LaunchEntryPath = $scriptPath
	}
}
catch
{
	$isSavedPs1Launch = $false
}

if (-not $isSavedPs1Launch)
{
	try
	{
		if (-not [string]::IsNullOrWhiteSpace($entryPath) -and
			(Test-Path -LiteralPath $entryPath) -and
			([string]::Equals([System.IO.Path]::GetExtension($entryPath), '.exe', [System.StringComparison]::OrdinalIgnoreCase)))
		{
			$entryName = [System.IO.Path]::GetFileNameWithoutExtension($entryPath)
			if ($entryName -notmatch '^(powershell|pwsh|powershell_ise|powershellstudio)$')
			{
				$isExeLaunch = $true
				$script:LaunchSourceType = 'Exe'
				$script:LaunchEntryPath = $entryPath
			}
		}
	}
	catch
	{
		$isExeLaunch = $false
	}
}

if (-not $isSavedPs1Launch -and -not $isExeLaunch)
{
	$script:LaunchSourceType = 'Inline'
	$script:IsInlineLaunch = $true
}

# Only attempt self-relaunch when this is a real saved, interactive .ps1 or packaged .exe launch.
# If the script was started from irm|iex or some other host/launcher, leave execution alone.
if (-not $IsAdmin -and -not $script:LaneSqlEnableWorkerMode -and [Environment]::UserInteractive -and ($isSavedPs1Launch -or $isExeLaunch))
{
	if ($script:ShowConsole) { Write-Host "Requesting elevated launch..." -ForegroundColor DarkYellow }
	try
	{
		if ($isSavedPs1Launch)
		{
			$argList = @(
				"-NoProfile"
				"-ExecutionPolicy", "Bypass"
				"-File", $scriptPath
			)
			
			if ($args -and $args.Count -gt 0)
			{
				foreach ($a in $args)
				{
					$argList += "$a"
				}
			}
			
			$psExe = $null
			try { $psExe = (Get-Command powershell.exe -ErrorAction SilentlyContinue).Source }
			catch { }
			if ([string]::IsNullOrWhiteSpace($psExe))
			{
				try { $psExe = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Source }
				catch { }
			}
			if ([string]::IsNullOrWhiteSpace($psExe) -or -not (Test-Path -LiteralPath $psExe))
			{
				$psExe = "powershell.exe"
			}
			
			Start-Process -FilePath $psExe -ArgumentList $argList -Verb RunAs | Out-Null
			exit 0
		}
		
		if ($isExeLaunch)
		{
			$exeArgList = @()
			if ($launchArgs -and $launchArgs.Count -gt 0) { $exeArgList = @($launchArgs) }
			Start-Process -FilePath $entryPath -ArgumentList $exeArgList -Verb RunAs | Out-Null
			exit 0
		}
	}
	catch
	{
		if ($script:ShowConsole) { Write-Host "Elevation canceled or unavailable. Continuing." -ForegroundColor DarkYellow }
	}
}

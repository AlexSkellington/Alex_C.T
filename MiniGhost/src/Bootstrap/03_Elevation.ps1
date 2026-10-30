# ===================================================================================================
#                                 SECTION: Elevation Bootstrap
# ---------------------------------------------------------------------------------------------------
# Description:
#   Relaunches MiniGhost as administrator before any maintenance logic runs.
#   Supports packaged EXE launches and .ps1 launches (the launcher and Run_MiniGhost.bat both run a
#   downloaded .ps1). An inline irm | iex run without admin rights is NOT relaunched: that needed
#   writing MiniGhost's own code to a new .ps1 in %TEMP% and starting it elevated - the dropper
#   pattern antivirus flags - so the user is told to start it as administrator instead.
# ===================================================================================================

function Invoke_Elevation_Bootstrap
{
	param (
		[string[]]$PassthroughArgs
	)
	
	try
	{
		$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
		$isAdministrator = ($null -ne $currentIdentity) -and ([Security.Principal.WindowsPrincipal]::new($currentIdentity)).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
	}
	catch
	{
		$isAdministrator = $false
	}
	
	if ($isAdministrator)
	{
		return
	}
	
	$launchPath = $null
	$workingDirectory = (Get-Location).Path
	$argumentList = @()
	$elevationStarted = $false
	
	try
	{
		$currentProcessPath = $null
		try
		{
			$currentProcessPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
		}
		catch { }
		
		$currentProcessLeaf = if ([string]::IsNullOrWhiteSpace($currentProcessPath)) { "" } else { [System.IO.Path]::GetFileName($currentProcessPath) }
		$preferredWindowsPowerShellPath = Join-Path $PSHOME 'powershell.exe'
		$systemWindowsPowerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
		$fallbackPwshPath = $null
		$shellHostPath = $null
		
		if (-not [string]::IsNullOrWhiteSpace($currentProcessPath))
		{
			if ($currentProcessLeaf -match '^(?i)powershell\.exe$')
			{
				$shellHostPath = $currentProcessPath
			}
			elseif ($currentProcessLeaf -match '^(?i)pwsh\.exe$')
			{
				$fallbackPwshPath = $currentProcessPath
			}
		}
		
		if ([string]::IsNullOrWhiteSpace($shellHostPath))
		{
			foreach ($candidate in @(
					$preferredWindowsPowerShellPath,
					$systemWindowsPowerShellPath,
					(Join-Path $PSHOME 'pwsh.exe'),
					$fallbackPwshPath,
					$currentProcessPath
				))
			{
				if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate))
				{
					$shellHostPath = $candidate
					break
				}
			}
		}
		
		$scriptFilePath = $PSCommandPath
		if ([string]::IsNullOrWhiteSpace($scriptFilePath))
		{
			try { $scriptFilePath = $MyInvocation.MyCommand.Path }
			catch { $scriptFilePath = $null }
		}
		
		if (-not [string]::IsNullOrWhiteSpace($scriptFilePath))
		{
			try
			{
				$scriptFilePath = (Resolve-Path -LiteralPath $scriptFilePath -ErrorAction Stop).ProviderPath
			}
			catch { }
		}
		
		$isScriptLaunch = -not [string]::IsNullOrWhiteSpace($scriptFilePath) -and ([System.IO.Path]::GetExtension($scriptFilePath) -ieq '.ps1') -and (Test-Path -LiteralPath $scriptFilePath)
		
		if ($isScriptLaunch)
		{
			if (-not [string]::IsNullOrWhiteSpace($shellHostPath))
			{
				$launchPath = $shellHostPath
				$workingDirectory = Split-Path -Path $scriptFilePath -Parent
				$argumentList = @('-NoProfile')
				if ([System.IO.Path]::GetFileName($shellHostPath) -match '^(?i)powershell\.exe$')
				{
					$argumentList += '-STA'
				}
				
				$argumentList += @('-ExecutionPolicy', 'Bypass', '-File', $scriptFilePath)
				if ($PassthroughArgs) { $argumentList += $PassthroughArgs }
			}
		}
		elseif ($currentProcessLeaf -and ($currentProcessLeaf -notmatch '^(?i)(powershell|pwsh|powershell_ise)\.exe$'))
		{
			$launchPath = $currentProcessPath
			if (-not [string]::IsNullOrWhiteSpace($currentProcessPath))
			{
				$workingDirectory = Split-Path -Path $currentProcessPath -Parent
			}
			if ($PassthroughArgs) { $argumentList += $PassthroughArgs }
		}
		# Inline (irm | iex) without admin: no relaunch - see the header
		
		if (-not [string]::IsNullOrWhiteSpace($launchPath))
		{
			$processArgumentList = $null
			if ($argumentList -and $argumentList.Count -gt 0)
			{
				$quotedArguments = foreach ($argument in $argumentList)
				{
					$text = if ($null -eq $argument) { '' } else { [string]$argument }
					if ($text -notmatch '[\s"]')
					{
						$text
						continue
					}
					
					$escapedText = $text -replace '(\\*)"', '$1$1\"'
					$escapedText = $escapedText -replace '(\\+)$', '$1$1'
					'"' + $escapedText + '"'
				}
				
				$processArgumentList = ($quotedArguments -join ' ')
			}
			
			$startProcessParams = @{
				FilePath         = $launchPath
				WorkingDirectory = $workingDirectory
				Verb             = 'RunAs'
				WindowStyle      = 'Minimized'
			}
			
			if (-not [string]::IsNullOrWhiteSpace($processArgumentList))
			{
				$startProcessParams['ArgumentList'] = $processArgumentList
			}
			
			Start-Process @startProcessParams | Out-Null
			$elevationStarted = $true
		}
	}
	catch [System.ComponentModel.Win32Exception]
	{
		if ($_.Exception.NativeErrorCode -ne 1223)
		{
			$elevationStarted = $false
		}
	}
	catch
	{
		$elevationStarted = $false
	}
	
	if ($elevationStarted)
	{
		exit
	}
	
	[System.Windows.Forms.MessageBox]::Show(
		"MiniGhost needs administrator rights.`r`n`r`nStart it with Run_MiniGhost.bat, or open PowerShell with right-click > Run as administrator and run it there.",
		"Administrator Required",
		[System.Windows.Forms.MessageBoxButtons]::OK,
		[System.Windows.Forms.MessageBoxIcon]::Warning
	) | Out-Null
	exit
}

Invoke_Elevation_Bootstrap -PassthroughArgs $args

# No Set-ExecutionPolicy here: MiniGhost loads no script files and starts no child PowerShell (the
# elevated relaunch passes -ExecutionPolicy Bypass itself), so it changed nothing - it only made the
# script look like it disables security settings.
Import-Module -Name Microsoft.PowerShell.Utility
Add-Type -AssemblyName System.Drawing

Write-Host "Script starting, pls wait..." -ForegroundColor Yellow

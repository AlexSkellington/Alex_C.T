#######################################################################################################
#                                                                                                     #
#                                     TBS MAINTENANCE SCRIPT                                          #
#                                                                                                     #
#                                        Author: Alex_C.T                                             #
#                                                                                                     #
#  > Edit only in consultation with Alex_C.T                                                          #
#  > This script performs advanced maintenance and diagnostics on TBS systems                         #
#                                                                                                     #
#######################################################################################################

$script:ShowConsole = $false
$script:ConsoleHidden = $false
$script:RestoreConsoleOnExit = $false

try
{
	$__scriptPath = $PSCommandPath
	if ([string]::IsNullOrWhiteSpace($__scriptPath))
	{
		try { $__scriptPath = $MyInvocation.MyCommand.Path }
		catch { $__scriptPath = $null }
	}
}
catch
{
	$__scriptPath = $null
}

$__entryPath = $null
$__launchArgs = @()
$__isSavedPs1Launch = $false
$__isExeLaunch = $false
$__isInlineLikeLaunch = $false

try
{
	$__cmdLineArgs = [Environment]::GetCommandLineArgs()
	if ($__cmdLineArgs -and $__cmdLineArgs.Count -gt 0)
	{
		$__entryPath = $__cmdLineArgs[0]
		if ($__cmdLineArgs.Count -gt 1)
		{
			$__launchArgs = @($__cmdLineArgs[1 .. ($__cmdLineArgs.Count - 1)])
		}
	}
}
catch
{
	$__entryPath = $null
	$__launchArgs = @()
}

try
{
	if (-not [string]::IsNullOrWhiteSpace($__scriptPath) -and
		(Test-Path -LiteralPath $__scriptPath) -and
		([string]::Equals([System.IO.Path]::GetExtension($__scriptPath), '.ps1', [System.StringComparison]::OrdinalIgnoreCase)))
	{
		$__isSavedPs1Launch = $true
	}
}
catch
{
	$__isSavedPs1Launch = $false
}

if (-not $__isSavedPs1Launch)
{
	try
	{
		if (-not [string]::IsNullOrWhiteSpace($__entryPath) -and
			(Test-Path -LiteralPath $__entryPath) -and
			([string]::Equals([System.IO.Path]::GetExtension($__entryPath), '.exe', [System.StringComparison]::OrdinalIgnoreCase)))
		{
			$__entryName = [System.IO.Path]::GetFileNameWithoutExtension($__entryPath)
			if ($__entryName -notmatch '^(powershell|pwsh|powershell_ise|powershellstudio)$')
			{
				$__isExeLaunch = $true
			}
		}
	}
	catch
	{
		$__isExeLaunch = $false
	}
}

$__isInlineLikeLaunch = (-not $__isSavedPs1Launch -and -not $__isExeLaunch)
$script:ShowConsole = [bool](@($__launchArgs | Where-Object { $_ -in @('-ShowConsole', '/ShowConsole', '-KeepConsole', '/KeepConsole') }).Count -gt 0)
$script:LaneSqlEnableWorkerMode = $false
$script:LaneSqlEnableWorkerLane = $null
$script:LaneSqlEnableWorkerMachine = $null
$script:LaneSqlEnableWorkerTcpPort = '1433'
$script:LaneSqlEnableWorkerStartTime = '04:15'

if ($__launchArgs -and $__launchArgs.Count -gt 0)
{
	for ($__argIndex = 0; $__argIndex -lt $__launchArgs.Count; $__argIndex++)
	{
		$__argValue = [string]$__launchArgs[$__argIndex]
		switch -Regex ($__argValue)
		{
			'^-LaneSqlEnableWorker$'
			{
				$script:LaneSqlEnableWorkerMode = $true
				continue
			}
			'^-LaneSqlEnableWorkerLane$'
			{
				if (($__argIndex + 1) -lt $__launchArgs.Count)
				{
					$script:LaneSqlEnableWorkerLane = [string]$__launchArgs[$__argIndex + 1]
					$__argIndex++
				}
				continue
			}
			'^-LaneSqlEnableWorkerMachine$'
			{
				if (($__argIndex + 1) -lt $__launchArgs.Count)
				{
					$script:LaneSqlEnableWorkerMachine = [string]$__launchArgs[$__argIndex + 1]
					$__argIndex++
				}
				continue
			}
			'^-LaneSqlEnableWorkerTcpPort$'
			{
				if (($__argIndex + 1) -lt $__launchArgs.Count)
				{
					$script:LaneSqlEnableWorkerTcpPort = [string]$__launchArgs[$__argIndex + 1]
					$__argIndex++
				}
				continue
			}
			'^-LaneSqlEnableWorkerStartTime$'
			{
				if (($__argIndex + 1) -lt $__launchArgs.Count)
				{
					$script:LaneSqlEnableWorkerStartTime = [string]$__launchArgs[$__argIndex + 1]
					$__argIndex++
				}
				continue
			}
		}
	}
}

if (-not ([System.Management.Automation.PSTypeName]'ConsoleWindowHelper').Type)
{
	Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class ConsoleWindowHelper {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
}

Write-Host "Script starting, pls wait..." -ForegroundColor Yellow

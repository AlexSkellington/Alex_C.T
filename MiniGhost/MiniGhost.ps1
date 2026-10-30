#######################################################################################################
#                                                                                                     #
#                                        MiniGhost SCRIPT                                             #
#                                                                                                     #
#                                        Author: Alex_C.T                                             #
#                                                                                                     #
#  > Edit only in consultation with Alex_C.T                                                          #
#  > This script performs advanced maintenance and diagnostics on TBS systems                         #
#                                                                                                     #
#######################################################################################################

# ===================================================================================================
#                                       SECTION: Parameters
# ---------------------------------------------------------------------------------------------------
# Description:
#   Defines the script parameters, allowing users to run the script in silent mode.
# ===================================================================================================

# Script build version (cunsult with Alex_C.T before changing this)
$VersionNumber = "1.4.2"
$VersionDate = "2026-10-30"

# Retrieve Major, Minor, Build, and Revision version numbers of PowerShell
$major = $PSVersionTable.PSVersion.Major
$minor = $PSVersionTable.PSVersion.Minor
$build = $PSVersionTable.PSVersion.Build
$revision = $PSVersionTable.PSVersion.Revision

# Combine them into a single version string
$PowerShellVersion = "$major.$minor.$build.$revision"

# ===================================================================================================
#                           SECTION: Import Necessary Assemblies and Modules
# ---------------------------------------------------------------------------------------------------
# Description:
#   Imports required .NET assemblies for creating and managing Windows Forms and graphical components
#   and imports necessary PowerShell modules required for the script's operation.
# ===================================================================================================

# Add the WinForms assembly early so the elevation gate can show a warning dialog if needed.
Add-Type -AssemblyName System.Windows.Forms

# No hidden console: the console is no longer hidden through user32 ShowWindow (that needed C# compiled
# at run time, and hiding a window is a classic malware trait antivirus scores). The launcher and
# the elevated relaunch start the console minimized instead - visible in the taskbar.

# Prefer the standard WinForms startup order, but don't fail if the current PowerShell
# session has already created a window handle before MiniGhost starts.
[System.Windows.Forms.Application]::EnableVisualStyles()
try
{
	[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
}
catch
{
	$innerException = $_.Exception.InnerException
	$alreadyInitialized = ($innerException -is [System.InvalidOperationException]) -or ($_.Exception.Message -like '*SetCompatibleTextRenderingDefault must be called before the first IWin32Window object is created*')
	if (-not $alreadyInitialized)
	{
		throw
	}
}

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

# ===================================================================================================
#                                   SECTION: Initialize Variables
# ---------------------------------------------------------------------------------------------------
# Description:
#   Initializes all necessary variables and paths required for the script's operation, including
#   dynamic detection of the main Storeman folder, Office subpaths, INI files, encoding settings,
#   counters, and interop types for MailSlot messaging.
# ===================================================================================================

# ---------------------------------------------------------------------------------------------------
# Script-Scoped Result and Process Tracking Structures
# ---------------------------------------------------------------------------------------------------
$script:FunctionResults = @{ }

# Get the current machine name
$currentMachineName = $env:COMPUTERNAME

# Initialize script-scoped variables for new store number and new machine name
$script:newStoreNumber = $null
$script:newMachineName = $null
$script:NodeRole = "Unknown"

# ---------------------------------------------------------------------------------------------------
# Encoding Settings
# ---------------------------------------------------------------------------------------------------
$script:ansiPcEncoding = [System.Text.Encoding]::GetEncoding(1252) # Windows-1252 legacy files
$script:utf8NoBOM = New-Object System.Text.UTF8Encoding($false) # UTF-8 no BOM (for output)
$script:utf8NoBOM = $utf8NoBOM
$script:ansiPcEncoding = $ansiPcEncoding

# ---------------------------------------------------------------------------------------------------
# Locate Base Path: Storeman Folder Detection (case-insensitive)
# ---------------------------------------------------------------------------------------------------
$BasePath = $null
$targetSubPathPattern = 'Office\Dbs\INFO_*901_WIN.INI'
$storemanDirs = @()
$fixedDrives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 -and $_.Root -match '^[A-Z]:\\$' }

foreach ($drive in $fixedDrives)
{
	# Case-insensitive match for any *storeman* variation in directory name
	$dirs = Get-ChildItem -Path "$($drive.Root)" -Directory -ErrorAction SilentlyContinue |
	Where-Object { $_.Name -imatch 'storeman' } |
	ForEach-Object {
		$candidatePath = Join-Path $_.FullName '\'
		$files = Get-ChildItem -Path $candidatePath -Filter 'Startup.ini' -ErrorAction SilentlyContinue
		if ($files) { $_ }
	}
	if ($dirs) { $storemanDirs += $dirs }
}

if ($storemanDirs.Count -gt 1)
{
	# Prefer a path that is actually shared as a Windows share
	$shares = Get-SmbShare -ErrorAction SilentlyContinue
	foreach ($dir in $storemanDirs)
	{
		if ($shares.Path -contains $dir.FullName)
		{
			$BasePath = $dir.FullName
			break
		}
	}
	if (-not $BasePath) { $BasePath = $storemanDirs[0].FullName }
}
elseif ($storemanDirs.Count -eq 1)
{
	$BasePath = $storemanDirs[0].FullName
}

# Final fallback: Default to C:\storeman if none found
if (-not $BasePath)
{
	$fallback = "C:\storeman"
	$candidatePath = Join-Path $fallback '\'
	$files = Get-ChildItem -Path $candidatePath -Filter 'Startup.ini' -ErrorAction SilentlyContinue
	if ($files) { $BasePath = $fallback }
	else
	{
		Write-Warning "Could not locate any storeman folder containing 'storeman\Startup.ini'.`nRunning with limited functionality."
		$BasePath = $fallback
	}
}

Write-Host "Selected (storeman) folder: '$BasePath'" -ForegroundColor Magenta

# ---------------------------------------------------------------------------------------------------
# Build All Core Paths and File Locations
# ---------------------------------------------------------------------------------------------------

# Storeman root paths
$OfficePath = Join-Path $BasePath "Office"
$LoadPath = Join-Path $OfficePath "Load"
$StartupIniPath = Join-Path $BasePath "Startup.ini"
$GlobalSmsStartIniPath = Join-Path $BasePath "SMSStart.ini"
$SystemIniPath = Join-Path $OfficePath "system.ini"
$GasInboxPath = Join-Path $OfficePath "XchGAS\INBOX"
$DbsPath = Join-Path $OfficePath "Dbs"
# SmsHttps.ini (auto)
$SmsHttpsIniPath = Join-Path $BasePath "SmsHttps64\SmsHttps.INI"
if (-not (Test-Path $SmsHttpsIniPath)) { $SmsHttpsIniPath = Join-Path $BasePath "SmsHttps64\SmsHttps.ini" }
if (-not (Test-Path $SmsHttpsIniPath)) { $SmsHttpsIniPath = Join-Path $BasePath "SmsHttps32\SmsHttps.ini" }
if (-not (Test-Path $SmsHttpsIniPath)) { $SmsHttpsIniPath = Join-Path $BasePath "SmsHttps\SmsHttps.ini" }
if (-not (Test-Path $SmsHttpsIniPath)) { $SmsHttpsIniPath = $null }

$TempDir = [System.IO.Path]::GetTempPath()

# Initialize variables for the INFO_*901 files
$WinIniPath = $null
$SmsStartIniPath = $null # INFO_*901_SMSStart.ini inside Dbs

# ---------------------------------------------------------------------------------------------------
# Find INFO_*901_WIN.INI
# ---------------------------------------------------------------------------------------------------
try
{
	$WinIniMatch = Get-ChildItem -Path $DbsPath -Filter 'INFO_*901_WIN.INI' -ErrorAction Stop |
	Select-Object -First 1
	if ($WinIniMatch)
	{
		$WinIniPath = $WinIniMatch.FullName
	}
}
catch { }

# ---------------------------------------------------------------------------------------------------
# Find INFO_*901_SMSStart.ini
# ---------------------------------------------------------------------------------------------------
try
{
	$SmsStartIniMatch = Get-ChildItem -Path $DbsPath -Filter 'INFO_*901_SMSStart.ini' -ErrorAction Stop |
	Select-Object -First 1
	
	if ($SmsStartIniMatch)
	{
		$SmsStartIniPath = $SmsStartIniMatch.FullName
	}
}
catch { }

# Initialize a hashtable to track the status of each operation
$operationStatus = @{
	"StoreNumberChange" = @{ Status = "Pending"; Message = ""; Details = "" }
	"MachineNameChange" = @{ Status = "Pending"; Message = ""; Details = "" }
	"OldXFoldersDeletion" = @{ Status = "Pending"; Message = ""; Details = "" }
	"StartupIniUpdate"  = @{ Status = "Pending"; Message = ""; Details = "" }
	"IPConfiguration"   = @{ Status = "Pending"; Message = ""; Details = "" }
	"TableTruncation"   = @{ Status = "Pending"; Message = ""; Details = "" }
	"RegistryCleanup"   = @{ Status = "Pending"; Message = ""; Details = "" }
	"SQLDatabaseUpdate" = @{ Status = "Pending"; Message = ""; Details = "" }
}

# ===================================================================================================
#                                 FUNCTION: Draw_MiniGhost_Logo
# ---------------------------------------------------------------------------------------------------
# Description:
#   Draws the MiniGhost logo (violet gradient disc with a white ghost) into the square (X, Y, Size)
#   of any Graphics - the window icon and the header badge use it.
#
# Notes:
#   - Coordinates are computed in expression mode and passed to .NET methods, never as
#     'New-Object T($a + $b, ...)' (command mode turns that into $a + @($b, ...)).
# ===================================================================================================

function Draw_MiniGhost_Logo
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.Drawing.Graphics]$Graphics,
		[float]$X = 0,
		[float]$Y = 0,
		[float]$Size = 32
	)
	
	$Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
	$disc = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
		[System.Drawing.RectangleF]::new($X, $Y, $Size, $Size),
		[System.Drawing.Color]::FromArgb(126, 87, 214),
		[System.Drawing.Color]::FromArgb(79, 70, 189),
		[float]90)
	$Graphics.FillEllipse($disc, $X, $Y, $Size, $Size)
	$disc.Dispose()
	
	# Ghost: round head, straight sides, three scallops along the bottom
	$w = $Size * 0.48
	$h = $Size * 0.56
	$left = $X + ($Size - $w) / 2
	$top = $Y + $Size * 0.2
	$bump = $w / 3
	$ghost = New-Object System.Drawing.Drawing2D.GraphicsPath
	$ghost.AddArc($left, $top, $w, $w, [float]180, [float]180)
	$ghost.AddLine(($left + $w), ($top + $w / 2), ($left + $w), ($top + $h - $bump / 2))
	for ($i = 0; $i -lt 3; $i++)
	{
		$ghost.AddArc(($left + $w - ($i + 1) * $bump), ($top + $h - $bump), $bump, $bump, [float]0, [float]180)
	}
	$ghost.CloseFigure()
	$white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
	$Graphics.FillPath($white, $ghost)
	$white.Dispose()
	$ghost.Dispose()
	
	$eye = $Size * 0.085
	$eyeY = $top + $w * 0.38
	$dark = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(79, 70, 189))
	$Graphics.FillEllipse($dark, ($left + $w * 0.3 - $eye / 2), $eyeY, $eye, ($eye * 1.3))
	$Graphics.FillEllipse($dark, ($left + $w * 0.7 - $eye / 2), $eyeY, $eye, ($eye * 1.3))
	$dark.Dispose()
}

# ===================================================================================================
# 									FUNCTION: Get_NEW_Machine_Name
# ---------------------------------------------------------------------------------------------------
# Description:
# Prompts user for a new machine name via GUI.
# Supports:
#   - Classic: PREFIX + 3 digits, with optional hyphen before digits (e.g. LANE003, POS-001, SERVER-001)
#   - Prefixed: optional 1-4 digit store + PREFIX + optional hyphen + 3 digits (e.g. 0231LANE006, 0231POS-001)
# If prefixed format used -> offers to sync store number
# Returns uppercase validated name or $null if cancelled
# ===================================================================================================

function Get_NEW_Machine_Name
{
	while ($true)
	{
		$nodeRole = "Unknown"
		if (-not [string]::IsNullOrWhiteSpace([string]$script:NodeRole))
		{
			$nodeRole = ([string]$script:NodeRole).Trim()
		}
		
		$machineNamePromptText = if ($nodeRole -match '^(?i)(StoreServer|HostServer)$')
		{
			"Server examples:`nSERVER001  -  SERVER-001  -  0231SERVER001`nRe-enter the current machine name to force a full INI + SQL sync."
		}
		elseif ($nodeRole -match '^(?i)Lane$')
		{
			"Lane examples:`nPOS001  -  POS-001  -  LN026001  -  0384LANE001  -  001`nRe-enter the current machine name to force a full INI + SQL sync."
		}
		else
		{
			"Lane examples:`nPOS001  -  POS-001  -  LN026001  -  0384LANE001  -  001`nServer examples:`nSERVER001  -  SERVER-001  -  0231SERVER001`nRe-enter the current machine name to force a full INI + SQL sync."
		}
		$promptLineCount = ($machineNamePromptText -split "`n").Count
		$labelHeight = [Math]::Max(54, (18 * $promptLineCount) + 6)
		$formClientWidth = 430
		$formClientHeight = $labelHeight + 112
		
		$form = New-Object System.Windows.Forms.Form
		$form.Text = "Enter New Machine Name"
		$form.ClientSize = New-Object System.Drawing.Size($formClientWidth, $formClientHeight)
		$form.StartPosition = "CenterParent"
		$form.FormBorderStyle = 'FixedDialog'
		$form.MaximizeBox = $false
		$form.MinimizeBox = $false
		$form.ShowInTaskbar = $false
		
		$label = New-Object System.Windows.Forms.Label
		$label.Text = $machineNamePromptText
		$label.Location = New-Object System.Drawing.Point(10, 15)
		$label.Size = New-Object System.Drawing.Size(($formClientWidth - 20), $labelHeight)
		$label.AutoSize = $false
		
		$textBox = New-Object System.Windows.Forms.TextBox
		$textBox.Location = New-Object System.Drawing.Point(10, ($label.Bottom + 10))
		$textBox.Size = New-Object System.Drawing.Size(($formClientWidth - 20), 22)
		
		$okBtn = New-Object System.Windows.Forms.Button
		$okBtn.Text = "OK"
		$okBtn.Location = New-Object System.Drawing.Point(([int](($formClientWidth / 2) - 75)), ($textBox.Bottom + 14))
		$okBtn.Size = New-Object System.Drawing.Size(70, 28)
		$okBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
		
		$cancelBtn = New-Object System.Windows.Forms.Button
		$cancelBtn.Text = "Cancel"
		$cancelBtn.Location = New-Object System.Drawing.Point(($okBtn.Right + 15), $okBtn.Top)
		$cancelBtn.Size = New-Object System.Drawing.Size(70, 28)
		$cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		
		$form.Controls.AddRange(@($label, $textBox, $okBtn, $cancelBtn))
		$form.AcceptButton = $okBtn
		$form.CancelButton = $cancelBtn
		
		Set_Modern_Dialog_Style -Form $form -PrimaryButton $okBtn
		$dialogResult = $form.ShowDialog()
		
		# PS 5.1-safe: read textbox BEFORE disposing
		$userInputRaw = $textBox.Text
		$form.Dispose()
		
		if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
		{
			return $null
		}
		
		$userInput = $userInputRaw
		if ($null -eq $userInput) { $userInput = "" }
		$userInput = $userInput.Trim().ToUpper()
		
		if ([string]::IsNullOrEmpty($userInput))
		{
			[void][System.Windows.Forms.MessageBox]::Show("Machine name cannot be empty.", "Error")
			continue
		}
		
		if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
		
		$currentMachineForStyle = Get_Effective_Machine_Name
		$currentStore = $null
		if ($script:FunctionResults.ContainsKey('StoreNumber') -and $script:FunctionResults['StoreNumber'])
		{
			$currentStore = ($script:FunctionResults['StoreNumber'].ToString()).Trim()
		}
		else
		{
			$currentStore = Get_Store_Number_From_INI
			if ($currentStore) { $currentStore = ($currentStore.ToString()).Trim() }
		}
		
		$machineContext = Get_Machine_Name_Context -MachineName $userInput -CurrentMachineName $currentMachineForStyle
		if (-not $machineContext.Success)
		{
			[void][System.Windows.Forms.MessageBox]::Show(
				$machineContext.ErrorMessage,
				"Invalid Machine Name",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			continue
		}
		
		$normalizedName = $machineContext.MachineName
		$storePrefixFromName = $machineContext.StorePrefix
		$currentStoreComparable = $null
		if ($currentStore -match '^\d{3,4}$')
		{
			$currentStoreComparable = [int]$currentStore
		}
		
		if ($storePrefixFromName -and ($storePrefixFromName -match '^\d{3,4}$'))
		{
			$storePrefixComparable = [int]$storePrefixFromName
			if (($null -eq $currentStoreComparable) -or ($storePrefixComparable -ne $currentStoreComparable))
			{
				$displayStore = $currentStore
				if ([string]::IsNullOrEmpty($displayStore)) { $displayStore = "UNKNOWN" }
				
				$sync = [System.Windows.Forms.MessageBox]::Show(
					"Machine name has store prefix '$storePrefixFromName',`nbut current store is '$displayStore'.`n`nUpdate store number to '$storePrefixFromName'?",
					"Store Number Mismatch",
					[System.Windows.Forms.MessageBoxButtons]::YesNo,
					[System.Windows.Forms.MessageBoxIcon]::Question
				)
				
				if ($sync -eq [System.Windows.Forms.DialogResult]::Yes)
				{
					$iniParams = @{
						newStoreNumber = $storePrefixFromName
						MachineName    = $normalizedName
						OldMachineName = $currentMachineForStyle
					}
					if ($currentStore -match '^\d{3,4}$') { $iniParams["OldStoreNumber"] = $currentStore }
					
					$success = Update_INIs @iniParams
					
					if ($success)
					{
						$script:newStoreNumber = $storePrefixFromName
						$script:FunctionResults['StoreNumber'] = $storePrefixFromName
						
						if ($storeNumberLabel)
						{
							$storeNumberLabel.Text = "$storePrefixFromName (updated)"
							$storeNumberLabel.Refresh()
						}
					}
					else
					{
						[void][System.Windows.Forms.MessageBox]::Show("Failed to update store number in INI.", "Error")
						continue
					}
				}
			}
		}
		
		return $normalizedName
	}
}

# ===================================================================================================
# 									FUNCTION: Get_NEW_Store_Number
# ---------------------------------------------------------------------------------------------------
# Description:
# Prompts the user via a GUI to enter a valid store number.
# Accepts exactly 3 or 4 digits (no leading zeros required, not all zeros).
# Returns the entered value as-is (no padding) or $null if cancelled.
# ===================================================================================================

function Get_NEW_Store_Number
{
	while ($true)
	{
		# Decide required store length based on CURRENT store number length (3 vs 4)
		$requiredLen = 4
		
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('StoreNumber') -and $script:FunctionResults['StoreNumber'])
		{
			$cur = ($script:FunctionResults['StoreNumber'].ToString()).Trim()
			if ($cur -match '^\d{3,4}$') { $requiredLen = $cur.Length }
		}
		else
		{
			$iniStore = Get_Store_Number_From_INI # returns as-is (3 or 4) if your function is set that way
			if ($iniStore)
			{
				$iniStore = ($iniStore.ToString()).Trim()
				if ($iniStore -match '^\d{3,4}$') { $requiredLen = $iniStore.Length }
			}
		}
		
		$storeNumberForm = New-Object System.Windows.Forms.Form
		$storeNumberForm.Text = "Enter New Store Number"
		$storeNumberForm.ClientSize = New-Object System.Drawing.Size(340, 135)
		$storeNumberForm.StartPosition = "CenterParent"
		$storeNumberForm.FormBorderStyle = 'FixedDialog'
		$storeNumberForm.MaximizeBox = $false
		$storeNumberForm.MinimizeBox = $false
		$storeNumberForm.ShowInTaskbar = $false
		
		$label = New-Object System.Windows.Forms.Label
		$label.Text = "New Store Number (exactly $requiredLen digits, e.g. " + ($(if ($requiredLen -eq 3) { "123" }
				else { "4123" })) + "):"
		$label.Location = New-Object System.Drawing.Point(10, 20)
		$label.Size = New-Object System.Drawing.Size(315, 40)
		$label.AutoSize = $false
		
		$textBox = New-Object System.Windows.Forms.TextBox
		$textBox.Location = New-Object System.Drawing.Point(10, 65)
		$textBox.Size = New-Object System.Drawing.Size(320, 20)
		$textBox.MaxLength = $requiredLen
		
		$okButton = New-Object System.Windows.Forms.Button
		$okButton.Text = "OK"
		$okButton.Location = New-Object System.Drawing.Point(90, 98)
		$okButton.Size = New-Object System.Drawing.Size(75, 23)
		
		$cancelButton = New-Object System.Windows.Forms.Button
		$cancelButton.Text = "Cancel"
		$cancelButton.Location = New-Object System.Drawing.Point(175, 98)
		$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
		
		$dialogSelection = $null
		$okButton.Add_Click({
				$script:MiniGhostTempDialogSelection = 'OK'
				$storeNumberForm.Close()
			})
		$cancelButton.Add_Click({
				$script:MiniGhostTempDialogSelection = 'Cancel'
				$storeNumberForm.Close()
			})
		
		$storeNumberForm.Controls.AddRange(@($label, $textBox, $okButton, $cancelButton))
		$storeNumberForm.AcceptButton = $okButton
		$storeNumberForm.CancelButton = $cancelButton
		
		$script:MiniGhostTempDialogSelection = $null
		Set_Modern_Dialog_Style -Form $storeNumberForm -PrimaryButton $okButton
		[void]$storeNumberForm.ShowDialog()
		
		# IMPORTANT: read before Dispose (PS WinForms reliability)
		$userInput = [string]$textBox.Text
		$dialogSelection = $script:MiniGhostTempDialogSelection
		$script:MiniGhostTempDialogSelection = $null
		$storeNumberForm.Dispose()
		
		if ($dialogSelection -ne 'OK')
		{
			return $null
		}
		
		if ($null -eq $userInput) { $userInput = "" }
		$userInput = $userInput.Trim()
		
		if ([string]::IsNullOrEmpty($userInput))
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Store number cannot be empty.",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			continue
		}
		
		# Build regex based on required length; reject all zeros for that length
		$re = '^(?!0{' + $requiredLen + '})\d{' + $requiredLen + '}$'
		
		if ($userInput -match $re)
		{
			if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
			$script:FunctionResults['StoreNumber'] = $userInput
			return $userInput # return as-is, no padding
		}
		else
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Invalid store number.`n`nMust be exactly $requiredLen digits (numeric).`nNot allowed: all zeros, too short, too long.`nExample: " + ($(if ($requiredLen -eq 3) { "123" }
						else { "4123" })),
				"Invalid Input",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
		}
	}
}

# ===================================================================================================
#                                 FUNCTION: New_MiniGhost_Action_Card
# ---------------------------------------------------------------------------------------------------
# Description:
#   Returns an action card for the main window: coloured dot, title, subtitle, and a host panel
#   for its buttons (in the card's .Tag).
#
# Parameters:
#   - Title, Subtitle [string]: card header text
#   - Accent          [Color]: colour of the header dot
# ===================================================================================================

function New_MiniGhost_Action_Card
{
	param ([string]$Title, [string]$Subtitle, [System.Drawing.Color]$Accent)

	$card = New_MiniGhost_Card -Margin (New-Object System.Windows.Forms.Padding(6))
	$card.Dock = 'Fill'

	# Buttons go into the host panel (11_Layout_And_Tooltips); the header docks above it.
	# ($host is a read-only automatic variable - hence $buttonHost)
	$buttonHost = New-Object System.Windows.Forms.Panel
	$buttonHost.Dock = 'Fill'
	$buttonHost.BackColor = [System.Drawing.Color]::White
	$buttonHost.Padding = New-Object System.Windows.Forms.Padding(12, 4, 12, 12)
	$card.Controls.Add($buttonHost)

	$head = New-Object System.Windows.Forms.Panel
	$head.Dock = 'Top'
	$head.Height = 58
	$head.BackColor = [System.Drawing.Color]::White
	$head.Tag = $Accent
	$head.add_Paint({
			param ($s, $e)
			try
			{
				$e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
				$dot = New-Object System.Drawing.SolidBrush($s.Tag)
				$e.Graphics.FillEllipse($dot, 14, 16, 10, 10)
				$dot.Dispose()
			}
			catch { }
		})
	$t = New-Object System.Windows.Forms.Label
	$t.Text = $Title
	$t.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
	$t.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
	$t.AutoSize = $true
	$t.Location = New-Object System.Drawing.Point(30, 9)
	$head.Controls.Add($t)
	$st = New-Object System.Windows.Forms.Label
	$st.Text = $Subtitle
	$st.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
	$st.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 110)
	$st.AutoSize = $true
	$st.Location = New-Object System.Drawing.Point(31, 32)
	$head.Controls.Add($st)
	$card.Controls.Add($head)

	$card.Tag = $buttonHost
	return $card
}

# ===================================================================================================
#                                    FUNCTION: New_MiniGhost_Card
# ---------------------------------------------------------------------------------------------------
# Description:
#   Returns a white card panel with a 1 px light border (main-window sections).
# ===================================================================================================

function New_MiniGhost_Card
{
	param ([System.Windows.Forms.Padding]$Margin)
	$card = New-Object System.Windows.Forms.Panel
	$card.BackColor = [System.Drawing.Color]::White
	$card.Margin = $Margin
	$card.Padding = New-Object System.Windows.Forms.Padding(1)
	$card.add_Paint({
			param ($s, $e)
			try
			{
				$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(229, 229, 229))
				$e.Graphics.DrawRectangle($pen, 0, 0, ($s.Width - 1), ($s.Height - 1))
				$pen.Dispose()
			}
			catch { }
		})
	return $card
}

# ===================================================================================================
#                                  FUNCTION: Set_Label_Text
# ---------------------------------------------------------------------------------------------------
# Description:
# Updates a WinForms label and optionally refreshes its parent layout.
# ===================================================================================================

function Set_Label_Text
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		$Label,
		[Parameter(Mandatory = $true)]
		[string]$Text,
		[Parameter(Mandatory = $false)]
		[switch]$RefreshParent,
		[Parameter(Mandatory = $false)]
		[switch]$ProcessEvents
	)
	
	if (-not $Label) { return }
	
	$Label.Text = $Text
	$Label.Refresh()
	
	if ($RefreshParent -and $Label.Parent)
	{
		$Label.Parent.PerformLayout()
		$Label.Parent.Refresh()
	}
	
	if ($ProcessEvents)
	{
		[System.Windows.Forms.Application]::DoEvents()
	}
}

# ===================================================================================================
#                                FUNCTION: Set_Modern_Button_Style
# ---------------------------------------------------------------------------------------------------
# Description:
#   Gives a button the flat MiniGhost look: Primary (filled accent), Secondary (white with a border),
#   Action (main-window card button: left-aligned, coloured stripe) or Danger (Action in red text).
#
# Parameters:
#   - Button      [Button]: the button to style
#   - Kind        [string]: Primary | Secondary | Action | Danger (default Secondary)
#   - AccentColor [Color]: stripe colour for Action / Danger (default accent blue)
#
# Notes:
#   - Hover / pressed colours come from FlatAppearance, so no mouse handlers are needed.
#   - The stripe is drawn by a Paint handler that reads the colour from the button's .Tag (handlers
#     run in their own scope - nothing from this function is visible there).
# ===================================================================================================

function Set_Modern_Button_Style
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.Button]$Button,
		[Parameter(Mandatory = $false)]
		[ValidateSet('Primary', 'Secondary', 'Action', 'Danger')]
		[string]$Kind = 'Secondary',
		[Parameter(Mandatory = $false)]
		[System.Drawing.Color]$AccentColor = [System.Drawing.Color]::FromArgb(0, 103, 192)
	)
	
	$Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	$Button.UseVisualStyleBackColor = $false
	$Button.Cursor = [System.Windows.Forms.Cursors]::Hand
	$Button.FlatAppearance.BorderSize = 1
	
	switch ($Kind)
	{
		'Primary'
		{
			$Button.BackColor = [System.Drawing.Color]::FromArgb(0, 103, 192)
			$Button.ForeColor = [System.Drawing.Color]::White
			$Button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5)
			$Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 103, 192)
			$Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(25, 117, 197)
			$Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(0, 84, 158)
		}
		'Secondary'
		{
			$Button.BackColor = [System.Drawing.Color]::White
			$Button.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
			$Button.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
			$Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(209, 209, 209)
			$Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
			$Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(235, 235, 235)
		}
		default
		{
			# Action / Danger: card buttons in the main window
			$Button.BackColor = [System.Drawing.Color]::White
			$Button.ForeColor = if ($Kind -eq 'Danger') { [System.Drawing.Color]::FromArgb(196, 43, 28) } else { [System.Drawing.Color]::FromArgb(32, 32, 32) }
			$Button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
			$Button.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
			$Button.Padding = New-Object System.Windows.Forms.Padding(14, 0, 8, 0)
			$Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
			$Button.FlatAppearance.MouseOverBackColor = if ($Kind -eq 'Danger') { [System.Drawing.Color]::FromArgb(253, 240, 238) } else { [System.Drawing.Color]::FromArgb(240, 246, 253) }
			$Button.FlatAppearance.MouseDownBackColor = if ($Kind -eq 'Danger') { [System.Drawing.Color]::FromArgb(249, 224, 220) } else { [System.Drawing.Color]::FromArgb(224, 236, 250) }
			$Button.Tag = $AccentColor
			$Button.Add_Paint({
					param ($s, $e)
					try
					{
						if ($s.Tag -is [System.Drawing.Color])
						{
							$stripe = New-Object System.Drawing.SolidBrush($s.Tag)
							$e.Graphics.FillRectangle($stripe, 0, 0, 4, $s.Height)
							$stripe.Dispose()
						}
					}
					catch { }
				})
		}
	}
}

# ===================================================================================================
#                                FUNCTION: Set_Modern_Dialog_Style
# ---------------------------------------------------------------------------------------------------
# Description:
#   Restyles a dialog and every control in it to match the main window: white background, Segoe UI,
#   flat buttons (the given primary button filled with the accent colour), bordered text boxes.
#
# Parameters:
#   - Form          [Form]: the dialog
#   - PrimaryButton [Button]: the default action (OK / Reboot Now); other buttons get Secondary
#
# Notes:
#   - Layout is not changed: positions and sizes stay as the dialog defined them.
#   - Monospaced fonts (report text) are kept.
# ===================================================================================================

function Set_Modern_Dialog_Style
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.Form]$Form,
		[Parameter(Mandatory = $false)]
		[System.Windows.Forms.Button]$PrimaryButton
	)
	
	$Form.BackColor = [System.Drawing.Color]::White
	$Form.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
	$Form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
	if ($script:MiniGhostIcon) { try { $Form.Icon = $script:MiniGhostIcon; $Form.ShowIcon = $true } catch { } }
	
	$pending = New-Object System.Collections.Generic.Queue[System.Windows.Forms.Control]
	foreach ($c in @($Form.Controls)) { $pending.Enqueue($c) }
	while ($pending.Count -gt 0)
	{
		$c = $pending.Dequeue()
		foreach ($child in @($c.Controls)) { $pending.Enqueue($child) }
		try
		{
			if ($c -is [System.Windows.Forms.Button])
			{
				if ($PrimaryButton -and [object]::ReferenceEquals($c, $PrimaryButton)) { Set_Modern_Button_Style -Button $c -Kind Primary }
				else { Set_Modern_Button_Style -Button $c -Kind Secondary }
			}
			elseif ($c -is [System.Windows.Forms.TextBox])
			{
				if ($c.Font.Name -notmatch '(?i)consolas|courier|lucida console') { $c.Font = New-Object System.Drawing.Font("Segoe UI", 10) }
				if ($c.ReadOnly -and $c.BorderStyle -eq [System.Windows.Forms.BorderStyle]::None)
				{
					$c.BackColor = [System.Drawing.Color]::White
				}
				else
				{
					$c.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
				}
			}
			elseif ($c -is [System.Windows.Forms.Label])
			{
				$c.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
				$c.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
			}
			elseif ($c -is [System.Windows.Forms.FlowLayoutPanel] -and $c.Dock -eq [System.Windows.Forms.DockStyle]::Bottom)
			{
				# button bar: light footer like Windows 11 dialogs
				$c.BackColor = [System.Drawing.Color]::FromArgb(243, 243, 243)
			}
		}
		catch { }
	}
}

# ===================================================================================================
#                                FUNCTION: Set_Operation_Status
# ---------------------------------------------------------------------------------------------------
# Description:
# Central helper to update the shared operation-status table only when the requested key exists.
# ===================================================================================================

function Set_Operation_Status
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus,
		[Parameter(Mandatory = $true)]
		[string]$Key,
		[Parameter(Mandatory = $true)]
		[string]$Status,
		[Parameter(Mandatory = $false)]
		[string]$Message = "",
		[Parameter(Mandatory = $false)]
		[string]$Details = ""
	)
	
	if ($OperationStatus -and $OperationStatus.ContainsKey($Key))
	{
		$OperationStatus[$Key].Status = $Status
		$OperationStatus[$Key].Message = $Message
		$OperationStatus[$Key].Details = $Details
	}
}

# ===================================================================================================
#                              FUNCTION: Show_Action_Prompt_With_Log
# ---------------------------------------------------------------------------------------------------
# Description:
#   Message dialog with one or two action buttons and an optional View Log button that opens the
#   given log text in Show_Text_Report_Form.
#
# Returns:
#   "Primary", "Secondary" or "Closed" (window closed with X).
#
# Notes:
#   - Click handlers run in their own scope: a plain $selection = ... inside them would set a new
#     local variable and the function would always return "Closed". The choice is carried on the
#     form's .Tag instead.
# ===================================================================================================

function Show_Action_Prompt_With_Log
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$Text,
		[Parameter(Mandatory = $true)]
		[string]$Title,
		[Parameter(Mandatory = $false)]
		[string]$PrimaryButtonText = "OK",
		[Parameter(Mandatory = $false)]
		[string]$SecondaryButtonText,
		[Parameter(Mandatory = $false)]
		[string]$LogText,
		[Parameter(Mandatory = $false)]
		[string]$LogTitle = "Details"
	)
	
	$promptForm = New-Object System.Windows.Forms.Form
	$promptForm.Text = $Title
	$promptForm.Size = New-Object System.Drawing.Size(620, 230)
	$promptForm.StartPosition = "CenterParent"
	$promptForm.FormBorderStyle = "FixedDialog"
	$promptForm.MaximizeBox = $false
	$promptForm.MinimizeBox = $false
	$promptForm.ShowInTaskbar = $false
	
	$messageTextBox = New-Object System.Windows.Forms.TextBox
	$messageTextBox.Multiline = $true
	$messageTextBox.ReadOnly = $true
	$messageTextBox.BorderStyle = "None"
	$messageTextBox.BackColor = $promptForm.BackColor
	$messageTextBox.TabStop = $false
	$messageTextBox.Text = $Text
	$messageTextBox.SetBounds(16, 16, 570, 110)
	
	$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
	$buttonPanel.Dock = "Bottom"
	$buttonPanel.Height = 48
	$buttonPanel.FlowDirection = "RightToLeft"
	$buttonPanel.WrapContents = $false
	$buttonPanel.Padding = New-Object System.Windows.Forms.Padding(0, 6, 12, 6)
	
	$promptForm.Tag = "Closed"
	
	$primaryButton = New-Object System.Windows.Forms.Button
	$primaryButton.Text = $PrimaryButtonText
	$primaryButton.AutoSize = $true
	$primaryButton.Add_Click({
			$promptForm.Tag = "Primary"
			$promptForm.Close()
		})
	$buttonPanel.Controls.Add($primaryButton)
	$promptForm.AcceptButton = $primaryButton
	$promptForm.CancelButton = $primaryButton
	
	if (-not [string]::IsNullOrWhiteSpace($SecondaryButtonText))
	{
		$secondaryButton = New-Object System.Windows.Forms.Button
		$secondaryButton.Text = $SecondaryButtonText
		$secondaryButton.AutoSize = $true
		$secondaryButton.Add_Click({
				$promptForm.Tag = "Secondary"
				$promptForm.Close()
			})
		$buttonPanel.Controls.Add($secondaryButton)
		$promptForm.CancelButton = $secondaryButton
	}
	
	if (-not [string]::IsNullOrWhiteSpace($LogText))
	{
		$viewLogButton = New-Object System.Windows.Forms.Button
		$viewLogButton.Text = "View Log"
		$viewLogButton.AutoSize = $true
		$viewLogButton.Add_Click({
				Show_Text_Report_Form -Title $LogTitle -Text $LogText
			})
		$buttonPanel.Controls.Add($viewLogButton)
	}
	
	$promptForm.Controls.Add($messageTextBox)
	$promptForm.Controls.Add($buttonPanel)
	Set_Modern_Dialog_Style -Form $promptForm -PrimaryButton $primaryButton
	[void]$promptForm.ShowDialog()
	$selection = [string]$promptForm.Tag
	$promptForm.Dispose()
	return $selection
}

# ===================================================================================================
#                                 FUNCTION: Show_App_Message
# ---------------------------------------------------------------------------------------------------
# Description:
# Small wrapper around WinForms MessageBox to reduce repeated enum boilerplate.
# ===================================================================================================

function Show_App_Message
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$Text,
		[Parameter(Mandatory = $true)]
		[string]$Title,
		[Parameter(Mandatory = $false)]
		[ValidateSet('OK', 'OKCancel', 'YesNo', 'YesNoCancel')]
		[string]$Buttons = 'OK',
		[Parameter(Mandatory = $false)]
		[ValidateSet('None', 'Information', 'Warning', 'Error', 'Question', 'Asterisk', 'Exclamation', 'Hand', 'Stop')]
		[string]$Icon = 'Information'
	)
	
	switch -Regex ($Icon)
	{
		'^(?i)Asterisk$'    { $Icon = 'Information'; break }
		'^(?i)Exclamation$' { $Icon = 'Warning'; break }
		'^(?i)Hand$'        { $Icon = 'Error'; break }
		'^(?i)Stop$'        { $Icon = 'Error'; break }
	}
	
	$buttonValue = [System.Enum]::Parse([System.Windows.Forms.MessageBoxButtons], $Buttons)
	$iconValue = [System.Enum]::Parse([System.Windows.Forms.MessageBoxIcon], $Icon)
	return [System.Windows.Forms.MessageBox]::Show($Text, $Title, $buttonValue, $iconValue)
}

# ===================================================================================================
#                                FUNCTION: Show_Text_Report_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Shows a read-only text report (summary, failure log) in a resizable window with a Close button.
# ===================================================================================================

function Show_Text_Report_Form
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$Title,
		[Parameter(Mandatory = $true)]
		[string]$Text,
		[Parameter(Mandatory = $false)]
		[int]$Width = 720,
		[Parameter(Mandatory = $false)]
		[int]$Height = 500
	)
	
	$reportForm = New-Object System.Windows.Forms.Form
	$reportForm.Text = $Title
	$reportForm.Size = New-Object System.Drawing.Size($Width, $Height)
	$reportForm.StartPosition = "CenterParent"
	$reportForm.ShowInTaskbar = $false
	
	$closeButton = New-Object System.Windows.Forms.Button
	$closeButton.Text = "Close"
	$closeButton.Dock = "Bottom"
	$closeButton.Height = 34
	$closeButton.Add_Click({ $reportForm.Close() })
	
	$textBox = New-Object System.Windows.Forms.TextBox
	$textBox.Multiline = $true
	$textBox.ReadOnly = $true
	$textBox.ScrollBars = "Vertical"
	$textBox.Dock = "Fill"
	$textBox.Font = New-Object System.Drawing.Font("Consolas", 9)
	$textBox.Text = $Text
	
	$reportForm.AcceptButton = $closeButton
	$reportForm.Controls.Add($textBox)
	$reportForm.Controls.Add($closeButton)
	Set_Modern_Dialog_Style -Form $reportForm -PrimaryButton $closeButton
	[void]$reportForm.ShowDialog()
}

# ===================================================================================================
#                              FUNCTION: Get_Effective_Machine_Name
# ---------------------------------------------------------------------------------------------------
# Description:
# Returns the pending/new machine name when one exists; otherwise returns the current computer name.
# ===================================================================================================

function Get_Effective_Machine_Name
{
	[CmdletBinding()]
	param ()
	
	if ($script:newMachineName -and -not [string]::IsNullOrWhiteSpace([string]$script:newMachineName))
	{
		return [string]$script:newMachineName
	}
	
	return $env:COMPUTERNAME
}

# ===================================================================================================
#                           FUNCTION: Get_Machine_Name_Context
# ---------------------------------------------------------------------------------------------------
# Description:
# Normalizes a requested machine name and derives the terminal behavior used by INI + SQL sync paths.
# ===================================================================================================

function Get_Machine_Name_Context
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[string]$CurrentMachineName = $env:COMPUTERNAME,
		[Parameter(Mandatory = $false)]
		[ValidatePattern('^\d{3}$')]
		[string]$ServerTerminal = '901'
	)
	
	$normalizedHost = [string]$MachineName
	if ($null -eq $normalizedHost) { $normalizedHost = "" }
	$normalizedHost = $normalizedHost.Trim()
	$normalizedHost = $normalizedHost -replace '^[\\\/]+', ''
	if ($normalizedHost -match '[\\\/]') { $normalizedHost = ($normalizedHost -split '[\\\/]')[0] }
	if ($normalizedHost -match '\.') { $normalizedHost = ($normalizedHost -split '\.')[0] }
	$normalizedHost = $normalizedHost.Trim().ToUpper()
	
	$currentHost = [string]$CurrentMachineName
	if ($null -eq $currentHost) { $currentHost = "" }
	$currentHost = $currentHost.Trim()
	$currentHost = $currentHost -replace '^[\\\/]+', ''
	if ($currentHost -match '[\\\/]') { $currentHost = ($currentHost -split '[\\\/]')[0] }
	if ($currentHost -match '\.') { $currentHost = ($currentHost -split '\.')[0] }
	$currentHost = $currentHost.Trim().ToUpper()
	
	$nodeRole = "Unknown"
	try
	{
		if (-not [string]::IsNullOrWhiteSpace([string]$script:NodeRole))
		{
			$nodeRole = ([string]$script:NodeRole).Trim()
		}
	}
	catch { }
	
	$isServerNodeRole = ($nodeRole -match '^(?i)(StoreServer|HostServer)$')
	$isLaneNodeRole = ($nodeRole -match '^(?i)Lane$')
	$machineNameFormatText = if ($isServerNodeRole)
	{
		"Valid examples:`n  SERVER001`n  SERVER-001`n  0231SERVER001`n`nRule: [letters] + [optional store digits before or after the prefix] + [optional hyphen] + [3-digit terminal]."
	}
	elseif ($isLaneNodeRole)
	{
		"Valid examples:`n  POS001`n  POS-001`n  LN026001`n  0384LANE001`n  001`n`nRule: [letters] + [optional store digits before or after the prefix] + [optional hyphen] + [3-digit terminal]."
	}
	else
	{
		"Lane examples:`n  POS001`n  POS-001`n  LN026001`n  0384LANE001`n  001`n`nServer examples:`n  SERVER001`n  SERVER-001`n  0231SERVER001`n`nRule: [letters] + [optional store digits before or after the prefix] + [optional hyphen] + [3-digit terminal]."
	}
	$namePrefix = $null
	$nameSeparator = ""
	$storePrefixRaw = $null
	$storePlacement = "None"
	$terminalRaw = $null
	$inputWasTerminalOnly = $false
	
	if ([string]::IsNullOrWhiteSpace($normalizedHost))
	{
		return [PSCustomObject]@{
			Success            = $false
			ErrorMessage       = "Machine name cannot be empty."
			MachineName        = $null
			CurrentMachineName = $currentHost
			IsSameMachineName  = $false
			IsServerMoniker    = $false
			MachineNumber      = $null
			NodeRole           = $nodeRole
			NamePrefix         = $null
			NameSeparator      = ""
			StorePrefix        = $null
			StorePlacement     = "None"
			TerminalNumber     = $null
			InputWasTerminalOnly = $false
		}
	}
	
	if ($normalizedHost -match '^\d{3}$')
	{
		if ($isServerNodeRole)
		{
			return [PSCustomObject]@{
				Success              = $false
				ErrorMessage         = "Terminal-only input is only supported for lane-style machine names."
				MachineName          = $normalizedHost
				CurrentMachineName   = $currentHost
				IsSameMachineName    = $false
				IsServerMoniker      = $true
				MachineNumber        = $null
				NodeRole             = $nodeRole
				NamePrefix           = $null
				NameSeparator        = ""
				StorePrefix          = $null
				StorePlacement       = "None"
				TerminalNumber       = $normalizedHost
				InputWasTerminalOnly = $true
			}
		}
		
		$currentStore = $null
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('StoreNumber') -and $script:FunctionResults['StoreNumber'])
		{
			$currentStore = ($script:FunctionResults['StoreNumber'].ToString()).Trim()
		}
		else
		{
			$currentStore = Get_Store_Number_From_INI
			if ($currentStore) { $currentStore = ($currentStore.ToString()).Trim() }
		}
		
		if ($currentHost -match '^(?<StorePrefix>\d{3,4})(?<NamePrefix>[A-Z]+)(?<NameSeparator>-?)(?<Terminal>\d{3})$')
		{
			$storePrefixRaw = $matches['StorePrefix']
			$namePrefix = $matches['NamePrefix']
			$nameSeparator = $matches['NameSeparator']
			$storePlacement = "Leading"
		}
		elseif ($currentHost -match '^(?<NamePrefix>[A-Z]+)(?<StorePrefix>\d{3,4})(?<NameSeparator>-?)(?<Terminal>\d{3})$')
		{
			$namePrefix = $matches['NamePrefix']
			$storePrefixRaw = $matches['StorePrefix']
			$nameSeparator = $matches['NameSeparator']
			$storePlacement = "Embedded"
		}
		elseif ($currentHost -match '^(?<NamePrefix>[A-Z]+)(?<NameSeparator>-?)(?<Terminal>\d{3})$')
		{
			$namePrefix = $matches['NamePrefix']
			$nameSeparator = $matches['NameSeparator']
			$storePlacement = "None"
		}
		else
		{
			return [PSCustomObject]@{
				Success              = $false
				ErrorMessage         = "Terminal-only input requires the current machine name to already use a lane naming style with a reusable prefix."
				MachineName          = $normalizedHost
				CurrentMachineName   = $currentHost
				IsSameMachineName    = $false
				IsServerMoniker      = $false
				MachineNumber        = $null
				NodeRole             = $nodeRole
				NamePrefix           = $null
				NameSeparator        = ""
				StorePrefix          = $null
				StorePlacement       = "None"
				TerminalNumber       = $normalizedHost
				InputWasTerminalOnly = $true
			}
		}
		
		if (($storePlacement -ne "None") -and [string]::IsNullOrWhiteSpace($storePrefixRaw) -and ($currentStore -match '^\d{3,4}$'))
		{
			$storePrefixRaw = $currentStore
		}
		
		if (($storePlacement -ne "None") -and [string]::IsNullOrWhiteSpace($storePrefixRaw))
		{
			return [PSCustomObject]@{
				Success              = $false
				ErrorMessage         = "Terminal-only input could not determine which store number to preserve from the current naming style."
				MachineName          = $normalizedHost
				CurrentMachineName   = $currentHost
				IsSameMachineName    = $false
				IsServerMoniker      = $false
				MachineNumber        = $null
				NodeRole             = $nodeRole
				NamePrefix           = $namePrefix
				NameSeparator        = $nameSeparator
				StorePrefix          = $null
				StorePlacement       = $storePlacement
				TerminalNumber       = $normalizedHost
				InputWasTerminalOnly = $true
			}
		}
		
		$terminalRaw = $normalizedHost
		$inputWasTerminalOnly = $true
		
		switch ($storePlacement)
		{
			"Leading"  { $normalizedHost = "$storePrefixRaw$namePrefix$nameSeparator$terminalRaw" }
			"Embedded" { $normalizedHost = "$namePrefix$storePrefixRaw$nameSeparator$terminalRaw" }
			default    { $normalizedHost = "$namePrefix$nameSeparator$terminalRaw" }
		}
	}
	elseif ($normalizedHost -match '^(?<StorePrefix>\d{3,4})(?<NamePrefix>[A-Z]+)(?<NameSeparator>-?)(?<Terminal>\d{3})$')
	{
		$storePrefixRaw = $matches['StorePrefix']
		$namePrefix = $matches['NamePrefix']
		$nameSeparator = $matches['NameSeparator']
		$terminalRaw = $matches['Terminal']
		$storePlacement = "Leading"
	}
	elseif ($normalizedHost -match '^(?<NamePrefix>[A-Z]+)(?<StorePrefix>\d{3,4})(?<NameSeparator>-?)(?<Terminal>\d{3})$')
	{
		$namePrefix = $matches['NamePrefix']
		$storePrefixRaw = $matches['StorePrefix']
		$nameSeparator = $matches['NameSeparator']
		$terminalRaw = $matches['Terminal']
		$storePlacement = "Embedded"
	}
	elseif ($normalizedHost -match '^(?<NamePrefix>[A-Z]+)(?<NameSeparator>-?)(?<Terminal>\d{3})$')
	{
		$namePrefix = $matches['NamePrefix']
		$nameSeparator = $matches['NameSeparator']
		$terminalRaw = $matches['Terminal']
		$storePlacement = "None"
	}
	else
	{
		return [PSCustomObject]@{
			Success              = $false
			ErrorMessage         = "Invalid format.`n`n$machineNameFormatText"
			MachineName          = $normalizedHost
			CurrentMachineName   = $currentHost
			IsSameMachineName    = $false
			IsServerMoniker      = $isServerNodeRole
			MachineNumber        = $null
			NodeRole             = $nodeRole
			NamePrefix           = $null
			NameSeparator        = ""
			StorePrefix          = $null
			StorePlacement       = "None"
			TerminalNumber       = $null
			InputWasTerminalOnly = $false
		}
	}
	
	if ($normalizedHost.Length -gt 15)
	{
		return [PSCustomObject]@{
			Success              = $false
			ErrorMessage         = "Machine name '$normalizedHost' exceeds the Windows 15-character computer-name limit."
			MachineName          = $normalizedHost
			CurrentMachineName   = $currentHost
			IsSameMachineName    = $false
			IsServerMoniker      = $isServerNodeRole
			MachineNumber        = $null
			NodeRole             = $nodeRole
			NamePrefix           = $namePrefix
			NameSeparator        = $nameSeparator
			StorePrefix          = $storePrefixRaw
			StorePlacement       = $storePlacement
			TerminalNumber       = $terminalRaw
			InputWasTerminalOnly = $inputWasTerminalOnly
		}
	}
	
	$isServerMoniker = $isServerNodeRole
	if ($normalizedHost -match '(?i)SERVER')
	{
		$isServerMoniker = $true
	}
	
	$machineNumber = $null
	if ($isServerMoniker)
	{
		$machineNumber = $ServerTerminal
	}
	else
	{
		$machineNumber = ([int]$terminalRaw).ToString("D3")
		if ($machineNumber -eq '000')
		{
			return [PSCustomObject]@{
				Success              = $false
				ErrorMessage         = "Invalid terminal number extracted from '$normalizedHost' (000 is not allowed)."
				MachineName          = $normalizedHost
				CurrentMachineName   = $currentHost
				IsSameMachineName    = $false
				IsServerMoniker      = $false
				MachineNumber        = $null
				NodeRole             = $nodeRole
				NamePrefix           = $namePrefix
				NameSeparator        = $nameSeparator
				StorePrefix          = $storePrefixRaw
				StorePlacement       = $storePlacement
				TerminalNumber       = $terminalRaw
				InputWasTerminalOnly = $inputWasTerminalOnly
			}
		}
	}
	
	return [PSCustomObject]@{
		Success              = $true
		ErrorMessage         = $null
		MachineName          = $normalizedHost
		CurrentMachineName   = $currentHost
		IsSameMachineName    = ($normalizedHost -eq $currentHost)
		IsServerMoniker      = $isServerMoniker
		MachineNumber        = $machineNumber
		NodeRole             = $nodeRole
		NamePrefix           = $namePrefix
		NameSeparator        = $nameSeparator
		StorePrefix          = $storePrefixRaw
		StorePlacement       = $storePlacement
		TerminalNumber       = $terminalRaw
		InputWasTerminalOnly = $inputWasTerminalOnly
	}
}

# ===================================================================================================
#                                      FUNCTION: Get_Store_Name_From_INI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the store name from the system.ini file.
#   Stores the result in $script:FunctionResults['StoreName'].
# ===================================================================================================

function Get_Store_Name_From_INI
{
	param (
		[string]$INIPath = $SystemIniPath
	)
	
	# Initialize StoreName
	$script:FunctionResults['StoreName'] = "N/A"
	
	if (Test-Path $INIPath)
	{
		# First NAME= line only, and everything after the first '=' (a name may contain '=';
		# several NAME= lines used to produce an array that printed as one run-together label)
		$storeName = Select-String -Path $INIPath -Pattern "^NAME=" | Select-Object -First 1 | ForEach-Object {
			$_.Line.Split('=', 2)[1].Trim()
		}
		if ($storeName)
		{
			$script:FunctionResults['StoreName'] = $storeName
			# Write-Log "Store name found in system.ini: $storeName" "green"
		}
		else
		{
			# Write-Log "Store name not found in system.ini." "yellow"
		}
	}
	else
	{
		# Write-Log "INI file not found: $INIPath" "yellow"
	}
	
	# Update the storeNameLabel in the GUI
	if (-not $SilentMode -and $storeNameLabel -ne $null)
	{
		$storeNameLabel.Text = "$($script:FunctionResults['StoreName'])"
		$form.Refresh()
		[System.Windows.Forms.Application]::DoEvents()
	}
}

# ===================================================================================================
#                                       FUNCTION: Get_Store_Number_From_INI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the store number from the startup.ini file.
# ===================================================================================================

function Get_Store_Number_From_INI
{
	[CmdletBinding()]
	param (
		[switch]$UpdateLabel
	)
	
	# Ensure FunctionResults exists (PS 5.1 safe)
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
	
	# Default/fallback
	$script:FunctionResults['StoreNumber'] = "N/A"
	$foundStore = $null
	
	if (Test-Path $startupIniPath)
	{
		try
		{
			$iniContent = Get-Content $startupIniPath -ErrorAction Stop
			
			foreach ($line in $iniContent)
			{
				if ($line -match '^\s*STORE\s*=\s*(\d{3,4})\s*$')
				{
					$foundStore = $matches[1]
					$script:FunctionResults['StoreNumber'] = $foundStore
					break
				}
			}
		}
		catch
		{
			Write-Warning ("Failed to read startup.ini: {0}" -f $_)
		}
	}
	
	# Update label ONLY if requested
	if ($UpdateLabel -and (-not $SilentMode) -and $script:storeNumberLabel)
	{
		Set_Label_Text -Label $script:storeNumberLabel -Text "$($script:FunctionResults['StoreNumber'])" -RefreshParent -ProcessEvents
	}
	
	if ($foundStore) { return $foundStore }
	return $null
}

# ===================================================================================================
#                           FUNCTION: Resolve_Startup_Ini_Path
# ---------------------------------------------------------------------------------------------------
# Description:
# Resolves the active Startup.ini path using an explicit value first, then existing script/global
# variables, and finally BasePath when available.
# ===================================================================================================

function Resolve_Startup_Ini_Path
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath
	)
	
	$resolvedPath = $StartupIniPath
	
	try
	{
		if ([string]::IsNullOrWhiteSpace($resolvedPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name StartupIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value))
				{
					$resolvedPath = [string]$v.Value
					break
				}
			}
		}
		
		if ([string]::IsNullOrWhiteSpace($resolvedPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name BasePath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value))
				{
					$candidatePath = Join-Path ([string]$v.Value) "Startup.ini"
					if (Test-Path $candidatePath)
					{
						$resolvedPath = $candidatePath
						break
					}
				}
			}
		}
		
		if (-not [string]::IsNullOrWhiteSpace($resolvedPath) -and (Test-Path $resolvedPath))
		{
			return $resolvedPath
		}
	}
	catch { }
	
	return $null
}

# ===================================================================================================
# 									FUNCTION: Update_INIs
# ---------------------------------------------------------------------------------------------------
# Description:
# Updates ALL required INI files with the new store number.
# Supports both 3-digit and 4-digit store numbers.
# NOTE: No backups are created.
#
# Always-best-effort behavior:
# - Does NOT prompt
# - If MachineName missing -> uses $env:COMPUTERNAME
# - If terminal digits cannot be extracted -> still updates other INIs; skips TER insertion where unsafe
# - If old store cannot be detected -> still forces STORE= and REDIR lines; skips token-replace mapping
# - Startup.ini / Server.ini / SMSStart.ini / INFO_*901_WIN.ini preserve ORIGINAL encoding + ORIGINAL newline style
# - Updates Server.ini when present (same folder as Startup.ini first, then common roots)
#
# IMPORTANT SAFETY RULE:
# - If MachineName contains "SERVER" (anywhere, case-insensitive), we MUST NOT change terminal numbers.
#   That means:
#     * Do NOT derive terminal from name
#     * Do NOT rewrite TER= lines in Startup.ini / Server.ini / SMSStart.ini
#     * DO still update STORE= and REDIRMAIL/REDIRMSG to <STORE>901
#
# SmsHttps.INI behavior:
# - SERVER moniker:
#     * KEEP multiple processor entries
#     * ONLY update the store prefix in:
#         - processor key (e.g., 0231901 -> 0242901)
#         - REDIRMAIL/REDIRMSG values (e.g., 0231901 -> 0242901)
#     * Clear LicenseGUID (always)
# - NON-SERVER moniker:
#     * (Safe default here) we DO NOT touch processors (to avoid accidental changes)
#     * Clear LicenseGUID (always)
#
# NEW:
# - SMSStart.ini may contain PARAMETERS=/ini=Startup902 (or Startup902.ini).
#   We resolve that referenced INI and update it too (Store/Redir/token-replace).
#   We DO NOT force TER changes in the referenced Startup*.ini (we generally leave TER as-is).
# ===================================================================================================

function Update_INIs
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidatePattern('^(?!0+$)\d{3,4}$')]
		[string]$newStoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[string]$OldStoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineName,
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath,
		[Parameter(Mandatory = $false)]
		[string]$GlobalSmsStartIniPath,
		[Parameter(Mandatory = $false)]
		[string]$WinIniPath,
		[Parameter(Mandatory = $false)]
		[string]$SmsHttpsIniPath,
		# Accepted for the existing callers; SmsHttps.INI is only updated when it already exists.
		[Parameter(Mandatory = $false)]
		[switch]$CreateSmsHttpsIfMissing
	)
	
	$success = $true
	
	# ------------------------------------------------------------------------------------------------
	# Prefer already-built variables for paths (no hardcoded probing)
	# ------------------------------------------------------------------------------------------------
	try
	{
		if ([string]::IsNullOrWhiteSpace($StartupIniPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name StartupIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $StartupIniPath = [string]$v.Value; break }
			}
		}
		if ([string]::IsNullOrWhiteSpace($GlobalSmsStartIniPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name GlobalSmsStartIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $GlobalSmsStartIniPath = [string]$v.Value; break }
			}
		}
		if ([string]::IsNullOrWhiteSpace($WinIniPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name WinIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $WinIniPath = [string]$v.Value; break }
			}
		}
		if ([string]::IsNullOrWhiteSpace($SmsHttpsIniPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name SmsHttpsIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $SmsHttpsIniPath = [string]$v.Value; break }
			}
		}
	}
	catch { }
	
	# Optional: reuse already-built core paths if present (does not rebuild new ones)
	$BasePath = $null
	$DbsPath = $null
	$SystemIniVar = $null
	try
	{
		foreach ($sc in @('Script', 'Global'))
		{
			if (-not $BasePath)
			{
				$v = Get-Variable -Name BasePath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and $v.Value) { $BasePath = [string]$v.Value }
			}
			if (-not $DbsPath)
			{
				$v = Get-Variable -Name DbsPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and $v.Value) { $DbsPath = [string]$v.Value }
			}
			if (-not $SystemIniVar)
			{
				$v = Get-Variable -Name SystemIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and $v.Value) { $SystemIniVar = [string]$v.Value }
			}
		}
	}
	catch { }
	
	# If caller didn't supply MachineName, use current machine name (no prompt, no rename)
	if ([string]::IsNullOrWhiteSpace($MachineName))
	{
		$MachineName = $env:COMPUTERNAME
	}
	
	# ------------------------------------------------------------------------------------------------
	# Normalize NEW store (preserve user width for STORE=/REDIR lines)
	# ------------------------------------------------------------------------------------------------
	$newStoreTrim = ($newStoreNumber + "").Trim()
	$newStoreInt = [int]$newStoreTrim
	$newStore3 = $newStoreInt.ToString("D3")
	$newStore4 = $newStoreInt.ToString("D4")
	
	# ------------------------------------------------------------------------------------------------
	# Normalize machine name (strip UNC/path + domain) + ALWAYS uppercase
	# ------------------------------------------------------------------------------------------------
	$mn = ($MachineName + "").Trim()
	$mn = $mn -replace '^[\\\/]+', ''
	if ($mn -match '[\\\/]') { $mn = ($mn -split '[\\\/]')[0] }
	if ($mn -match '\.') { $mn = ($mn -split '\.')[0] }
	$mn = $mn.Trim().ToUpper()
	
	# ------------------------------------------------------------------------------------------------
	# SERVER MONIKER RULE: if name contains SERVER anywhere -> NEVER change terminal numbers
	# ------------------------------------------------------------------------------------------------
	$isServerMoniker = $false
	if ($mn -match '(?i)SERVER') { $isServerMoniker = $true }
	elseif ($script:NodeRole -and ($script:NodeRole -match '^(?i)(StoreServer|HostServer)$')) { $isServerMoniker = $true }
	
	# ------------------------------------------------------------------------------------------------
	# Derive terminal ONLY for NON-server moniker machines
	# ------------------------------------------------------------------------------------------------
	$newTerminal = $null
	if (-not $isServerMoniker -and ($mn -match '(\d{1,3})$'))
	{
		$newTerminal = ([int]$Matches[1]).ToString("D3")
		if ($newTerminal -eq "000") { $newTerminal = $null }
	}
	
	# ===============================================================================================
	# 1) Resolve + Update Startup.ini (required)  (no hardcoded probing; must already be set)
	# ===============================================================================================
	if ([string]::IsNullOrWhiteSpace($StartupIniPath) -and $BasePath)
	{
		try
		{
			$p = Join-Path $BasePath "Startup.ini"
			if (Test-Path $p) { $StartupIniPath = $p }
		}
		catch { }
	}
	
	if ([string]::IsNullOrWhiteSpace($StartupIniPath) -or -not (Test-Path $StartupIniPath))
	{
		Write-Host "startup.ini not found. Ensure `$StartupIniPath is already set (or pass -StartupIniPath)." -ForegroundColor Red
		return $false
	}
	
	$oldStoreDetected = $null
	$startupDir = $null
	try { $startupDir = Split-Path -Path $StartupIniPath -Parent }
	catch { $startupDir = $null }
	
	try
	{
		$bytes = [System.IO.File]::ReadAllBytes($StartupIniPath)
		
		$encStartup = $null
		if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
		{ $encStartup = New-Object System.Text.UTF8Encoding($true) }
		elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
		{ $encStartup = [System.Text.Encoding]::Unicode }
		elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
		{ $encStartup = [System.Text.Encoding]::BigEndianUnicode }
		else
		{ $encStartup = [System.Text.Encoding]::Default }
		
		$textStartup = $encStartup.GetString($bytes)
		
		$nlStartup = "`r`n"
		if ($textStartup -notmatch "`r`n" -and $textStartup -match "`n") { $nlStartup = "`n" }
		
		$startupLines = $textStartup -split "\r?\n", -1
		
		# Detect OLD store number (prefer STORE=)
		foreach ($l in $startupLines)
		{
			if ($l -match '^\s*STORE\s*=\s*(\d{3,4})\s*$') { $oldStoreDetected = $matches[1]; break }
		}
		if (-not $oldStoreDetected)
		{
			foreach ($l in $startupLines)
			{
				if ($l -match '^\s*SERVERNAME\s*=\s*(\d{3,4})[A-Za-z]') { $oldStoreDetected = $matches[1]; break }
			}
		}
		if (-not $oldStoreDetected)
		{
			foreach ($l in $startupLines)
			{
				if ($l -match '\\\\(\d{3,4})[A-Za-z]') { $oldStoreDetected = $matches[1]; break }
			}
		}
		
		# If user provided OldStoreNumber, override detection (if valid)
		if (-not [string]::IsNullOrWhiteSpace($OldStoreNumber))
		{
			$os = $OldStoreNumber.Trim()
			if ($os -match '^(?!0+$)\d{3,4}$') { $oldStoreDetected = $os }
		}
		
		$doTokenReplace = $true
		if (-not $oldStoreDetected) { $doTokenReplace = $false }
		
		$oldTokens = @()
		$newTokens = @()
		if ($doTokenReplace)
		{
			$oldStoreInt = [int]$oldStoreDetected
			$oldTokens = @($oldStoreInt.ToString("D4"), $oldStoreInt.ToString("D3"))
			$newTokens = @($newStore4, $newStore3)
		}
		
		for ($i = 0; $i -lt $startupLines.Count; $i++)
		{
			$line = $startupLines[$i]
			
			$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
			$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
			
			# TER only when NON-server and we have a derived terminal
			if (-not $isServerMoniker -and $newTerminal)
			{
				$line = $line -replace '^\s*TER\s*=\s*\d{1,4}\s*$', ("TER=" + $newTerminal)
			}
			
			# Always uppercase SERVERNAME value if present
			if ($line -match '^(?<ws>\s*)(?i:SERVERNAME)\s*=\s*(?<val>.*)\s*$')
			{
				$ws = $matches['ws']
				$val = ($matches['val'] + "").Trim()
				if ($val.Length -gt 0) { $line = $ws + "SERVERNAME=" + $val.ToUpper() }
			}
			
			# Token replace mapping (SAFE: do not rewrite terminal suffixes like SERVER001)
			if ($doTokenReplace)
			{
				for ($t = 0; $t -lt $oldTokens.Count; $t++)
				{
					$oldEsc = [regex]::Escape($oldTokens[$t])
					$newTok = $newTokens[$t]
					
					$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
					$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
					$line = $line -replace ("(?<![\dA-Za-z.])" + $oldEsc + "(?!\d|\.\d)"), $newTok
				}
			}
			
			$startupLines[$i] = $line
		}
		
		[System.IO.File]::WriteAllText($StartupIniPath, ($startupLines -join $nlStartup), $encStartup)
		Write-Host "Updated startup.ini" -ForegroundColor Green
		
		if (-not $oldStoreDetected) { $oldStoreDetected = $newStoreTrim }
	}
	catch
	{
		$success = $false
		Write-Host "Failed updating startup.ini: $($_.Exception.Message)" -ForegroundColor Red
		if (-not $oldStoreDetected) { $oldStoreDetected = $newStoreTrim }
	}
	
	# ===============================================================================================
	# 1B) Update Server.ini (optional) (relative to Startup.ini / BasePath only)
	# ===============================================================================================
	$serverIniPath = $null
	try
	{
		if ($startupDir)
		{
			$p = Join-Path $startupDir "Server.ini"
			if (Test-Path $p) { $serverIniPath = $p }
		}
		if (-not $serverIniPath -and $BasePath)
		{
			$p = Join-Path $BasePath "Server.ini"
			if (Test-Path $p) { $serverIniPath = $p }
		}
	}
	catch { $serverIniPath = $null }
	
	if (-not [string]::IsNullOrWhiteSpace($serverIniPath) -and (Test-Path $serverIniPath))
	{
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($serverIniPath)
			
			$encServer = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{ $encServer = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{ $encServer = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{ $encServer = [System.Text.Encoding]::BigEndianUnicode }
			else
			{ $encServer = [System.Text.Encoding]::Default }
			
			$textServer = $encServer.GetString($bytes)
			$nlServer = "`r`n"
			if ($textServer -notmatch "`r`n" -and $textServer -match "`n") { $nlServer = "`n" }
			
			$serverLines = $textServer -split "\r?\n", -1
			
			$doTokenReplaceS = $false
			$oldTokensS = @()
			$newTokensS = @()
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$doTokenReplaceS = $true
				$osi = [int]$oldStoreDetected
				$oldTokensS = @($osi.ToString("D4"), $osi.ToString("D3"))
				$newTokensS = @($newStore4, $newStore3)
			}
			
			for ($i = 0; $i -lt $serverLines.Count; $i++)
			{
				$line = $serverLines[$i]
				
				$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
				$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
				
				if (-not $isServerMoniker -and $newTerminal)
				{
					$line = $line -replace '^\s*TER\s*=\s*\d{1,4}\s*$', ("TER=" + $newTerminal)
				}
				
				if ($line -match '^(?<ws>\s*)(?i:SERVERNAME)\s*=\s*(?<val>.*)\s*$')
				{
					$ws = $matches['ws']
					$val = ($matches['val'] + "").Trim()
					if ($val.Length -gt 0) { $line = $ws + "SERVERNAME=" + $val.ToUpper() }
				}
				
				if ($doTokenReplaceS)
				{
					for ($t = 0; $t -lt $oldTokensS.Count; $t++)
					{
						$oldEsc = [regex]::Escape($oldTokensS[$t])
						$newTok = $newTokensS[$t]
						
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
						$line = $line -replace ("(?<![\dA-Za-z.])" + $oldEsc + "(?!\d|\.\d)"), $newTok
					}
				}
				
				$serverLines[$i] = $line
			}
			
			[System.IO.File]::WriteAllText($serverIniPath, ($serverLines -join $nlServer), $encServer)
			Write-Host "Updated Server.ini" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating Server.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	
	# ===============================================================================================
	# 1C) Update Office\System.ini (optional) using existing $SystemIniPath if available
	# ===============================================================================================
	$systemIniPath = $null
	if (-not [string]::IsNullOrWhiteSpace($SystemIniVar) -and (Test-Path $SystemIniVar))
	{
		$systemIniPath = $SystemIniVar
	}
	elseif ($startupDir)
	{
		$p = Join-Path $startupDir "Office\System.ini"
		if (Test-Path $p) { $systemIniPath = $p }
	}
	
	if (-not [string]::IsNullOrWhiteSpace($systemIniPath) -and (Test-Path $systemIniPath))
	{
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($systemIniPath)
			
			$encSys = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{ $encSys = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{ $encSys = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{ $encSys = [System.Text.Encoding]::BigEndianUnicode }
			else
			{ $encSys = [System.Text.Encoding]::Default }
			
			$textSys = $encSys.GetString($bytes)
			
			$nlSys = "`r`n"
			if ($textSys -notmatch "`r`n" -and $textSys -match "`n") { $nlSys = "`n" }
			
			$linesSys = $textSys -split "\r?\n", -1
			
			$newStoreIsOne = $false
			try { if ([int]$newStoreTrim -eq 1) { $newStoreIsOne = $true } }
			catch { $newStoreIsOne = $false }
			
			$oldTokensSys = @()
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$osi = [int]$oldStoreDetected
				$oldTokensSys = @($osi.ToString("D4"), $osi.ToString("D3"))
			}
			
			$inSmsSection = $false
			$changedSys = $false
			
			for ($i = 0; $i -lt $linesSys.Count; $i++)
			{
				$line = $linesSys[$i]
				
				if ($line -match '^\s*\[(?<sec>.+?)\]\s*$')
				{
					$sec = $matches['sec'].Trim()
					$inSmsSection = ($sec -ieq 'SMS')
					continue
				}
				
				if (-not $inSmsSection) { continue }
				
				if ($line -match '^(?<ws>\s*)(?i:Name)\s*=\s*(?<val>.*)\s*$')
				{
					$ws = $matches['ws']
					$val = $matches['val']
					$origVal = $val
					
					$val2 = ($val + "").Trim()
					
					if ($newStoreIsOne)
					{
						$val2 = [regex]::Replace($val2, '\s+(?:0001|001|01|1)\s*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
						foreach ($tok in @('0001', '001', '01', '1'))
						{
							$val2 = [regex]::Replace($val2, '\s+' + [regex]::Escape($tok) + '\s*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
						}
						$val2 = $val2.Trim()
					}
					else
					{
						$replaced = $false
						foreach ($tok in $oldTokensSys)
						{
							if (-not [string]::IsNullOrWhiteSpace($tok) -and ($val2 -match ("(?<!\d)" + [regex]::Escape($tok) + "(?!\d)")))
							{
								$val2 = [regex]::Replace($val2, "(?<!\d)" + [regex]::Escape($tok) + "(?!\d)", $newStoreTrim)
								$replaced = $true
							}
						}
						
						if (-not $replaced -and ($val2 -match '\s+\d{3,4}\s*$'))
						{
							$val2 = [regex]::Replace($val2, '\s+\d{3,4}\s*$', (" " + $newStoreTrim))
							$replaced = $true
						}
						
						if (-not $replaced)
						{
							$val2 = ($val2.TrimEnd() + " " + $newStoreTrim).Trim()
						}
					}
					
					if ($val2 -ne $origVal)
					{
						$linesSys[$i] = ($ws + "Name=" + $val2)
						$changedSys = $true
					}
					
					break
				}
			}
			
			if ($changedSys)
			{
				[System.IO.File]::WriteAllText($systemIniPath, ($linesSys -join $nlSys), $encSys)
				Write-Host "Updated Office\System.ini (SMS Name)" -ForegroundColor Green
			}
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating Office\System.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	
	# ===============================================================================================
	# 2) Update Global SMSStart.ini (optional)
	#     + NEW: detect PARAMETERS=/ini=Startup### and update referenced INI too
	#     + NEW: update [TIMESYNC] SERVER=... (store prefix) + ENABLE=1 (insert if missing)
	# ===============================================================================================
	if (-not [string]::IsNullOrWhiteSpace($GlobalSmsStartIniPath) -and (Test-Path $GlobalSmsStartIniPath))
	{
		$referencedIniNames = @()
		
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($GlobalSmsStartIniPath)
			
			$encSmsStart = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{ $encSmsStart = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{ $encSmsStart = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{ $encSmsStart = [System.Text.Encoding]::BigEndianUnicode }
			else
			{ $encSmsStart = [System.Text.Encoding]::Default }
			
			$textSmsStart = $encSmsStart.GetString($bytes)
			$nlSmsStart = "`r`n"
			if ($textSmsStart -notmatch "`r`n" -and $textSmsStart -match "`n") { $nlSmsStart = "`n" }
			
			$globalLines = $textSmsStart -split "\r?\n", -1
			
			$doTokenReplace = $false
			$oldTokens = @()
			$newTokens = @()
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$doTokenReplace = $true
				$osi = [int]$oldStoreDetected
				$oldTokens = @($osi.ToString("D4"), $osi.ToString("D3"))
				$newTokens = @($newStore4, $newStore3)
			}
			
			$out = @()
			$inSmsStartSection = $false
			$terFound = $false
			$terLine = $null
			if (-not $isServerMoniker -and $newTerminal) { $terLine = "TER=$newTerminal" }
			
			$inTimeSyncSection = $false
			$tsEnableFound = $false
			
			foreach ($raw in $globalLines)
			{
				$line = $raw
				
				# Detect PARAMETERS=/ini=Startup902 (or Startup902.ini)
				if ($line -match '^\s*(?i:PARAMETERS)\s*=\s*(?<p>.*)\s*$')
				{
					$p = $matches['p']
					if (-not [string]::IsNullOrWhiteSpace($p))
					{
						if ($p -match '(?i)(?:^|\s)/ini\s*=\s*(?<ini>[^\s"]+)')
						{
							$iniName = $matches['ini'].Trim().Trim('"').Trim("'")
							if (-not [string]::IsNullOrWhiteSpace($iniName))
							{
								$referencedIniNames += $iniName
							}
						}
					}
				}
				
				if ($line -match '^\s*\[(.+?)\]\s*$')
				{
					# Leaving [SMSSTART] -> ensure TER inserted if needed
					if ($inSmsStartSection -and -not $terFound -and $terLine)
					{
						$out += $terLine
						$terFound = $true
					}
					
					# Leaving [TIMESYNC] -> ensure ENABLE=1 if missing
					if ($inTimeSyncSection -and -not $tsEnableFound)
					{
						$out += "ENABLE=1"
						$tsEnableFound = $true
					}
					
					$sectionName = $matches[1].Trim()
					$inSmsStartSection = ($sectionName -ieq 'SMSSTART')
					if ($inSmsStartSection) { $terFound = $false }
					
					$inTimeSyncSection = ($sectionName -ieq 'TIMESYNC')
					if ($inTimeSyncSection) { $tsEnableFound = $false }
					
					$out += $line
					continue
				}
				
				if ($inSmsStartSection)
				{
					$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
					$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
					
					# TER only when allowed (non-server)
					if ($terLine -and ($line -match '^\s*(?i:TER)\s*='))
					{
						$line = $terLine
						$terFound = $true
					}
					
					if ($doTokenReplace)
					{
						for ($t = 0; $t -lt $oldTokens.Count; $t++)
						{
							$oldEsc = [regex]::Escape($oldTokens[$t])
							$newTok = $newTokens[$t]
							$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
							$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
							$line = $line -replace ("(?<![\dA-Za-z.])" + $oldEsc + "(?!\d|\.\d)"), $newTok
						}
					}
				}
				elseif ($inTimeSyncSection)
				{
					# Force ENABLE=1 (and mark found)
					if ($line -match '^\s*(?i:ENABLE)\s*=')
					{
						$line = "ENABLE=1"
						$tsEnableFound = $true
					}
					
					# Update SERVER= store prefix (and force uppercase)
					if ($line -match '^\s*(?i:SERVER)\s*=\s*(?<sv>.*)\s*$')
					{
						$sv = ($matches['sv'] + "").Trim()
						
						if ($sv -match '^(?<st>\d{3,4})(?<rest>[A-Za-z].*)$')
						{
							$st = $matches['st']
							$rest = $matches['rest']
							$newSt = ([int]$newStoreTrim).ToString(("D{0}" -f $st.Length))
							$sv = $newSt + $rest
						}
						else
						{
							if ($doTokenReplace)
							{
								for ($t = 0; $t -lt $oldTokens.Count; $t++)
								{
									$oldEsc = [regex]::Escape($oldTokens[$t])
									$newTok = $newTokens[$t]
									$sv = $sv -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
									$sv = $sv -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
									$sv = $sv -replace ("(?<![\dA-Za-z.])" + $oldEsc + "(?!\d|\.\d)"), $newTok
								}
							}
						}
						
						$line = "SERVER=" + $sv.ToUpper()
					}
				}
				
				$out += $line
			}
			
			# If file ended while still in sections
			if ($inSmsStartSection -and -not $terFound -and $terLine)
			{
				$out += $terLine
			}
			if ($inTimeSyncSection -and -not $tsEnableFound)
			{
				$out += "ENABLE=1"
			}
			
			[System.IO.File]::WriteAllText($GlobalSmsStartIniPath, ($out -join $nlSmsStart), $encSmsStart)
			Write-Host "Updated SMSStart.ini (SMSSTART + TIMESYNC)" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating SMSStart.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
		
		# ---- Update referenced Startup*.ini files (PARAMETERS=/ini=...) ----
		if ($referencedIniNames -and $referencedIniNames.Count -gt 0)
		{
			$smsStartDir = $null
			try { $smsStartDir = Split-Path -Path $GlobalSmsStartIniPath -Parent }
			catch { $smsStartDir = $null }
			
			foreach ($iniRef in ($referencedIniNames | Select-Object -Unique))
			{
				try
				{
					$iniTarget = ($iniRef + "").Trim()
					if ([string]::IsNullOrWhiteSpace($iniTarget)) { continue }
					if ($iniTarget -notmatch '\.ini$') { $iniTarget = ($iniTarget + ".ini") }
					
					$candidatePaths = @()
					
					if ($iniTarget -match '^[A-Za-z]:\\' -or $iniTarget -match '^\\\\')
					{
						$candidatePaths += $iniTarget
					}
					else
					{
						if ($smsStartDir) { $candidatePaths += (Join-Path $smsStartDir $iniTarget) }
						if ($BasePath) { $candidatePaths += (Join-Path $BasePath    $iniTarget) }
						if ($startupDir) { $candidatePaths += (Join-Path $startupDir  $iniTarget) }
					}
					
					$resolved = $null
					foreach ($cp in $candidatePaths)
					{
						if (-not [string]::IsNullOrWhiteSpace($cp) -and (Test-Path $cp))
						{
							$resolved = $cp
							break
						}
					}
					
					if (-not $resolved) { continue }
					
					$b = [System.IO.File]::ReadAllBytes($resolved)
					
					$encRef = $null
					if ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
					{ $encRef = New-Object System.Text.UTF8Encoding($true) }
					elseif ($b.Length -ge 2 -and $b[0] -eq 0xFF -and $b[1] -eq 0xFE)
					{ $encRef = [System.Text.Encoding]::Unicode }
					elseif ($b.Length -ge 2 -and $b[0] -eq 0xFE -and $b[1] -eq 0xFF)
					{ $encRef = [System.Text.Encoding]::BigEndianUnicode }
					else
					{ $encRef = [System.Text.Encoding]::Default }
					
					$txt = $encRef.GetString($b)
					$nlRef = "`r`n"
					if ($txt -notmatch "`r`n" -and $txt -match "`n") { $nlRef = "`n" }
					
					$lines = $txt -split "\r?\n", -1
					
					$doTR = $false
					$ot = @()
					$nt = @()
					if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
					{
						$doTR = $true
						$osi = [int]$oldStoreDetected
						$ot = @($osi.ToString("D4"), $osi.ToString("D3"))
						$nt = @($newStore4, $newStore3)
					}
					
					for ($i = 0; $i -lt $lines.Count; $i++)
					{
						$line = $lines[$i]
						
						$line = [regex]::Replace($line, '^\s*(?i:STORE)\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim))
						$line = [regex]::Replace($line, '^\s*(?i:Store)\s*=\s*\d{1,4}\s*$', ("Store=" + $newStoreTrim))
						
						$line = [regex]::Replace(
							$line,
							'^\s*(?i:REDIRMAIL|REDIRMSG|RedirMail|RedirMsg)\s*=\s*\d{1,4}(901)\s*$',
							{
								param ($m)
								$key = $m.Value.Split('=')[0].Trim()
								return ($key + "=" + $newStoreTrim + "901")
							}
						)
						
						if ($line -match '^(?<ws>\s*)(?i:SERVERNAME)\s*=\s*(?<val>.*)\s*$')
						{
							$ws = $matches['ws']
							$val = ($matches['val'] + "").Trim()
							if ($val.Length -gt 0) { $line = $ws + "SERVERNAME=" + $val.ToUpper() }
						}
						
						# Do NOT force TER changes in referenced Startup*.ini
						
						if ($doTR)
						{
							for ($t = 0; $t -lt $ot.Count; $t++)
							{
								$oldEsc = [regex]::Escape($ot[$t])
								$newTok = $nt[$t]
								$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
								$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
								$line = $line -replace ("(?<![\dA-Za-z.])" + $oldEsc + "(?!\d|\.\d)"), $newTok
							}
						}
						
						$lines[$i] = $line
					}
					
					[System.IO.File]::WriteAllText($resolved, ($lines -join $nlRef), $encRef)
					Write-Host ("Updated referenced INI: " + $resolved) -ForegroundColor Green
				}
				catch
				{
					$success = $false
					Write-Host ("Failed updating referenced INI '" + $iniRef + "': " + $_.Exception.Message) -ForegroundColor Red
				}
			}
		}
	}
	
	# ===============================================================================================
	# 3) Update INFO_*901_WIN.ini (optional) (use variable; small fallback via $DbsPath only)
	# ===============================================================================================
	if ([string]::IsNullOrWhiteSpace($WinIniPath) -and $DbsPath -and (Test-Path $DbsPath))
	{
		try
		{
			$f = Get-ChildItem -Path $DbsPath -File -Filter "INFO_*901_WIN.ini" -ErrorAction SilentlyContinue | Select-Object -First 1
			if ($f) { $WinIniPath = $f.FullName }
		}
		catch { }
	}
	
	if (-not [string]::IsNullOrWhiteSpace($WinIniPath) -and (Test-Path $WinIniPath))
	{
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($WinIniPath)
			
			$encWin = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{ $encWin = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{ $encWin = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{ $encWin = [System.Text.Encoding]::BigEndianUnicode }
			else
			{ $encWin = [System.Text.Encoding]::Default }
			
			$textWin = $encWin.GetString($bytes)
			$nlWin = "`r`n"
			if ($textWin -notmatch "`r`n" -and $textWin -match "`n") { $nlWin = "`n" }
			
			$winLines = $textWin -split "\r?\n", -1
			
			$doTokenReplaceW = $false
			$oldTokensW = @()
			$newTokensW = @()
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$doTokenReplaceW = $true
				$osi = [int]$oldStoreDetected
				$oldTokensW = @($osi.ToString("D4"), $osi.ToString("D3"))
				$newTokensW = @($newStore4, $newStore3)
			}
			
			for ($j = 0; $j -lt $winLines.Count; $j++)
			{
				$line = $winLines[$j]
				$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
				$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
				
				if ($doTokenReplaceW)
				{
					for ($t = 0; $t -lt $oldTokensW.Count; $t++)
					{
						$oldEsc = [regex]::Escape($oldTokensW[$t])
						$newTok = $newTokensW[$t]
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
						$line = $line -replace ("(?<![\dA-Za-z.])" + $oldEsc + "(?!\d|\.\d)"), $newTok
					}
				}
				
				$winLines[$j] = $line
			}
			
			[System.IO.File]::WriteAllText($WinIniPath, ($winLines -join $nlWin), $encWin)
			Write-Host "Updated INFO_*901_WIN.ini" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating INFO_*901_WIN.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	
	# ===============================================================================================
	# 4) Update SmsHttps.INI (FORCED)
	#   - SERVER moniker:
	#       * KEEP multiple processors
	#       * ONLY update store prefix in processor keys + REDIRMAIL/REDIRMSG values
	#       * Always clear LicenseGUID
	#   - NON-SERVER:
	#       * Enforce EXACTLY ONE processor entry keyed <STORE><TER> (store width inferred from file when possible)
	#       * Purge other processor lines
	#       * Force REDIRMAIL/REDIRMSG=<STORE>901
	#       * Always clear LicenseGUID
	#   - ALWAYS writes the file (even if identical), so it "triggers" every run.
	# ===============================================================================================
	
	if (-not [string]::IsNullOrWhiteSpace($SmsHttpsIniPath) -and (Test-Path $SmsHttpsIniPath))
	{
		try
		{
			# If terminal couldn't be derived from machine name, try TER= in Startup.ini (NON-server only)
			if (-not $isServerMoniker -and -not $newTerminal -and -not [string]::IsNullOrWhiteSpace($StartupIniPath) -and (Test-Path $StartupIniPath))
			{
				try
				{
					$terLine = (Get-Content -Path $StartupIniPath -ErrorAction Stop | Where-Object { $_ -match '^\s*TER\s*=\s*\d{1,4}\s*$' } | Select-Object -First 1)
					if ($terLine -and ($terLine -match '^\s*TER\s*=\s*(\d{1,4})\s*$'))
					{
						$tmpTer = $matches[1]
						if ($tmpTer -match '^\d{1,3}$')
						{
							$newTerminal = ([int]$tmpTer).ToString("D3")
							if ($newTerminal -eq "000") { $newTerminal = $null }
						}
					}
				}
				catch { }
			}
			
			$bytes = [System.IO.File]::ReadAllBytes($SmsHttpsIniPath)
			
			$enc = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $enc = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $enc = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { $enc = [System.Text.Encoding]::BigEndianUnicode }
			else { $enc = [System.Text.Encoding]::Default }
			
			$text = $enc.GetString($bytes)
			
			$nl = "`r`n"
			if ($text -notmatch "`r`n" -and $text -match "`n") { $nl = "`n" }
			
			$lines = $text -split "\r?\n", -1
			
			# --- Ensure/clear [GENERAL] LicenseGUID ---
			$genStart = -1
			for ($i = 0; $i -lt $lines.Length; $i++)
			{
				if ($lines[$i] -match '^\s*\[\s*GENERAL\s*\]\s*$') { $genStart = $i; break }
			}
			if ($genStart -lt 0)
			{
				$lines = @("[GENERAL]", "LicenseGUID=", "") + $lines
			}
			else
			{
				$genEnd = $lines.Length
				for ($i = $genStart + 1; $i -lt $lines.Length; $i++)
				{
					if ($lines[$i] -match '^\s*\[.*\]\s*$') { $genEnd = $i; break }
				}
				
				$lgFound = $false
				for ($i = $genStart + 1; $i -lt $genEnd; $i++)
				{
					if ($lines[$i] -match '^(?<ws>\s*)(?i:LicenseGUID)\s*=\s*(?<val>.*)\s*$')
					{
						$lines[$i] = ($matches['ws'] + "LicenseGUID=")
						$lgFound = $true
						break
					}
				}
				if (-not $lgFound)
				{
					$pre = @()
					$post = @()
					for ($k = 0; $k -lt $genEnd; $k++) { $pre += $lines[$k] }
					for ($k = $genEnd; $k -lt $lines.Length; $k++) { $post += $lines[$k] }
					$lines = @($pre + "LicenseGUID=" + $post)
				}
			}
			
			# --- Find [PROCESSORS] section (create if missing) ---
			$secStart = -1
			for ($i = 0; $i -lt $lines.Length; $i++)
			{
				if ($lines[$i] -match '^\s*\[\s*PROCESSORS?\s*\]\s*$') { $secStart = $i; break }
			}
			
			if ($secStart -lt 0)
			{
				# Always create it (forced behavior)
				$lines = @($lines + "" + "[PROCESSORS]")
				$secStart = ($lines.Length - 1)
			}
			
			$secEnd = $lines.Length
			for ($i = $secStart + 1; $i -lt $lines.Length; $i++)
			{
				if ($lines[$i] -match '^\s*\[.*\]\s*$') { $secEnd = $i; break }
			}
			
			# --- Infer store width from existing processor keys when possible (6=3+3, 7=4+3) ---
			$storeWidth = $newStoreTrim.Length
			try
			{
				for ($i = $secStart + 1; $i -lt $secEnd; $i++)
				{
					if ($lines[$i] -match '^\s*(\d{6,7})\s*=')
					{
						$keyLen = $matches[1].Length
						if ($keyLen -eq 7) { $storeWidth = 4; break }
						if ($keyLen -eq 6) { $storeWidth = 3; break }
					}
				}
			}
			catch { }
			
			$storeNorm = ([int]$newStoreTrim).ToString(("D{0}" -f $storeWidth))
			$redirKey901 = $storeNorm + "901"
			
			# =========================
			# SERVER: keep many processors; update store prefix in key + REDIR store prefix
			# =========================
			if ($isServerMoniker)
			{
				for ($i = $secStart + 1; $i -lt $secEnd; $i++)
				{
					$line = $lines[$i]
					if ($line -match '^(?<ws>\s*)(?<store>\d{3,4})(?<term>\d{3})\s*=\s*(?<rhs>.*)$')
					{
						$ws = $matches['ws']
						$st = $matches['store']
						$term = $matches['term']
						$rhs = $matches['rhs']
						
						$newStorePart = ([int]$newStoreTrim).ToString(("D{0}" -f $st.Length))
						$newKey = $newStorePart + $term
						
						$rhs2 = $rhs
						$rhs2 = [regex]::Replace($rhs2, '(?i)\b(REDIRMAIL)\s*=\s*(\d{3,4})(\d{3})\b', {
								param ($m)
								$k = $m.Groups[1].Value
								$st2 = $m.Groups[2].Value
								$tr2 = $m.Groups[3].Value
								$nst2 = ([int]$newStoreTrim).ToString(("D{0}" -f $st2.Length))
								return ($k + "=" + $nst2 + $tr2)
							})
						$rhs2 = [regex]::Replace($rhs2, '(?i)\b(REDIRMSG)\s*=\s*(\d{3,4})(\d{3})\b', {
								param ($m)
								$k = $m.Groups[1].Value
								$st2 = $m.Groups[2].Value
								$tr2 = $m.Groups[3].Value
								$nst2 = ([int]$newStoreTrim).ToString(("D{0}" -f $st2.Length))
								return ($k + "=" + $nst2 + $tr2)
							})
						
						$lines[$i] = ($ws + $newKey + "=" + $rhs2)
					}
				}
			}
			# =========================
			# NON-SERVER: enforce ONE processor keyed <STORE><TER> and force REDIRs to <STORE>901
			# =========================
			else
			{
				if ($newTerminal)
				{
					$newKey = $storeNorm + $newTerminal
					
					# Collect existing processor entries (numeric)
					$entries = @()
					for ($i = $secStart + 1; $i -lt $secEnd; $i++)
					{
						if ($lines[$i] -match '^(?<ws>\s*)(?<key>\d{6,7})\s*=\s*(?<rhs>.*)$')
						{
							$entries += [pscustomobject]@{
								Index = $i
								Ws    = $matches['ws']
								Key   = $matches['key']
								Rhs   = $matches['rhs']
							}
						}
					}
					
					# Pick a source rhs (same logic as you had, but without requiring old info to exist)
					$source = $null
					foreach ($e in $entries)
					{
						if ($e.Key.Length -ge 3 -and $e.Key.Substring($e.Key.Length - 3) -eq $newTerminal) { $source = $e; break }
					}
					if (-not $source -and $entries.Count -gt 0) { $source = $entries[0] }
					
					$rhsUpdated = ""
					$wsToUse = ""
					if ($source)
					{
						$rhsUpdated = $source.Rhs
						$wsToUse = $source.Ws
					}
					
					# Replace any old processor keys inside RHS -> newKey
					foreach ($e in $entries)
					{
						$rhsUpdated = $rhsUpdated -replace ("(?<!\d)" + [regex]::Escape($e.Key) + "(?!\d)"), $newKey
					}
					
					# Strip existing REDIRMAIL/REDIRMSG then enforce <STORE>901
					$rhsUpdated = [regex]::Replace($rhsUpdated, '\bREDIRMAIL\s*=\s*\d{6,7}', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
					$rhsUpdated = [regex]::Replace($rhsUpdated, '\bREDIRMSG\s*=\s*\d{6,7}', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
					$rhsUpdated = $rhsUpdated -replace ',{2,}', ','
					$rhsUpdated = $rhsUpdated.Trim().Trim(',').Trim()
					
					if ([string]::IsNullOrWhiteSpace($rhsUpdated))
					{
						$rhsUpdated = "REDIRMAIL=$redirKey901,REDIRMSG=$redirKey901,TARGETSEND=,TARGETRECV=,DEADLOCKPRIORITY="
					}
					else
					{
						$rhsUpdated = "REDIRMAIL=$redirKey901,REDIRMSG=$redirKey901," + $rhsUpdated
					}
					
					$newProcessorLine = $wsToUse + $newKey + "=" + $rhsUpdated
					
					# Purge all numeric processor lines; insert ONE just after section header
					$out = @()
					for ($i = 0; $i -lt $lines.Length; $i++)
					{
						$out += $lines[$i]
						
						if ($i -eq $secStart)
						{
							$out += $newProcessorLine
							continue
						}
						
						if ($i -gt $secStart -and $i -lt $secEnd)
						{
							if ($lines[$i] -match '^\s*\d{6,7}\s*=') { $out = $out[0 .. ($out.Count - 2)]; continue } # remove what we just added
						}
					}
					
					$lines = $out
				}
			}
			
			# ---- FORCED write (always) ----
			[System.IO.File]::WriteAllText($SmsHttpsIniPath, ($lines -join $nl), $enc)
			Write-Host ("Updated SmsHttps.INI (forced): " + $SmsHttpsIniPath) -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating SmsHttps.INI: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	return $success
}

# ===================================================================================================
#                          FUNCTION: Update_Startup_Ini_For_Machine
# ---------------------------------------------------------------------------------------------------
# Description:
# Applies machine-specific Startup.ini changes that should happen both after a rename and during a
# same-name sync. This keeps the DBSERVER host update in one place and preserves encoding/newlines.
# ===================================================================================================

function Update_Startup_Ini_For_Machine
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[psobject]$MachineContext,
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath
	)
	
	$resolvedStartupIniPath = Resolve_Startup_Ini_Path -StartupIniPath $StartupIniPath
	if ([string]::IsNullOrWhiteSpace($resolvedStartupIniPath))
	{
		return [PSCustomObject]@{
			Success       = $false
			Path          = $null
			DbServerValue = $null
			TerValue      = $null
			Details       = "Startup.ini path could not be resolved."
			ErrorMessage  = "startup.ini not found."
		}
	}
	
	try
	{
		$bytes = [System.IO.File]::ReadAllBytes($resolvedStartupIniPath)
		
		$encStartup = $null
		if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
		{ $encStartup = New-Object System.Text.UTF8Encoding($true) }
		elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
		{ $encStartup = [System.Text.Encoding]::Unicode }
		elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
		{ $encStartup = [System.Text.Encoding]::BigEndianUnicode }
		else
		{ $encStartup = [System.Text.Encoding]::Default }
		
		$textStartup = $encStartup.GetString($bytes)
		$nlStartup = "`r`n"
		if ($textStartup -notmatch "`r`n" -and $textStartup -match "`n") { $nlStartup = "`n" }
		
		$startupLines = $textStartup -split "\r?\n", -1
		
		$existingDbLine = $null
		foreach ($line in $startupLines)
		{
			if ($line -match '^\s*(?i:DBSERVER)\s*=')
			{
				$existingDbLine = $line
				break
			}
		}
		
		$existingDbServer = $null
		if ($existingDbLine -match '^\s*(?i:DBSERVER)\s*=\s*(.+?)\s*$')
		{
			$existingDbServer = $matches[1].Trim()
		}
		
		$instance = $null
		if ($existingDbServer -and ($existingDbServer -match '\\'))
		{
			$parts = $existingDbServer.Split('\')
			if ($parts.Count -ge 2 -and $parts[1]) { $instance = $parts[1] }
		}
		
		$dbServerValue = if ($instance)
		{
			"DBSERVER=$($MachineContext.MachineName)\$instance"
		}
		else
		{
			"DBSERVER=$($MachineContext.MachineName)"
		}
		
		$terValue = $null
		if (-not $MachineContext.IsServerMoniker -and $MachineContext.MachineNumber)
		{
			$terValue = "TER=$($MachineContext.MachineNumber)"
		}
		
		$updatedLines = @()
		$dbServerUpdated = $false
		$terUpdated = $false
		
		foreach ($line in $startupLines)
		{
			$updatedLine = $line
			
			if (-not $MachineContext.IsServerMoniker -and $terValue -and ($updatedLine -match '^\s*(?i:TER)\s*='))
			{
				$updatedLine = $terValue
				$terUpdated = $true
			}
			
			if ($updatedLine -match '^\s*(?i:DBSERVER)\s*=')
			{
				$updatedLine = $dbServerValue
				$dbServerUpdated = $true
			}
			
			$updatedLines += $updatedLine
		}
		
		if (-not $dbServerUpdated)
		{
			$updatedLines += $dbServerValue
			$dbServerUpdated = $true
		}
		
		[System.IO.File]::WriteAllText($resolvedStartupIniPath, ($updatedLines -join $nlStartup), $encStartup)
		
		$detailParts = @()
		$detailParts += "DBSERVER set to '$dbServerValue'."
		if ($MachineContext.IsServerMoniker)
		{
			$detailParts += "TER left unchanged due to SERVER moniker."
		}
		elseif ($terUpdated)
		{
			$detailParts += "TER set to '$terValue'."
		}
		else
		{
			$detailParts += "TER line was not present, so only DBSERVER was enforced here."
		}
		
		return [PSCustomObject]@{
			Success       = $true
			Path          = $resolvedStartupIniPath
			DbServerValue = $dbServerValue
			TerValue      = $terValue
			Details       = ($detailParts -join ' ')
			ErrorMessage  = $null
		}
	}
	catch
	{
		return [PSCustomObject]@{
			Success       = $false
			Path          = $resolvedStartupIniPath
			DbServerValue = $null
			TerValue      = $null
			Details       = "Startup.ini machine sync failed."
			ErrorMessage  = $_.Exception.Message
		}
	}
}

# ===================================================================================================
#                          FUNCTION: Ensure_Database_Connection_Context
# ---------------------------------------------------------------------------------------------------
# Description:
# Lazily refreshes DBSERVER / DBNAME / ConnectionString from Startup.ini so SQL-specific work can
# initialize only when needed instead of front-loading that work during GUI startup.
# ===================================================================================================

function Ensure_Database_Connection_Context
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath
	)
	
	if (-not $script:FunctionResults)
	{
		$script:FunctionResults = @{ }
	}
	
	$connectionStringValue = $null
	if ($script:FunctionResults.ContainsKey('ConnectionString') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['ConnectionString']))
	{
		$connectionStringValue = [string]$script:FunctionResults['ConnectionString']
	}
	
	$dbNameValue = $null
	if ($script:FunctionResults.ContainsKey('DBNAME') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['DBNAME']))
	{
		$dbNameValue = [string]$script:FunctionResults['DBNAME']
	}
	
	$dbServerValue = $null
	if ($script:FunctionResults.ContainsKey('DBSERVER') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['DBSERVER']))
	{
		$dbServerValue = [string]$script:FunctionResults['DBSERVER']
	}
	
	if ([string]::IsNullOrWhiteSpace($connectionStringValue) -or [string]::IsNullOrWhiteSpace($dbNameValue) -or [string]::IsNullOrWhiteSpace($dbServerValue))
	{
		Get_Database_Connection_String -StartupIniPath $StartupIniPath
		
		if ($script:FunctionResults.ContainsKey('ConnectionString') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['ConnectionString']))
		{
			$connectionStringValue = [string]$script:FunctionResults['ConnectionString']
		}
	}
	
	$script:connectionString = $connectionStringValue
	return $connectionStringValue
}

# ===================================================================================================
#                                       FUNCTION: Execute_SQL_Commands
# ---------------------------------------------------------------------------------------------------
# Description:
#   Executes a given SQL command using the provided connection string.
# ===================================================================================================

function Execute_SQL_Commands
{
	param (
		[string]$commandText
	)
	
	$script:LastSqlExecutionError = $null
	$resolvedConnectionString = Ensure_Database_Connection_Context
	if ([string]::IsNullOrWhiteSpace($resolvedConnectionString))
	{
		$script:LastSqlExecutionError = "Database connection string is not available."
		return $false
	}
	
	$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$sqlConnection.ConnectionString = $resolvedConnectionString
	$sqlCommand = $sqlConnection.CreateCommand()
	$sqlCommand.CommandText = $commandText
	
	try
	{
		$sqlConnection.Open()
		$sqlCommand.ExecuteNonQuery() | Out-Null
		return $true # Indicate success
	}
	catch
	{
		$script:LastSqlExecutionError = $_.Exception.Message
		return $false # Indicate failure
	}
	finally
	{
		$sqlConnection.Close()
	}
}

# ===================================================================================================
#                               FUNCTION: Get_Database_Connection_String
# ---------------------------------------------------------------------------------------------------
# Description:
#   Reads DBSERVER / DBNAME from Startup.ini, finds the SQL instance that answers (configured
#   server first, then local aliases and local instances), and stores the configured and runtime
#   server, DBNAME and connection string in $script:FunctionResults.
#
# Parameters:
#   - StartupIniPath [string]: Startup.ini to read. When empty, Resolve_Startup_Ini_Path finds it
#     (script StartupIniPath, then BasePath\Startup.ini).
# ===================================================================================================

function Get_Database_Connection_String
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath
	)
	
	if (-not $script:FunctionResults)
	{
		$script:FunctionResults = @{ }
	}
	
	# The parameter shadows the script-level $StartupIniPath (names are case-insensitive), so an
	# empty parameter must be resolved explicitly - reading $startupIniPath here returns the parameter.
	$effectiveStartupIniPath = $StartupIniPath
	if ([string]::IsNullOrWhiteSpace($effectiveStartupIniPath))
	{
		$effectiveStartupIniPath = Resolve_Startup_Ini_Path
	}
	
	if (-not $effectiveStartupIniPath)
	{
		return
	}
	
	# Read the Startup.ini file
	try
	{
		$content = Get-Content -Path $effectiveStartupIniPath -ErrorAction Stop
		
		# Extract DBSERVER
		$dbServerLine = $content | Where-Object { $_ -match '^DBSERVER=' }
		if ($dbServerLine)
		{
			$dbServer = $dbServerLine -replace '^DBSERVER=', ''
			$dbServer = $dbServer.Trim()
			if (-not $dbServer)
			{
				$dbServer = "localhost"
			}
		}
		else
		{
			$dbServer = "localhost"
		}
		
		# Extract DBNAME
		$dbNameLine = $content | Where-Object { $_ -match '^DBNAME=' }
		if ($dbNameLine)
		{
			$dbName = $dbNameLine -replace '^DBNAME=', ''
			$dbName = $dbName.Trim()
			if (-not $dbName)
			{
				return
			}
		}
		else
		{
			return
		}
	}
	catch
	{
		return
	}
	
	$configuredDbServer = $dbServer
	$resolvedRuntimeDbServer = $configuredDbServer
	$connectionString = "Server=$configuredDbServer;Database=$dbName;Integrated Security=True;Application Name=MiniGhost;"
	$candidateServers = New-Object System.Collections.Generic.List[string]
	$addCandidate = {
		param (
			[string]$Candidate
		)
		
		if ([string]::IsNullOrWhiteSpace($Candidate))
		{
			return
		}
		
		$normalizedCandidate = $Candidate.Trim()
		if ([string]::IsNullOrWhiteSpace($normalizedCandidate))
		{
			return
		}
		
		foreach ($existingCandidate in $candidateServers)
		{
			if ([string]::Equals($existingCandidate, $normalizedCandidate, [System.StringComparison]::OrdinalIgnoreCase))
			{
				return
			}
		}
		
		[void]$candidateServers.Add($normalizedCandidate)
	}
	
	$configuredHost = $configuredDbServer
	$instanceName = $null
	if ($configuredDbServer -match '\\')
	{
		$configuredParts = $configuredDbServer.Split('\', 2)
		$configuredHost = $configuredParts[0].Trim()
		if ($configuredParts.Count -ge 2 -and -not [string]::IsNullOrWhiteSpace($configuredParts[1]))
		{
			$instanceName = $configuredParts[1].Trim()
		}
	}
	
	$localAliases = @('localhost', '.', '(local)')
	if (-not [string]::IsNullOrWhiteSpace($env:COMPUTERNAME))
	{
		$localAliases += $env:COMPUTERNAME
	}
	
	$isConfiguredLocal = $false
	foreach ($alias in $localAliases)
	{
		if ([string]::Equals($configuredHost, $alias, [System.StringComparison]::OrdinalIgnoreCase))
		{
			$isConfiguredLocal = $true
			break
		}
	}
	
	if ($isConfiguredLocal)
	{
		& $addCandidate $configuredDbServer
	}
	
	if ($instanceName)
	{
		foreach ($alias in $localAliases)
		{
			& $addCandidate ("{0}\{1}" -f $alias, $instanceName)
		}
	}
	else
	{
		foreach ($alias in $localAliases)
		{
			& $addCandidate $alias
		}
		
		foreach ($registryPath in @(
				'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL',
				'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL'
			))
		{
			if (-not (Test-Path $registryPath)) { continue }
			
			try
			{
				$itemProps = Get-ItemProperty -Path $registryPath -ErrorAction Stop
				foreach ($prop in $itemProps.PSObject.Properties)
				{
					if ($prop.Name -in @('PSPath', 'PSParentPath', 'PSChildName', 'PSDrive', 'PSProvider'))
					{
						continue
					}
					
					if ([string]::IsNullOrWhiteSpace($prop.Name))
					{
						continue
					}
					
					$instanceCandidate = $prop.Name.Trim()
					foreach ($alias in $localAliases)
					{
						& $addCandidate ("{0}\{1}" -f $alias, $instanceCandidate)
					}
				}
			}
			catch { }
		}
	}
	
	& $addCandidate $configuredDbServer
	
	foreach ($candidateServer in $candidateServers)
	{
		$testConnection = $null
		try
		{
			$testConnectionString = "Server=$candidateServer;Database=$dbName;Integrated Security=True;Connect Timeout=3;Application Name=MiniGhost;"
			$testConnection = New-Object System.Data.SqlClient.SqlConnection($testConnectionString)
			$testConnection.Open()
			
			$resolvedRuntimeDbServer = $candidateServer
			$connectionString = "Server=$candidateServer;Database=$dbName;Integrated Security=True;Application Name=MiniGhost;"
			break
		}
		catch { }
		finally
		{
			if ($testConnection)
			{
				$testConnection.Close()
				$testConnection.Dispose()
			}
		}
	}
	
	# Store configured and runtime database context in the FunctionResults hashtable
	$script:FunctionResults['ConfiguredDBSERVER'] = $configuredDbServer
	$script:FunctionResults['DBSERVER'] = $configuredDbServer
	$script:FunctionResults['RuntimeDBSERVER'] = $resolvedRuntimeDbServer
	$script:FunctionResults['DBNAME'] = $dbName
	$script:FunctionResults['ConnectionString'] = $connectionString
}

# ===================================================================================================
#                                       FUNCTION: RPT_Blank
# ---------------------------------------------------------------------------------------------------
# Description:
#   Truncates the specified list of tables in the SQL database.
# ===================================================================================================

function RPT_Blank
{
	param (
		[string[]]$tables
	)
	
	# Initialize an array to store failed truncate commands
	$failedTruncateTables = @()
	
	# Prepare and execute Truncate Commands
	foreach ($table in $tables)
	{
		$command = "TRUNCATE TABLE $table;"
		if (-not (Execute_SQL_Commands -commandText $command))
		{
			$failedTruncateTables += $table # Add failed table to the array
		}
	}
	
	# Return the list of failed tables
	return $failedTruncateTables
}

# ===================================================================================================
#                               FUNCTION: Update_SQL_Database
# ---------------------------------------------------------------------------------------------------
# Description:
#   Runs the newer protected SQL machine/store sync routine used by rename flows so the standalone
#   "Update SQL Database" button follows the same safer TER/STO/LNK/RUN/STD update path.
#
# Parameters:
#   - CurrentMachineName : Current hostname to use if NewMachineName is empty
#   - NewMachineName     : Optional override hostname (ex: after rename)
#   - OperationStatus    : Shared status hashtable (ex: $OperationStatus["SQLDatabaseUpdate"])
# ===================================================================================================

function Update_SQL_Database
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)]
		[string]$CurrentMachineName,
		[Parameter()]
		[string]$NewMachineName,
		[Parameter(Mandatory)]
		[hashtable]$OperationStatus
	)
	
	try
	{
		$storeNumberFromINI = Get_Store_Number_From_INI
		if ($null -eq $storeNumberFromINI)
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Store number not found in startup.ini.",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Store number not found in startup.ini."
			$OperationStatus["SQLDatabaseUpdate"].Details = ""
			return
		}
		
		$storeNumber = ($storeNumberFromINI.ToString()).Trim()
		
		$machineName = if (-not [string]::IsNullOrWhiteSpace($NewMachineName)) { $NewMachineName }
		else { $CurrentMachineName }
		
		$machineContext = Get_Machine_Name_Context -MachineName $machineName -CurrentMachineName $CurrentMachineName
		if (-not $machineContext.Success)
		{
			[System.Windows.Forms.MessageBox]::Show(
				$machineContext.ErrorMessage,
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Invalid machine name."
			$OperationStatus["SQLDatabaseUpdate"].Details = $machineContext.ErrorMessage
			return
		}
		
		$machineName = $machineContext.MachineName
		$machineNumber = $machineContext.MachineNumber
		
		$sqlUpdateResult = Update_SQL_Tables_For_Machine_Name_Change `
															 -storeNumber $storeNumber `
															 -machineName $machineName `
															 -machineNumber $machineNumber `
															 -OldMachineName $CurrentMachineName
		
		if ($sqlUpdateResult -and $sqlUpdateResult.Success)
		{
			$details = "Protected SQL sync completed."
			if ($sqlUpdateResult.ContainsKey('RunTabUpdated'))
			{
				$details += " RUN updated: $($sqlUpdateResult.RunTabUpdated)."
			}
			if ($sqlUpdateResult.ContainsKey('LocalTerminalDB'))
			{
				$details += " Local terminal DB: $($sqlUpdateResult.LocalTerminalDB)."
			}
			
			$OperationStatus["SQLDatabaseUpdate"].Status = "Successful"
			$OperationStatus["SQLDatabaseUpdate"].Message = "SQL database updated successfully."
			$OperationStatus["SQLDatabaseUpdate"].Details = $details
			[System.Windows.Forms.MessageBox]::Show(
				"SQL database updated successfully.",
				"Success",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Information
			)
		}
		else
		{
			$failedSqlCommands = @()
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('FailedCommands') -and $sqlUpdateResult.FailedCommands)
			{
				$failedSqlCommands = @($sqlUpdateResult.FailedCommands)
			}
			
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Failed to update SQL database."
			$OperationStatus["SQLDatabaseUpdate"].Details = if ($failedSqlCommands.Count -gt 0) { "Failed SQL commands are listed below." }
			else { "SQL update result indicated failure." }
			[System.Windows.Forms.MessageBox]::Show(
				"Failed to update SQL database.",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			
			if ($failedSqlCommands.Count -gt 0)
			{
				Show_Text_Report_Form -Title "Failed SQL Commands" -Text ($failedSqlCommands -join "`r`n`r`n") -Width 640 -Height 420
			}
		}
	}
	catch
	{
		$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
		$OperationStatus["SQLDatabaseUpdate"].Message = "Unhandled error during SQL update."
		$OperationStatus["SQLDatabaseUpdate"].Details = $_.Exception.Message
		
		[System.Windows.Forms.MessageBox]::Show(
			"Unhandled error:`r`n$($_.Exception.Message)",
			"Error",
			[System.Windows.Forms.MessageBoxButtons]::OK,
			[System.Windows.Forms.MessageBoxIcon]::Error
		)
	}
}

# ===================================================================================================
#                       FUNCTION: Update_SQL_Tables_For_Machine_Name_Change
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates TER_TAB, STO_TAB, LNK_TAB, RUN_TAB, and STD_TAB after machine name change / store change.
#
# Auto DB context detection (based on Startup.ini DBNAME already gathered via Get_Database_Connection_String):
#   - LANESQL             => LocalTerminalDB (lane/POS/SCO cleanup in same store is allowed)
#   - STORESQL or HOSTSQL => ServerStoreDB  (do NOT lane-cleanup other terminals)
#
# STD_TAB enhancement:
#   - Also updates F1531 (CompanyName) store suffix when present (e.g., "Company Name 231" or "Company #0231")
#
# FIXES:
#   - Non-server runs preserve/copy TER_TAB 901 into the NEW store before deleting old-store rows.
#   - STO_TAB now always ensures the 901 row exists.
#   - LNK_TAB now preserves/copies 901 into the NEW store before cleanup, so the server side is not left empty.
# ===================================================================================================

function Update_SQL_Tables_For_Machine_Name_Change
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidatePattern('^(?!0+$)\d{3,4}$')]
		[string]$storeNumber,
		[Parameter(Mandatory = $true)]
		[string]$machineName,
		[Parameter(Mandatory = $true)]
		[string]$machineNumber,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineName,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineNumber,
		[Parameter(Mandatory = $false)]
		[ValidateSet('Auto', 'LocalTerminalDB', 'ServerStoreDB')]
		[string]$DatabaseContext = 'Auto'
	)
	
	# ---------------------------
	# Constants / normalization
	# ---------------------------
	$hostStoreNumber = "999"
	$serverTerminal = "901"
	$protectedMinTerminal = 940
	
	$storeNumber = (($storeNumber | ForEach-Object { if ($null -eq $_) { "" }
				else { $_.ToString() } }).Trim())
	$machineName = (($machineName | ForEach-Object { if ($null -eq $_) { "" }
				else { $_.ToString() } }).Trim())
	$machineNumber = (($machineNumber | ForEach-Object { if ($null -eq $_) { "" }
				else { $_.ToString() } }).Trim())
	
	# Detect SERVER moniker (example: 0242SERVER001)
	$isServerMoniker = $false
	if (-not [string]::IsNullOrWhiteSpace($machineName) -and ($machineName -match '(?i)SERVER'))
	{
		$isServerMoniker = $true
	}
	elseif ($script:NodeRole -and ($script:NodeRole -match '^(?i)(StoreServer|HostServer)$'))
	{
		$isServerMoniker = $true
	}
	
	$null = Ensure_Database_Connection_Context
	
	# ---------------------------
	# Pull DBNAME (already gathered by Get_Database_Connection_String) for better Auto context
	# ---------------------------
	$dbNameDetected = $null
	try
	{
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('DBNAME') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['DBNAME']))
		{
			$dbNameDetected = ([string]$script:FunctionResults['DBNAME']).Trim()
		}
	}
	catch { }
	
	# Fallback: parse from ConnectionString if present
	if ([string]::IsNullOrWhiteSpace($dbNameDetected))
	{
		try
		{
			if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('ConnectionString'))
			{
				$cs = [string]$script:FunctionResults['ConnectionString']
				if ($cs -match '(?i)\bDatabase\s*=\s*([^;]+)') { $dbNameDetected = $Matches[1].Trim() }
				elseif ($cs -match '(?i)\bInitial\s+Catalog\s*=\s*([^;]+)') { $dbNameDetected = $Matches[1].Trim() }
			}
		}
		catch { }
	}
	
	# ---------------------------
	# Normalize NEW/OLD server host string (for matching TER_TAB.F1125 and replacing STO_TAB.F1018)
	# ---------------------------
	$newServerNameNorm = $machineName
	if ($null -eq $newServerNameNorm) { $newServerNameNorm = "" }
	$newServerNameNorm = $newServerNameNorm.Trim()
	$newServerNameNorm = $newServerNameNorm -replace '^[\\\/]+', ''
	if ($newServerNameNorm -match '[\\\/]') { $newServerNameNorm = ($newServerNameNorm -split '[\\\/]')[0] }
	if ($newServerNameNorm -match '\.') { $newServerNameNorm = ($newServerNameNorm -split '\.')[0] }
	$newServerNameNorm = $newServerNameNorm.Trim().ToUpper()
	
	$oldServerNameNorm = $OldMachineName
	if ([string]::IsNullOrWhiteSpace($oldServerNameNorm)) { $oldServerNameNorm = $env:COMPUTERNAME }
	if ($null -eq $oldServerNameNorm) { $oldServerNameNorm = "" }
	$oldServerNameNorm = $oldServerNameNorm.Trim()
	$oldServerNameNorm = $oldServerNameNorm -replace '^[\\\/]+', ''
	if ($oldServerNameNorm -match '[\\\/]') { $oldServerNameNorm = ($oldServerNameNorm -split '[\\\/]')[0] }
	if ($oldServerNameNorm -match '\.') { $oldServerNameNorm = ($oldServerNameNorm -split '\.')[0] }
	$oldServerNameNorm = $oldServerNameNorm.Trim().ToUpper()
	
	# ---------------------------
	# Normalize NEW terminal to 3 digits
	# ---------------------------
	$newTerminal = $null
	if ($machineNumber -match '^\d{1,3}$')
	{
		$newTerminal = ([int]$machineNumber).ToString("D3")
		if ($newTerminal -eq "000") { $newTerminal = $null }
	}
	else
	{
		$tmp = $machineNumber
		if ($null -eq $tmp) { $tmp = "" }
		$tmp = $tmp.Trim()
		$tmp = $tmp -replace '^[\\\/]+', ''
		if ($tmp -match '[\\\/]') { $tmp = ($tmp -split '[\\\/]')[0] }
		if ($tmp -match '\.') { $tmp = ($tmp -split '\.')[0] }
		$tmp = $tmp.Trim().ToUpper()
		
		if ($tmp -match '(\d{1,3})$')
		{
			$newTerminal = ([int]$Matches[1]).ToString("D3")
			if ($newTerminal -eq "000") { $newTerminal = $null }
		}
	}
	
	# Determine OLD terminal for safe RUN_TAB swap (used only when NOT server)
	if ([string]::IsNullOrWhiteSpace($OldMachineName)) { $OldMachineName = $env:COMPUTERNAME }
	
	$oldTerminal = $null
	
	if (-not [string]::IsNullOrWhiteSpace($OldMachineNumber))
	{
		if ($OldMachineNumber -match '^\d{1,3}$')
		{
			$oldTerminal = ([int]$OldMachineNumber).ToString("D3")
			if ($oldTerminal -eq "000") { $oldTerminal = $null }
		}
		else
		{
			$tmpOld = $OldMachineNumber
			if ($null -eq $tmpOld) { $tmpOld = "" }
			$tmpOld = $tmpOld.Trim()
			$tmpOld = $tmpOld -replace '^[\\\/]+', ''
			if ($tmpOld -match '[\\\/]') { $tmpOld = ($tmpOld -split '[\\\/]')[0] }
			if ($tmpOld -match '\.') { $tmpOld = ($tmpOld -split '\.')[0] }
			$tmpOld = $tmpOld.Trim().ToUpper()
			
			if ($tmpOld -match '(\d{1,3})$')
			{
				$oldTerminal = ([int]$Matches[1]).ToString("D3")
				if ($oldTerminal -eq "000") { $oldTerminal = $null }
			}
		}
	}
	
	if (-not $oldTerminal)
	{
		$tmpOld2 = $OldMachineName
		if ($null -eq $tmpOld2) { $tmpOld2 = "" }
		$tmpOld2 = $tmpOld2.Trim()
		$tmpOld2 = $tmpOld2 -replace '^[\\\/]+', ''
		if ($tmpOld2 -match '[\\\/]') { $tmpOld2 = ($tmpOld2 -split '[\\\/]')[0] }
		if ($tmpOld2 -match '\.') { $tmpOld2 = ($tmpOld2 -split '\.')[0] }
		$tmpOld2 = $tmpOld2.Trim().ToUpper()
		
		if ($tmpOld2 -match '(\d{1,3})$')
		{
			$oldTerminal = ([int]$Matches[1]).ToString("D3")
			if ($oldTerminal -eq "000") { $oldTerminal = $null }
		}
	}
	
	# ---------------------------
	# Decide DB context
	# ---------------------------
	$isLocalTerminalDB = $false
	
	if ($DatabaseContext -eq 'LocalTerminalDB')
	{
		$isLocalTerminalDB = $true
	}
	elseif ($DatabaseContext -eq 'ServerStoreDB')
	{
		$isLocalTerminalDB = $false
	}
	else
	{
		# Auto: prefer DBNAME
		if (-not [string]::IsNullOrWhiteSpace($dbNameDetected) -and ($dbNameDetected -match '^(?i)LANESQL$'))
		{
			$isLocalTerminalDB = $true
		}
		elseif (-not [string]::IsNullOrWhiteSpace($dbNameDetected) -and ($dbNameDetected -match '^(?i)(STORESQL|HOSTSQL)$'))
		{
			$isLocalTerminalDB = $false
		}
		else
		{
			# Fallback heuristic: if targeting this box, assume local terminal DB
			$execHostNorm = $env:COMPUTERNAME
			if ($null -eq $execHostNorm) { $execHostNorm = "" }
			$execHostNorm = ($execHostNorm + "").Trim()
			$execHostNorm = $execHostNorm -replace '^[\\\/]+', ''
			if ($execHostNorm -match '[\\\/]') { $execHostNorm = ($execHostNorm -split '[\\\/]')[0] }
			if ($execHostNorm -match '\.') { $execHostNorm = ($execHostNorm -split '\.')[0] }
			$execHostNorm = $execHostNorm.Trim().ToUpper()
			
			$targetHostNorm = $machineName
			if ($null -eq $targetHostNorm) { $targetHostNorm = "" }
			$targetHostNorm = ($targetHostNorm + "").Trim()
			$targetHostNorm = $targetHostNorm -replace '^[\\\/]+', ''
			if ($targetHostNorm -match '[\\\/]') { $targetHostNorm = ($targetHostNorm -split '[\\\/]')[0] }
			if ($targetHostNorm -match '\.') { $targetHostNorm = ($targetHostNorm -split '\.')[0] }
			$targetHostNorm = $targetHostNorm.Trim().ToUpper()
			
			$isLocalTerminalDB = ($targetHostNorm -eq $execHostNorm)
		}
	}
	
	# SAFETY: Non-server mode must have a terminal, otherwise we can't safely upsert/cleanup.
	if (-not $isServerMoniker -and -not $newTerminal)
	{
		return @{
			Success		    = $false
			FailedCommands  = @("Refusing to run: could not derive a valid NEW terminal from -machineNumber/-machineName for non-server mode.")
			IsServerMoniker = $false
			RunTabUpdated   = $false
			OldTerminalUsed = $oldTerminal
			NewTerminalUsed = $newTerminal
			DatabaseContext = $DatabaseContext
			LocalTerminalDB = $isLocalTerminalDB
			DBNAME		    = $dbNameDetected
		}
	}
	
	# ---------------------------
	# Table names
	# ---------------------------
	$terTableName = "TER_TAB"
	$stoTableName = "STO_TAB"
	$lnkTableName = "LNK_TAB"
	$runTableName = "RUN_TAB"
	$stdTableName = "STD_TAB"
	
	# ===============================================================================================
	# TER_TAB
	# ===============================================================================================
	
	$createViewCommandTer = @"
IF OBJECT_ID('Ter_Load', 'V') IS NOT NULL DROP VIEW Ter_Load;
EXEC(N'
CREATE VIEW Ter_Load AS
SELECT F1056, F1057, F1058, F1125, F1169
FROM $terTableName;
');
"@
	
	$moveProtectedTer_ToNewStore = @"
DECLARE @NewStore  varchar(10) = '$storeNumber';
DECLARE @HostStore varchar(10) = '$hostStoreNumber';
DECLARE @ProtectedMin int = $protectedMinTerminal;

INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169)
SELECT @NewStore, t.F1057, t.F1058, t.F1125, t.F1169
FROM $terTableName t
WHERE t.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) >= @ProtectedMin
  AND NOT EXISTS (
      SELECT 1
      FROM $terTableName x
      WHERE x.F1056 = @NewStore
        AND x.F1057 = t.F1057
  );

DELETE t
FROM $terTableName t
WHERE t.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) >= @ProtectedMin;
"@
	
	$moveServerNodes902_920_ToNewStore_AndUpdate = @"
DECLARE @NewStore  varchar(10) = '$storeNumber';
DECLARE @HostStore varchar(10) = '$hostStoreNumber';
DECLARE @NewServer varchar(200) = '$newServerNameNorm';
DECLARE @OldServer varchar(200) = '$oldServerNameNorm';

;WITH Candidates AS
(
    SELECT t.*
    FROM $terTableName t
    WHERE t.F1056 NOT IN (@NewStore, @HostStore)
      AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920
      AND t.F1125 IS NOT NULL
      AND
      (
          UPPER(t.F1125) LIKE '%\\' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%\\' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%@' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%@' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%' + @NewServer + '%'
      )
)
INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169)
SELECT @NewStore, c.F1057, c.F1058, c.F1125, c.F1169
FROM Candidates c
WHERE NOT EXISTS (
    SELECT 1
    FROM $terTableName x
    WHERE x.F1056 = @NewStore
      AND x.F1057 = c.F1057
);

UPDATE t
SET
    t.F1125 = CASE
                WHEN t.F1125 IS NULL THEN t.F1125
                WHEN LEFT(LTRIM(RTRIM(t.F1125)),1)='@' THEN '@' + @NewServer
                ELSE REPLACE(UPPER(t.F1125), @OldServer, @NewServer)
              END,
    t.F1169 = CASE
                WHEN t.F1169 IS NULL THEN t.F1169
                ELSE '\\' + @NewServer + '\storeman\office\XF' + @NewStore + '901\'
              END
FROM $terTableName t
WHERE t.F1056 = @NewStore
  AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920;

DELETE t
FROM $terTableName t
WHERE t.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920
  AND t.F1125 IS NOT NULL
  AND
  (
      UPPER(t.F1125) LIKE '%\\' + @OldServer + '%'
      OR UPPER(t.F1125) LIKE '%\\' + @NewServer + '%'
      OR UPPER(t.F1125) LIKE '%@' + @OldServer + '%'
      OR UPPER(t.F1125) LIKE '%@' + @NewServer + '%'
      OR UPPER(t.F1125) LIKE '%' + @OldServer + '%'
      OR UPPER(t.F1125) LIKE '%' + @NewServer + '%'
  );
"@
	
	$deleteOtherStoresTer = @"
DELETE FROM $terTableName
WHERE F1056 NOT IN ('$storeNumber', '$hostStoreNumber');
"@
	
	$deleteNonServerRowsForNewStore = @"
DECLARE @NewStore varchar(10) = '$storeNumber';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';
DECLARE @ProtectedMin int = $protectedMinTerminal;
DECLARE @NewServer varchar(200) = '$newServerNameNorm';
DECLARE @OldServer varchar(200) = '$oldServerNameNorm';

DELETE t
FROM $terTableName t
WHERE t.F1056 = @NewStore
  AND t.F1057 <> @ServerTerm
  AND
  (
      (
          TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) IS NULL
          OR TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) < @ProtectedMin
      )
      AND
      (
          NOT (
              TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920
              AND t.F1125 IS NOT NULL
              AND
              (
                  UPPER(t.F1125) LIKE '%\\' + @NewServer + '%'
                  OR UPPER(t.F1125) LIKE '%\\' + @OldServer + '%'
                  OR UPPER(t.F1125) LIKE '%@' + @NewServer + '%'
                  OR UPPER(t.F1125) LIKE '%@' + @OldServer + '%'
                  OR UPPER(t.F1125) LIKE '%' + @NewServer + '%'
                  OR UPPER(t.F1125) LIKE '%' + @OldServer + '%'
              )
          )
      )
  );
"@
	
	$upsertLaneTer = @"
IF EXISTS (SELECT 1 FROM $terTableName WHERE F1056='$storeNumber' AND F1057='$newTerminal')
BEGIN
    UPDATE $terTableName
    SET F1058='Terminal $newTerminal',
        F1125='\\$machineName\storeman\office\XF$storeNumber$newTerminal\',
        F1169='\\$machineName\storeman\office\XF$storeNumber$serverTerminal\'
    WHERE F1056='$storeNumber' AND F1057='$newTerminal';
END
ELSE
BEGIN
    INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169) VALUES
    ('$storeNumber', '$newTerminal',
     'Terminal $newTerminal',
     '\\$machineName\storeman\office\XF$storeNumber$newTerminal\',
     '\\$machineName\storeman\office\XF$storeNumber$serverTerminal\');
END
"@
	
	$preserve901ForNonServerStoreChange = @"
DECLARE @NewStore   varchar(10) = '$storeNumber';
DECLARE @HostStore  varchar(10) = '$hostStoreNumber';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';

IF NOT EXISTS (
    SELECT 1
    FROM $terTableName
    WHERE F1056 = @NewStore
      AND F1057 = @ServerTerm
)
BEGIN
    INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169)
    SELECT TOP 1
           @NewStore,
           @ServerTerm,
           CASE
               WHEN NULLIF(LTRIM(RTRIM(t.F1058)), '') IS NULL THEN 'Server'
               ELSE t.F1058
           END,
           t.F1125,
           t.F1169
    FROM $terTableName t
    WHERE t.F1057 = @ServerTerm
      AND t.F1056 <> @HostStore
    ORDER BY
        CASE WHEN t.F1125 IS NOT NULL AND LTRIM(RTRIM(t.F1125)) <> '' THEN 0 ELSE 1 END,
        t.F1056;
END

UPDATE $terTableName
SET F1058 = CASE
                WHEN NULLIF(LTRIM(RTRIM(F1058)), '') IS NULL THEN 'Server'
                ELSE F1058
            END
WHERE F1056 = @NewStore
  AND F1057 = @ServerTerm;
"@
	
	$cleanupTer_ForLocalTerminalDB = @"
DECLARE @NewStore varchar(10) = '$storeNumber';
DECLARE @KeepTerm varchar(10) = '$newTerminal';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';
DECLARE @ProtectedMin int = $protectedMinTerminal;

DELETE t
FROM $terTableName t
WHERE t.F1056 = @NewStore
  AND t.F1057 NOT IN (@KeepTerm, @ServerTerm)
  AND (TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) IS NULL OR TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) < @ProtectedMin);
"@
	
	$upsertServer901Ter_ForNewStoreOnly = @"
IF EXISTS (SELECT 1 FROM $terTableName WHERE F1056='$storeNumber' AND F1057='$serverTerminal')
BEGIN
    UPDATE $terTableName
    SET F1058='Server',
        F1125='@$machineName',
        F1169=NULL
    WHERE F1056='$storeNumber' AND F1057='$serverTerminal';
END
ELSE
BEGIN
    INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169) VALUES
    ('$storeNumber', '$serverTerminal',
     'Server',
     '@$machineName',
     NULL);
END
"@
	
	$dropViewCommandTer = @"
IF OBJECT_ID('Ter_Load', 'V') IS NOT NULL DROP VIEW Ter_Load;
"@
	
	# ===============================================================================================
	# STO_TAB
	# ===============================================================================================
	
	$createViewCommandSto = @"
IF OBJECT_ID('Sto_Load', 'V') IS NOT NULL DROP VIEW Sto_Load;
EXEC(N'
CREATE VIEW Sto_Load AS
SELECT F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691
FROM $stoTableName;
');
"@
	
	$upsertStoTerminal = @"
MERGE INTO $stoTableName AS target
USING (VALUES
    ('$newTerminal', 'Terminal $newTerminal', 1, 1, 1, NULL, NULL, NULL, NULL)
) AS source (F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691)
ON target.F1000 = source.F1000
WHEN MATCHED THEN
    UPDATE SET
        F1018 = source.F1018,
        F1180 = source.F1180,
        F1181 = source.F1181,
        F1182 = source.F1182
WHEN NOT MATCHED THEN
    INSERT (F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691)
    VALUES (source.F1000, source.F1018, source.F1180, source.F1181, source.F1182, source.F1937, source.F1965, source.F1966, source.F2691);
"@
	
	$upsertStoServer901 = @"
MERGE INTO $stoTableName AS target
USING (VALUES
    ('$serverTerminal', 'Server', 1, 1, 1, NULL, NULL, NULL, NULL)
) AS source (F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691)
ON target.F1000 = source.F1000
WHEN MATCHED THEN
    UPDATE SET
        F1018 = CASE
                    WHEN NULLIF(LTRIM(RTRIM(target.F1018)), '') IS NULL THEN source.F1018
                    ELSE target.F1018
                END,
        F1180 = ISNULL(target.F1180, source.F1180),
        F1181 = ISNULL(target.F1181, source.F1181),
        F1182 = ISNULL(target.F1182, source.F1182)
WHEN NOT MATCHED THEN
    INSERT (F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691)
    VALUES (source.F1000, source.F1018, source.F1180, source.F1181, source.F1182, source.F1937, source.F1965, source.F1966, source.F2691);
"@
	
	$cleanupSto_ForLocalTerminalDB = @"
DECLARE @KeepTerm varchar(10) = '$newTerminal';

DELETE FROM $stoTableName
WHERE F1000 <> @KeepTerm
  AND F1000 NOT IN ('DSM','PAL','RAL','XAL','901','999');
"@
	
	$cleanupSto_ForServer = @"
;WITH StoNums AS
(
    SELECT
        F1000,
        F1018,
        TRY_CONVERT(int, LTRIM(RTRIM(F1000))) AS NumVal
    FROM $stoTableName
    WHERE ISNUMERIC(F1000) = 1
)
DELETE s
FROM $stoTableName s
JOIN StoNums n ON n.F1000 = s.F1000
WHERE
    n.F1000 NOT IN ('901','999')
    AND
    (
        (n.NumVal BETWEEN 1 AND 99)
        OR
        (
            n.NumVal BETWEEN 902 AND 920
            AND (n.F1018 IS NULL OR UPPER(n.F1018) NOT LIKE '%SERVER%')
        )
    );

DECLARE @NewServer varchar(200) = '$newServerNameNorm';
DECLARE @OldServer varchar(200) = '$oldServerNameNorm';

UPDATE $stoTableName
SET F1018 = REPLACE(UPPER(F1018), @OldServer, @NewServer)
WHERE ISNUMERIC(F1000) = 1
  AND TRY_CONVERT(int, LTRIM(RTRIM(F1000))) BETWEEN 902 AND 920
  AND F1018 IS NOT NULL
  AND UPPER(F1018) LIKE '%' + @OldServer + '%';
"@
	
	$dropViewCommandSto = @"
IF OBJECT_ID('Sto_Load', 'V') IS NOT NULL DROP VIEW Sto_Load;
"@
	
	# ===============================================================================================
	# LNK_TAB
	# ===============================================================================================
	
	$createViewCommandLnk = @"
IF OBJECT_ID('Lnk_Load', 'V') IS NOT NULL DROP VIEW Lnk_Load;
EXEC(N'
CREATE VIEW Lnk_Load AS
SELECT F1000, F1056, F1057
FROM $lnkTableName;
');
"@
	
	$preserve901LnkForNewStore = @"
DECLARE @NewStore   varchar(10) = '$storeNumber';
DECLARE @HostStore  varchar(10) = '$hostStoreNumber';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';

INSERT INTO $lnkTableName (F1000, F1056, F1057)
SELECT DISTINCT
       l.F1000,
       @NewStore,
       @ServerTerm
FROM $lnkTableName l
WHERE l.F1057 = @ServerTerm
  AND l.F1056 <> @HostStore
  AND l.F1000 IS NOT NULL
  AND NOT EXISTS (
      SELECT 1
      FROM $lnkTableName x
      WHERE x.F1000 = l.F1000
        AND x.F1056 = @NewStore
        AND x.F1057 = @ServerTerm
  );

IF NOT EXISTS (
    SELECT 1
    FROM $lnkTableName
    WHERE F1056 = @NewStore
      AND F1057 = @ServerTerm
)
BEGIN
    INSERT INTO $lnkTableName (F1000, F1056, F1057)
    VALUES (@ServerTerm, @NewStore, @ServerTerm);
END
"@
	
	$rebuildLnkForTerminal = @"
DECLARE @NewStore  varchar(10) = '$storeNumber';
DECLARE @HostStore varchar(10) = '$hostStoreNumber';
DECLARE @Term      varchar(10) = '$newTerminal';
DECLARE @ProtectedMin int = $protectedMinTerminal;

DECLARE @NewServer varchar(200) = '$newServerNameNorm';
DECLARE @OldServer varchar(200) = '$oldServerNameNorm';

DECLARE @SourceStore varchar(10);

SELECT TOP 1 @SourceStore = F1056
FROM $lnkTableName
WHERE F1057 = @Term
  AND F1056 <> @HostStore
  AND F1056 <> @NewStore
ORDER BY F1056;

IF @SourceStore IS NULL
BEGIN
    SELECT TOP 1 @SourceStore = F1056
    FROM $lnkTableName
    WHERE F1057 = @Term
      AND F1056 = @NewStore;
END

DECLARE @Codes TABLE (F1000 varchar(50) NOT NULL PRIMARY KEY);

IF @SourceStore IS NOT NULL
BEGIN
    INSERT INTO @Codes(F1000)
    SELECT DISTINCT F1000
    FROM $lnkTableName
    WHERE F1056 = @SourceStore
      AND F1057 = @Term
      AND F1000 IS NOT NULL;
END

IF EXISTS (SELECT 1 FROM $terTableName WHERE F1056=@NewStore AND TRY_CONVERT(int, LTRIM(RTRIM(F1057))) BETWEEN 902 AND 920)
BEGIN
    DECLARE @ServerTerms TABLE (Term varchar(10) NOT NULL PRIMARY KEY);

    INSERT INTO @ServerTerms(Term)
    SELECT DISTINCT t.F1057
    FROM $terTableName t
    WHERE t.F1056 = @NewStore
      AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920
      AND t.F1125 IS NOT NULL
      AND
      (
          UPPER(t.F1125) LIKE '%\\' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%\\' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%@' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%@' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%' + @OldServer + '%'
      );

    INSERT INTO $lnkTableName (F1000, F1056, F1057)
    SELECT l.F1000, @NewStore, l.F1057
    FROM $lnkTableName l
    JOIN @ServerTerms st ON st.Term = l.F1057
    WHERE l.F1056 NOT IN (@NewStore, @HostStore)
      AND NOT EXISTS (
          SELECT 1
          FROM $lnkTableName x
          WHERE x.F1000 = l.F1000
            AND x.F1056 = @NewStore
            AND x.F1057 = l.F1057
      );

    DELETE l
    FROM $lnkTableName l
    JOIN @ServerTerms st ON st.Term = l.F1057
    WHERE l.F1056 NOT IN (@NewStore, @HostStore);
END

INSERT INTO $lnkTableName (F1000, F1056, F1057)
SELECT l.F1000, @NewStore, l.F1057
FROM $lnkTableName l
WHERE l.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(l.F1057))) >= @ProtectedMin
  AND NOT EXISTS (
      SELECT 1
      FROM $lnkTableName x
      WHERE x.F1000 = l.F1000
        AND x.F1056 = @NewStore
        AND x.F1057 = l.F1057
  );

DELETE l
FROM $lnkTableName l
WHERE l.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(l.F1057))) >= @ProtectedMin;

DELETE FROM $lnkTableName
WHERE F1056 NOT IN (@NewStore, @HostStore);

IF NOT EXISTS (SELECT 1 FROM @Codes)
BEGIN
    INSERT INTO @Codes(F1000) VALUES (@Term);
    IF NOT EXISTS (SELECT 1 FROM @Codes WHERE F1000='DSM') INSERT INTO @Codes(F1000) VALUES ('DSM');
    IF NOT EXISTS (SELECT 1 FROM @Codes WHERE F1000='PAL') INSERT INTO @Codes(F1000) VALUES ('PAL');
    IF NOT EXISTS (SELECT 1 FROM @Codes WHERE F1000='RAL') INSERT INTO @Codes(F1000) VALUES ('RAL');
    IF NOT EXISTS (SELECT 1 FROM @Codes WHERE F1000='XAL') INSERT INTO @Codes(F1000) VALUES ('XAL');
END

INSERT INTO $lnkTableName (F1000, F1056, F1057)
SELECT c.F1000, @NewStore, @Term
FROM @Codes c
WHERE NOT EXISTS (
    SELECT 1
    FROM $lnkTableName x
    WHERE x.F1000 = c.F1000
      AND x.F1056 = @NewStore
      AND x.F1057 = @Term
);
"@
	
	$cleanupLnk_ForLocalTerminalDB = @"
DECLARE @NewStore varchar(10) = '$storeNumber';
DECLARE @KeepTerm varchar(10) = '$newTerminal';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';
DECLARE @ProtectedMin int = $protectedMinTerminal;

DELETE l
FROM $lnkTableName l
WHERE l.F1056 = @NewStore
  AND l.F1057 NOT IN (@KeepTerm, @ServerTerm)
  AND (TRY_CONVERT(int, LTRIM(RTRIM(l.F1057))) IS NULL OR TRY_CONVERT(int, LTRIM(RTRIM(l.F1057))) < @ProtectedMin);
"@
	
	$dropViewCommandLnk = @"
IF OBJECT_ID('Lnk_Load', 'V') IS NOT NULL DROP VIEW Lnk_Load;
"@
	
	# ===============================================================================================
	# RUN_TAB (ONLY when NOT server)
	# ===============================================================================================
	
	$createViewCommandRun = @"
IF OBJECT_ID('Run_Load', 'V') IS NOT NULL DROP VIEW Run_Load;
EXEC(N'
CREATE VIEW Run_Load AS
SELECT F1102, F1000, F1104
FROM $runTableName;
');
"@
	
	$updateRunTab_SafeSwap_OldToNew = @"
DECLARE @OldTerm varchar(10) = '$oldTerminal';
DECLARE @NewTerm varchar(10) = '$newTerminal';
DECLARE @Server  varchar(10) = '$serverTerminal';

IF @OldTerm IS NOT NULL AND @OldTerm <> '' AND @OldTerm <> @NewTerm
BEGIN
    UPDATE $runTableName
    SET F1000 = @NewTerm
    WHERE F1000 = @OldTerm;

    UPDATE $runTableName
    SET F1104 = @NewTerm
    WHERE F1104 = @OldTerm
      AND F1104 <> @Server;
END
"@
	
	$dropViewCommandRun = @"
IF OBJECT_ID('Run_Load', 'V') IS NOT NULL DROP VIEW Run_Load;
"@
	
	# ===============================================================================================
	# STD_TAB (store key + CompanyName suffix fix in F1531)
	# ===============================================================================================
	
	$createViewCommandStd = @"
IF OBJECT_ID('Std_Load', 'V') IS NOT NULL DROP VIEW Std_Load;
EXEC(N'
CREATE VIEW Std_Load AS
SELECT F1056, F1531
FROM $stdTableName;
');
"@
	
	$fixStdTabStoreKey_AndCompanySuffix = @"
DECLARE @NewStore  varchar(10) = '$storeNumber';
DECLARE @HostStore varchar(10) = '$hostStoreNumber';

-- =========================================
-- 1) Force store key: keep ONLY (@NewStore + @HostStore)
--    - If @NewStore row exists: delete all other non-host rows
--    - Else: convert ONE non-host row to @NewStore, then delete the rest
-- =========================================
IF EXISTS (SELECT 1 FROM $stdTableName WHERE F1056 = @NewStore)
BEGIN
    DELETE FROM $stdTableName
    WHERE F1056 <> @NewStore
      AND F1056 <> @HostStore;
END
ELSE
BEGIN
    DECLARE @PickOld varchar(10) = NULL;

    SELECT TOP 1 @PickOld = F1056
    FROM $stdTableName
    WHERE F1056 <> @HostStore
    ORDER BY F1056;

    IF @PickOld IS NOT NULL
    BEGIN
        UPDATE $stdTableName
        SET F1056 = @NewStore
        WHERE F1056 = @PickOld;
    END

    DELETE FROM $stdTableName
    WHERE F1056 <> @NewStore
      AND F1056 <> @HostStore;
END

-- =========================================
-- 2) Update F1531 ONLY IF a store number is already present at the end
--    (space/#/- + 3 or 4 digits). If no suffix exists -> do nothing.
--    Robust against NBSP/TAB/CR/LF at the end.
-- =========================================
;WITH X AS
(
    SELECT
        s.F1056,
        s.F1531,
        Clean = RTRIM(
                    REPLACE(REPLACE(REPLACE(REPLACE(
                        LTRIM(RTRIM(s.F1531)),
                        CHAR(160), ' '),
                        CHAR(9),  ' '),
                        CHAR(13), ' '),
                        CHAR(10), ' ')
                )
    FROM STD_TAB s
    WHERE s.F1056 = @NewStore
      AND s.F1531 IS NOT NULL
)
UPDATE X
SET F1531 =
    CASE
        WHEN Clean LIKE '%[ #\-][0-9][0-9][0-9][0-9]'
             AND TRY_CONVERT(int, RIGHT(Clean, 4)) IS NOT NULL
             AND RIGHT(Clean, 4) NOT IN ('0901','0999')
        THEN LEFT(Clean, LEN(Clean) - 4) + @NewStore

        WHEN Clean LIKE '%[ #\-][0-9][0-9][0-9]'
             AND TRY_CONVERT(int, RIGHT(Clean, 3)) IS NOT NULL
             AND RIGHT(Clean, 3) NOT IN ('901','999')
        THEN LEFT(Clean, LEN(Clean) - 3) + @NewStore

        WHEN Clean LIKE '%[ #\-][0-9][0-9]'
             AND TRY_CONVERT(int, RIGHT(Clean, 2)) IS NOT NULL
        THEN LEFT(Clean, LEN(Clean) - 2) + @NewStore

        WHEN Clean LIKE '%[ #\-][0-9]'
             AND TRY_CONVERT(int, RIGHT(Clean, 1)) IS NOT NULL
        THEN LEFT(Clean, LEN(Clean) - 1) + @NewStore

        ELSE F1531
    END
WHERE
    Clean LIKE '%[ #\-][0-9]'
    OR Clean LIKE '%[ #\-][0-9][0-9]'
    OR Clean LIKE '%[ #\-][0-9][0-9][0-9]'
    OR Clean LIKE '%[ #\-][0-9][0-9][0-9][0-9]';
"@
	
	$dropViewCommandStd = @"
IF OBJECT_ID('Std_Load', 'V') IS NOT NULL DROP VIEW Std_Load;
"@
	
	# ===============================================================================================
	# Execute SQL commands (ORDER MATTERS)
	# ===============================================================================================
	
	$sqlCommands = @(
		$createViewCommandTer,
		$moveProtectedTer_ToNewStore
	)
	
	if ($isServerMoniker)
	{
		$sqlCommands += $moveServerNodes902_920_ToNewStore_AndUpdate
		$sqlCommands += $deleteOtherStoresTer
		$sqlCommands += $deleteNonServerRowsForNewStore
		$sqlCommands += $upsertServer901Ter_ForNewStoreOnly
	}
	else
	{
		$sqlCommands += $preserve901ForNonServerStoreChange
		$sqlCommands += $deleteOtherStoresTer
		$sqlCommands += $upsertLaneTer
		
		if ($isLocalTerminalDB)
		{
			$sqlCommands += $cleanupTer_ForLocalTerminalDB
		}
	}
	
	$sqlCommands += @(
		$dropViewCommandTer,
		
		# STO_TAB
		$createViewCommandSto,
		$upsertStoServer901
	)
	
	if ($isServerMoniker)
	{
		$sqlCommands += $cleanupSto_ForServer
	}
	else
	{
		$sqlCommands += $upsertStoTerminal
		if ($isLocalTerminalDB)
		{
			$sqlCommands += $cleanupSto_ForLocalTerminalDB
		}
	}
	
	$sqlCommands += @(
		$dropViewCommandSto,
		
		# LNK_TAB
		$createViewCommandLnk,
		$preserve901LnkForNewStore,
		$rebuildLnkForTerminal
	)
	
	if (-not $isServerMoniker -and $isLocalTerminalDB)
	{
		$sqlCommands += $cleanupLnk_ForLocalTerminalDB
	}
	
	$sqlCommands += @(
		$dropViewCommandLnk
	)
	
	if (-not $isServerMoniker)
	{
		$sqlCommands += @(
			$createViewCommandRun,
			$updateRunTab_SafeSwap_OldToNew,
			$dropViewCommandRun
		)
	}
	
	$sqlCommands += @(
		$createViewCommandStd,
		$fixStdTabStoreKey_AndCompanySuffix,
		$dropViewCommandStd
	)
	
	$allSqlSuccessful = $true
	$failedSqlCommands = @()
	
	foreach ($command in $sqlCommands)
	{
		if (-not (Execute_SQL_Commands -commandText $command))
		{
			$allSqlSuccessful = $false
			$failureMessage = if (-not [string]::IsNullOrWhiteSpace($script:LastSqlExecutionError)) { $script:LastSqlExecutionError } else { "Unknown SQL execution failure." }
			$failedSqlCommands += ("Error: " + $failureMessage + "`r`nSQL:`r`n" + $command)
		}
	}
	
	return @{
		Success		    = $allSqlSuccessful
		FailedCommands  = $failedSqlCommands
		IsServerMoniker = $isServerMoniker
		RunTabUpdated   = (-not $isServerMoniker)
		OldTerminalUsed = $oldTerminal
		NewTerminalUsed = $newTerminal
		DatabaseContext = $DatabaseContext
		LocalTerminalDB = $isLocalTerminalDB
		DBNAME		    = $dbNameDetected
		ConfiguredDBSERVER = if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('ConfiguredDBSERVER')) { [string]$script:FunctionResults['ConfiguredDBSERVER'] } else { $null }
		RuntimeDBSERVER = if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('RuntimeDBSERVER')) { [string]$script:FunctionResults['RuntimeDBSERVER'] } else { $null }
	}
}

# ===================================================================================================
#                             FUNCTION: Update_SQL_Tables_For_Store_Number_Change
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates STD_TAB in the SQL database after store number change.
# ===================================================================================================

function Update_SQL_Tables_For_Store_Number_Change
{
	param (
		[string]$storeNumber # New store number to apply
	)
	
	# =============================
	# TABLE NAMES
	# =============================
	$stdTableName = "STD_TAB"
	$terTableName = "TER_TAB"
	
	# =============================
	# SQL FOR STD_TAB
	# =============================
	
	# Create view (required by Storeman load process)
	$createViewCommandStd = @"
IF OBJECT_ID('Std_Load', 'V') IS NOT NULL DROP VIEW Std_Load;
EXEC(N'
CREATE VIEW Std_Load AS
SELECT F1056
FROM $stdTableName;
');
"@
	
	# Update STD_TAB.F1056 to the new store number
	$updateStdTabCommand = @"
UPDATE $stdTableName
SET F1056 = '$storeNumber';
"@
	
	# Drop the view
	$dropViewCommandStd = "DROP VIEW Std_Load;"
	
	
	# =============================
	# SQL FOR TER_TAB
	# =============================
	# Only update rows where F1057 = '901'
	# This ensures ONLY the back-office record gets changed
	# and does NOT break lane entries.
	# =============================
	$updateTerTabCommand = @"
UPDATE $terTableName
SET F1056 = '$storeNumber'
WHERE F1057 = '901';
"@
	
	
	# =============================
	# EXECUTION PIPELINE
	# =============================
	$sqlCommands = @(
		$createViewCommandStd,
		$updateStdTabCommand,
		$dropViewCommandStd,
		$updateTerTabCommand # <-- NEW COMMAND
	)
	
	$allSqlSuccessful = $true
	$failedSqlCommands = @()
	
	foreach ($command in $sqlCommands)
	{
		# Execute_SQL_Commands is part of your existing framework
		# so no changes here.
		if (-not (Execute_SQL_Commands -commandText $command))
		{
			$allSqlSuccessful = $false
			$failureMessage = if (-not [string]::IsNullOrWhiteSpace($script:LastSqlExecutionError)) { $script:LastSqlExecutionError } else { "Unknown SQL execution failure." }
			$failedSqlCommands += ("Error: " + $failureMessage + "`r`nSQL:`r`n" + $command)
		}
	}
	
	# =============================
	# RETURN STRUCTURED RESULT
	# =============================
	return @{
		Success	       = $allSqlSuccessful
		FailedCommands = $failedSqlCommands
	}
}

# ===================================================================================================
#                          FUNCTION: Convert_Prefix_Length_To_Subnet_Mask
# ---------------------------------------------------------------------------------------------------
# Description:
#   Converts an IPv4 prefix length (24) to a dotted subnet mask (255.255.255.0).
#
# Returns:
#   The dotted mask, or $null when the prefix length is not 0-32.
# ===================================================================================================

function Convert_Prefix_Length_To_Subnet_Mask
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		$PrefixLength
	)
	
	$prefix = 0
	if ($null -eq $PrefixLength -or -not [int]::TryParse(([string]@($PrefixLength)[0]).Trim(), [ref]$prefix)) { return $null }
	if ($prefix -lt 0 -or $prefix -gt 32) { return $null }
	
	$bits = ('1' * $prefix).PadRight(32, '0')
	$octets = for ($i = 0; $i -lt 32; $i += 8) { [Convert]::ToInt32($bits.Substring($i, 8), 2) }
	return ($octets -join '.')
}

# ===================================================================================================
#                          FUNCTION: Convert_Subnet_Mask_To_Prefix_Length
# ---------------------------------------------------------------------------------------------------
# Description:
#   Converts a dotted IPv4 subnet mask (255.255.255.0) to its prefix length (24).
#
# Returns:
#   The prefix length (0-32), or $null when the mask is not a valid contiguous IPv4 mask.
# ===================================================================================================

function Convert_Subnet_Mask_To_Prefix_Length
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$SubnetMask
	)
	
	if ([string]::IsNullOrWhiteSpace($SubnetMask) -or $SubnetMask.Trim() -notmatch '^\d{1,3}(\.\d{1,3}){3}$') { return $null }
	
	$bits = ""
	foreach ($octet in $SubnetMask.Trim().Split('.'))
	{
		$value = [int]$octet
		if ($value -gt 255) { return $null }
		$bits += [Convert]::ToString($value, 2).PadLeft(8, '0')
	}
	
	# A mask is ones followed by zeros only
	if ($bits -notmatch '^1*0*$') { return $null }
	$firstZero = $bits.IndexOf('0')
	if ($firstZero -lt 0) { return 32 }
	return $firstZero
}

# ===================================================================================================
#                                       FUNCTION: Get_Active_IP_Config
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the active IP configuration for network adapters that are up and have valid IPv4 addresses.
#   Optimized for performance by prefiltering and minimizing pipeline overhead.
#   Attempts WMI first for compatibility, falls back to NetAdapter methods if WMI fails.
# ===================================================================================================

function Get_Active_IP_Config
{
	try
	{
		# WMI attempt: Get default route from Win32_IP4RouteTable
		$defaultRoute = Get-WmiObject -Class Win32_IP4RouteTable -ErrorAction Stop |
		Where-Object { $_.Destination -eq '0.0.0.0' -and $_.Mask -eq '0.0.0.0' -and $_.NextHop -ne '0.0.0.0' } |
		Sort-Object -Property Metric1 |
		Select-Object -First 1
		
		if ($defaultRoute -and $defaultRoute.InterfaceIndex)
		{
			$idx = [int]$defaultRoute.InterfaceIndex
			
			# Check adapter status via Win32_NetworkAdapter (NetConnectionStatus 2 = Connected).
			# The route's InterfaceIndex is the adapter's InterfaceIndex - not its DeviceID / Index
			# (those number adapters differently, so filtering on them picked an unrelated adapter).
			$ad = Get-WmiObject -Class Win32_NetworkAdapter -Filter "InterfaceIndex = $idx" -ErrorAction SilentlyContinue | Select-Object -First 1
			if ($ad -and $ad.NetConnectionStatus -eq 2)
			{
				# Get config via Win32_NetworkAdapterConfiguration
				$wmiCfg = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "InterfaceIndex = $idx" -ErrorAction SilentlyContinue | Select-Object -First 1
				if ($wmiCfg -and $wmiCfg.IPEnabled -and $wmiCfg.DefaultIPGateway)
				{
					# Filter for valid IPv4 (non-APIPA)
					$validIPs = $wmiCfg.IPAddress | Where-Object { $_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' -and $_ -notmatch '^169\.254\.' }
					if ($validIPs)
					{
						# Construct a custom object similar to Get-NetIPConfiguration for consistency
						$validIPs = @($validIPs)
						$ipv4Mask = @($wmiCfg.IPSubnet | Where-Object { $_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' }) | Select-Object -First 1
						$ipv4Address = [PSCustomObject]@{
							IPAddress    = $validIPs[0]
							PrefixLength = Convert_Subnet_Mask_To_Prefix_Length -SubnetMask $ipv4Mask
						}
						$cfg = [PSCustomObject]@{
							InterfaceAlias	   = $ad.NetConnectionID
							InterfaceIndex	   = $idx
							IPv4Address	       = $ipv4Address
							IPv4DefaultGateway = [PSCustomObject]@{ NextHop = $wmiCfg.DefaultIPGateway[0] }
							DNSServer		   = $wmiCfg.DNSServerSearchOrder
							# Add more properties if needed for parity
						}
						return @($cfg) # Return as array for consistent .Count / [0]
					}
				}
			}
		}
	}
	catch
	{
		# Ignore and fall back to original methods
	}
	
	# =========================================================================================
	# FALLBACK: Original approach (NetRoute/NetAdapter/NetIPConfiguration)
	# =========================================================================================
	
	try
	{
		$defaultRoute = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -AddressFamily IPv4 -ErrorAction Stop |
		Where-Object { $_.NextHop -and $_.NextHop -ne '0.0.0.0' } |
		Sort-Object -Property RouteMetric, InterfaceMetric |
		Select-Object -First 1
		
		if ($defaultRoute -and $defaultRoute.InterfaceIndex)
		{
			$idx = [int]$defaultRoute.InterfaceIndex
			
			# Ensure adapter is Up (do NOT restrict to -Physical here; active interface might be Wi-Fi/VPN/etc.)
			$ad = Get-NetAdapter -InterfaceIndex $idx -ErrorAction SilentlyContinue
			if ($ad -and $ad.Status -eq 'Up')
			{
				# Make sure there is at least one valid IPv4 (non-APIPA) on this interface
				$ipRow = Get-NetIPAddress -InterfaceIndex $idx -AddressFamily IPv4 -ErrorAction SilentlyContinue |
				Where-Object { $_.IPAddress -and $_.IPAddress -notlike '169.254.*' -and $_.IPAddress -ne '0.0.0.0' } |
				Select-Object -First 1
				
				if ($ipRow)
				{
					$cfg = Get-NetIPConfiguration -InterfaceIndex $idx -ErrorAction SilentlyContinue
					if ($cfg -and $cfg.IPv4Address)
					{
						# Validate IPv4Address (can be array). If any non-APIPA exists, return it.
						$hasValid = $false
						foreach ($addr in @($cfg.IPv4Address))
						{
							if ($addr -and $addr.IPAddress -and ($addr.IPAddress -notlike '169.254.*'))
							{
								$hasValid = $true
								break
							}
						}
						
						if ($hasValid)
						{
							return @($cfg) # return as array for consistent .Count / [0]
						}
					}
				}
			}
		}
	}
	catch
	{
		# ignore and fall back
	}
	
	# =========================================================================================
	# SECONDARY FALLBACK: Physical + Up adapters -> Get-NetIPConfiguration
	# =========================================================================================
	
	# Prefilter by adapter status to reduce objects early
	$adapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object Status -eq 'Up'
	if (-not $adapters) { return $null }
	
	# Get only configs for up adapters
	$adapterNames = $adapters | Select-Object -ExpandProperty Name
	$ipConfigs = Get-NetIPConfiguration -InterfaceAlias $adapterNames -ErrorAction SilentlyContinue
	
	# Filter for valid IPv4 (not APIPA/169.254.x.x, not null)
	$validConfigs = $ipConfigs | Where-Object {
		$_.IPv4Address -and
		$_.IPv4Address.IPAddress -and
		($_.IPv4Address.IPAddress -notlike '169.254*')
	}
	
	if ($validConfigs) { return @($validConfigs) }
	return $null
}

# ===================================================================================================
#                                       FUNCTION: Remove_GT_Registry_Values
# ---------------------------------------------------------------------------------------------------
# Description:
#   Removes all registry values starting with 'GT' from specified registry paths.
# ===================================================================================================

function Remove_GT_Registry_Values
{
	# Define registry paths for 32-bit and 64-bit
	$regPath32 = "HKLM:\SOFTWARE\Store Management\Counters"
	$regPath64 = "HKLM:\SOFTWARE\Wow6432Node\Store Management\Counters"
	
	$is64bit = [System.Environment]::Is64BitOperatingSystem
	
	$totalDeletedCount = 0
	$success = $true
	$status = "Successful"
	$message = ""
	
	# Build list of candidate paths (prefer Wow6432Node on 64-bit, but include both if present)
	$pathsToCheck = @()
	
	if ($is64bit)
	{
		if (Test-Path $regPath64) { $pathsToCheck += $regPath64 }
		if (Test-Path $regPath32) { $pathsToCheck += $regPath32 }
	}
	else
	{
		if (Test-Path $regPath32) { $pathsToCheck += $regPath32 }
	}
	
	if (-not $pathsToCheck -or $pathsToCheck.Count -eq 0)
	{
		return @{
			Success	     = $false
			Status	     = "Failed"
			DeletedCount = 0
			Message	     = "No valid registry paths found for the current environment."
		}
	}
	
	foreach ($path in $pathsToCheck)
	{
		try
		{
			# Get value names from the key by inspecting item properties
			# (Get-ItemProperty returns a PSCustomObject with properties = value names + PS metadata)
			$itemProps = Get-ItemProperty -Path $path -ErrorAction Stop
			
			# Collect only real registry value names starting with GT (exclude PS metadata props)
			$gtNames = @()
			foreach ($p in $itemProps.PSObject.Properties)
			{
				if ($p.Name -eq 'PSPath') { continue }
				if ($p.Name -eq 'PSParentPath') { continue }
				if ($p.Name -eq 'PSChildName') { continue }
				if ($p.Name -eq 'PSDrive') { continue }
				if ($p.Name -eq 'PSProvider') { continue }
				
				if ($p.Name -like 'GT*')
				{
					$gtNames += $p.Name
				}
			}
			
			# Delete each GT* value
			foreach ($name in $gtNames)
			{
				try
				{
					Remove-ItemProperty -Path $path -Name $name -ErrorAction Stop
					$totalDeletedCount++
				}
				catch
				{
					# Keep going but mark overall as partial failure
					$success = $false
					$status = "Failed"
					if ($message) { $message += " | " }
					$message += "Failed to delete '$name' in '$path': $_"
				}
			}
		}
		catch
		{
			$success = $false
			$status = "Failed"
			if ($message) { $message += " | " }
			$message += "Error accessing registry path '$path': $_"
		}
	}
	
	return @{
		Success	     = $success
		Status	     = $status
		DeletedCount = $totalDeletedCount
		Message	     = $message
	}
}

# ===================================================================================================
#                             FUNCTION: Remove_MiniGhost_Temp_Copies
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deletes the copies of MiniGhost left in %TEMP% (MiniGhost.ps1 and the
#   MiniGhost_<version>_Elevated.ps1 that versions before 1.4.2 wrote for irm | iex launches).
#
# Parameters:
#   - TempPath [string]: folder to clean (default: the user's %TEMP%)
#
# Returns:
#   Number of files deleted.
#
# Notes:
#   - Top level of the temp folder only. The old cleanup searched all of %TEMP% recursively for
#     MiniGhost.ps1 on every exit (slow on a full temp folder) and never matched the _Elevated copy.
# ===================================================================================================

function Remove_MiniGhost_Temp_Copies
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$TempPath = [System.IO.Path]::GetTempPath()
	)
	
	$deleted = 0
	if ([string]::IsNullOrWhiteSpace($TempPath) -or -not (Test-Path -LiteralPath $TempPath)) { return $deleted }
	
	foreach ($file in @(Get-ChildItem -LiteralPath $TempPath -File -Filter 'MiniGhost*.ps1' -ErrorAction SilentlyContinue))
	{
		if ($file.Name -notmatch '^(?i)MiniGhost(_[\d.]+_Elevated)?\.ps1$') { continue }
		try
		{
			Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
			$deleted++
		}
		catch { }
	}
	return $deleted
}

# ===================================================================================================
# FUNCTION: Remove_Old_XZ_Folders
# ---------------------------------------------------------------------------------------------------
# Description:
# Removes old X* and Z* folders under \storeman\office that do NOT match the NEW store and allowed
# terminal numbers.
#
# Strict folder name pattern acted upon:
#   ^(X|Z)[A-Za-z](store 3-4 digits)(terminal 3 digits)$
# Examples: XF0231006, XW0242901, ZF1234006, ZW1234901
#
# Keeps ONLY folders that match:
#   - Store = new store (3-digit OR 4-digit representation)
#   - Terminal IN: current lane terminal (from MachineName), 900, 901
#
# Safeguard: refuses to run if extracted terminal is 901 unless -AllowBackoffice is passed.
# Office folder: -OfficePath, else the detected storeman Office folder ($OfficePath), else
# \\localhost\storeman\office, C:\storeman\office, D:\storeman\office.
# Automatic: no prompts.
# ===================================================================================================

function Remove_Old_XZ_Folders
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus,
		[Parameter(Mandatory = $false)]
		[switch]$AllowBackoffice,
		[Parameter(Mandatory = $false)]
		[string]$OfficePath
	)
	
	# -----------------------------
	# Init / status bucket
	# -----------------------------
	$deletedFolders = @()
	$failedToDeleteFolders = @()
	$anyFoldersSeen = $false
	$anyCandidatesMatched = $false
	
	if ($OperationStatus)
	{
		if (-not $OperationStatus.ContainsKey("OldXFoldersDeletion"))
		{
			$OperationStatus["OldXFoldersDeletion"] = [pscustomobject]@{
				Status  = ""
				Message = ""
				Details = ""
			}
		}
	}
	
	# -----------------------------
	# Validate store number
	# -----------------------------
	$storeNumberTrim = $StoreNumber
	if ($null -eq $storeNumberTrim) { $storeNumberTrim = "" }
	$storeNumberTrim = $storeNumberTrim.Trim()
	
	if ($storeNumberTrim -notmatch '^(?!0+$)\d{3,4}$')
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Invalid StoreNumber '$StoreNumber' (must be 3 or 4 digits, not all zeros)."
			$OperationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
		}
		Write-Host "Failed: Invalid StoreNumber '$StoreNumber'." -ForegroundColor Red
		return
	}
	
	$storeInt = [int]$storeNumberTrim
	$store3 = $storeInt.ToString("D3")
	$store4 = $storeInt.ToString("D4")
	
	# -----------------------------
	# Pick base path
	# -----------------------------
	# The detected storeman folder comes first: it can be on another drive or named differently
	# (e.g. E:\Storeman), where the fixed paths below would miss it or pick a stale copy.
	$possibleBasePaths = @()
	if (-not [string]::IsNullOrWhiteSpace($OfficePath)) { $possibleBasePaths += $OfficePath }
	else
	{
		# (the parameter hides the script-level $OfficePath, so read that one explicitly)
		foreach ($sc in @('Script', 'Global'))
		{
			$v = Get-Variable -Name OfficePath -Scope $sc -ErrorAction SilentlyContinue
			if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $possibleBasePaths += [string]$v.Value; break }
		}
		$possibleBasePaths += @("\\localhost\storeman\office", "C:\storeman\office", "D:\storeman\office")
	}
	$basePath = $null
	foreach ($p in $possibleBasePaths)
	{
		if (Test-Path $p) { $basePath = $p; break }
	}
	
	if (-not $basePath)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "None of the base paths exist: $($possibleBasePaths -join ', ')"
			$OperationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
		}
		Write-Host "Failed: None of the base paths exist: $($possibleBasePaths -join ', ')" -ForegroundColor Red
		return
	}
	
	# -----------------------------
	# Extract machine terminal (last 1-3 digits, padded to 3)
	# -----------------------------
	$mn = $MachineName
	if ($null -eq $mn) { $mn = "" }
	$mn = $mn.Trim()
	
	$mn = $mn -replace '^[\\\/]+', '' # strip leading \\ or /
	if ($mn -match '[\\\/]') { $mn = ($mn -split '[\\\/]')[0] } # keep host portion
	if ($mn -match '\.') { $mn = ($mn -split '\.')[0] } # drop domain
	$mn = $mn.Trim().ToUpper()
	
	if ($mn -notmatch '(\d{1,3})$')
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "MachineName '$MachineName' does not end with 1-3 digits; cannot extract terminal number."
			$OperationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
		}
		Write-Host "Failed: MachineName '$MachineName' cannot provide terminal digits." -ForegroundColor Red
		return
	}
	
	$machineNumber = ([int]$Matches[1]).ToString("D3")
	
	# Safeguard for backoffice unless explicitly allowed
	if ($machineNumber -eq "901" -and -not $AllowBackoffice)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Refusing to run because extracted terminal number is 901 (backoffice). Use -AllowBackoffice to override."
			$OperationStatus["OldXFoldersDeletion"].Details = "This could delete many lane folders. Override only if intentional."
		}
		Write-Host "Failed: Refusing to run on terminal 901 (backoffice safeguard). Use -AllowBackoffice to override." -ForegroundColor Red
		return
	}
	
	# Keep set: current lane + 900 + 901
	$keepStores = @($store3, $store4)
	$keepTerminals = @($machineNumber, "900", "901")
	
	# -----------------------------
	# Enumerate X* and Z* folders
	# -----------------------------
	$folders = @()
	
	$xFolders = Get-ChildItem -Path $basePath -Directory -Filter "X*" -ErrorAction SilentlyContinue
	if ($xFolders) { $folders += $xFolders }
	
	$zFolders = Get-ChildItem -Path $basePath -Directory -Filter "Z*" -ErrorAction SilentlyContinue
	if ($zFolders) { $folders += $zFolders }
	
	if (-not $folders -or $folders.Count -eq 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "No Folders Found"
			$OperationStatus["OldXFoldersDeletion"].Message = "No X* or Z* folders were found under '$basePath'."
			$OperationStatus["OldXFoldersDeletion"].Details = "Nothing to delete."
		}
		Write-Host "Info: No X* or Z* folders found under '$basePath'." -ForegroundColor Cyan
		return
	}
	
	$anyFoldersSeen = $true
	
	foreach ($folder in $folders)
	{
		$folderName = $folder.Name
		
		# STRICT pattern only: (X|Z)(letter)(store 3-4 digits)(terminal 3 digits)
		if ($folderName -match '^(?<prefix>[XZ][A-Za-z])(?<store>\d{3,4})(?<terminal>\d{3})$')
		{
			$folderStore = $matches['store']
			$folderTerminal = $matches['terminal']
			
			$storeOk = ($keepStores -contains $folderStore)
			$terminalOk = ($keepTerminals -contains $folderTerminal)
			
			# Delete if it does NOT match new store AND allowed terminal(s)
			if (-not ($storeOk -and $terminalOk))
			{
				$anyCandidatesMatched = $true
				
				$maxRetries = 3
				$retryCount = 0
				$deleted = $false
				
				while ($retryCount -lt $maxRetries -and -not $deleted)
				{
					try
					{
						Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction Stop
						$deletedFolders += $folderName
						$deleted = $true
					}
					catch
					{
						$retryCount++
						Start-Sleep -Seconds 2
						
						if ($retryCount -ge $maxRetries)
						{
							$failedToDeleteFolders += $folderName
							Write-Host "Failed to delete folder: $folderName. Error: $($_.Exception.Message)" -ForegroundColor Red
						}
					}
				}
			}
		}
	}
	
	# -----------------------------
	# Results
	# -----------------------------
	$resultMessage = ""
	if ($deletedFolders.Count -gt 0)
	{
		$resultMessage += "Deleted folders:`n$($deletedFolders -join "`n")`n"
	}
	if ($failedToDeleteFolders.Count -gt 0)
	{
		$resultMessage += "Failed to delete folders:`n$($failedToDeleteFolders -join "`n")`n"
	}
	
	if (-not $anyCandidatesMatched)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "No Matching Folders"
			$OperationStatus["OldXFoldersDeletion"].Message = "No X*/Z* folders needed deletion. Kept stores: $($keepStores -join ', ') terminals: $($keepTerminals -join ', ')."
			$OperationStatus["OldXFoldersDeletion"].Details = "Nothing to delete."
		}
		Write-Host "Info: Nothing to delete. Kept stores: $($keepStores -join ', ') terminals: $($keepTerminals -join ', ')." -ForegroundColor Cyan
		return
	}
	
	if ($deletedFolders.Count -gt 0 -and $failedToDeleteFolders.Count -eq 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Successful"
			$OperationStatus["OldXFoldersDeletion"].Message = "Old X*/Z* folders deleted successfully."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Success: Old X*/Z* folders deleted successfully." -ForegroundColor Green
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Green }
	}
	elseif ($deletedFolders.Count -gt 0 -and $failedToDeleteFolders.Count -gt 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Partial Failure"
			$OperationStatus["OldXFoldersDeletion"].Message = "Some old X*/Z* folders could not be deleted."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Warning: Some old X*/Z* folders could not be deleted." -ForegroundColor Yellow
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Yellow }
	}
	else
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Failed to delete any old X*/Z* folders."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Error: Failed to delete any old X*/Z* folders." -ForegroundColor Red
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Red }
	}
}

# ===================================================================================================
#                         FUNCTION: Invoke_Machine_Name_Change_Workflow
# ---------------------------------------------------------------------------------------------------
# Description:
# Coordinates the machine-name action from the GUI. If the requested name matches the current name,
# the workflow skips Rename-Computer and still performs the full INI + SQL sync.
# ===================================================================================================

function Invoke_Machine_Name_Change_Workflow
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$RequestedMachineName,
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus
	)
	
	$oldMachineName = $env:COMPUTERNAME
	$machineContext = Get_Machine_Name_Context -MachineName $RequestedMachineName -CurrentMachineName $oldMachineName
	if (-not $machineContext.Success)
	{
		Show_App_Message -Text $machineContext.ErrorMessage -Title "Error" -Icon Error | Out-Null
		Set_Operation_Status -OperationStatus $OperationStatus -Key "MachineNameChange" -Status "Failed" -Message "Invalid machine name." -Details $machineContext.ErrorMessage
		return
	}
	
	$confirmMessage = $null
	$confirmTitle = $null
	if ($machineContext.IsSameMachineName)
	{
		$confirmMessage = "The requested machine name matches the current machine name '$($machineContext.MachineName)'.`n`nRun a full INI + SQL sync anyway?"
		$confirmTitle = "Confirm Machine Sync"
	}
	else
	{
		$confirmMessage = "Are you sure you want to change the machine name to '$($machineContext.MachineName)' and run the full INI + SQL sync?"
		$confirmTitle = "Confirm Machine Name Change"
	}
	
	$result = Show_App_Message -Text $confirmMessage -Title $confirmTitle -Buttons YesNo -Icon Question
	if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }
	
	$currentStoreNumber = Get_Store_Number_From_INI
	if (-not $currentStoreNumber)
	{
		Show_App_Message -Text "Store number not found in startup.ini." -Title "Error" -Icon Error | Out-Null
		Set_Operation_Status -OperationStatus $OperationStatus -Key "MachineNameChange" -Status "Failed" -Message "Store number not found in startup.ini." -Details ""
		return
	}
	
	$startupIniPath = Resolve_Startup_Ini_Path -StartupIniPath $StartupIniPath
	$renamePerformed = $false
	$script:newMachineName = $machineContext.MachineName
	
	try
	{
		if (-not $machineContext.IsSameMachineName)
		{
			Rename-Computer -NewName $machineContext.MachineName -Force -ErrorAction Stop
			$renamePerformed = $true
			
			Set_Label_Text -Label $script:machineNameLabel -Text "$env:COMPUTERNAME -> $script:newMachineName (after reboot)" -RefreshParent -ProcessEvents
			Set_Operation_Status -OperationStatus $OperationStatus -Key "MachineNameChange" -Status "Successful" -Message "Machine name changed successfully." -Details "Machine name changed to '$script:newMachineName'."
			
			Remove_Old_XZ_Folders -MachineName $script:newMachineName -StoreNumber $currentStoreNumber
		}
		
		$syncResult = Invoke_Machine_Name_Sync -StoreNumber $currentStoreNumber -MachineName $machineContext.MachineName -OldMachineName $oldMachineName -MachineContext $machineContext -StartupIniPath $startupIniPath -OperationStatus $OperationStatus
		$syncFailureLogText = $null
		if (-not $syncResult.Success)
		{
			$syncFailureLogParts = @(
				"Machine Name:`r`n$($machineContext.MachineName)",
				"Store Number:`r`n$currentStoreNumber"
			)
			
			if (-not $syncResult.IniSyncSuccess -or -not $syncResult.StartupIniSuccess)
			{
				$iniSectionParts = @()
				
				if ($syncResult.IniSyncError)
				{
					$iniSectionParts += "Update_INIs Error:`r`n$($syncResult.IniSyncError)"
				}
				
				if ($syncResult.StartupIniError)
				{
					$iniSectionParts += "Startup.ini Error:`r`n$($syncResult.StartupIniError)"
				}
				
				if ($syncResult.StartupIniDetails)
				{
					$iniSectionParts += "Additional Details:`r`n$($syncResult.StartupIniDetails)"
				}
				elseif ($OperationStatus -and $OperationStatus.ContainsKey("StartupIniUpdate") -and $OperationStatus["StartupIniUpdate"].Details)
				{
					$iniSectionParts += "Operation Status Details:`r`n$($OperationStatus["StartupIniUpdate"].Details)"
				}
				
				if ($iniSectionParts.Count -gt 0)
				{
					$syncFailureLogParts += "INI Sync:`r`n$($iniSectionParts -join "`r`n`r`n")"
				}
			}
			
			if (-not $syncResult.SqlSyncSuccess)
			{
				$sqlSectionParts = @()
				
				if ($syncResult.SqlSyncError)
				{
					$sqlSectionParts += "SQL Sync Error:`r`n$($syncResult.SqlSyncError)"
				}
				
				if ($syncResult.SqlResult)
				{
					if ($syncResult.SqlResult.ContainsKey('DatabaseContext') -and $syncResult.SqlResult.DatabaseContext)
					{
						$sqlSectionParts += "Database Context:`r`n$($syncResult.SqlResult.DatabaseContext)"
					}
					
					if ($syncResult.SqlResult.ContainsKey('DBNAME') -and $syncResult.SqlResult.DBNAME)
					{
						$sqlSectionParts += "DBNAME:`r`n$($syncResult.SqlResult.DBNAME)"
					}
					
					if ($syncResult.SqlResult.ContainsKey('ConfiguredDBSERVER') -and $syncResult.SqlResult.ConfiguredDBSERVER)
					{
						$sqlSectionParts += "Configured DBSERVER:`r`n$($syncResult.SqlResult.ConfiguredDBSERVER)"
					}
					elseif ($syncResult.ConfiguredDbServer)
					{
						$sqlSectionParts += "Configured DBSERVER:`r`n$($syncResult.ConfiguredDbServer)"
					}
					
					if ($syncResult.SqlResult.ContainsKey('RuntimeDBSERVER') -and $syncResult.SqlResult.RuntimeDBSERVER)
					{
						$sqlSectionParts += "Runtime SQL Target:`r`n$($syncResult.SqlResult.RuntimeDBSERVER)"
					}
					elseif ($syncResult.RuntimeDbServer)
					{
						$sqlSectionParts += "Runtime SQL Target:`r`n$($syncResult.RuntimeDbServer)"
					}
					
					if ($syncResult.SqlResult.ContainsKey('FailedCommands') -and $syncResult.SqlResult.FailedCommands)
					{
						$sqlSectionParts += "Failed SQL Commands:`r`n$($syncResult.SqlResult.FailedCommands -join "`r`n`r`n")"
					}
				}
				elseif ($OperationStatus -and $OperationStatus.ContainsKey("SQLDatabaseUpdate") -and $OperationStatus["SQLDatabaseUpdate"].Details)
				{
					$sqlSectionParts += "Operation Status Details:`r`n$($OperationStatus["SQLDatabaseUpdate"].Details)"
				}
				
				if ($sqlSectionParts.Count -gt 0)
				{
					$syncFailureLogParts += "SQL Sync:`r`n$($sqlSectionParts -join "`r`n`r`n")"
				}
			}
			
			$syncFailureLogText = ($syncFailureLogParts -join "`r`n`r`n").Trim()
		}
		
		if ($machineContext.IsSameMachineName)
		{
			if ($syncResult.Success)
			{
				Set_Operation_Status -OperationStatus $OperationStatus -Key "MachineNameChange" -Status "Successful" -Message "Machine settings synced successfully." -Details "Requested machine name already matched '$script:newMachineName'; full INI + SQL sync was run without renaming."
				Set_Label_Text -Label $script:machineNameLabel -Text "$script:newMachineName (sync refreshed)" -RefreshParent
				Show_App_Message -Text "Full INI + SQL sync completed for '$script:newMachineName'. No reboot is required because the machine name did not change." -Title "Sync Complete" -Icon Information | Out-Null
			}
			else
			{
				Set_Operation_Status -OperationStatus $OperationStatus -Key "MachineNameChange" -Status "Failed" -Message "Machine sync completed with errors." -Details "Requested machine name already matched '$script:newMachineName', but one or more INI/SQL updates failed."
				if ($syncFailureLogText)
				{
					[void](Show_Action_Prompt_With_Log -Text "Same-name sync ran, but one or more INI/SQL updates failed.`r`n`r`nUse View Log to see the failure reasons, or use Summary for the full operation status." -Title "Sync Completed With Errors" -PrimaryButtonText "OK" -LogText $syncFailureLogText -LogTitle "Sync Failure Log")
				}
				else
				{
					Show_App_Message -Text "Same-name sync ran, but one or more INI/SQL updates failed. Review the summary for details." -Title "Sync Completed With Errors" -Icon Warning | Out-Null
				}
			}
		}
		elseif ($renamePerformed)
		{
			$rebootMessage = if ($syncResult.Success)
			{
				"Machine name changed successfully to '$script:newMachineName'. The system will need to reboot for changes to take effect. Do you want to reboot now?"
			}
			else
			{
				"Machine name changed to '$script:newMachineName', but one or more INI/SQL updates failed. A reboot is still required for the rename. Do you want to reboot now?"
			}
			
			$rebootIcon = if ($syncResult.Success)
			{
				[System.Windows.Forms.MessageBoxIcon]::Information
			}
			else
			{
				[System.Windows.Forms.MessageBoxIcon]::Warning
			}
			
			$rebootResult = $null
			if (-not $syncResult.Success -and $syncFailureLogText)
			{
				$rebootResult = Show_Action_Prompt_With_Log -Text ($rebootMessage + "`r`n`r`nUse View Log to see the failure reasons.") -Title "Machine Rename Complete" -PrimaryButtonText "Reboot Now" -SecondaryButtonText "Later" -LogText $syncFailureLogText -LogTitle "Sync Failure Log"
			}
			else
			{
				$rebootResult = Show_App_Message -Text $rebootMessage -Title "Machine Rename Complete" -Buttons YesNo -Icon $rebootIcon.ToString()
			}
			
			if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes -or $rebootResult -eq "Primary")
			{
				Restart-Computer -Force
			}
		}
	}
	catch
	{
		$errorMessage = $_.Exception.Message
		Show_App_Message -Text "Error changing machine name: $errorMessage" -Title "Error" -Icon Error | Out-Null
		Set_Operation_Status -OperationStatus $OperationStatus -Key "MachineNameChange" -Status "Failed" -Message "Error changing machine name." -Details "Error: $errorMessage"
	}
}

# ===================================================================================================
#                            FUNCTION: Invoke_Machine_Name_Sync
# ---------------------------------------------------------------------------------------------------
# Description:
# Runs the best-effort INI + Startup.ini + SQL sync for a machine name target. This is used after an
# actual rename and also when the requested name already matches the current machine name.
# ===================================================================================================

function Invoke_Machine_Name_Sync
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidatePattern('^(?!0+$)\d{3,4}$')]
		[string]$StoreNumber,
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineName = $env:COMPUTERNAME,
		[Parameter(Mandatory = $false)]
		[psobject]$MachineContext,
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath,
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus
	)
	
	if (-not $MachineContext -or -not $MachineContext.Success)
	{
		$MachineContext = Get_Machine_Name_Context -MachineName $MachineName -CurrentMachineName $OldMachineName
	}
	
	if (-not $MachineContext.Success)
	{
		if ($OperationStatus -and $OperationStatus.ContainsKey("StartupIniUpdate"))
		{
			$OperationStatus["StartupIniUpdate"].Status = "Failed"
			$OperationStatus["StartupIniUpdate"].Message = "INI sync could not start."
			$OperationStatus["StartupIniUpdate"].Details = $MachineContext.ErrorMessage
		}
		if ($OperationStatus -and $OperationStatus.ContainsKey("SQLDatabaseUpdate"))
		{
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "SQL sync could not start."
			$OperationStatus["SQLDatabaseUpdate"].Details = $MachineContext.ErrorMessage
		}
		
		return [PSCustomObject]@{
			Success             = $false
			MachineContext      = $MachineContext
			IniSyncSuccess      = $false
			StartupIniSuccess   = $false
			SqlSyncSuccess      = $false
			StartupIniPath      = $null
			IniSyncError        = $MachineContext.ErrorMessage
			StartupIniError     = $null
			SqlSyncError        = $MachineContext.ErrorMessage
			StartupIniDetails   = $null
			SqlResult           = $null
		}
	}
	
	$resolvedStartupIniPath = Resolve_Startup_Ini_Path -StartupIniPath $StartupIniPath
	$null = Ensure_Database_Connection_Context -StartupIniPath $resolvedStartupIniPath
	
	$configuredDbServer = $null
	if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('ConfiguredDBSERVER') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['ConfiguredDBSERVER']))
	{
		$configuredDbServer = [string]$script:FunctionResults['ConfiguredDBSERVER']
	}
	elseif ($script:FunctionResults -and $script:FunctionResults.ContainsKey('DBSERVER') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['DBSERVER']))
	{
		$configuredDbServer = [string]$script:FunctionResults['DBSERVER']
	}
	
	$runtimeDbServer = $null
	if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('RuntimeDBSERVER') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['RuntimeDBSERVER']))
	{
		$runtimeDbServer = [string]$script:FunctionResults['RuntimeDBSERVER']
	}
	
	$iniSyncSuccess = $false
	$iniSyncError = $null
	try
	{
		$iniSyncSuccess = [bool](Update_INIs -newStoreNumber $StoreNumber -MachineName $MachineContext.MachineName -OldStoreNumber $StoreNumber -OldMachineName $OldMachineName -CreateSmsHttpsIfMissing -StartupIniPath $resolvedStartupIniPath)
		if (-not $iniSyncSuccess)
		{
			$iniSyncError = "Update_INIs reported a failure."
		}
	}
	catch
	{
		$iniSyncError = $_.Exception.Message
	}
	
	$startupIniResult = Update_Startup_Ini_For_Machine -MachineContext $MachineContext -StartupIniPath $resolvedStartupIniPath
	$startupIniSuccess = $startupIniResult.Success
	$startupIniError = $startupIniResult.ErrorMessage
	
	if ($OperationStatus -and $OperationStatus.ContainsKey("StartupIniUpdate"))
	{
		if ($iniSyncSuccess -and $startupIniSuccess)
		{
			$OperationStatus["StartupIniUpdate"].Status = "Successful"
			$OperationStatus["StartupIniUpdate"].Message = "INI files updated successfully."
			$OperationStatus["StartupIniUpdate"].Details = $startupIniResult.Details
		}
		else
		{
			$detailParts = @()
			if ($iniSyncError) { $detailParts += $iniSyncError }
			if ($startupIniError) { $detailParts += $startupIniError }
			if ($startupIniResult.Details) { $detailParts += $startupIniResult.Details }
			
			$OperationStatus["StartupIniUpdate"].Status = "Failed"
			$OperationStatus["StartupIniUpdate"].Message = "INI sync completed with errors."
			$OperationStatus["StartupIniUpdate"].Details = ($detailParts -join " ")
		}
	}
	
	$sqlUpdateResult = $null
	$sqlSyncSuccess = $false
	$sqlSyncError = $null
	try
	{
		$sqlUpdateResult = Update_SQL_Tables_For_Machine_Name_Change `
															 -storeNumber $StoreNumber `
															 -machineName $MachineContext.MachineName `
															 -machineNumber $MachineContext.MachineNumber `
															 -OldMachineName $OldMachineName
		
		if ($sqlUpdateResult -and $sqlUpdateResult.Success)
		{
			$sqlSyncSuccess = $true
		}
		else
		{
			$sqlSyncError = "SQL update result indicated failure."
		}
	}
	catch
	{
		$sqlSyncError = $_.Exception.Message
	}
	
	if ($OperationStatus -and $OperationStatus.ContainsKey("SQLDatabaseUpdate"))
	{
		if ($sqlSyncSuccess)
		{
			$detailParts = @("Protected SQL sync completed.")
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('RunTabUpdated'))
			{
				$detailParts += "RUN updated: $($sqlUpdateResult.RunTabUpdated)."
			}
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('LocalTerminalDB'))
			{
				$detailParts += "Local terminal DB: $($sqlUpdateResult.LocalTerminalDB)."
			}
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('ConfiguredDBSERVER') -and $sqlUpdateResult.ConfiguredDBSERVER)
			{
				$detailParts += "Configured DBSERVER: $($sqlUpdateResult.ConfiguredDBSERVER)."
			}
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('RuntimeDBSERVER') -and $sqlUpdateResult.RuntimeDBSERVER)
			{
				$detailParts += "Runtime SQL target: $($sqlUpdateResult.RuntimeDBSERVER)."
			}
			
			$OperationStatus["SQLDatabaseUpdate"].Status = "Successful"
			$OperationStatus["SQLDatabaseUpdate"].Message = "SQL tables updated successfully after machine sync."
			$OperationStatus["SQLDatabaseUpdate"].Details = ($detailParts -join " ")
		}
		else
		{
			$detailParts = @()
			if ($sqlSyncError) { $detailParts += $sqlSyncError }
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('ConfiguredDBSERVER') -and $sqlUpdateResult.ConfiguredDBSERVER)
			{
				$detailParts += "Configured DBSERVER: $($sqlUpdateResult.ConfiguredDBSERVER)."
			}
			elseif ($configuredDbServer)
			{
				$detailParts += "Configured DBSERVER: $configuredDbServer."
			}
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('RuntimeDBSERVER') -and $sqlUpdateResult.RuntimeDBSERVER)
			{
				$detailParts += "Runtime SQL target: $($sqlUpdateResult.RuntimeDBSERVER)."
			}
			elseif ($runtimeDbServer)
			{
				$detailParts += "Runtime SQL target: $runtimeDbServer."
			}
			
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Failed to update SQL tables after machine sync."
			$OperationStatus["SQLDatabaseUpdate"].Details = ($detailParts -join " ")
		}
	}
	
	return [PSCustomObject]@{
		Success             = ($iniSyncSuccess -and $startupIniSuccess -and $sqlSyncSuccess)
		MachineContext      = $MachineContext
		IniSyncSuccess      = $iniSyncSuccess
		StartupIniSuccess   = $startupIniSuccess
		SqlSyncSuccess      = $sqlSyncSuccess
		StartupIniPath      = $resolvedStartupIniPath
		IniSyncError        = $iniSyncError
		StartupIniError     = $startupIniError
		SqlSyncError        = $sqlSyncError
		StartupIniDetails   = $startupIniResult.Details
		SqlResult           = $sqlUpdateResult
		ConfiguredDbServer  = $configuredDbServer
		RuntimeDbServer     = $runtimeDbServer
	}
}

# ===================================================================================================
#                                       SECTION: Main Script Execution
# ---------------------------------------------------------------------------------------------------
# Description:
#   Orchestrates the execution flow of the script, initializing variables, processing items, and handling user interactions.
# ===================================================================================================

# Get current IP configuration
# Always an array: the WMI path returns a [pscustomobject], and in PowerShell 5.1 indexing a single
# pscustomobject ($x[0]) gives $null - the IP showed "not found" and Configure Network had no adapter.
$currentConfigs = @(Get_Active_IP_Config | Where-Object { $_ })

# Get the store number
$currentStoreNumber = Get_Store_Number_From_INI -UpdateLabel
if (-not $currentStoreNumber) { $currentStoreNumber = $script:FunctionResults['StoreNumber'] }

# Get the store name
Get_Store_Name_From_INI
$storeName = $script:FunctionResults['StoreName']

# Detect the machine role from DBNAME before showing the UI.
$dbNameForRole = $null
try
{
	if (Test-Path $startupIniPath)
	{
		$dbNameLine = Get-Content -LiteralPath $startupIniPath -ErrorAction Stop | Where-Object { $_ -match '^\s*(?i:DBNAME)\s*=' } | Select-Object -First 1
		if ($dbNameLine -and ($dbNameLine -match '^\s*(?i:DBNAME)\s*=\s*(.+?)\s*$'))
		{
			$dbNameForRole = $matches[1].Trim().ToUpper()
			$script:FunctionResults['DBNAME'] = $dbNameForRole
		}
	}
}
catch { }

$script:NodeRole = "Unknown"
if ($dbNameForRole -match '^(?i)LANESQL$')
{
	$script:NodeRole = "Lane"
}
elseif ($dbNameForRole -match '^(?i)STORESQL$')
{
	$script:NodeRole = "StoreServer"
}
elseif ($dbNameForRole -match '^(?i)HOSTSQL$')
{
	$script:NodeRole = "HostServer"
}

# Database connection details remain lazy; DBNAME is read here only for node-role detection.
$connectionString = $null

# Set the old machine name variable
$oldMachineName = $currentMachineName

# Indicate the script has started
Write-Host "Script started" -ForegroundColor Green

# ===================================================================================================
#                                  SECTION: GUI - Main Window
# ---------------------------------------------------------------------------------------------------
# Description:
#   First of the GUI files (src\GUI\01..11 run in order and share script-scope variables): the main
#   $form with its header bar, the "Current Info" tiles, the three action cards the buttons go into
#   (filled by 11_Layout_And_Tooltips) and the footer.
#
# Variables (read by the later GUI files and by functions that update the window):
#   - $form, $script:MainForm, $toolTip, $script:MiniGhostIcon
#   - $script:machineNameLabel, $script:storeNameLabel, $script:storeNumberLabel,
#     $script:ipAddressLabel        value labels of the info tiles (text = the value only)
#   - $script:CardIdentity, $script:CardDatabase, $script:CardSystem   button hosts of the cards
#   - $ui* colours
#
# Notes:
#   - Dialog functions name their own window $form: anything that can run while a dialog is open
#     (handlers, timers) must use $script:MainForm.
#   - Paint handlers only use their sender ($s) and colours stored on .Tag - they run in their
#     own scope.
# ===================================================================================================

# ---- Palette (same family as the TBS Maintenance Script) -----------------------------------------
$uiBg = [System.Drawing.Color]::FromArgb(243, 243, 243)
$uiCard = [System.Drawing.Color]::White
$uiBorder = [System.Drawing.Color]::FromArgb(229, 229, 229)
$uiText = [System.Drawing.Color]::FromArgb(32, 32, 32)
$uiTextMuted = [System.Drawing.Color]::FromArgb(110, 110, 110)
$uiAccentIdentity = [System.Drawing.Color]::FromArgb(0, 103, 192)  # blue
$uiAccentDatabase = [System.Drawing.Color]::FromArgb(126, 87, 214) # violet
$uiAccentSystem = [System.Drawing.Color]::FromArgb(230, 126, 34)   # orange

# ---- Main form -----------------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$script:MainForm = $form
$form.Text = "MiniGhost $VersionNumber"
$form.StartPosition = "CenterScreen"
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$form.MinimumSize = New-Object System.Drawing.Size(820, 560)
$form.Size = New-Object System.Drawing.Size(940, 620)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MaximizeBox = $true
$form.MinimizeBox = $true
$form.BackColor = $uiBg
$form.ForeColor = $uiText
try
{
	$prop = $form.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
	if ($prop) { $prop.SetValue($form, $true, $null) }
}
catch { }

# Window / taskbar icon drawn from the logo
$script:MiniGhostIcon = $null
try
{
	$iconBmp = New-Object System.Drawing.Bitmap(32, 32)
	$iconG = [System.Drawing.Graphics]::FromImage($iconBmp)
	Draw_MiniGhost_Logo -Graphics $iconG -X 1 -Y 1 -Size 30
	$iconG.Dispose()
	$script:MiniGhostIcon = [System.Drawing.Icon]::FromHandle($iconBmp.GetHicon())
	$form.Icon = $script:MiniGhostIcon
}
catch { }

# Global tooltip provider (per-button tooltips)
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 20000
$toolTip.InitialDelay = 350
$toolTip.ReshowDelay = 150
$toolTip.ShowAlways = $true

# ---- Root layout: header / info / actions / footer -----------------------------------------------
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = 'Fill'
$mainLayout.ColumnCount = 1
$mainLayout.RowCount = 4
$mainLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$mainLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$mainLayout.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize
[void]$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 72)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 92)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$form.Controls.Add($mainLayout)

# ---- Header bar ----------------------------------------------------------------------------------
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = 'Fill'
$headerPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 36)
$headerPanel.add_Paint({
		param ($s, $e)
		try
		{
			$rect = New-Object System.Drawing.Rectangle(0, 0, $s.Width, $s.Height)
			$bg = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, [System.Drawing.Color]::FromArgb(26, 26, 32), [System.Drawing.Color]::FromArgb(48, 44, 66), [System.Drawing.Drawing2D.LinearGradientMode]::Horizontal)
			$e.Graphics.FillRectangle($bg, $rect)
			$bg.Dispose()
			$stripe = New-Object System.Drawing.Rectangle(0, 0, 4, $s.Height)
			$sb = New-Object System.Drawing.Drawing2D.LinearGradientBrush($stripe, [System.Drawing.Color]::FromArgb(126, 87, 214), [System.Drawing.Color]::FromArgb(0, 120, 215), [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
			$e.Graphics.FillRectangle($sb, $stripe)
			$sb.Dispose()
		}
		catch { }
	})

$logoPanel = New-Object System.Windows.Forms.Panel
$logoPanel.Location = New-Object System.Drawing.Point(20, 16)
$logoPanel.Size = New-Object System.Drawing.Size(40, 40)
$logoPanel.BackColor = [System.Drawing.Color]::Transparent
$logoPanel.add_Paint({
		param ($s, $e)
		try { Draw_MiniGhost_Logo -Graphics $e.Graphics -X 1 -Y 1 -Size ([float]([Math]::Min($s.Width, $s.Height) - 2)) }
		catch { }
	})
$headerPanel.Controls.Add($logoPanel)

$bannerLabel = New-Object System.Windows.Forms.Label
$bannerLabel.Text = "MiniGhost"
$bannerLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 16)
$bannerLabel.ForeColor = [System.Drawing.Color]::White
$bannerLabel.BackColor = [System.Drawing.Color]::Transparent
$bannerLabel.AutoSize = $true
$bannerLabel.Location = New-Object System.Drawing.Point(70, 10)
$headerPanel.Controls.Add($bannerLabel)

$bannerSubLabel = New-Object System.Windows.Forms.Label
$bannerSubLabel.Text = "Re-identify a cloned TBS machine: name, store, network, INIs and SQL"
$bannerSubLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$bannerSubLabel.ForeColor = [System.Drawing.Color]::FromArgb(185, 185, 200)
$bannerSubLabel.BackColor = [System.Drawing.Color]::Transparent
$bannerSubLabel.AutoSize = $true
$bannerSubLabel.Location = New-Object System.Drawing.Point(72, 42)
$headerPanel.Controls.Add($bannerSubLabel)

# Right side: node role + version pills
$pillPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$pillPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
$pillPanel.WrapContents = $false
$pillPanel.AutoSize = $true
$pillPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$pillPanel.BackColor = [System.Drawing.Color]::Transparent
$pillPanel.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Right'
$pillPanel.Location = New-Object System.Drawing.Point(600, 23)

$versionPill = New-Object System.Windows.Forms.Label
$versionPill.Text = "v$VersionNumber"
$versionPill.AutoSize = $true
$versionPill.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$versionPill.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 230)
$versionPill.BackColor = [System.Drawing.Color]::FromArgb(62, 60, 80)
$versionPill.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
$versionPill.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)

$rolePill = New-Object System.Windows.Forms.Label
$roleText = switch -Regex ([string]$script:NodeRole)
{
	'^(?i)Lane$' { 'LANE'; break }
	'^(?i)StoreServer$' { 'STORE SERVER'; break }
	'^(?i)HostServer$' { 'HOST SERVER'; break }
	default { 'ROLE UNKNOWN' }
}
$rolePill.Text = $roleText
$rolePill.AutoSize = $true
$rolePill.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$rolePill.ForeColor = [System.Drawing.Color]::White
$rolePill.BackColor = switch -Regex ([string]$script:NodeRole)
{
	'^(?i)Lane$' { [System.Drawing.Color]::FromArgb(0, 103, 192); break }
	'^(?i)(StoreServer|HostServer)$' { [System.Drawing.Color]::FromArgb(16, 124, 64); break }
	default { [System.Drawing.Color]::FromArgb(110, 110, 120) }
}
$rolePill.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
$rolePill.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$toolTip.SetToolTip($rolePill, "Node role from Startup.ini DBNAME (LANESQL = lane, STORESQL = store server, HOSTSQL = host server). It decides which machine-name formats are accepted.")

$pillPanel.Controls.Add($versionPill)
$pillPanel.Controls.Add($rolePill)
$headerPanel.Controls.Add($pillPanel)
# keep the pills right-aligned when the window is resized (Anchor alone does not move an AutoSize flow panel reliably)
$headerPanel.add_Resize({
		param ($s, $e)
		try
		{
			foreach ($c in @($s.Controls))
			{
				if ($c -is [System.Windows.Forms.FlowLayoutPanel]) { $c.Left = [Math]::Max(300, $s.ClientSize.Width - $c.Width - 20); $c.Top = [int](($s.ClientSize.Height - $c.Height) / 2) }
			}
		}
		catch { }
	})
$mainLayout.Controls.Add($headerPanel, 0, 0)

# ---- Current Info: four tiles (caption + value) --------------------------------------------------
$infoCard = New_MiniGhost_Card -Margin (New-Object System.Windows.Forms.Padding(16, 14, 16, 0))
$infoCard.Dock = 'Fill'

$infoLayout = New-Object System.Windows.Forms.TableLayoutPanel
$infoLayout.Dock = 'Fill'
$infoLayout.ColumnCount = 4
$infoLayout.RowCount = 2
$infoLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$infoLayout.Padding = New-Object System.Windows.Forms.Padding(8, 10, 8, 8)
$infoLayout.BackColor = [System.Drawing.Color]::White
foreach ($pct in @(25, 25, 22, 28)) { [void]$infoLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, $pct))) }
[void]$infoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))
[void]$infoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$infoCard.Controls.Add($infoLayout)

$currentIP = if ($currentConfigs -and @($currentConfigs).Count -gt 0) { [string](@(@($currentConfigs)[0].IPv4Address)[0].IPAddress) }
else { "Not found" }
if ([string]::IsNullOrWhiteSpace($currentIP)) { $currentIP = "Not found" }

$infoTiles = @(
	@{ Caption = 'MACHINE NAME'; Value = [string]$currentMachineName; Var = 'machineNameLabel' },
	@{ Caption = 'STORE NAME'; Value = [string]$storeName; Var = 'storeNameLabel' },
	@{ Caption = 'STORE NUMBER'; Value = [string]$currentStoreNumber; Var = 'storeNumberLabel' },
	@{ Caption = 'IP ADDRESS'; Value = $currentIP; Var = 'ipAddressLabel' }
)
for ($col = 0; $col -lt $infoTiles.Count; $col++)
{
	$tile = $infoTiles[$col]
	$cap = New-Object System.Windows.Forms.Label
	$cap.Text = $tile.Caption
	$cap.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8)
	$cap.ForeColor = $uiTextMuted
	$cap.Dock = 'Fill'
	$cap.Margin = New-Object System.Windows.Forms.Padding(10, 0, 6, 0)
	$cap.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
	$infoLayout.Controls.Add($cap, $col, 0)
	
	$val = New-Object System.Windows.Forms.Label
	$val.Text = $(if ([string]::IsNullOrWhiteSpace($tile.Value)) { 'N/A' } else { $tile.Value })
	$val.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
	$val.ForeColor = $uiText
	$val.Dock = 'Fill'
	$val.AutoSize = $false
	$val.AutoEllipsis = $true
	$val.Margin = New-Object System.Windows.Forms.Padding(10, 0, 6, 0)
	$val.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
	$infoLayout.Controls.Add($val, $col, 1)
	Set-Variable -Scope Script -Name $tile.Var -Value $val
}
# the full text is in the tooltip when a long value is cut with "..."
foreach ($v in @($script:machineNameLabel, $script:storeNameLabel, $script:storeNumberLabel, $script:ipAddressLabel))
{
	$v.add_TextChanged({ param ($s, $e) try { $toolTip.SetToolTip($s, $s.Text) } catch { } })
	$toolTip.SetToolTip($v, $v.Text)
}
$mainLayout.Controls.Add($infoCard, 0, 1)

# ---- Actions: three cards (Identity / Database / System) -----------------------------------------
$actionsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$actionsLayout.Dock = 'Fill'
$actionsLayout.ColumnCount = 3
$actionsLayout.RowCount = 1
$actionsLayout.Margin = New-Object System.Windows.Forms.Padding(10, 8, 10, 6)
$actionsLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$actionsLayout.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize
foreach ($pct in @(33.33, 33.33, 33.34)) { [void]$actionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, $pct))) }
[void]$actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$mainLayout.Controls.Add($actionsLayout, 0, 2)

$cardIdentityPanel = New_MiniGhost_Action_Card -Title 'Identity' -Subtitle 'Machine name, store number, network' -Accent $uiAccentIdentity
$cardDatabasePanel = New_MiniGhost_Action_Card -Title 'Database' -Subtitle 'SQL tables for this store / node' -Accent $uiAccentDatabase
$cardSystemPanel = New_MiniGhost_Action_Card -Title 'System' -Subtitle 'Registry, results, restart' -Accent $uiAccentSystem
$actionsLayout.Controls.Add($cardIdentityPanel, 0, 0)
$actionsLayout.Controls.Add($cardDatabasePanel, 1, 0)
$actionsLayout.Controls.Add($cardSystemPanel, 2, 0)
$script:CardIdentity = $cardIdentityPanel.Tag
$script:CardDatabase = $cardDatabasePanel.Tag
$script:CardSystem = $cardSystemPanel.Tag

# ---- Footer --------------------------------------------------------------------------------------
$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Dock = 'Fill'
$footerPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$footerPanel.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 238)

$footerLeft = New-Object System.Windows.Forms.Label
$footerLeft.Dock = 'Left'
$footerLeft.AutoSize = $false
$footerLeft.Width = 420
$footerLeft.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$footerLeft.Padding = New-Object System.Windows.Forms.Padding(16, 0, 0, 0)
$footerLeft.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$footerLeft.ForeColor = $uiTextMuted
$footerLeft.Text = "Every action asks before it changes anything."

$footerRight = New-Object System.Windows.Forms.Label
$footerRight.Dock = 'Fill'
$footerRight.AutoSize = $false
$footerRight.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$footerRight.Padding = New-Object System.Windows.Forms.Padding(0, 0, 16, 0)
$footerRight.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$footerRight.ForeColor = $uiTextMuted
$footerRight.Text = "Created by Alex_C.T   |   Version $VersionNumber   |   Revised $VersionDate   |   PowerShell $PowerShellVersion"

$footerPanel.Controls.Add($footerRight)
$footerPanel.Controls.Add($footerLeft)
$mainLayout.Controls.Add($footerPanel, 0, 3)

# ===================================================================================================
#									SECTION: GUI Buttons
# ---------------------------------------------------------------------------------------------------
# Description:
#   Creates and configures buttons on the GUI form for various operations.
# ===================================================================================================

############################################################################
# 1) Change Machine Name Button (UPDATED: supports same-name full sync workflow)
############################################################################
$changeMachineNameButton = New-Object System.Windows.Forms.Button
$changeMachineNameButton.Text = "Change Machine Name"
$changeMachineNameButton.Location = New-Object System.Drawing.Point(10, 120)
$changeMachineNameButton.Size = New-Object System.Drawing.Size(150, 35)
$changeMachineNameButton.Add_Click({
		if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))
		{
			[System.Windows.Forms.MessageBox]::Show(
				"This script must be run as an administrator.",
				"Insufficient Privileges",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			return
		}
		
		$newMachineNameInput = Get_NEW_Machine_Name
		if ($newMachineNameInput -eq $null)
		{
			Set_Operation_Status -OperationStatus $operationStatus -Key "MachineNameChange" -Status "Cancelled" -Message "Machine name change was cancelled by the user." -Details ""
			Show_App_Message -Text "Machine name change was cancelled." -Title "Cancelled" -Icon Information | Out-Null
			return
		}
		
		Invoke_Machine_Name_Change_Workflow -RequestedMachineName $newMachineNameInput -OperationStatus $operationStatus
	})

############################################################################
# 2) Configure Network Button
############################################################################
$configureNetworkButton = New-Object System.Windows.Forms.Button
$configureNetworkButton.Text = "Configure Network"
$configureNetworkButton.Location = New-Object System.Drawing.Point(170, 120)
$configureNetworkButton.Size = New-Object System.Drawing.Size(150, 35)
$configureNetworkButton.Add_Click({
		# Implement the ConfigureNetwork function with GUI elements
		
		if ($currentConfigs -eq $null -or $currentConfigs.Count -eq 0)
		{
			[System.Windows.Forms.MessageBox]::Show("No valid active IP configuration found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			return
		}
		
		# If multiple adapters are found, ask the user to select one
		if ($currentConfigs.Count -gt 1)
		{
			# Create a form to select the network adapter
			$adapterForm = New-Object System.Windows.Forms.Form
			$adapterForm.Text = "Select Network Adapter"
			$adapterForm.Size = New-Object System.Drawing.Size(400, 200)
			$adapterForm.StartPosition = "CenterParent"
			
			$label = New-Object System.Windows.Forms.Label
			$label.Text = "Select the network adapter:"
			$label.Location = New-Object System.Drawing.Point(10, 10)
			$label.Size = New-Object System.Drawing.Size(380, 20)
			$label.AutoSize = $true
			
			$listBox = New-Object System.Windows.Forms.ListBox
			$listBox.Location = New-Object System.Drawing.Point(10, 40)
			$listBox.Size = New-Object System.Drawing.Size(360, 80)
			
			for ($i = 0; $i -lt $currentConfigs.Count; $i++)
			{
				$adapterName = $currentConfigs[$i].InterfaceAlias
				$ipAddress = $currentConfigs[$i].IPv4Address.IPAddress
				$listBox.Items.Add("$adapterName - IP: $ipAddress")
			}
			
			$okButton = New-Object System.Windows.Forms.Button
			$okButton.Text = "OK"
			$okButton.Location = New-Object System.Drawing.Point(150, 130)
			$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
			
			$adapterForm.AcceptButton = $okButton
			$adapterForm.Controls.AddRange(@($label, $listBox, $okButton))
			Set_Modern_Dialog_Style -Form $adapterForm -PrimaryButton $okButton
			$adapterForm.ShowDialog() | Out-Null
			
			$selectedIndex = $listBox.SelectedIndex
			if ($selectedIndex -ge 0 -and $selectedIndex -lt $currentConfigs.Count)
			{
				$currentConfig = $currentConfigs[$selectedIndex]
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("No network adapter selected.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return
			}
		}
		else
		{
			# Only one adapter found
			$currentConfig = $currentConfigs[0]
		}
		
		# Proceed with configuring network
		# Display current IP Address
		# first address / gateway: an adapter can carry several (Split on an array would fail)
		$currentIP = [string](@($currentConfig.IPv4Address)[0].IPAddress)
		$currentGateway = [string](@($currentConfig.IPv4DefaultGateway)[0].NextHop)
		$adapterName = $currentConfig.InterfaceAlias
		
		# Ask user for network type
		$networkTypeResult = [System.Windows.Forms.MessageBox]::Show("Will the adapter use DHCP?", "Network Configuration", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
		
		if ($networkTypeResult -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			# Configure DHCP
			try
			{
				# netsh never throws: check its exit code so a failure is not reported as success
				# ("DHCP is already enabled on this interface" also exits 1 - that one is fine)
				$netshOutput = netsh interface ip set address name="$adapterName" source=dhcp 2>&1
				if ($LASTEXITCODE -ne 0 -and (($netshOutput | Out-String) -notmatch '(?i)already')) { throw ("netsh (address) exit code {0}: {1}" -f $LASTEXITCODE, (($netshOutput | Out-String).Trim())) }
				$netshOutput = netsh interface ip set dns name="$adapterName" source=dhcp 2>&1
				if ($LASTEXITCODE -ne 0 -and (($netshOutput | Out-String) -notmatch '(?i)already')) { throw ("netsh (dns) exit code {0}: {1}" -f $LASTEXITCODE, (($netshOutput | Out-String).Trim())) }
				[System.Windows.Forms.MessageBox]::Show("DHCP configuration applied.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
				$operationStatus["IPConfiguration"].Status = "Successful"
				$operationStatus["IPConfiguration"].Message = "DHCP configuration applied."
				$operationStatus["IPConfiguration"].Details = "Adapter: $adapterName"
			}
			catch
			{
				[System.Windows.Forms.MessageBox]::Show("Failed to configure DHCP: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				$operationStatus["IPConfiguration"].Status = "Failed"
				$operationStatus["IPConfiguration"].Message = "Failed to configure DHCP."
				$operationStatus["IPConfiguration"].Details = "Error: $_"
			}
			
		}
		elseif ($networkTypeResult -eq [System.Windows.Forms.DialogResult]::No)
		{
			# Configure Static IP
			# Loop until valid input or user cancels
			$validInput = $false
			while (-not $validInput)
			{
				# Ask for last octet
				$gatewayParts = $currentGateway.Split('.')
				if ($gatewayParts.Length -lt 4)
				{
					[System.Windows.Forms.MessageBox]::Show("Invalid gateway format: $currentGateway", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
					$operationStatus["IPConfiguration"].Status = "Failed"
					$operationStatus["IPConfiguration"].Message = "Invalid gateway format."
					$operationStatus["IPConfiguration"].Details = "Gateway: $currentGateway"
					return
				}
				$gatewayPrefix = $gatewayParts[0 .. 2] -join '.'
				$ipForm = New-Object System.Windows.Forms.Form
				$ipForm.Text = "Enter Last Octet for Static IP"
				$ipForm.Size = New-Object System.Drawing.Size(350, 170)
				$ipForm.StartPosition = "CenterParent"
				
				$label = New-Object System.Windows.Forms.Label
				$label.Text = "Enter the last octet (1-254) of the static IP address $gatewayPrefix."
				$label.Location = New-Object System.Drawing.Point(10, 20)
				$label.Size = New-Object System.Drawing.Size(320, 40)
				$label.AutoSize = $true
				
				$textBox = New-Object System.Windows.Forms.TextBox
				$textBox.Location = New-Object System.Drawing.Point(10, 70)
				$textBox.Size = New-Object System.Drawing.Size(320, 20)
				
				$okButton = New-Object System.Windows.Forms.Button
				$okButton.Text = "OK"
				$okButton.Location = New-Object System.Drawing.Point(130, 100)
				$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
				
				$ipForm.AcceptButton = $okButton
				
				$ipForm.Controls.AddRange(@($label, $textBox, $okButton))
				Set_Modern_Dialog_Style -Form $ipForm -PrimaryButton $okButton
				$dialogResult = $ipForm.ShowDialog()
				
				if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK)
				{
					$lastOctet = $textBox.Text.Trim()
					# Validate last octet
					if ($lastOctet -match '^\d{1,3}$' -and [int]$lastOctet -ge 1 -and [int]$lastOctet -le 254)
					{
						$ipAddress = "$gatewayPrefix.$lastOctet"
						
						# Check if IP is in use
						$pingResult = Test-Connection -ComputerName $ipAddress -Count 1 -Quiet
						if ($pingResult)
						{
							[System.Windows.Forms.MessageBox]::Show("The IP address '$ipAddress' is already in use. Please choose a different one.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
						}
						else
						{
							try
							{
								# Keep the adapter's current subnet size (fall back to /24 when it is unknown)
								$subnetMask = '255.255.255.0'
								$currentPrefix = 0
								if ([int]::TryParse([string](@($currentConfig.IPv4Address)[0].PrefixLength), [ref]$currentPrefix) -and $currentPrefix -ge 8 -and $currentPrefix -le 30)
								{
									$subnetMask = Convert_Prefix_Length_To_Subnet_Mask -PrefixLength $currentPrefix
								}
								
								# netsh never throws: check its exit code so a failure is not reported as success
								$netshOutput = netsh interface ip set address name="$adapterName" static $ipAddress $subnetMask $currentGateway 2>&1
								if ($LASTEXITCODE -ne 0) { throw ("netsh exit code {0}: {1}" -f $LASTEXITCODE, (($netshOutput | Out-String).Trim())) }
								
								# Update the label
								$ipAddressLabel.Text = "$currentIP -> $ipAddress"
								$operationStatus["IPConfiguration"].Status = "Successful"
								$operationStatus["IPConfiguration"].Message = "Static IP configuration applied."
								$operationStatus["IPConfiguration"].Details = "Adapter: $adapterName, IP: $ipAddress"
								
								# Set DNS
								Set-DnsClientServerAddress -InterfaceAlias $adapterName -ServerAddresses ("8.8.8.8", "8.8.4.4")
								[System.Windows.Forms.MessageBox]::Show("Static IP configuration applied.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
								$operationStatus["IPConfiguration"].Status = "Successful"
								$operationStatus["IPConfiguration"].Message = "Static IP configuration applied."
								$operationStatus["IPConfiguration"].Details = "Adapter: $adapterName, IP: $ipAddress"
								$validInput = $true
							}
							catch
							{
								[System.Windows.Forms.MessageBox]::Show("Failed to configure static IP: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
								$operationStatus["IPConfiguration"].Status = "Failed"
								$operationStatus["IPConfiguration"].Message = "Failed to configure static IP."
								$operationStatus["IPConfiguration"].Details = "Error: $_"
								$validInput = $true # Exit loop since an error occurred
							}
						}
					}
					else
					{
						[System.Windows.Forms.MessageBox]::Show("Invalid input. Please enter a valid number between 1 and 254.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
					}
				}
				else
				{
					# User canceled the dialog
					$operationStatus["IPConfiguration"].Status = "Skipped"
					$operationStatus["IPConfiguration"].Message = "User canceled static IP configuration."
					$operationStatus["IPConfiguration"].Details = "" # Empty details
					$validInput = $true # Exit loop
				}
			}
			
		}
		else
		{
			$operationStatus["IPConfiguration"].Status = "Skipped"
			$operationStatus["IPConfiguration"].Message = "User chose not to configure network settings."
		}
	})

############################################################################
# 3) Update Store Number Button (UPDATED: always runs Update_INIs using latest intended MachineName)
############################################################################
$updateStoreNumberButton = New-Object System.Windows.Forms.Button
$updateStoreNumberButton.Text = "Update Store Number"
$updateStoreNumberButton.Location = New-Object System.Drawing.Point(330, 120)
$updateStoreNumberButton.Size = New-Object System.Drawing.Size(150, 35)
$updateStoreNumberButton.Add_Click({
		
		# Get old store number
		$oldStoreNumber = Get_Store_Number_From_INI
		
		if ($oldStoreNumber -eq $null)
		{
			Show_App_Message -Text "Store number not found in startup.ini." -Title "Error" -Icon Error | Out-Null
			Set_Operation_Status -OperationStatus $operationStatus -Key "StoreNumberChange" -Status "Failed" -Message "Store number not found." -Details "startup.ini not found or store number not defined."
			return
		}
		
		$newStoreNumberInput = Get_NEW_Store_Number
		if ($newStoreNumberInput -eq $null) { return }
		$newStoreNumberInput = ([string]$newStoreNumberInput).Trim()
		if ($newStoreNumberInput -notmatch '^\d{3,4}$')
		{
			Show_App_Message -Text "The new store number input was invalid. Please enter exactly 3 or 4 digits." -Title "Invalid Store Number" -Icon Error | Out-Null
			Set_Operation_Status -OperationStatus $operationStatus -Key "StoreNumberChange" -Status "Failed" -Message "Invalid store number input." -Details "Received value: '$newStoreNumberInput'."
			return
		}
		
		$warningResult = Show_App_Message -Text ("You are about to change the store number from '{0}' to '{1}'. Do you want to proceed?" -f $oldStoreNumber, $newStoreNumberInput) -Title "Warning" -Buttons YesNo -Icon Warning
		
		if ($warningResult -ne [System.Windows.Forms.DialogResult]::Yes)
		{
			Set_Operation_Status -OperationStatus $operationStatus -Key "StoreNumberChange" -Status "Cancelled" -Message "Store number change was cancelled by the user." -Details "Old store number remains '$oldStoreNumber'."
			return
		}
		
		# Prefer the intended machine name if one is already pending from the rename workflow.
		$effectiveMachineName = Get_Effective_Machine_Name
		
		# Always run Update_INIs (best-effort) so INIs are consistent even if other steps later fail
		$iniOk = $false
		try
		{
			$iniOk = Update_INIs -newStoreNumber $newStoreNumberInput -MachineName $effectiveMachineName -OldStoreNumber $oldStoreNumber -OldMachineName $env:COMPUTERNAME -CreateSmsHttpsIfMissing
		}
		catch
		{
			$iniOk = $false
		}
		
		if ($iniOk)
		{
			$script:newStoreNumber = $newStoreNumberInput
			Set_Label_Text -Label $storeNumberLabel -Text "$oldStoreNumber -> $script:newStoreNumber"
			Set_Operation_Status -OperationStatus $operationStatus -Key "StoreNumberChange" -Status "Successful" -Message "Store number updated in INIs." -Details "Store number changed to '$script:newStoreNumber'."
			Show_App_Message -Text "Store number successfully changed to '$script:newStoreNumber'." -Title "Store Number Updated" -Icon Information | Out-Null
		}
		else
		{
			Set_Operation_Status -OperationStatus $operationStatus -Key "StoreNumberChange" -Status "Failed" -Message "Update_INIs failed." -Details "Best-effort update encountered an error."
			Show_App_Message -Text "Update_INIs failed (best-effort). Check console/log output." -Title "Error" -Icon Error | Out-Null
		}
		
		# SQL update (keep your existing logic)
		$sqlUpdateResult = Update_SQL_Tables_For_Store_Number_Change -storeNumber $newStoreNumberInput
		
		if ($sqlUpdateResult.Success)
		{
			$operationStatus["SQLDatabaseUpdate"].Status = "Successful"
			$operationStatus["SQLDatabaseUpdate"].Message = "STD_TAB updated successfully after store number change."
			$operationStatus["SQLDatabaseUpdate"].Details = "STD_TAB updated with new store number."
		}
		else
		{
			$operationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$operationStatus["SQLDatabaseUpdate"].Message = "Failed to update STD_TAB after store number change."
			$operationStatus["SQLDatabaseUpdate"].Details = "Failed commands: $($sqlUpdateResult.FailedCommands -join ', ')"
			
			# The INI result was already shown; without this the SQL failure was only visible in Summary
			$sqlFailureLog = (@($sqlUpdateResult.FailedCommands) -join "`r`n`r`n")
			if ([string]::IsNullOrWhiteSpace($sqlFailureLog)) { $sqlFailureLog = "SQL update result indicated failure." }
			[void](Show_Action_Prompt_With_Log -Text "The store number SQL update (STD_TAB / TER_TAB 901) failed.`r`n`r`nUse View Log to see the failed commands, then run Update SQL Database or fix the database connection." -Title "SQL Update Failed" -PrimaryButtonText "OK" -LogText $sqlFailureLog -LogTitle "Failed SQL Commands")
		}
		
		<# Run Update_INIs again at the end so INIs always reflect the latest state
		try
		{
			$null = Update_INIs -newStoreNumber $newStoreNumberInput -MachineName $effectiveMachineName -OldStoreNumber $oldStoreNumber -OldMachineName $env:COMPUTERNAME -CreateSmsHttpsIfMissing
		}
		catch { }#>
	})

############################################################################
# 4) Truncate Tables Button
############################################################################
$truncateTablesButton = New-Object System.Windows.Forms.Button
$truncateTablesButton.Text = "Truncate Tables"
$truncateTablesButton.Location = New-Object System.Drawing.Point(10, 160)
$truncateTablesButton.Size = New-Object System.Drawing.Size(150, 35)
$truncateTablesButton.Add_Click({
		# Define the tables to truncate
		$tablesToTruncate = @(
			"TRS_LOG", "TRS_FIN", "TRS_DPT", "TRS_SUB", "TRS_CLK",
			"TRS_CLT", "TRS_VND", "TRS_ITM",
			"RPT_CLT_D", "RPT_CLT_W", "RPT_CLT_M", "RPT_CLT_N", "RPT_CLT_Y", "RPT_CLT_P", "RPT_CLT_F", "RPT_CLT_ITM_D", "RPT_CLT_ITM_N",
			"RPT_FIN", "RPT_DPT", "RPT_SUB", "RPT_HOU", "RPT_VND",
			"COST_REV", "POS_REV", "OBJ_REV", "PRICE_REV", "REV_HDR", "SAL_REG_SAV", "SAL_HDR_SAV", "SAL_TTL_SAV", "SAL_DET_SAV",
			"RPT_CLK_D", "RPT_CLK_W", "RPT_CLK_M", "RPT_CLK_Y", "RPT_CLK_P", "RPT_CLK_F", "RPT_CLK_N",
			"RPT_ITM_D", "RPT_ITM_W", "RPT_ITM_M", "RPT_ITM_Y", "RPT_ITM_P", "RPT_ITM_F", "RPT_ITM_N",
			"DATA_REG",
			"RPT_CHK", "RENT_TAB", "TIM_TAB", "GAS_INVENT", "GAS_COUNT", "GAS_TRANS",
			"SAL_BAT", "SAL_HDR", "SAL_REG", "SAL_DET", "SAL_TTL",
			"REC_BAT", "REC_HDR", "REC_REG", "REC_TTL", "INV_HDR", "INV_REG", "INV_TTL",
			"RPT_PAY_N", "RPT_PAY_M", "TRS_PAY", "OFR_TAB"
		)
		
		$result = [System.Windows.Forms.MessageBox]::Show("Do you want to truncate the specified tables?", "Truncate Tables", [System.Windows.Forms.MessageBoxButtons]::YesNo)
		if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			# Truncate tables
			$failedTruncateTables = RPT_Blank -tables $tablesToTruncate
			
			if ($failedTruncateTables.Count -eq 0)
			{
				[System.Windows.Forms.MessageBox]::Show("Tables truncated successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
				$operationStatus["TableTruncation"].Status = "Successful"
				$operationStatus["TableTruncation"].Message = "Tables truncated successfully."
				$operationStatus["TableTruncation"].Details = "All tables were truncated successfully."
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("Failed to truncate some tables: $($failedTruncateTables -join ', ')", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				$operationStatus["TableTruncation"].Status = "Failed"
				$operationStatus["TableTruncation"].Message = "Failed to truncate some tables."
				$operationStatus["TableTruncation"].Details = "Failed tables: $($failedTruncateTables -join ', ')"
			}
		}
		else
		{
			$operationStatus["TableTruncation"].Status = "Skipped"
			$operationStatus["TableTruncation"].Message = "User chose not to truncate tables."
		}
	})

############################################################################
# 6) Registry Cleanup Button
############################################################################
$registryCleanupButton = New-Object System.Windows.Forms.Button
$registryCleanupButton.Text = "Registry Cleanup"
$registryCleanupButton.Location = New-Object System.Drawing.Point(330, 160)
$registryCleanupButton.Size = New-Object System.Drawing.Size(150, 35)
$registryCleanupButton.Add_Click({
		$result = [System.Windows.Forms.MessageBox]::Show("Do you want to delete all registry values starting with 'GT'?", "Registry Cleanup", [System.Windows.Forms.MessageBoxButtons]::YesNo)
		if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			$gtRegistryCleanupResult = Remove_GT_Registry_Values
			
			$registryStatus = $gtRegistryCleanupResult.Status
			
			if ($registryStatus -eq 'Successful')
			{
				$operationStatus["RegistryCleanup"].Status = "Successful"
				$operationStatus["RegistryCleanup"].Message = "GT registry values removed successfully."
				$operationStatus["RegistryCleanup"].Details = "$($gtRegistryCleanupResult.DeletedCount) 'GT' registry keys were deleted."
				[System.Windows.Forms.MessageBox]::Show("GT registry values removed successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			}
			elseif ($registryStatus -eq 'Skipped')
			{
				$operationStatus["RegistryCleanup"].Status = "Skipped"
				$operationStatus["RegistryCleanup"].Message = "User chose not to delete GT registry values."
				$operationStatus["RegistryCleanup"].Details = ""
			}
			else
			{
				$operationStatus["RegistryCleanup"].Status = "Failed"
				$operationStatus["RegistryCleanup"].Message = "Failed to remove GT registry values."
				$operationStatus["RegistryCleanup"].Details = $gtRegistryCleanupResult.Message
				[System.Windows.Forms.MessageBox]::Show("Failed to remove GT registry values.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			}
		}
		else
		{
			$operationStatus["RegistryCleanup"].Status = "Skipped"
			$operationStatus["RegistryCleanup"].Message = "User chose not to delete GT registry values."
		}
	})

############################################################################
# 9) Update Database Button (simple click call)
############################################################################
$updateSQLDatabaseButton = New-Object System.Windows.Forms.Button
$updateSQLDatabaseButton.Text = "Update SQL Database"
$updateSQLDatabaseButton.Location = New-Object System.Drawing.Point(10, 240)
$updateSQLDatabaseButton.Size = New-Object System.Drawing.Size(150, 35)

$updateSQLDatabaseButton.Add_Click({
		Update_SQL_Database `
								 -CurrentMachineName $currentMachineName `
								 -NewMachineName $script:newMachineName `
								 -OperationStatus $operationStatus
	})

############################################################################
# 10) Summary Button
############################################################################
$summaryButton = New-Object System.Windows.Forms.Button
$summaryButton.Text = "Show Summary"
$summaryButton.Location = New-Object System.Drawing.Point(170, 240)
$summaryButton.Size = New-Object System.Drawing.Size(150, 35)
$summaryButton.Add_Click({
		# Display the summary in a new form
		# Build the summary text
		$summaryText = ""
		foreach ($operationKey in $operationStatus.Keys)
		{
			$operation = $operationStatus[$operationKey]
			$status = $operation.Status
			$message = $operation.Message
			$details = $operation.Details
			
			$summaryText += "${operationKey}: $status`r`n"
			if ($message -ne "")
			{
				$summaryText += "  Message: $message`r`n"
			}
			if ($details -ne "")
			{
				$summaryText += "  Details: $details`r`n"
			}
			$summaryText += "`r`n"
		}
		
		Show_Text_Report_Form -Title "Operation Summary" -Text $summaryText -Width 600 -Height 400
	})

############################################################################
# 11) Reboot Button
############################################################################
$rebootButton = New-Object System.Windows.Forms.Button
$rebootButton.Text = "Reboot System"
$rebootButton.Location = New-Object System.Drawing.Point(330, 240)
$rebootButton.Size = New-Object System.Drawing.Size(150, 35)
$rebootButton.Add_Click({
		$rebootResult = [System.Windows.Forms.MessageBox]::Show("Do you want to reboot now?", "Reboot", [System.Windows.Forms.MessageBoxButtons]::YesNo)
		if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			# Clean the temp copies first: nothing after Restart-Computer -Force is guaranteed to run
			[void](Remove_MiniGhost_Temp_Copies -TempPath $TempDir)
			Restart-Computer -Force
		}
	})

############################################################################
# Handle form closing event (X button)
############################################################################
# Closes straight away (no "Are you sure?" prompt since v1.4.1): every action already asks before it
# changes anything, and nothing runs in the background that closing could cut off.
$form.add_FormClosing({
		# Clean the temp copies of MiniGhost
		[void](Remove_MiniGhost_Temp_Copies -TempPath $TempDir)
	})

# ===================================================================================================
#                              SECTION: GUI - Button Layout and Tooltips
# ---------------------------------------------------------------------------------------------------
# Description:
#   Styles the eight action buttons (created with their click handlers in 02..09) and places them
#   in the three cards of 01_Main_Form, then sets their tooltips.
# ===================================================================================================

# Card -> buttons, top to bottom; Danger buttons get red text
$actionGroups = @(
	@{ Host = $script:CardIdentity; Accent = $uiAccentIdentity; Buttons = @($changeMachineNameButton, $updateStoreNumberButton, $configureNetworkButton) },
	@{ Host = $script:CardDatabase; Accent = $uiAccentDatabase; Buttons = @($updateSQLDatabaseButton, $truncateTablesButton) },
	@{ Host = $script:CardSystem; Accent = $uiAccentSystem; Buttons = @($registryCleanupButton, $summaryButton, $rebootButton) }
)
$dangerButtons = @($truncateTablesButton, $rebootButton)

foreach ($group in $actionGroups)
{
	# Dock=Top stacks the LAST added control at the top, so add bottom-up
	$ordered = @($group.Buttons)
	[array]::Reverse($ordered)
	foreach ($btn in $ordered)
	{
		$kind = if ($dangerButtons -contains $btn) { 'Danger' } else { 'Action' }
		Set_Modern_Button_Style -Button $btn -Kind $kind -AccentColor $group.Accent
		$btn.Dock = 'Fill'
		$btn.Margin = New-Object System.Windows.Forms.Padding(0)
		$btn.MinimumSize = New-Object System.Drawing.Size(0, 0)
		
		# wrapper = button + 8 px gap below it (docked controls ignore Margin)
		$slot = New-Object System.Windows.Forms.Panel
		$slot.Dock = 'Top'
		$slot.Height = 54
		$slot.BackColor = [System.Drawing.Color]::White
		$slot.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
		$slot.Controls.Add($btn)
		$group.Host.Controls.Add($slot)
	}
}

# Tooltips (one per button)
$toolTip.SetToolTip($updateStoreNumberButton, "Change the store number in Startup.ini, Server.ini, SMSStart.ini, System.ini, INFO_*901_WIN.INI and SmsHttps.INI, then update STD_TAB / TER_TAB 901.")
$toolTip.SetToolTip($configureNetworkButton, "Switch the active network adapter to DHCP, or to a static IP in the gateway's subnet (the address is pinged first).")
$changeMachineNameToolTipText = "Rename this computer or re-enter the current machine name to force a full INI + SQL sync."
if ($script:NodeRole -match '^(?i)Lane$')
{
	$changeMachineNameToolTipText += " Lane formats support POS001, POS-001, LN026001, 0384LANE001, or just 001 to reuse the current style."
}
elseif ($script:NodeRole -match '^(?i)(StoreServer|HostServer)$')
{
	$changeMachineNameToolTipText += " Server formats support SERVER001, SERVER-001, or 0231SERVER001."
}
else
{
	$changeMachineNameToolTipText += " Accepted formats depend on node role and are shown in the rename dialog."
}
$toolTip.SetToolTip($changeMachineNameButton, $changeMachineNameToolTipText)
$toolTip.SetToolTip($truncateTablesButton, "Empty the transaction, report and sales tables (TRS_*, RPT_*, SAL_*, REC_*, INV_* ...) - asks first. Cannot be undone.")
$toolTip.SetToolTip($registryCleanupButton, "Remove the 'GT*' counter values under Store Management\Counters (after confirmation).")
$toolTip.SetToolTip($updateSQLDatabaseButton, "Bring TER_TAB, STO_TAB, LNK_TAB, RUN_TAB and STD_TAB in line with this machine's name and store number.")
$toolTip.SetToolTip($summaryButton, "See which operations of this session succeeded, failed or were skipped.")
$toolTip.SetToolTip($rebootButton, "Restart the computer (asks first). Needed after a machine rename.")

# ===================================================================================================
#                                       SECTION: Show the Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Displays the main GUI form to the user.
# ===================================================================================================

[void]$form.ShowDialog()

# Indicate the script is closing
Write-Host "Script closing..." -ForegroundColor Yellow

# Close the console to aviod duplicate logging to the richbox
exit

#######################################################################################################
#
#                                     CHANGELOG
#
#######################################################################################################

# ===================================================================================================
#  v1.4.2  -  2026-10-30
#  Antivirus: no in-memory run, no self-copy, no hidden window
# ===================================================================================================
#
#  [CHG]  The bit.ly launcher (MiniGhost_V38.ps1) saves MiniGhost to %TEMP%\MiniGhost.ps1 and starts
#         that file as administrator in a minimized window - it no longer runs downloaded code with
#         Invoke-Expression (Defender blocks that as Trojan:Win32/Commando).
#  [CHG]  MiniGhost no longer writes its own code to a new .ps1 and relaunches it elevated (an
#         inline irm | iex run without admin rights now says to start it as administrator).
#  [CHG]  The console is no longer hidden (no user32 ShowWindow, no C# compiled at run time); the
#         launcher and the elevated relaunch start it minimized instead.

# ===================================================================================================
#  v1.4.1  -  2026-10-30
#  Closes without asking
# ===================================================================================================
#
#  [CHG]  Closing the window no longer asks "Are you sure you want to exit?" - it closes at once
#         (and still removes MiniGhost's temp copies).

# ===================================================================================================
#  v1.4.0  -  2026-10-30
#  New look: same buttons and behaviour, modern window and dialogs
# ===================================================================================================
#
#  [CHG]  Main window: dark header with the MiniGhost logo, node-role and version badges; the
#         current machine name, store name, store number and IP as four tiles; the eight actions
#         in three cards (Identity / Database / System) - Truncate Tables and Reboot in red; a
#         footer with author, version and PowerShell version. Window and taskbar icon.
#  [CHG]  Dialogs (store number, machine name, network, prompts, reports) use the same flat style
#         with the default button highlighted. The failed-SQL list uses the report window.
#  [CHG]  Tile values show the value only ("POS005 -> POS006 (after reboot)"); a cut-off value is
#         shown in full on hover. Button tooltips say what each action changes.
#  [FIX]  The current IP showed "not found" and Configure Network found no adapter whenever the
#         WMI lookup answered (since v1.3.9 that is the normal case): PowerShell 5.1 cannot index
#         a single [pscustomobject]. The adapter list is always an array now.
#  [CHG]  No Set-ExecutionPolicy call (it changed nothing in MiniGhost and is a pattern antivirus
#         heuristics look for). Every build is scanned by Microsoft Defender in CI.
#  [NEW]  tests\Smoke.ps1 runs MiniGhost off-screen (read-only start) and saves screenshots of
#         the window and dialogs; CI runs it on every push.

# ===================================================================================================
#  v1.3.9  -  2026-10-30
#  Review pass: database connection, network, INI rewrites, dialogs, temp cleanup
# ===================================================================================================
#
#  Fixes
#  [FIX]  Truncate Tables, Update SQL Database and the store-number SQL step could not connect on a
#         fresh launch ("Database connection string is not available"): the connection lookup read
#         its own empty parameter instead of the detected Startup.ini. It finds Startup.ini now.
#  [FIX]  Rename with INI/SQL errors -> "Reboot Now" did nothing: the dialog always answered
#         "Closed". It returns the button that was clicked.
#  [FIX]  Store number change rewrote IP addresses whose octet equalled the old store number
#         (store 100: 10.0.100.5 -> 10.0.242.5), and a TIMESYNC SERVER= IP address got its first
#         octet replaced by the store number. IP addresses are left alone.
#  [FIX]  Network detection (WMI path) matched the route's interface index against the adapter's
#         DeviceID, which numbers adapters differently - it could show another adapter's IP. The
#         subnet prefix length was also miscalculated (33 for 255.255.255.0).
#  [FIX]  Configure Network reported success when netsh failed (native commands never throw); the
#         exit code is checked. A static IP keeps the adapter's current subnet size instead of
#         always /24.
#  [FIX]  A failed store-number SQL update was only visible in Summary; it is shown with the failed
#         commands.
#  [FIX]  Old X*/Z* folder cleanup looks in the detected storeman Office folder first (it only knew
#         \\localhost\storeman, C:\storeman and D:\storeman).
#  [FIX]  Store name: only the first NAME= line, and names containing '=' are kept whole.
#
#  Faster
#  [PERF] Closing MiniGhost no longer searches all of %TEMP% recursively; it deletes only its own
#         copies at the top of %TEMP% - including the _Elevated copy the old cleanup never matched.
#         Reboot System cleans up before restarting (nothing after Restart-Computer is sure to run).
#
#  Code
#  [CHG]  Removed the unused Delete_Files_And_Folders (280 lines) and a no-op SmsHttps block.
#  [NEW]  tests\MiniGhost_SelfTest.ps1 and a CI workflow that builds and tests on every push.

# ===================================================================================================
#  v1.3.8  -  2026-04-15
#  Sources split into one file per function (MiniGhost\src) with a build step - no behaviour change
# ===================================================================================================
#
#  [CHG]  MiniGhost.ps1 is now built from MiniGhost\src by MiniGhost\build\Build.ps1 (same code,
#         functions grouped by category). Delete_Files/Folders is named Delete_Files_And_Folders.

# ===================================================================================================
#                              FUNCTION: Export_VNC_Files_For_All_Nodes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generates UltraVNC (.vnc) connection files for every lane, scale, and backoffice discovered
#   in the store, saved under Desktop\Lanes, Desktop\Scales, Desktop\Backoffices. Default VNC
#   password is 4330df922eb03b6e (UltraVNC encrypted form) unless a per-machine password is
#   available; Ishida scales also get an Ishida_Wrapper_#.vnc. Scales with missing IP info are
#   skipped and logged.
#
# Parameters:
#   - LaneNumToMachineName       [hashtable] (required): LaneNumber => MachineName
#   - ScaleCodeToIPInfo          [hashtable] (required): ScaleCode => scale object (with IP)
#   - BackofficeNumToMachineName [hashtable] (required): BackofficeTerminal => MachineName
#   - AllVNCPasswords            [hashtable]: MachineName => VNC password
#   - RunInBackground            [switch]: run as a background job
#   - SilentBackground           [switch]: suppress log window output
#   - InternalBackgroundJob      [switch]: internal flag - already inside a background job
# ===================================================================================================

function Export_VNC_Files_For_All_Nodes
{
	param (
		[Parameter(Mandatory = $true)]
		[hashtable]$LaneNumToMachineName,
		[Parameter(Mandatory = $true)]
		[hashtable]$ScaleCodeToIPInfo,
		[Parameter(Mandatory = $true)]
		[hashtable]$BackofficeNumToMachineName,
		[Parameter(Mandatory = $false)]
		[hashtable]$AllVNCPasswords,
		[Parameter(Mandatory = $false)]
		[switch]$RunInBackground,
		[Parameter(Mandatory = $false)]
		[switch]$SilentBackground,
		[Parameter(Mandatory = $false)]
		[switch]$InternalBackgroundJob
	)

	$log = {
		param (
			[string]$Message,
			[string]$Color = "Black"
		)

		if (-not $SilentBackground)
		{
			Write_Log $Message $Color
		}
	}

	$validLaneMap = @{ }
	if ($LaneNumToMachineName)
	{
		foreach ($kv in $LaneNumToMachineName.GetEnumerator())
		{
			$machineName = [string]$kv.Value
			if (-not [string]::IsNullOrWhiteSpace($machineName))
			{
				$validLaneMap[[string]$kv.Key] = $machineName.Trim()
			}
		}
	}

	$validScaleMap = @{ }
	if ($ScaleCodeToIPInfo)
	{
		foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
		{
			$scaleObj = $kv.Value
			$ip = $null
			if ($null -ne $scaleObj)
			{
				if (($scaleObj.PSObject.Properties.Name -contains 'FullIP') -and -not [string]::IsNullOrWhiteSpace([string]$scaleObj.FullIP))
				{
					$ip = [string]$scaleObj.FullIP
				}
				elseif (
					($scaleObj.PSObject.Properties.Name -contains 'IPNetwork') -and
					($scaleObj.PSObject.Properties.Name -contains 'IPDevice') -and
					-not [string]::IsNullOrWhiteSpace([string]$scaleObj.IPNetwork) -and
					-not [string]::IsNullOrWhiteSpace([string]$scaleObj.IPDevice)
				)
				{
					$ip = "$($scaleObj.IPNetwork)$($scaleObj.IPDevice)"
				}
			}

			if (-not [string]::IsNullOrWhiteSpace($ip))
			{
				$validScaleMap[[string]$kv.Key] = $scaleObj
			}
		}
	}

	$validBackofficeMap = @{ }
	if ($BackofficeNumToMachineName)
	{
		foreach ($kv in $BackofficeNumToMachineName.GetEnumerator())
		{
			$machineName = [string]$kv.Value
			if (-not [string]::IsNullOrWhiteSpace($machineName))
			{
				$validBackofficeMap[[string]$kv.Key] = $machineName.Trim()
			}
		}
	}

	$LaneNumToMachineName = $validLaneMap
	$ScaleCodeToIPInfo = $validScaleMap
	$BackofficeNumToMachineName = $validBackofficeMap

	if ($RunInBackground -and -not $InternalBackgroundJob)
	{
		$hasAnyNodes = (
			($LaneNumToMachineName -and $LaneNumToMachineName.Count -gt 0) -or
			($ScaleCodeToIPInfo -and $ScaleCodeToIPInfo.Count -gt 0) -or
			($BackofficeNumToMachineName -and $BackofficeNumToMachineName.Count -gt 0)
		)

		if (-not $hasAnyNodes) { return }

		if (
			$script:VncExportTask -and
			$script:VncExportTask.Job -and
			$script:VncExportTask.Job.State -in @('Running', 'NotStarted')
		)
		{
			return
		}

		try
		{
			$writeLogDefinition = 'param([string]$Message, $Color, [switch]$IncludeTimestamp)'
			$getAllVncPasswordsDefinition = ${function:Get_All_VNC_Passwords}.ToString()
			$exportVncDefinition = ${function:Export_VNC_Files_For_All_Nodes}.ToString()
			# Get_All_VNC_Passwords probes the nodes with Test-Hosts_Parallel - the job needs it too
			# (missing before v2.7.0: the background export failed whenever it had to look up passwords)
			$testHostsDefinition = ${function:Test-Hosts_Parallel}.ToString()

			$job = Start-Job -Name 'VNCExportJob' -ScriptBlock {
				param (
					[string]$WriteLogDefinition,
					[string]$GetAllVncPasswordsDefinition,
					[string]$ExportVncDefinition,
					[string]$TestHostsDefinition,
					[hashtable]$LaneMap,
					[hashtable]$ScaleMap,
					[hashtable]$BackofficeMap,
					[hashtable]$AllVncPasswords,
					[string]$UltraVncFolder,
					[int]$AnsiCodePage
				)

				Add-Type -AssemblyName System.Windows.Forms 2>$null
				Add-Type -AssemblyName System.Drawing 2>$null

				Set-Item -Path function:Write_Log -Value ([scriptblock]::Create($WriteLogDefinition))
				Set-Item -Path function:Get_All_VNC_Passwords -Value ([scriptblock]::Create($GetAllVncPasswordsDefinition))
				Set-Item -Path function:Export_VNC_Files_For_All_Nodes -Value ([scriptblock]::Create($ExportVncDefinition))
				Set-Item -Path function:Test-Hosts_Parallel -Value ([scriptblock]::Create($TestHostsDefinition))

				$script:UltraVncFolder = $UltraVncFolder
				$script:ansiPcEncoding = [System.Text.Encoding]::GetEncoding($AnsiCodePage)
				$script:FunctionResults = @{ }

				try
				{
					Export_VNC_Files_For_All_Nodes `
												   -LaneNumToMachineName $LaneMap `
												   -ScaleCodeToIPInfo $ScaleMap `
												   -BackofficeNumToMachineName $BackofficeMap `
												   -AllVNCPasswords $AllVncPasswords `
												   -SilentBackground `
												   -InternalBackgroundJob

					$passwordCache = $null
					if ($script:FunctionResults.ContainsKey('AllVNCPasswords'))
					{
						$passwordCache = $script:FunctionResults['AllVNCPasswords']
					}

					[pscustomobject]@{
						Success		    = $true
						AllVNCPasswords = $passwordCache
						Error		    = $null
					}
				}
				catch
				{
					[pscustomobject]@{
						Success		    = $false
						AllVNCPasswords = $null
						Error		    = $_.Exception.Message
					}
				}
			} -ArgumentList @(
				$writeLogDefinition,
				$getAllVncPasswordsDefinition,
				$exportVncDefinition,
				$testHostsDefinition,
				$LaneNumToMachineName,
				$ScaleCodeToIPInfo,
				$BackofficeNumToMachineName,
				$AllVNCPasswords,
				$script:UltraVncFolder,
				1252
			)

			$script:VncExportTask = @{
				Job	    = $job
				Started = Get-Date
			}

			if (-not $script:VncExportReaper)
			{
				$script:VncExportReaper = New-Object System.Windows.Forms.Timer
				$script:VncExportReaper.Interval = 1000
				$script:VncExportReaper.Add_Tick({
						if (-not $script:VncExportTask)
						{
							try
							{
								if ($script:VncExportReaper)
								{
									$script:VncExportReaper.Stop()
									$script:VncExportReaper.Dispose()
									$script:VncExportReaper = $null
								}
							}
							catch { }
							return
						}

						$taskState = $script:VncExportTask
						if (-not $taskState.Job -or $taskState.Job.State -in @('Running', 'NotStarted')) { return }

						$result = $null
						try { $result = Receive-Job -Job $taskState.Job -ErrorAction SilentlyContinue }
						catch { }
						try { Remove-Job -Job $taskState.Job -Force -ErrorAction SilentlyContinue | Out-Null }
						catch { }

						$resultObj = $null
						if ($result)
						{
							$resultObj = @($result) | Select-Object -Last 1
						}

						if (
							$resultObj -and
							$resultObj.Success -and
							$resultObj.AllVNCPasswords -and
							($resultObj.AllVNCPasswords -is [hashtable]) -and
							$resultObj.AllVNCPasswords.Count -gt 0
						)
						{
							$script:FunctionResults['AllVNCPasswords'] = $resultObj.AllVNCPasswords
						}

						$script:VncExportTask = $null

						try
						{
							if ($script:VncExportReaper)
							{
								$script:VncExportReaper.Stop()
								$script:VncExportReaper.Dispose()
								$script:VncExportReaper = $null
							}
						}
						catch { }
					})
			}

			$script:VncExportReaper.Start()
		}
		catch
		{
			try { if ($job) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null } }
			catch { }
			$script:VncExportTask = $null
		}

		return
	}

	& $log "`r`n==================== Starting Export_VNCFiles_ForAllNodes ====================`r`n" "blue"
	$DefaultVNCPassword = "4330df922eb03b6e"
	$desktop = [Environment]::GetFolderPath("Desktop")
	$lanesDir = Join-Path $desktop "Lanes"
	$scalesDir = Join-Path $desktop "Scales"
	$backofficesDir = Join-Path $desktop "BackOffices"
	$hasLaneNodes = ($LaneNumToMachineName -and $LaneNumToMachineName.Count -gt 0)
	$hasScaleNodes = ($ScaleCodeToIPInfo -and $ScaleCodeToIPInfo.Count -gt 0)
	$hasBackofficeNodes = ($BackofficeNumToMachineName -and $BackofficeNumToMachineName.Count -gt 0)

	if (-not ($hasLaneNodes -or $hasScaleNodes -or $hasBackofficeNodes))
	{
		& $log "No lanes, scales, or backoffices were found. Nothing to export." "yellow"
		& $log "`r`n==================== Export_VNCFiles_ForAllNodes Completed ====================" "blue"
		return
	}

	# ---- If passwords not provided, scan them ----
	if (-not $AllVNCPasswords -or $AllVNCPasswords.Count -eq 0)
	{
		& $log "Gathering VNC passwords for all machines`r`n" "magenta"
		$AllVNCPasswords = Get_All_VNC_Passwords -LaneNumToMachineName $LaneNumToMachineName -ScaleCodeToIPInfo $ScaleCodeToIPInfo -BackofficeNumToMachineName $BackofficeNumToMachineName
	}

	# --- Shared VNC file content with token ---
	$vncTemplate = @"
[connection]
host=%%HOST%%
port=5900
proxyhost=
proxyport=0
password=%%PASSWORD%%
[options]
use_encoding_0=1
use_encoding_1=1
use_encoding_2=1
use_encoding_3=0
use_encoding_4=1
use_encoding_5=1
use_encoding_6=1
use_encoding_7=1
use_encoding_8=1
use_encoding_9=1
use_encoding_10=1
use_encoding_11=0
use_encoding_12=0
use_encoding_13=0
use_encoding_14=0
use_encoding_15=0
use_encoding_16=1
use_encoding_17=1
use_encoding_18=1
use_encoding_19=1
use_encoding_20=0
use_encoding_21=0
use_encoding_22=0
use_encoding_23=0
use_encoding_24=0
use_encoding_25=1
use_encoding_26=1
use_encoding_27=1
use_encoding_28=0
use_encoding_29=1
preferred_encoding=10
restricted=0
AllowUntrustedServers=0
viewonly=0
nostatus=0
nohotkeys=0
showtoolbar=1
fullscreen=0
SavePos=0
SaveSize=0
GNOME=0
directx=0
autoDetect=0
8bit=0
shared=1
swapmouse=0
belldeiconify=0
BlockSameMouse=0
emulate3=1
JapKeyboard=0
emulate3timeout=100
emulate3fuzz=4
disableclipboard=0
localcursor=1
Scaling=0
AutoScaling=1
AutoScalingEven=0
AutoScalingLimit=0
scale_num=1
scale_den=1
cursorshape=1
noremotecursor=0
compresslevel=6
quality=8
ServerScale=1
Reconnect=3
EnableCache=0
EnableZstd=1
QuickOption=1
UseDSMPlugin=0
UseProxy=0
sponsor=0
allowMonitorSpanning=0
ChangeServerRes=0
extendDisplay=0
showExtend=0
use_virt=0
useAllMonitors=0
requestedWidth=0
requestedHeight=0
DSMPlugin=NoPlugin
folder=$script:UltraVncFolder
prefix=vnc_
imageFormat=.jpeg
InfoMsg=
AutoReconnect=3
ExitCheck=0
FileTransferTimeout=30
ListenPort=5500
KeepAliveInterval=5
ThrottleMouse=0
AutoAcceptIncoming=0
AutoAcceptNoDSM=0
RequireEncryption=0
PreemptiveUpdates=0
"@

	$resolveVncPassword = {
		param (
			[string]$LookupKey
		)

		if (
			$AllVNCPasswords -and
			-not [string]::IsNullOrWhiteSpace($LookupKey) -and
			$AllVNCPasswords.ContainsKey($LookupKey) -and
			-not [string]::IsNullOrWhiteSpace([string]$AllVNCPasswords[$LookupKey])
		)
		{
			return [string]$AllVNCPasswords[$LookupKey]
		}

		return $DefaultVNCPassword
	}

	$writeExports = {
		param (
			[System.Collections.ArrayList]$Exports,
			[string]$TargetDir,
			[string]$CategoryLabel
		)

		if (-not $Exports -or $Exports.Count -eq 0)
		{
			& $log "No $CategoryLabel VNC files to export. Skipping folder creation.`r`n" "gray"
			return 0
		}

		[void][System.IO.Directory]::CreateDirectory($TargetDir)

		foreach ($item in $Exports)
		{
			$content = $vncTemplate.Replace('%%HOST%%', $item.Host).Replace('%%PASSWORD%%', $item.Password)
			if ($item.AllowUntrustedServers)
			{
				$content = $content -replace 'AllowUntrustedServers=0', 'AllowUntrustedServers=1'
			}

			[System.IO.File]::WriteAllText($item.FilePath, $content, $script:ansiPcEncoding)
			& $log "Created: $($item.FilePath)" "green"
		}

		& $log "$($Exports.Count) $CategoryLabel VNC files written to $TargetDir`r`n" "blue"
		return $Exports.Count
	}

	# ---- Lanes ---- #
	& $log "-------------------- Exporting Lane VNC Files --------------------" "blue"
	$laneExports = New-Object System.Collections.ArrayList
	$dedupedLanes = @{ }
	foreach ($kv in $LaneNumToMachineName.GetEnumerator())
	{
		$machineName = [string]$kv.Value
		if ([string]::IsNullOrWhiteSpace($machineName)) { continue }

		# Only write one file per machine name
		$machineKey = $machineName.Trim().ToLowerInvariant()
		if (-not $dedupedLanes.ContainsKey($machineKey))
		{
			$dedupedLanes[$machineKey] = $true
			if ($machineName.ToUpper() -match '^(POS|SCO)\d+$')
			{
				$fileName = "$($machineName.ToUpper()).vnc"
			}
			else
			{
				$fileName = "Lane_$($kv.Key).vnc"
			}

			[void]$laneExports.Add([pscustomobject]@{
					FilePath = [System.IO.Path]::Combine($lanesDir, $fileName)
					Host	 = $machineName
					Password = (& $resolveVncPassword $machineName)
					AllowUntrustedServers = $false
				})
		}
	}
	$laneCount = & $writeExports $laneExports $lanesDir 'lane'

	# ---- Scales ---- #
	& $log "-------------------- Exporting Scale VNC Files --------------------" "blue"
	$scaleExports = New-Object System.Collections.ArrayList
	$dedupedScales = @{ }
	foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
	{
		$scaleObj = $kv.Value
		# Make sure to dedupe by IP address
		$ip = if ($scaleObj.FullIP) { $scaleObj.FullIP }
		elseif ($scaleObj.IPNetwork -and $scaleObj.IPDevice) { "$($scaleObj.IPNetwork)$($scaleObj.IPDevice)" }
		else { $null }
		if ($ip)
		{
			$ipKey = $ip.Trim().ToLowerInvariant()
			if ($dedupedScales.ContainsKey($ipKey)) { continue }
			$dedupedScales[$ipKey] = $true

			$octets = $ip -split '\.'
			$lastOctet = $octets[-1]
			$brandRaw = ($scaleObj.ScaleBrand -as [string]).Trim()
			$model = ($scaleObj.ScaleModel -as [string]).Trim()
			$brand = if ($brandRaw)
			{
				($brandRaw -split ' ' | ForEach-Object {
						if ($_.Length -gt 0) { $_.Substring(0, 1).ToUpper() + $_.Substring(1).ToLower() }
						else { $_ }
					}) -join ' '
			}
			else { "" }
			if ($brand -and $model)
			{
				$fileName = "$brand($model)_${lastOctet}.vnc"
			}
			elseif ($brand)
			{
				$fileName = "$brand(Unknown)_${lastOctet}.vnc"
			}
			else
			{
				$fileName = "Scale_${lastOctet}.vnc"
			}

			[void]$scaleExports.Add([pscustomobject]@{
					FilePath = [System.IO.Path]::Combine($scalesDir, $fileName)
					Host	 = $ip
					Password = (& $resolveVncPassword $ip)
					AllowUntrustedServers = ($brand -like '*Ishida*')
				})
		}
	}
	$scaleCount = & $writeExports $scaleExports $scalesDir 'scale'

	# ---- Backoffices ---- #
	& $log "-------------------- Exporting Backoffice VNC Files --------------------" "blue"
	$boExports = New-Object System.Collections.ArrayList
	$dedupedBOs = @{ }
	foreach ($kv in $BackofficeNumToMachineName.GetEnumerator())
	{
		$boName = [string]$kv.Value
		$terminal = $kv.Key
		if ([string]::IsNullOrWhiteSpace($boName)) { continue }

		$boKey = $boName.Trim().ToLowerInvariant()
		if (-not $dedupedBOs.ContainsKey($boKey))
		{
			$dedupedBOs[$boKey] = $true
			$fileName = "Backoffice_${terminal}.vnc"

			[void]$boExports.Add([pscustomobject]@{
					FilePath = [System.IO.Path]::Combine($backofficesDir, $fileName)
					Host	 = $boName
					Password = (& $resolveVncPassword $boName)
					AllowUntrustedServers = $false
				})
		}
	}
	$boCount = & $writeExports $boExports $backofficesDir 'backoffice'

	& $log "VNC file export complete!" "green"
	& $log "`r`n==================== Export_VNCFiles_ForAllNodes Completed ====================" "blue"
}

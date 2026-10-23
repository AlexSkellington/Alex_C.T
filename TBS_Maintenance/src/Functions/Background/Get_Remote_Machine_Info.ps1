# ===================================================================================================
#                              FUNCTION: Get_Remote_Machine_Info
# ---------------------------------------------------------------------------------------------------
# Description:
#   Collects hardware/software info from selected lanes, scales and backoffices: one runspace
#   thread per node (10 at a time) tries WMI, then CIM, then REG.EXE, each bounded by a timeout, with an
#   INI-based fallback for lanes/backoffices. Writes per-group info files and populates
#   $script:LaneHardwareInfo / $script:ScaleHardwareInfo / $script:BackofficeHardwareInfo.
#
# Parameters:
#   - RunInBackground       [switch]: run as a background job
#   - SilentBackground      [switch]: background mode without box logging
#   - InternalBackgroundJob [switch]: called from inside the background job itself
#   - UseAllAvailableNodes  [switch]: probe every discovered node (skip the selection form)
#   - SelectedLanes / SelectedScales / SelectedBackoffices [object[]]: pre-selected nodes
#
# Variables:
#   - $script:FunctionResults: lane/scale/backoffice maps and StoreNumber (input)
#
# Writes:
#   - Info text files under Desktop\Lanes, Desktop\Scales and Desktop\BackOffices
#
# Notes:
#   - Requires admin rights on the remote machines for WMI/CIM/REG/SC operations.
# ===================================================================================================

function Get_Remote_Machine_Info
{
	param (
		[switch]$RunInBackground,
		[switch]$SilentBackground,
		[switch]$InternalBackgroundJob,
		[switch]$UseAllAvailableNodes,
		[object[]]$SelectedLanes,
		[object[]]$SelectedScales,
		[object[]]$SelectedBackoffices
	)

	$log = {
		param (
			[string]$Message,
			$Color
		)
		if (-not $SilentBackground)
		{
			Write_Log $Message $Color
		}
	}

	& $log "`r`n==================== Starting Get_Remote_Machine_Info ====================`r`n" "blue"

	# --------------------------- Tunables ---------------------------
	$maxConcurrentJobs = 10
	$wmiTimeoutSeconds = 5
	$cimTimeoutSeconds = 10
	$usePSRemotingFallback = $false # set $true to try PS-Session registry read before REG.EXE

	# --------------------------- Outputs (reset) ---------------------------
	$script:LaneHardwareInfo = $null
	$script:ScaleHardwareInfo = $null
	$script:BackofficeHardwareInfo = $null

	# --------------------------- Inputs from your environment ---------------------------
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
	$BackofficeNumToMachineName = $script:FunctionResults['BackofficeNumToMachineName']
	$StoreNumber = $script:FunctionResults['StoreNumber']
	$DbsPath = $script:DbsPath

	# --------------------------- Keep only discovered/usable nodes ---------------------------
	$availableLaneNodes = @{ }
	if ($LaneNumToMachineName)
	{
		foreach ($kv in $LaneNumToMachineName.GetEnumerator())
		{
			$machineName = [string]$kv.Value
			if (-not [string]::IsNullOrWhiteSpace($machineName))
			{
				$availableLaneNodes[[string]$kv.Key] = $machineName.Trim()
			}
		}
	}

	$availableScaleNodes = @{ }
	if ($ScaleCodeToIPInfo)
	{
		foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
		{
			$scaleObj = $kv.Value
			$scaleIp = $null
			if ($null -ne $scaleObj)
			{
				if (($scaleObj.PSObject.Properties.Name -contains 'FullIP') -and -not [string]::IsNullOrWhiteSpace([string]$scaleObj.FullIP))
				{
					$scaleIp = [string]$scaleObj.FullIP
				}
				elseif (
					($scaleObj.PSObject.Properties.Name -contains 'IPNetwork') -and
					($scaleObj.PSObject.Properties.Name -contains 'IPDevice') -and
					-not [string]::IsNullOrWhiteSpace([string]$scaleObj.IPNetwork) -and
					-not [string]::IsNullOrWhiteSpace([string]$scaleObj.IPDevice)
				)
				{
					$scaleIp = (([string]$scaleObj.IPNetwork).TrimEnd('.') + '.' + [string]$scaleObj.IPDevice)
				}
			}

			if (-not [string]::IsNullOrWhiteSpace($scaleIp))
			{
				$availableScaleNodes[[string]$kv.Key] = $scaleObj
			}
		}
	}

	$availableBackofficeNodes = @{ }
	if ($BackofficeNumToMachineName)
	{
		foreach ($kv in $BackofficeNumToMachineName.GetEnumerator())
		{
			$machineName = [string]$kv.Value
			if (-not [string]::IsNullOrWhiteSpace($machineName))
			{
				$availableBackofficeNodes[[string]$kv.Key] = $machineName.Trim()
			}
		}
	}

	$LaneNumToMachineName = $availableLaneNodes
	$ScaleCodeToIPInfo = $availableScaleNodes
	$BackofficeNumToMachineName = $availableBackofficeNodes

	$selectedLanes = @($SelectedLanes)
	$selectedScales = @($SelectedScales)
	$selectedBOs = @($SelectedBackoffices)

	if ($UseAllAvailableNodes -and
		(-not $selectedLanes -or $selectedLanes.Count -eq 0) -and
		(-not $selectedScales -or $selectedScales.Count -eq 0) -and
		(-not $selectedBOs -or $selectedBOs.Count -eq 0))
	{
		$selectedLanes = @()
		$selectedScales = @()
		$selectedBOs = @()

		if ($LaneNumToMachineName) { $selectedLanes = @($LaneNumToMachineName.Keys | Sort-Object) }
		if ($ScaleCodeToIPInfo) { $selectedScales = @($ScaleCodeToIPInfo.Keys | Sort-Object) }
		if ($BackofficeNumToMachineName) { $selectedBOs = @($BackofficeNumToMachineName.Keys | Sort-Object) }
	}

	if ((-not $UseAllAvailableNodes) -and
		(-not $selectedLanes -or $selectedLanes.Count -eq 0) -and
		(-not $selectedScales -or $selectedScales.Count -eq 0) -and
		(-not $selectedBOs -or $selectedBOs.Count -eq 0))
	{
		# --------------------------- Let user pick nodes ---------------------------
		$nodeSelection = Show_Node_Selection_Form -StoreNumber $StoreNumber `
												  -NodeTypes @("Lane", "Scale", "Backoffice") `
												  -Title "Select Nodes to Pull Hardware Info"

		if (-not $nodeSelection)
		{
			& $log "Get_Remote_Machine_Info cancelled by user." "yellow"
			& $log "==================== Get_Remote_Machine_Info Completed ====================" "blue"
			return $false
		}

		$selectedLanes = @($nodeSelection.Lanes)
		$selectedScales = @($nodeSelection.Scales)
		$selectedBOs = @($nodeSelection.Backoffices)
	}

	if ((-not $selectedLanes -or $selectedLanes.Count -eq 0) -and
		(-not $selectedScales -or $selectedScales.Count -eq 0) -and
		(-not $selectedBOs -or $selectedBOs.Count -eq 0))
	{
		& $log "No nodes selected. Operation aborted." "yellow"
		& $log "==================== Get_Remote_Machine_Info Completed ====================" "blue"
		return $false
	}

	if ($RunInBackground -and -not $InternalBackgroundJob)
	{
		if (
			$script:RemoteMachineInfoTask -and
			$script:RemoteMachineInfoTask.Job -and
			$script:RemoteMachineInfoTask.Job.State -in @('Running', 'NotStarted')
		)
		{
			& $log "Remote machine info export is already running in the background." "yellow"
			return $true
		}

		$job = $null
		try
		{
			$getRemoteMachineInfoDefinition = ${function:Get_Remote_Machine_Info}.ToString()
			$writeLogDefinition = 'param([string]$Message, $Color, [switch]$IncludeTimestamp)'

			$job = Start-Job -Name 'RemoteMachineInfoJob' -ScriptBlock {
				param (
					[string]$GetRemoteMachineInfoDefinition,
					[string]$WriteLogDefinition,
					[hashtable]$LaneMap,
					[hashtable]$ScaleMap,
					[hashtable]$BackofficeMap,
					[string]$StoreNumber,
					[string]$DbsPath,
					[bool]$SilentBackgroundMode,
					[object[]]$SelectedLanes,
					[object[]]$SelectedScales,
					[object[]]$SelectedBackoffices
				)

				Set-Item -Path function:Write_Log -Value ([scriptblock]::Create($WriteLogDefinition))
				Set-Item -Path function:Get_Remote_Machine_Info -Value ([scriptblock]::Create($GetRemoteMachineInfoDefinition))

				$script:FunctionResults = @{
					LaneNumToMachineName	   = $LaneMap
					ScaleCodeToIPInfo		   = $ScaleMap
					BackofficeNumToMachineName = $BackofficeMap
					StoreNumber			       = $StoreNumber
				}
				$script:DbsPath = $DbsPath

				try
				{
					Get_Remote_Machine_Info `
											-SelectedLanes $SelectedLanes `
											-SelectedScales $SelectedScales `
											-SelectedBackoffices $SelectedBackoffices `
											-InternalBackgroundJob `
											-SilentBackground:$SilentBackgroundMode | Out-Null

					[pscustomobject]@{
						Success			       = $true
						LaneHardwareInfo	   = $script:LaneHardwareInfo
						ScaleHardwareInfo	   = $script:ScaleHardwareInfo
						BackofficeHardwareInfo = $script:BackofficeHardwareInfo
						Error				   = $null
					}
				}
				catch
				{
					[pscustomobject]@{
						Success			       = $false
						LaneHardwareInfo	   = $null
						ScaleHardwareInfo	   = $null
						BackofficeHardwareInfo = $null
						Error				   = $_.Exception.Message
					}
				}
			} -ArgumentList @(
				$getRemoteMachineInfoDefinition,
				$writeLogDefinition,
				$LaneNumToMachineName,
				$ScaleCodeToIPInfo,
				$BackofficeNumToMachineName,
				$StoreNumber,
				$DbsPath,
				([bool]$SilentBackground),
				$selectedLanes,
				$selectedScales,
				$selectedBOs
			)

			$script:RemoteMachineInfoTask = @{
				Job	    = $job
				Started = Get-Date
				Silent  = [bool]$SilentBackground
			}

			if (-not $script:RemoteMachineInfoReaper)
			{
				$script:RemoteMachineInfoReaper = New-Object System.Windows.Forms.Timer
				$script:RemoteMachineInfoReaper.Interval = 1000
				$script:RemoteMachineInfoReaper.Add_Tick({
						if (-not $script:RemoteMachineInfoTask)
						{
							try
							{
								if ($script:RemoteMachineInfoReaper)
								{
									$script:RemoteMachineInfoReaper.Stop()
									$script:RemoteMachineInfoReaper.Dispose()
									$script:RemoteMachineInfoReaper = $null
								}
							}
							catch { }
							return
						}

						$taskState = $script:RemoteMachineInfoTask
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

						if ($resultObj -and $resultObj.Success)
						{
							$script:LaneHardwareInfo = $resultObj.LaneHardwareInfo
							$script:ScaleHardwareInfo = $resultObj.ScaleHardwareInfo
							$script:BackofficeHardwareInfo = $resultObj.BackofficeHardwareInfo
							if (-not $taskState.Silent)
							{
								Write_Log "Remote machine info export completed in background." "green"
							}
						}
						elseif (-not $taskState.Silent)
						{
							$errText = if ($resultObj -and $resultObj.Error) { $resultObj.Error }
							else { "Unknown error." }
							Write_Log "Remote machine info export failed in background: $errText" "red"
						}

						$script:RemoteMachineInfoTask = $null

						try
						{
							if ($script:RemoteMachineInfoReaper)
							{
								$script:RemoteMachineInfoReaper.Stop()
								$script:RemoteMachineInfoReaper.Dispose()
								$script:RemoteMachineInfoReaper = $null
							}
						}
						catch { }
					})
			}

			if (-not $SilentBackground)
			{
				& $log "Remote machine info export started in background." "yellow"
			}

			$script:RemoteMachineInfoReaper.Start()
		}
		catch
		{
			try { if ($job) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null } }
			catch { }
			$script:RemoteMachineInfoTask = $null
			& $log "Failed to start remote machine info export in background: $_" "red"
			return $false
		}

		return $true
	}


	# --------------------------- Per-node probe (runs in a runspace pool) ---------------------------
	# 1) ping, 2) WMI = CIM over DCOM, then CIM over WinRM - each with an operation timeout instead
	# of the old nested Start-Job per protocol (three PowerShell processes per node), 3) optional
	# PS-remoting registry read, 4) REG.EXE, 5) INI augmentation.
	$nodeProbeText = {
		param (
			$NodeName,
			$NodeNumber,
			$iniPath,
			$wmiTimeoutSeconds,
			$cimTimeoutSeconds,
			$usePSRemotingFallback
		)

		# --- 0) Result object scaffold ---
		$info = [PSCustomObject]@{
			Success			    = $false
			SystemManufacturer  = $null
			SystemProductName   = $null
			CPU				    = $null
			RAM				    = $null
			OSInfo			    = $null
			BIOS			    = $null
			Method			    = $null
			Error			    = $null
			MachineNameOverride = $null
		}

		# Win32 classes through one CIM session; $null when unreachable / no data
		$cimProbe = {
			param ([string]$Protocol, [int]$TimeoutSec)
			$session = $null
			try
			{
				$session = New-CimSession -ComputerName $NodeName -SessionOption (New-CimSessionOption -Protocol $Protocol) -OperationTimeoutSec $TimeoutSec -ErrorAction Stop
				$q = @{ CimSession = $session; OperationTimeoutSec = $TimeoutSec; ErrorAction = 'Stop' }
				$sys = Get-CimInstance @q -ClassName Win32_ComputerSystem
				if (-not ($sys -and $sys.Manufacturer -and $sys.Model)) { return $null }
				$cpu = Get-CimInstance @q -ClassName Win32_Processor | Select-Object -First 1
				$os = Get-CimInstance @q -ClassName Win32_OperatingSystem

				$biosVerOut = $null
				try
				{
					$bios = Get-CimInstance @q -ClassName Win32_BIOS
					if ($bios)
					{
						if ($bios.SMBIOSBIOSVersion) { $biosVerOut = $bios.SMBIOSBIOSVersion }
						elseif ($bios.BIOSVersion) { $biosVerOut = ($bios.BIOSVersion | Where-Object { $_ } | Select-Object -First 1) }
						elseif ($bios.Version) { $biosVerOut = $bios.Version }
						$dt = $null
						if ($bios.ReleaseDate -is [datetime]) { $dt = $bios.ReleaseDate }
						elseif ($bios.ReleaseDate) { try { $dt = [System.Management.ManagementDateTimeConverter]::ToDateTime([string]$bios.ReleaseDate) } catch { } }
						if ($dt)
						{
							if ($biosVerOut) { $biosVerOut = "$biosVerOut ($($dt.ToString('yyyy-MM-dd')))" }
							else { $biosVerOut = $dt.ToString('yyyy-MM-dd') }
						}
					}
				}
				catch { }

				[PSCustomObject]@{
					SystemManufacturer = $sys.Manufacturer
					SystemProductName  = $sys.Model
					CPU			       = $cpu.Name
					RAM			       = [math]::Round($sys.TotalPhysicalMemory / 1GB, 1)
					OSInfo			   = "$($os.Caption) ($($os.Version))"
					BIOS			   = $biosVerOut
					Method			   = $(if ($Protocol -eq 'Dcom') { 'WMI' } else { 'CIM' })
				}
			}
			catch { $null }
			finally { if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue } }
		}
		$applyCim = {
			param ($r)
			$info.SystemManufacturer = $r.SystemManufacturer
			$info.SystemProductName = $r.SystemProductName
			$info.CPU = $r.CPU
			$info.RAM = $r.RAM
			$info.OSInfo = $r.OSInfo
			if ($r.BIOS) { $info.BIOS = $r.BIOS }
			$info.Method = $r.Method
			$info.Success = $true
			$info.Error = $null
		}

		# --- 1) Ping (1.5 s; Test-Connection waited up to 4 s per switched-off node) ---
		$online = $false
		try
		{
			$pinger = New-Object System.Net.NetworkInformation.Ping
			try { $online = ($pinger.Send($NodeName, 1500).Status -eq [System.Net.NetworkInformation.IPStatus]::Success) }
			finally { $pinger.Dispose() }
		}
		catch { $online = $false }

		if (-not $online)
		{
			$info.Error = "Offline or unreachable."
			$info.Method = "Offline"
		}
		else
		{
			# --- 2) WMI (DCOM), then CIM (WinRM) ---
			$r = & $cimProbe 'Dcom' $wmiTimeoutSeconds
			if ($r) { & $applyCim $r }
			else
			{
				$info.Error = "WMI failed (access issue or null)"
				$r = & $cimProbe 'Wsman' $cimTimeoutSeconds
				if ($r) { & $applyCim $r }
				else { $info.Error = "CIM failed (access issue or null)" }
			}
		}
		# --- 3) PS-Remoting Registry (optional middle step) ---
		if (-not $info.Success -and $usePSRemotingFallback)
		{
			$regResult = $null
			try
			{
				$sessionOption = New-PSSessionOption -ConnectionTimeout 15000 -OpenTimeout 15000
				$session = New-PSSession -ComputerName $NodeName -SessionOption $sessionOption -ErrorAction SilentlyContinue
				if ($session)
				{
					$regResult = Invoke-Command -Session $session -ScriptBlock {
						$out = [PSCustomObject]@{
							SystemManufacturer = $null
							SystemProductName  = $null
							CPU			       = $null
							OSInfo			   = $null
							BIOS			   = $null
							Success		       = $false
							Error			   = $null
							Method			   = 'REG(PSRemoting)'
						}
						try
						{
							$manuf = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name SystemManufacturer -ErrorAction SilentlyContinue
							$prod = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name SystemProductName -ErrorAction SilentlyContinue
							$cpu = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0' -Name ProcessorNameString -ErrorAction SilentlyContinue
							$os = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue

							# BIOS pieces
							$biosVendor = (Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name BIOSVendor -ErrorAction SilentlyContinue).BIOSVendor
							$biosVerStr = (Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name BIOSVersion -ErrorAction SilentlyContinue).BIOSVersion
							if (-not $biosVerStr)
							{
								$biosVerStr = (Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name SystemBiosVersion -ErrorAction SilentlyContinue).SystemBiosVersion
							}
							$biosDate = (Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name BIOSReleaseDate -ErrorAction SilentlyContinue).BIOSReleaseDate

							# Build OSInfo
							$osInfo = $null
							if ($os -and $os.ProductName) { $osInfo = $os.ProductName }
							if ($os -and $os.DisplayVersion)
							{
								if ($osInfo) { $osInfo = "$osInfo ($($os.DisplayVersion))" }
								else { $osInfo = $os.DisplayVersion }
							}
							elseif ($os -and $os.CurrentBuild)
							{
								if ($osInfo) { $osInfo = "$osInfo (Build $($os.CurrentBuild))" }
								else { $osInfo = "Build $($os.CurrentBuild)" }
							}

							# BIOS assemble
							$biosOut = $null
							if ($biosVendor -and $biosVerStr) { $biosOut = "$biosVendor $biosVerStr" }
							elseif ($biosVerStr) { $biosOut = $biosVerStr }
							elseif ($biosVendor) { $biosOut = $biosVendor }
							if ($biosDate)
							{
								if ($biosOut) { $biosOut = "$biosOut ($biosDate)" }
								else { $biosOut = $biosDate }
							}

							$out.SystemManufacturer = $manuf.SystemManufacturer
							$out.SystemProductName = $prod.SystemProductName
							$out.CPU = $cpu.ProcessorNameString
							$out.OSInfo = $osInfo
							$out.BIOS = $biosOut
							$out.Success = $true
						}
						catch { $out.Error = $_.Exception.Message }
						return $out
					} 2>$null
					Remove-PSSession $session -ErrorAction SilentlyContinue
				}
			}
			catch
			{
				if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
				$regResult = [PSCustomObject]@{
					SystemManufacturer = $null
					SystemProductName  = $null
					CPU			       = $null
					OSInfo			   = $null
					BIOS			   = $null
					Success		       = $false
					Error			   = "PS-Remoting registry query failed"
					Method			   = 'REG(PSRemoting)'
				}
			}

			if ($regResult -and $regResult.Success)
			{
				if (-not $info.SystemManufacturer) { $info.SystemManufacturer = $regResult.SystemManufacturer }
				if (-not $info.SystemProductName) { $info.SystemProductName = $regResult.SystemProductName }
				if (-not $info.CPU) { $info.CPU = $regResult.CPU }
				if (-not $info.OSInfo) { $info.OSInfo = $regResult.OSInfo }
				if (-not $info.BIOS -and $regResult.BIOS) { $info.BIOS = $regResult.BIOS }
				$info.Method = $regResult.Method
				$info.Success = $true
				$info.Error = $null
			}
			elseif (-not $info.Error)
			{
				if ($regResult) { $info.Error = $regResult.Error; $info.Method = 'REG(PSRemoting)' }
			}
		}

		# --- 4) REG.EXE (starts RemoteRegistry if needed, then restores) ---
		if (-not $info.Success)
		{
			$wasRunning = $false
			$wasDisabled = $false
			try
			{
				$state = sc.exe "\\$NodeName" query RemoteRegistry 2>$null | Select-String 'STATE'
				$start = sc.exe "\\$NodeName" qc    RemoteRegistry 2>$null | Select-String 'START_TYPE'
				if ($state -and $state.Line -match 'RUNNING') { $wasRunning = $true }
				if ($start -and $start.Line -match 'DISABLED') { $wasDisabled = $true }
				if ($wasDisabled) { sc.exe "\\$NodeName" config RemoteRegistry start= demand 2>$null | Out-Null }
				if (-not $wasRunning) { sc.exe "\\$NodeName" start  RemoteRegistry 2>$null | Out-Null }

				# tiny inline getter: returns only the DATA (no "REG_SZ")
				$getVal = {
					param ($hive,
						$path,
						$name)
					$patternName = [regex]::Escape($name)
					$raw = reg.exe QUERY "\\$NodeName\$hive\$path" /v $name 2>$null
					if (-not $raw) { return $null }
					$line = $raw | Select-String -Pattern "^\s*$patternName\s+REG_\w+\s+.+$" | Select-Object -First 1
					if (-not $line) { return $null }
					$data = $line.Line -replace "^\s*$patternName\s+REG_\w+\s+", ""
					$data.Trim()
				}

				$manuf = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS'               'SystemManufacturer'
				$prod = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS'               'SystemProductName'
				$cpu = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\CentralProcessor\0' 'ProcessorNameString'
				$osN = & $getVal 'HKLM' 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'   'ProductName'
				$osB = & $getVal 'HKLM' 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'   'DisplayVersion'
				$osV = & $getVal 'HKLM' 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'   'CurrentBuild'

				# BIOS from REG.EXE
				$biosVendor = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS' 'BIOSVendor'
				$biosVerStr = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS' 'BIOSVersion'
				if (-not $biosVerStr)
				{
					$biosVerStr = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS' 'SystemBiosVersion'
				}
				$biosDate = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS' 'BIOSReleaseDate'
				if ($biosVerStr)
				{
					$biosVerStr = $biosVerStr -replace '\x00', ' '
					$biosVerStr = ($biosVerStr -split '\s{2,}' | Where-Object { $_ }) -join ' '
					$biosVerStr = $biosVerStr.Trim()
				}
				$biosOut = $null
				if ($biosVendor -and $biosVerStr) { $biosOut = "$biosVendor $biosVerStr" }
				elseif ($biosVerStr) { $biosOut = $biosVerStr }
				elseif ($biosVendor) { $biosOut = $biosVendor }
				if ($biosDate)
				{
					if ($biosOut) { $biosOut = "$biosOut ($biosDate)" }
					else { $biosOut = $biosDate }
				}

				# If CurrentBuild came as "0x... (7601)" keep the number in parentheses
				if ($osV -and $osV -match '\((\d+)\)') { $osV = $matches[1] }

				# build a clean OS string
				$osInfo = $null
				if ($osN) { $osInfo = $osN }
				if ($osB)
				{
					if ($osInfo) { $osInfo = "$osInfo ($osB)" }
					else { $osInfo = $osB }
				}
				elseif ($osV)
				{
					if ($osInfo) { $osInfo = "$osInfo (Build $osV)" }
					else { $osInfo = "Build $osV" }
				}

				if ($manuf -or $prod -or $cpu -or $osInfo -or $biosOut)
				{
					if (-not $info.SystemManufacturer) { $info.SystemManufacturer = $manuf }
					if (-not $info.SystemProductName) { $info.SystemProductName = $prod }
					if (-not $info.CPU) { $info.CPU = $cpu }
					if (-not $info.OSInfo) { $info.OSInfo = $osInfo }
					if (-not $info.BIOS -and $biosOut) { $info.BIOS = $biosOut }
					$info.Method = "REG.EXE"
					$info.Success = $true
					$info.Error = $null
				}
				else
				{
					$info.Method = "REG.EXE"
					$info.Success = $false
					$info.Error = "REG queries returned no data"
				}
			}
			catch
			{
				$info.Method = "REG.EXE"
				$info.Success = $false
				$info.Error = "REG queries failed"
			}
			finally
			{
				try
				{
					if (-not $wasRunning) { sc.exe "\\$NodeName" stop RemoteRegistry 2>$null | Out-Null }
					if ($wasDisabled) { sc.exe "\\$NodeName" config RemoteRegistry start= disabled 2>$null | Out-Null }
				}
				catch { }
			}
		}

		# A node that did not answer the ping and not REG.EXE either is reported as offline
		# (REG.EXE is still tried: some lanes block ICMP)
		if (-not $info.Success -and -not $online)
		{
			$info.Error = "Offline or unreachable."
			$info.Method = "Offline"
		}

		# --- 5) INI augmentation (optional; fills gaps if any fields still null) ---
		$returnInfo = [PSCustomObject]@{
			Machine    = $NodeName
			NodeNumber = $NodeNumber
			Info	   = $info
			IniFound   = $false
			IniPath    = $null
		}
		try
		{
			if ($iniPath)
			{
				$iniFolder = Split-Path $iniPath -Parent
				$iniLeaf = Split-Path $iniPath -Leaf
				$iniFile = Get-ChildItem -Path $iniFolder -Filter $iniLeaf -ErrorAction SilentlyContinue |
				Sort-Object LastWriteTime -Descending |
				Select-Object -First 1
				if ($iniFile -and (Test-Path $iniFile.FullName))
				{
					$returnInfo.IniFound = $true
					$returnInfo.IniPath = $iniFile.FullName
					$iniLines = Get-Content $iniFile.FullName -Encoding UTF8 -ErrorAction SilentlyContinue

					# tiny parser
					$sections = @{ }
					$curSec = ""
					foreach ($line in $iniLines)
					{
						if ($line -match '^\[(.+)\]$')
						{
							$curSec = $matches[1]; $sections[$curSec] = @{ }
						}
						elseif ($line -match '^\s*([^=]+?)\s*=\s*(.*)$' -and $curSec)
						{
							$sections[$curSec][$matches[1].Trim()] = $matches[2].Trim()
						}
					}
					if (-not $info.CPU -and $sections.ContainsKey('PROCESSOR') -and $sections['PROCESSOR'].ContainsKey('Cores'))
					{
						$cores = $sections['PROCESSOR']['Cores']
						$arch = $sections['PROCESSOR']['Architecture']
						if ($arch) { $info.CPU = "$cores cores ($arch)" }
						else { $info.CPU = "$cores cores" }
					}
					if ($info.RAM -eq $null -and $sections.ContainsKey('Memory') -and $sections['Memory'].ContainsKey('PhysicalMemory'))
					{
						$ramMb = $sections['Memory']['PhysicalMemory']
						if ($ramMb -match '^\d+$') { $info.RAM = [math]::Round([double]$ramMb/1024, 1) }
						else { $info.RAM = $ramMb }
					}
					if (-not $info.OSInfo -and $sections.ContainsKey('OperatingSystem') -and $sections['OperatingSystem'].ContainsKey('ProductName'))
					{
						$info.OSInfo = $sections['OperatingSystem']['ProductName']
					}
				}
			}
		}
		catch { }

		return $returnInfo

	}.ToString()

	# --------------------------- Output folders ---------------------------
	$desktop = [Environment]::GetFolderPath("Desktop")
	$lanesDir = Join-Path $desktop "Lanes"
	$scalesDir = Join-Path $desktop "Scales"
	$backofficesDir = Join-Path $desktop "BackOffices"

	# --------------------------- Process 3 categories uniformly ---------------------------
	foreach ($section in @(
			@{ Name = 'Lanes'; Selected = $selectedLanes; Dir = $lanesDir; FileName = 'Lanes_Info.txt' },
			@{ Name = 'Scales'; Selected = $selectedScales; Dir = $scalesDir; FileName = 'Scales_Info.txt' },
			@{ Name = 'BackOffices'; Selected = $selectedBOs; Dir = $backofficesDir; FileName = 'Backoffices_Info.txt' }
		))
	{
		if (-not $section.Selected -or $section.Selected.Count -eq 0) { continue }

		& $log "Processing $($section.Name) nodes..." "yellow"
		$results = @{ }
		$infoLines = New-Object System.Collections.Generic.List[string]

		# One node's result -> its info-file line + results entry
		$collectNode = {
			param ([string]$remoteName, $result)
			$info = if ($result) { $result.Info } else { $null }
			if (-not $info) { $info = [PSCustomObject]@{ Success = $false; Error = 'No answer within the time limit'; Method = 'Timeout'; BIOS = $null } }

			$line = "Machine Name: $remoteName |"
			if ($info.Success)
			{
				$line += " Manufacturer: $($info.SystemManufacturer) | Model: $($info.SystemProductName) | CPU: $($info.CPU)"
				if ($null -ne $info.RAM) { $line += " | RAM: $($info.RAM) GB" }
				if ($info.OSInfo) { $line += " | OS: $($info.OSInfo)" }
				& $log "Processed $remoteName ($($section.Name)): Success [$($info.Method)]" "green"
			}
			elseif ($info.Error)
			{
				$line += " [Hardware info unavailable] Error: $($info.Error)"
				& $log "Processed $remoteName ($($section.Name)): Error [$($info.Method)] - $($info.Error)" "red"
			}
			else
			{
				$line += " [No hardware info found]"
				& $log "Processed $remoteName ($($section.Name)): No info [$($info.Method)]" "yellow"
			}
			$biosText = if ($info.BIOS) { $info.BIOS } else { "Unknown" }
			$line += " | BIOS: $biosText | Method: $($info.Method)"
			$infoLines.Add($line)
			$results[$remoteName] = $info
		}

		$pool = [RunspaceFactory]::CreateRunspacePool(1, $maxConcurrentJobs)
		$pool.Open()
		$queued = @()
		foreach ($sel in $section.Selected)
		{
			# -------- Canonical per-node identity: $NodeNumber (code) + $NodeName (machine/IP) --------
			$NodeNumber = $null
			$NodeName = $null

			if ($section.Name -eq 'Lanes')
			{
				if ($sel.PSObject.Properties.Name -contains 'LaneNumber' -and $sel.LaneNumber)
				{
					$NodeNumber = "{0:D3}" -f [int]$sel.LaneNumber
				}
				else { $NodeNumber = "$sel" }
				$NodeName = $LaneNumToMachineName[$NodeNumber]
			}
			elseif ($section.Name -eq 'Scales')
			{
				$NodeNumber = "$sel"
				if ($ScaleCodeToIPInfo.ContainsKey($NodeNumber) -and
					$ScaleCodeToIPInfo[$NodeNumber].PSObject.Properties.Name -contains 'FullIP')
				{
					$NodeName = $ScaleCodeToIPInfo[$NodeNumber].FullIP
				}
			}
			elseif ($section.Name -eq 'BackOffices')
			{
				if ($sel.PSObject.Properties.Name -contains 'BONumber' -and $sel.BONumber)
				{
					$NodeNumber = "{0:D3}" -f [int]$sel.BONumber
				}
				else { $NodeNumber = "$sel" }
				$NodeName = $BackofficeNumToMachineName[$NodeNumber]
			}

			if (-not $NodeName)
			{
				& $log "Skipping $($section.Name) $NodeNumber (no NodeName/mapping found)" "red"
				continue
			}

			# -------- Optional INI path (used to augment fields) --------
			$iniPattern = "INFO_${StoreNumber}${NodeNumber}_SMSStart.ini"
			if ($section.Name -eq 'Lanes') { $iniPath = Join-Path "\\$NodeName\storeman\office\dbs" $iniPattern }
			elseif ($section.Name -eq 'BackOffices') { $iniPath = Join-Path $DbsPath $iniPattern }
			else { $iniPath = $null }

			$ps = [PowerShell]::Create()
			$ps.RunspacePool = $pool
			$null = $ps.AddScript($nodeProbeText).AddArgument($NodeName).AddArgument($NodeNumber).AddArgument($iniPath).AddArgument($wmiTimeoutSeconds).AddArgument($cimTimeoutSeconds).AddArgument($usePSRemotingFallback)
			$queued += [pscustomobject]@{ Name = [string]$NodeName; PS = $ps; Handle = $ps.BeginInvoke() }
		} # end foreach selected

		# Collect in order; 90 s for the first batch of nodes plus 30 s per further batch
		$deadline = (Get-Date).AddSeconds(90 + 30 * [Math]::Max(0, [Math]::Ceiling($queued.Count / $maxConcurrentJobs) - 1))
		$hung = $false
		foreach ($q in $queued)
		{
			$result = $null
			$left = [int][Math]::Max(0, ($deadline - (Get-Date)).TotalMilliseconds)
			if ($q.Handle.AsyncWaitHandle.WaitOne($left))
			{
				try { $result = @($q.PS.EndInvoke($q.Handle)) | Select-Object -Last 1 }
				catch { $result = $null }
				$q.PS.Dispose()
			}
			else
			{
				try { $null = $q.PS.BeginStop($null, $null) }
				catch { }
				$hung = $true
			}
			& $collectNode $q.Name $result
		}
		if ($hung) { try { $null = $pool.BeginClose($null, $null) } catch { } }
		else { try { $pool.Close(); $pool.Dispose() } catch { } }

		# Write file (stable sorting: IPs zero-padded; otherwise lexical)
		if ($infoLines.Count)
		{
			[void][System.IO.Directory]::CreateDirectory($section.Dir)

			$sortedLines = $infoLines | Sort-Object {
				if ($_ -match "^Machine Name: ([^|]+)")
				{
					$name = $matches[1].Trim()
					if ($name -match '^\d{1,3}(\.\d{1,3}){3}$')
					{
						$oct = $name -split '\.'
						'{0:D3}.{1:D3}.{2:D3}.{3:D3}' -f [int]$oct[0], [int]$oct[1], [int]$oct[2], [int]$oct[3]
					}
					else { $name }
				}
				else { $_ }
			}
			$filePath = Join-Path $section.Dir $section.FileName
			$sortedLines -join "`r`n" | Set-Content -Path $filePath -Encoding Default
			& $log "Exported $($section.Name) info to $filePath" "green"
		}

		# Publish programmatic map
		if ($section.Name -eq 'Lanes') { $script:LaneHardwareInfo = $results }
		elseif ($section.Name -eq 'Scales') { $script:ScaleHardwareInfo = $results }
		elseif ($section.Name -eq 'BackOffices') { $script:BackofficeHardwareInfo = $results }

		& $log "Completed processing $($section.Name).`r`n" "green"
	} # end sections

	& $log "==================== Get_Remote_Machine_Info Completed ====================" "blue"
}

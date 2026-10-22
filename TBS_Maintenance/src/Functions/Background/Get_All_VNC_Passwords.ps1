# ===================================================================================================
#                              FUNCTION: Get_All_VNC_Passwords
# ---------------------------------------------------------------------------------------------------
# Description:
#   Scans all lanes, backoffices and scales for the UltraVNC configuration file (UltraVNC.ini,
#   any case) in the standard install locations, extracts the encrypted VNC password and
#   returns a machine-name -> password hashtable. Uses a fixed default password for nodes
#   where no UltraVNC.ini can be read.
#
# Parameters:
#   - LaneNumToMachineName       [hashtable]: lane number -> machine name
#   - ScaleCodeToIPInfo          [hashtable]: scale code -> scale info object
#   - BackofficeNumToMachineName [hashtable]: backoffice number -> machine name
#
# Variables:
#   - $script:DefaultVNCPassword: fixed default for stock-password nodes (e.g. Ishida)
#
# Returns:
#   Hashtable MachineName => encrypted password (or $null when not found).
#
# Writes:
#   - $script:FunctionResults['AllVNCPasswords']
# ===================================================================================================

function Get_All_VNC_Passwords
{
	param (
		[Parameter(Mandatory = $false)]
		[hashtable]$LaneNumToMachineName,
		[Parameter(Mandatory = $false)]
		[hashtable]$ScaleCodeToIPInfo,
		[Parameter(Mandatory = $false)]
		[hashtable]$BackofficeNumToMachineName
	)
	
	# Default VNC password for lanes, backoffices, and scales with fixed passwords (e.g., Ishida)
	$DefaultVNCPassword = "4330df922eb03b6e"
	$script:DefaultVNCPassword = $DefaultVNCPassword
	
	$uvncFolders = @(
		"C$\Program Files\uvnc bvba\UltraVNC",
		"C$\Program Files (x86)\uvnc bvba\UltraVNC"
	)
	$AllVNCPasswords = @{ }
	
	# 1. Build main node list and tag brands for scales
	$NodeList = @()
	$BizerbaScales = @()
	if ($LaneNumToMachineName) { $NodeList += $LaneNumToMachineName.Values | Where-Object { $_ } }
	if ($ScaleCodeToIPInfo)
	{
		$uniqueScaleObjs = @{ }
		foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
		{
			$scaleObj = $kv.Value
			# Compute a unique identifier, e.g., IP or some combination
			$ip = $null
			if ($scaleObj.FullIP) { $ip = $scaleObj.FullIP }
			elseif ($scaleObj.IPNetwork -and $scaleObj.IPDevice) { $ip = "$($scaleObj.IPNetwork)$($scaleObj.IPDevice)" }
			# Add only once!
			if ($ip -and -not $uniqueScaleObjs.ContainsKey($ip))
			{
				$uniqueScaleObjs[$ip] = $scaleObj
			}
		}
		foreach ($ip in $uniqueScaleObjs.Keys)
		{
			$scaleObj = $uniqueScaleObjs[$ip]
			$ip = $null
			$isIshida = $false
			$isBizerba = $false
			if ($scaleObj -is [string]) { $ip = $scaleObj }
			elseif ($null -ne $scaleObj)
			{
				if (($scaleObj.PSObject.Properties.Name -contains "ScaleBrand") -and ($scaleObj.ScaleBrand -match "Ishida"))
				{
					$isIshida = $true
				}
				if (($scaleObj.PSObject.Properties.Name -contains "ScaleBrand") -and ($scaleObj.ScaleBrand -match "BIZERBA"))
				{
					$isBizerba = $true
				}
				if ($scaleObj.PSObject.Properties.Name -contains "FullIP" -and $scaleObj.FullIP)
				{
					$ip = $scaleObj.FullIP
				}
				elseif (($scaleObj.PSObject.Properties.Name -contains "IPNetwork") -and ($scaleObj.PSObject.Properties.Name -contains "IPDevice") -and $scaleObj.IPNetwork -and $scaleObj.IPDevice)
				{
					$ip = "$($scaleObj.IPNetwork)$($scaleObj.IPDevice)"
				}
			}
			if ($ip)
			{
				if ($isIshida)
				{
					Write_Log "Ishida scale [$ip] will use default VNC password.`r`n" "yellow"
					$AllVNCPasswords[$ip] = $DefaultVNCPassword
					continue
				}
				if ($isBizerba) { $BizerbaScales += @{ Host = $ip; Obj = $scaleObj }; continue }
				$NodeList += $ip
			}
		}
	}
	if ($BackofficeNumToMachineName) { $NodeList += $BackofficeNumToMachineName.Values | Where-Object { $_ } }
	$NodeList = $NodeList | Sort-Object -Unique
	
	if (($NodeList.Count -eq 0) -and ($BizerbaScales.Count -eq 0)) { throw "No machines provided for password extraction." }
	
	# 2. Ping regular nodes first (v2.6.9: parallel probe with a concurrency
	# cap - the old serial Test-Connection loop could block for minutes when
	# several nodes were down)
	$onlineProbe = @{}
	if ($NodeList.Count -gt 0) { $onlineProbe = Test-Hosts_Parallel -Targets $NodeList }
	$OnlineNodes = @()
	foreach ($node in $NodeList)
	{
		if ($onlineProbe[$node] -eq 'OK')
		{
			$OnlineNodes += $node
		}
		else
		{
			$AllVNCPasswords[$node] = $null
		}
	}
	
	# 3. Start jobs for online regular nodes
	$jobs = @()
	$MaxConcurrentJobs = 20 # Increased for faster processing
	
	foreach ($machineName in $OnlineNodes)
	{
		while ($jobs.Count -ge $MaxConcurrentJobs)
		{
			$done = Wait-Job -Job $jobs -Any -Timeout 10
			if ($done)
			{
				# Receive BEFORE removing: Remove-Job on a finished job discards
				# its output, which silently dropped these machines' VNC passwords.
				# -Any returns ONE job per wait, so only the reported jobs are
				# processed here; everything else stays in the pool for the
				# final Wait-Job/collect below (filtering by state here also
				# dropped jobs that finished between waits).
				$processedIds = @($done | ForEach-Object { $_.Id })
				foreach ($j in @($done))
				{
					$doneResult = Receive-Job $j
					if ($doneResult -and $doneResult.Machine) { $AllVNCPasswords[$doneResult.Machine] = $doneResult.Password }
					Remove-Job $j -Force
				}
				$jobs = @($jobs | Where-Object { $processedIds -notcontains $_.Id })
			}
		}
		$jobs += Start-Job -Name ("GetAllVNCPasswords_{0}" -f $machineName) -ArgumentList $machineName, $uvncFolders -ScriptBlock {
			param ($machineName,
				$uvncFolders)
			$password = $null
			foreach ($folder in $uvncFolders)
			{
				try
				{
					$iniFiles = Invoke-Command -ComputerName $machineName -TimeoutSec 30 -ScriptBlock {
						param ($dir)
						if (Test-Path $dir)
						{
							Get-ChildItem -Path $dir -Filter "*.ini" -File | ForEach-Object {
								if ($_.Name.ToLower() -eq "ultravnc.ini")
								{
									$_.FullName
								}
							}
						}
					} -ArgumentList ("C:\" + $folder.Substring(3)) -ErrorAction Stop
					foreach ($iniFile in $iniFiles)
					{
						$content = Invoke-Command -ComputerName $machineName -TimeoutSec 30 -ScriptBlock {
							param ($path)
							if (Test-Path $path)
							{
								Get-Content $path -ErrorAction Stop
							}
						} -ArgumentList $iniFile -ErrorAction Stop
						foreach ($line in $content)
						{
							if ($line -match '^\s*passwd\s*=\s*([0-9A-Fa-f]+)')
							{
								$password = $matches[1]
								break
							}
						}
						if ($password) { break }
					}
				}
				catch
				{
					try
					{
						$remotePath = "\\$machineName\$folder"
						if (Test-Path $remotePath -ErrorAction SilentlyContinue)
						{
							$iniFiles = Get-ChildItem -Path $remotePath -Filter "*.ini" -File -ErrorAction SilentlyContinue | Where-Object {
								$_.Name.ToLower() -eq "ultravnc.ini"
							}
							foreach ($iniFile in $iniFiles)
							{
								$content = Get-Content $iniFile.FullName -ErrorAction Stop
								foreach ($line in $content)
								{
									if ($line -match '^\s*passwd\s*=\s*([0-9A-Fa-f]+)')
									{
										$password = $matches[1]
										break
									}
								}
								if ($password) { break }
							}
						}
					}
					catch { continue }
				}
				if ($password) { break }
			}
			return @{ Machine = $machineName; Password = $password }
		}
	}
	
	# Only wait if there is anything to wait for: when the throttle loop above
	# has already drained all jobs, $jobs is empty/null and Wait-Job throws a
	# parameter validation error.
	if (@($jobs).Count -gt 0)
	{
		$null = Wait-Job -Job $jobs -Timeout 120
		foreach ($j in @($jobs))
		{
			if ($j.State -ne 'Completed') { Stop-Job -Job $j -ErrorAction SilentlyContinue }
		}
	}
	
	foreach ($j in $jobs)
	{
		$result = Receive-Job $j
		if ($result -and $result.Machine)
		{
			$AllVNCPasswords[$result.Machine] = $result.Password
		}
		Remove-Job $j -Force
	}
	
	# 4. Handle Bizerba scales (using cmdkey/bizuser logic, serially)
	foreach ($b in $BizerbaScales)
	{
		$scaleHost = $b.Host
		$uvncIniRelativePaths = @(
			"Program Files\uvnc bvba\UltraVNC\ultravnc.ini",
			"Program Files (x86)\uvnc bvba\UltraVNC\ultravnc.ini"
		)
		$passwords = @("bizerba", "biyerba")
		$username = "bizuser"
		$password = $null
		$fullIniPath = $null
		
		# Ping check first
		if (-not (Test-Connection -ComputerName $scaleHost -Count 1 -Quiet -ErrorAction SilentlyContinue))
		{
			$AllVNCPasswords[$scaleHost] = $null
			continue
		}
		
		foreach ($uvncIniRel in $uvncIniRelativePaths)
		{
			foreach ($pw in $passwords)
			{
				# Clear any lingering share session, but keep the persisted credential manager entry.
				try { $null = & net.exe use "\\$scaleHost\c$" '/delete' '/y' 2>$null }
				catch { }
				cmdkey /add:$scaleHost /user:$username /pass:$pw | Out-Null
				$shareIniPath = "\\$scaleHost\c$\$uvncIniRel"
				if (Test-Path $shareIniPath -ErrorAction SilentlyContinue)
				{
					try
					{
						$content = Get-Content $shareIniPath -ErrorAction Stop
						foreach ($line in $content)
						{
							if ($line -match '^\s*passwd\s*=\s*([0-9A-Fa-f]+)')
							{
								$password = $matches[1]
								$fullIniPath = $shareIniPath
								break
							}
						}
					}
					catch { }
				}
				# Drop any temporary SMB session, but do not delete the saved credential.
				try { $null = & net.exe use "\\$scaleHost\c$" '/delete' '/y' 2>$null }
				catch { }
				if ($password) { break }
			}
			if ($password) { break }
		}
		$AllVNCPasswords[$scaleHost] = $password
	}
	$script:FunctionResults['AllVNCPasswords'] = $AllVNCPasswords
	return $AllVNCPasswords
}

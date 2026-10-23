# ===================================================================================================
#                              FUNCTION: Get_All_VNC_Passwords
# ---------------------------------------------------------------------------------------------------
# Description:
#   Scans all lanes, backoffices and scales for the UltraVNC configuration file (UltraVNC.ini,
#   any case) in the standard install locations, extracts the encrypted VNC password and
#   returns a machine-name -> password hashtable. Uses a fixed default password for nodes
#   where no UltraVNC.ini can be read. Online lanes / backoffices / scales are read in
#   parallel in a runspace pool (20 threads, 2 minutes in total); Bizerba scales one by one
#   (they need the bizuser credential).
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
			elseif ($scaleObj.IPNetwork -and $scaleObj.IPDevice) { $ip = (([string]$scaleObj.IPNetwork).TrimEnd('.') + '.' + [string]$scaleObj.IPDevice) }
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
					$ip = (([string]$scaleObj.IPNetwork).TrimEnd('.') + '.' + [string]$scaleObj.IPDevice)
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

	# 3. Read ultravnc.ini of every online node in parallel - runspace threads, not one whole
	#    PowerShell process per node. The admin share first (fast, what works in the stores);
	#    WinRM only when the share cannot be read. (Before v2.7.7 WinRM came first and always
	#    failed: Invoke-Command has no -TimeoutSec in Windows PowerShell 5.1.)
	$readIniText = {
		param ($machineName, $uvncFolders)
		$findPasswd = {
			param ($lines)
			foreach ($line in $lines) { if ($line -match '^\s*passwd\s*=\s*([0-9A-Fa-f]+)') { return $matches[1] } }
			return $null
		}
		$password = $null
		foreach ($folder in $uvncFolders)
		{
			$iniPath = "\\$machineName\$folder\ultravnc.ini"
			try { if ([System.IO.File]::Exists($iniPath)) { $password = & $findPasswd ([System.IO.File]::ReadAllLines($iniPath)) } }
			catch { }
			if ($password) { break }
		}
		if (-not $password)
		{
			try
			{
				$opt = New-PSSessionOption -OpenTimeout 10000 -OperationTimeout 20000
				$lines = Invoke-Command -ComputerName $machineName -SessionOption $opt -ErrorAction Stop -ScriptBlock {
					param ($dirs)
					foreach ($d in $dirs) { $p = Join-Path $d 'ultravnc.ini'; if (Test-Path -LiteralPath $p) { Get-Content -LiteralPath $p; break } }
				} -ArgumentList (, @($uvncFolders | ForEach-Object { 'C:\' + $_.Substring(3) }))
				$password = & $findPasswd $lines
			}
			catch { }
		}
		return @{ Machine = $machineName; Password = $password }
	}.ToString()

	$MaxConcurrentJobs = 20
	if ($OnlineNodes.Count -gt 0)
	{
		$poolSize = [Math]::Min($MaxConcurrentJobs, [Math]::Max(1, $OnlineNodes.Count))
		$pool = [RunspaceFactory]::CreateRunspacePool(1, $poolSize)
		$pool.Open()
		$tasks = @(foreach ($machineName in $OnlineNodes)
			{
				$ps = [PowerShell]::Create()
				$ps.RunspacePool = $pool
				$null = $ps.AddScript($readIniText).AddArgument($machineName).AddArgument($uvncFolders)
				[pscustomobject]@{ Machine = $machineName; PS = $ps; Handle = $ps.BeginInvoke() }
			})
		$deadline = (Get-Date).AddSeconds(120)
		$hung = $false
		foreach ($task in $tasks)
		{
			$left = [int][Math]::Max(0, ($deadline - (Get-Date)).TotalMilliseconds)
			if ($task.Handle.AsyncWaitHandle.WaitOne($left))
			{
				try
				{
					$r = @($task.PS.EndInvoke($task.Handle)) | Select-Object -Last 1
					if ($r -and $r.Machine) { $AllVNCPasswords[$r.Machine] = $r.Password }
				}
				catch { $AllVNCPasswords[$task.Machine] = $null }
				$task.PS.Dispose()
			}
			else
			{
				# Still stuck after 2 minutes: give up on it without blocking on the stop
				$AllVNCPasswords[$task.Machine] = $null
				try { $null = $task.PS.BeginStop($null, $null) }
				catch { }
				$hung = $true
			}
		}
		if ($hung) { try { $null = $pool.BeginClose($null, $null) } catch { } }
		else { try { $pool.Close(); $pool.Dispose() } catch { } }
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

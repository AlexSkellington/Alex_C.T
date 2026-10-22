# ===================================================================================================
#                              FUNCTION: Sync_Selected_Node_Hosts
# ---------------------------------------------------------------------------------------------------
# Description:
#   Pick any subset of lanes/backoffices, resolve their IP/hostname mapping, update the local
#   hosts file (replacing old entries when IPs changed), and copy the finished file to all
#   selected nodes. Always includes the current machine; custom mappings are appended after a
#   blank line; each IP maps to one hostname (last selected wins).
#
# Parameters:
#   - StoreNumber  [string]: store to process (required)
#
# Variables:
#   - $script:FunctionResults: LaneNumToMachineName, BackofficeNumToMachineName
#                              (run Retrieve_Nodes first)
#
# Notes:
#   - Local IPv4 selection ignores VPN/virtual/tunnel adapters and prefers
#     Ethernet/Wi-Fi with a gateway.
#   - Requires admin for hosts file writes and copies.
# ===================================================================================================

function Sync_Selected_Node_Hosts
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	Write_Log "`r`n==================== Starting Sync_Selected_Node_Hosts ====================`r`n" "blue"

	if (-not $script:FunctionResults['LaneNumToMachineName'] -or -not $script:FunctionResults['BackofficeNumToMachineName'])
	{
		Write_Log "Node mappings not found. Please run Retrieve_Nodes first." "red"
		return
	}

	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes @("Lane", "Backoffice")
	if (-not $selection)
	{
		Write_Log "Node selection cancelled." "yellow"
		return
	}
	$lanes = $selection.Lanes
	$backoffices = $selection.Backoffices

	$ipToHost = @{ }
	$changedMappings = @()
	$finalRows = @()

	# ---------------- Local machine details (robust local IPv4 selection) ----------------
	$localHostname = $env:COMPUTERNAME
	$localIp = $null
	try
	{
		# Get all UP NICs; exclude obvious VPN/virtual/tunnel/loopback by name/description/type
		$badRegex = '(?i)(vpn|virtual|vmware|hyper[- ]?v|vbox|virtualbox|loopback|hamachi|zerotier|tailscale|anyconnect|forti|juniper|pulse|citrix|openvpn|wireguard|proton|nord|ppp|remote)'
		$nics = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
		Where-Object { $_.OperationalStatus -eq [System.Net.NetworkInformation.OperationalStatus]::Up }

		$candidates = @()
		foreach ($nic in $nics)
		{
			if ($nic.Name -match $badRegex -or $nic.Description -match $badRegex) { continue }

			$type = $nic.NetworkInterfaceType.ToString()
			if ($type -notin @('Ethernet', 'Wireless80211')) { continue } # prefer physical/LAN

			$props = $nic.GetIPProperties()
			$hasGw = $false
			foreach ($gw in $props.GatewayAddresses)
			{
				if ($gw.Address -and $gw.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { $hasGw = $true; break }
			}

			foreach ($ua in $props.UnicastAddresses)
			{
				if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
				$ip = $ua.Address.ToString()
				if ($ip -eq '0.0.0.0' -or $ip.StartsWith('127.') -or $ip.StartsWith('169.254.')) { continue }

				$isPrivate = ($ip -match '^(10\.)|(192\.168\.)|(172\.(1[6-9]|2[0-9]|3[0-1])\.)')
				$score = 0
				if ($type -eq 'Ethernet') { $score += 100 }
				elseif ($type -eq 'Wireless80211') { $score += 80 }
				else { $score += 10 }
				if ($hasGw) { $score += 50 }
				if ($isPrivate) { $score += 20 }

				$candidates += [PSCustomObject]@{
					IP		   = $ip
					NicName    = $nic.Name
					Desc	   = $nic.Description
					Type	   = $type
					HasGateway = $hasGw
					IsPrivate  = $isPrivate
					Score	   = $score
				}
			}
		}

		if ($candidates.Count -gt 0)
		{
			$chosen = $candidates | Sort-Object Score -Descending | Select-Object -First 1
			$localIp = $chosen.IP
			Write_Log ("Selected local IP: {0} via NIC '{1}' ({2}); HasGateway={3}; Private={4}" -f `
				$chosen.IP, $chosen.NicName, $chosen.Type, $chosen.HasGateway, $chosen.IsPrivate) "gray"
		}
	}
	catch
	{
		# swallow; fallback below
	}

	# Fallbacks if the scoring path didn't produce an IP
	if (-not $localIp)
	{
		try { $localIp = ([System.Net.Dns]::GetHostAddresses($localHostname) | Where-Object { $_.AddressFamily -eq 'InterNetwork' -and -not $_.ToString().StartsWith('169.254.') -and -not $_.ToString().StartsWith('127.') } | Select-Object -ExpandProperty IPAddressToString -First 1) }
		catch { }
	}
	if (-not $localIp)
	{
		try { $localIp = (Test-Connection -ComputerName $localHostname -Count 1 -ErrorAction Stop | Select-Object -ExpandProperty IPV4Address | Select-Object -First 1) }
		catch { }
	}
	if (-not $localIp)
	{
		Write_Log "Could not resolve a suitable local IPv4 address for $localHostname. Aborting." "red"
		return
	}
	$ipToHost[$localIp] = $localHostname

	# ---------------- Add lanes/backoffices ----------------
	$nodes = @()
	foreach ($lane in $lanes)
	{
		$hostname = $script:FunctionResults['LaneNumToMachineName'][$lane]
		if ($hostname) { $nodes += @{ Hostname = $hostname; LaneNum = $lane; Type = 'Lane' } }
	}
	foreach ($bo in $backoffices)
	{
		$hostname = $script:FunctionResults['BackofficeNumToMachineName'][$bo]
		if ($hostname) { $nodes += @{ Hostname = $hostname; LaneNum = $bo; Type = 'Backoffice' } }
	}

	foreach ($n in $nodes)
	{
		if ($n.Hostname -eq $localHostname) { continue } # already added as local
		$ip = $null
		try
		{
			# Prefer DNS A-record style resolution; fall back to ping
			$addr = [System.Net.Dns]::GetHostAddresses($n.Hostname) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1
			if ($addr) { $ip = $addr.IPAddressToString }
		}
		catch { }
		if (-not $ip)
		{
			try { $ip = (Test-Connection -ComputerName $n.Hostname -Count 1 -ErrorAction Stop | Select-Object -ExpandProperty IPV4Address | Select-Object -First 1) }
			catch { }
		}
		if (-not $ip)
		{
			Write_Log "$($n.Type) $($n.LaneNum): Could not resolve IP for $($n.Hostname). Skipping." "red"
			continue
		}
		$ipToHost[$ip] = $n.Hostname
	}

	if ($ipToHost.Count -eq 0)
	{
		Write_Log "No valid IP/hostname mappings found for selected nodes." "yellow"
		return
	}

	# ---------------- Load previous hosts file mappings ----------------
	$hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
	$oldLines = if (Test-Path $hostsPath) { Get-Content $hostsPath -Raw }
	else { "" }
	$oldMappings = @{ }
	foreach ($line in $oldLines -split "`r?`n")
	{
		if ($line -match '^\s*([0-9\.]+)\s+(\S+)\s*$') { $oldMappings[$matches[2].ToLower()] = $matches[1] }
	}

	# Preserve all lines before the first bare "IP hostname" mapping, minus trailing blank lines.
	# (A List: trimming an array with [0..($n-2)] looped forever on a one-line section, and a
	# one-line section became a string that "+" then glued the mappings onto.)
	$defaultSection = New-Object System.Collections.Generic.List[string]
	foreach ($line in $oldLines -split "`r?`n")
	{
		if ($line -match '^\s*[0-9\.]+\s+\S+\s*$') { break }
		$defaultSection.Add($line)
	}
	while ($defaultSection.Count -gt 0 -and [string]::IsNullOrWhiteSpace($defaultSection[$defaultSection.Count - 1])) { $defaultSection.RemoveAt($defaultSection.Count - 1) }
	$outputLines = New-Object System.Collections.Generic.List[string]
	$outputLines.AddRange($defaultSection)
	$outputLines.Add('') # one blank line after the defaults

	# ---------------- Build custom mappings (local first, then others by IP) ----------------
	$orderedMappings = @()
	$orderedMappings += @{ IP = $localIp; Hostname = $localHostname }
	foreach ($ip in ($ipToHost.Keys | Sort-Object))
	{
		if ($ip -eq $localIp) { continue } # already first
		$orderedMappings += @{ IP = $ip; Hostname = $ipToHost[$ip] }
	}

	foreach ($entry in $orderedMappings)
	{
		$hn = $entry.Hostname
		$ip = $entry.IP
		$oldIp = $oldMappings[$hn.ToLower()]
		$rowColor = if ($oldIp -and $oldIp -ne $ip) { "yellow" }
		else { "green" }
		$finalRows += @{ IP = $ip; Hostname = $hn; Color = $rowColor }
		$outputLines.Add("$ip`t$hn")
		if ($rowColor -eq "yellow") { $changedMappings += "$hn ($oldIp => $ip)" }
	}

	# ---------------- Write hosts file (local) ----------------
	# Stop here when it cannot be written (not elevated): copying would push the OLD file out
	try { Set-Content -Path $hostsPath -Value $outputLines.ToArray() -Encoding ascii -ErrorAction Stop }
	catch
	{
		Write_Log "Could not write the local hosts file ($hostsPath): $($_.Exception.Message). Run as Administrator. Nothing was copied." "red"
		Write_Log "`r`n==================== Sync_Selected_Node_Hosts Completed ====================" "blue"
		return
	}

	# ---------------- Copy hosts file to selected nodes (skip local) ----------------
	foreach ($entry in $orderedMappings)
	{
		$hn = $entry.Hostname
		$ip = $entry.IP
		if ($hn -eq $localHostname) { continue } # Don't network-copy to self
		$targetPath = "\\$hn\C$\Windows\System32\drivers\etc\hosts"
		try
		{
			Copy-Item -Path $hostsPath -Destination $targetPath -Force
			Write_Log "Copied hosts file to $hn [$ip]" "cyan"
		}
		catch
		{
			Write_Log "Failed to copy hosts file to $hn [$ip]: $_" "red"
		}
	}

	# ---------------- Table Output ----------------
	Write_Log "`r`nHost file mappings (sorted by IP, local host first):" "blue"
	$maxIpLen = ($finalRows | ForEach-Object { "$($_.IP)".Length } | Measure-Object -Maximum).Maximum
	$maxHostLen = ($finalRows | ForEach-Object { "$($_.Hostname)".Length } | Measure-Object -Maximum).Maximum
	Write_Log ("IP".PadRight($maxIpLen + 2) + "Hostname".PadRight($maxHostLen + 2) + "Changed") "blue"
	foreach ($r in $finalRows)
	{
		$line = "$($r.IP)".PadRight($maxIpLen + 2) + "$($r.Hostname)".PadRight($maxHostLen + 2)
		if ($r.Color -eq "yellow") { $line += "<CHANGED>" }
		Write_Log $line $r.Color
	}
	if ($changedMappings.Count)
	{
		Write_Log "Updated the following mappings:" "yellow"
		foreach ($chg in $changedMappings) { Write_Log "  $chg" "yellow" }
	}
	else
	{
		Write_Log "No mappings changed from previous hosts file." "green"
	}
	Write_Log "`r`n==================== Sync_Selected_Node_Hosts Completed ====================" "blue"
}

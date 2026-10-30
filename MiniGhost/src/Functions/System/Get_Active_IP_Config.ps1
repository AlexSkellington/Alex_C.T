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
			
			# Check adapter status via Win32_NetworkAdapter (NetConnectionStatus 2 = Connected)
			$ad = Get-WmiObject -Class Win32_NetworkAdapter -Filter "DeviceID = $idx" -ErrorAction SilentlyContinue
			if ($ad -and $ad.NetConnectionStatus -eq 2)
			{
				# Get config via Win32_NetworkAdapterConfiguration
				$wmiCfg = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "Index = $idx" -ErrorAction SilentlyContinue
				if ($wmiCfg -and $wmiCfg.IPEnabled -and $wmiCfg.DefaultIPGateway)
				{
					# Filter for valid IPv4 (non-APIPA)
					$validIPs = $wmiCfg.IPAddress | Where-Object { $_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' -and $_ -notmatch '^169\.254\.' }
					if ($validIPs)
					{
						# Construct a custom object similar to Get-NetIPConfiguration for consistency
						$ipv4Address = [PSCustomObject]@{
							IPAddress    = $validIPs[0]
							PrefixLength = ($wmiCfg.IPSubnet | Where-Object { $_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' })[0] | ForEach-Object {
								[int]([math]::Log([uint32]::MaxValue -bxor ([ipaddress]$_).Address, 2)) + 1
							}
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

# ---------------------------------------------------------------------------------------------------
# Lane SQL protocol settings - shared by both ways of enabling the lane SQL protocols:
#   - Enable_SQL_Protocols_On_Selected_Lanes (menu): runspace worker, remote registry of the lane
#   - Ensure_SQL_Enable_Scheduled_Task_For_File_Lane: the 4:15 AM task script, local registry
#   Embedded at build time right after the param() line of both scripts, so the two can never
#   drift apart again. Defines two scriptblocks:
#
#   & $SetSqlProtocolSettings <HKLM key> <tcp port>
#       For every SQL instance: mixed-mode logins (LoginMode=2, the Tecnica SQL-login fallback
#       needs it), TCP on with a static port and no dynamic ports, Named Pipes on, Shared Memory
#       on. Returns one object per instance: Instance, Service, Changed, Lines (Text / Color).
#
#   & $RestartSqlService <service> [<machine>]
#       Stops the service, waits for STOPPED, starts it, waits for RUNNING (60 s each at most).
#       No machine = this computer. Returns $true when the service reports RUNNING.
# ---------------------------------------------------------------------------------------------------
$SetSqlProtocolSettings = {
	param ([Microsoft.Win32.RegistryKey]$Hklm, [string]$TcpPort)

	$instances = @{ }
	foreach ($root in @('SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL', 'SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL'))
	{
		$k = $Hklm.OpenSubKey($root)
		if ($k)
		{
			foreach ($n in $k.GetValueNames())
			{
				$id = $k.GetValue($n)
				if ($id -and -not $instances.ContainsKey($n)) { $instances[$n] = $id }
			}
			$k.Close()
		}
	}

	foreach ($instanceName in $instances.Keys)
	{
		$instanceId = $instances[$instanceName]
		$lines = New-Object System.Collections.Generic.List[object]
		$changed = $false

		# Name, sub key, value, wanted value (each found in the 64-bit or the Wow6432Node hive)
		$settings = @(
			@('Mixed Mode Authentication', 'MSSQLServer', 'LoginMode', 2),
			@('TCP/IP protocol', 'MSSQLServer\SuperSocketNetLib\Tcp', 'Enabled', 1),
			@('Static TCP port', 'MSSQLServer\SuperSocketNetLib\Tcp\IPAll', 'TcpPort', $TcpPort),
			@('Named Pipes protocol', 'MSSQLServer\SuperSocketNetLib\Np', 'Enabled', 1),
			@('Shared Memory protocol', 'MSSQLServer\SuperSocketNetLib\Sm', 'Enabled', 1)
		)
		foreach ($s in $settings)
		{
			$label = $s[0]
			$found = $false
			foreach ($base in @('SOFTWARE\Microsoft\Microsoft SQL Server', 'SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server'))
			{
				$path = "$base\$instanceId\$($s[1])"
				$k = $Hklm.OpenSubKey($path, $true)
				if (-not $k) { continue }
				$found = $true
				try
				{
					if ($s[2] -eq 'TcpPort')
					{
						# Static port: TcpPort = port, TcpDynamicPorts empty
						$curPort = [string]$k.GetValue('TcpPort', '')
						$curDyn = [string]$k.GetValue('TcpDynamicPorts', '')
						if ($curPort -ne $TcpPort -or $curDyn -ne '')
						{
							$k.SetValue('TcpPort', $TcpPort, [Microsoft.Win32.RegistryValueKind]::String)
							$k.SetValue('TcpDynamicPorts', '', [Microsoft.Win32.RegistryValueKind]::String)
							$lines.Add([pscustomobject]@{ Text = "$label set to $TcpPort (was '$curPort', dynamic '$curDyn')."; Color = 'green' })
							$changed = $true
						}
						else { $lines.Add([pscustomobject]@{ Text = "$label already $TcpPort."; Color = 'gray' }) }
					}
					elseif ([int]$k.GetValue($s[2], 0) -ne $s[3])
					{
						$k.SetValue($s[2], $s[3], [Microsoft.Win32.RegistryValueKind]::DWord)
						$lines.Add([pscustomobject]@{ Text = "$label enabled."; Color = 'green' })
						$changed = $true
					}
					else { $lines.Add([pscustomobject]@{ Text = "$label already enabled."; Color = 'gray' }) }
				}
				finally { $k.Close() }
				break
			}
			if (-not $found) { $lines.Add([pscustomobject]@{ Text = "$label - registry key not found for $instanceId."; Color = 'yellow' }) }
		}

		$svc = if ($instanceName -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { "MSSQL`$$instanceName" }
		[pscustomobject]@{ Instance = $instanceName; Service = $svc; Changed = $changed; Lines = $lines.ToArray() }
	}
}

$RestartSqlService = {
	param ([string]$Service, [string]$Machine)

	$target = @()
	if (-not [string]::IsNullOrWhiteSpace($Machine)) { $target = @("\\$Machine") }
	# sc.exe returns at once; poll for STOPPED (1) / RUNNING (4)
	$null = & sc.exe @target stop $Service
	for ($i = 0; $i -lt 60; $i++)
	{
		if ((& sc.exe @target query $Service | Out-String) -match 'STATE\s+:\s+1\s') { break }
		Start-Sleep -Seconds 1
	}
	$null = & sc.exe @target start $Service
	for ($i = 0; $i -lt 60; $i++)
	{
		if ((& sc.exe @target query $Service | Out-String) -match 'STATE\s+:\s+4\s') { return $true }
		Start-Sleep -Seconds 1
	}
	return $false
}

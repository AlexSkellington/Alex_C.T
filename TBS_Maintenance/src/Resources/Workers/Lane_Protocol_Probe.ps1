# ---------------------------------------------------------------------------------------------------
# Lane SQL protocol probe - runspace worker for Start_Lane_Protocol_Jobs.
#   Embedded at build time into $script:LaneProtocolWorker; runs in the lane-protocol runspace pool
#   (no access to the main script's variables or functions).
#   Ping (700 ms) -> Offline when unreachable; then SQL login over tcp: -> 'TCP', over np: ->
#   'Named Pipes', else 'File' (the lane only works through the SQI file drop).
# ---------------------------------------------------------------------------------------------------
param([string]$machine, [string]$lane)

Add-Type -AssemblyName System.Data 2>$null
Add-Type -AssemblyName System 2>$null

$protocol = 'File'
$reachable = $true

if (-not [string]::IsNullOrWhiteSpace($machine))
{
	$machineNorm = [string]$machine
	$isLocalMachine = $false
	try
	{
		if ($machineNorm -match '^(?i)(localhost|\.|\(local\))$' -or $machineNorm -ieq $env:COMPUTERNAME) { $isLocalMachine = $true }
	}
	catch { }

	if (-not $isLocalMachine)
	{
		$reachable = $false
		try
		{
			$ping = New-Object System.Net.NetworkInformation.Ping
			$reply = $ping.Send($machineNorm, 700)
			if ($reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) { $reachable = $true }
		}
		catch { }
	}
}

if (-not $reachable)
{
	[PSCustomObject]@{ Lane = $lane; Protocol = 'Offline' }
	return
}

foreach ($probe in @(@{ Prefix = 'tcp'; Name = 'TCP' }, @{ Prefix = 'np'; Name = 'Named Pipes' }))
{
	$cn = $null
	try
	{
		$cs = "Data Source=$($probe.Prefix):$machine;Initial Catalog=master;Integrated Security=True;Connect Timeout=1"
		$cn = New-Object System.Data.SqlClient.SqlConnection $cs
		$cn.Open()
		$protocol = $probe.Name
	}
	catch { }
	finally
	{
		try { if ($cn) { $cn.Close() } } catch { }
		try { if ($cn) { $cn.Dispose() } } catch { }
	}
	if ($protocol -ne 'File') { break }
}

[PSCustomObject]@{ Lane = $lane; Protocol = $protocol }

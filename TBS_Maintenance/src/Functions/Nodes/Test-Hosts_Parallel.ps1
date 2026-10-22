# ===================================================================================================
#                              FUNCTION: Test-Hosts_Parallel
# ---------------------------------------------------------------------------------------------------
# Description:
#   Probes a list of hostnames/IPs in parallel with a per-probe timeout and returns
#   target -> 'OK' / 'FAIL' / 'ERROR'. Names are resolved with async DNS, then every address
#   is pinged with Ping.SendPingAsync, all in flight together (in batches), so a store full of
#   dead hosts costs about one timeout instead of one timeout per host.
#
# Parameters:
#   - Targets        [string[]]: hostnames or IPv4 addresses to probe (required, pipeline)
#   - TimeoutMs      [int]: per-probe timeout in milliseconds (default 400)
#   - MaxConcurrency [int]: max probes in flight per batch (default 64)
#
# Returns:
#   Hashtable target -> 'OK' (reachable), 'FAIL' (no reply in time) or 'ERROR' (name did
#   not resolve / probe failed).
#
# Notes:
#   - The previous Parallel.ForEach version ran a PowerShell scriptblock on thread-pool
#     threads, which have no runspace: it always threw and fell back to a serial
#     Test-Connection loop (8 dead hosts = 32 seconds).
# ===================================================================================================

function Test-Hosts_Parallel
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[string[]]$Targets,
		[int]$TimeoutMs = 400,
		[int]$MaxConcurrency = 64
	)

	begin { $all = New-Object System.Collections.Generic.List[string] }
	process { foreach ($t in $Targets) { if (-not [string]::IsNullOrWhiteSpace($t) -and -not $all.Contains($t)) { $all.Add($t) } } }
	end
	{
		$out = @{ }
		if ($all.Count -eq 0) { return $out }

		# Wait for a set of tasks without throwing (WaitAll throws when any task faulted)
		$waitTasks = {
			param ($Tasks, [int]$Ms)
			if (@($Tasks).Count -eq 0) { return }
			try { [void][System.Threading.Tasks.Task]::WaitAll([System.Threading.Tasks.Task[]]@($Tasks), $Ms) }
			catch { }
		}

		$batchSize = [Math]::Max(1, $MaxConcurrency)
		for ($start = 0; $start -lt $all.Count; $start += $batchSize)
		{
			$batch = $all.GetRange($start, [Math]::Min($batchSize, $all.Count - $start))

			# 1) Resolve names (IPs are used as-is)
			$address = @{ }
			$dns = @{ }
			foreach ($t in $batch)
			{
				$ip = $null
				if ([System.Net.IPAddress]::TryParse($t, [ref]$ip)) { $address[$t] = $ip }
				else
				{
					try { $dns[$t] = [System.Net.Dns]::GetHostAddressesAsync($t) }
					catch { $out[$t] = 'ERROR' }
				}
			}
			& $waitTasks $dns.Values ([Math]::Max(1500, $TimeoutMs * 3))
			foreach ($t in $dns.Keys)
			{
				$task = $dns[$t]
				$ipv4 = $null
				if ($task.Status -eq 'RanToCompletion') { $ipv4 = $task.Result | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1 }
				if ($ipv4) { $address[$t] = $ipv4 }
				else { $out[$t] = 'ERROR' }
			}

			# 2) Ping every resolved address at once (one Ping object per probe)
			$pings = @{ }
			foreach ($t in $address.Keys)
			{
				$p = New-Object System.Net.NetworkInformation.Ping
				try { $pings[$t] = @{ Ping = $p; Task = $p.SendPingAsync($address[$t], $TimeoutMs) } }
				catch { $p.Dispose(); $out[$t] = 'ERROR' }
			}
			& $waitTasks @($pings.Values | ForEach-Object { $_.Task }) ($TimeoutMs + 1000)
			foreach ($t in $pings.Keys)
			{
				$task = $pings[$t].Task
				if ($task.Status -eq 'RanToCompletion') { $out[$t] = if ($task.Result.Status -eq 'Success') { 'OK' } else { 'FAIL' } }
				elseif ($task.Status -eq 'Faulted') { $out[$t] = 'ERROR' }
				else { $out[$t] = 'FAIL' }
				try { $pings[$t].Ping.Dispose() }
				catch { }
			}
		}

		return $out
	}
}

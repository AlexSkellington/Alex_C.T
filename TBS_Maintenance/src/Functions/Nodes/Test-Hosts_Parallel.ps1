# ===================================================================================================
#                              FUNCTION: Test-Hosts_Parallel
# ---------------------------------------------------------------------------------------------------
# Description:
#   Probes a list of hostnames/IPs in parallel with a per-probe timeout and a concurrency
#   cap, and returns target -> 'OK' / 'FAIL' / 'ERROR'. Replaces serial Test-Connection
#   loops that could freeze the UI for minutes when several hosts are down. Falls back to a
#   serial Test-Connection loop if the parallel path is unavailable.
#
# Parameters:
#   - Targets        [string[]]: hostnames or IPv4 addresses to probe (required, pipeline)
#   - TimeoutMs      [int]: per-probe timeout in milliseconds (default 400)
#   - MaxConcurrency [int]: max probes in flight (default 10)
#
# Returns:
#   Hashtable target -> 'OK' (reachable), 'FAIL' (no reply in time) or 'ERROR'.
# ===================================================================================================

function Test-Hosts_Parallel
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[string[]]$Targets,
		[int]$TimeoutMs = 400,
		[int]$MaxConcurrency = 10
	)
	
	begin { $all = @() }
	process { $all += @($Targets) }
	end
	{
		$all = @($all | Where-Object { $_ })
		$out = @{}
		if ($all.Count -eq 0) { return $out }
		
		$done = $false
		try
		{
			$opts = New-Object System.Threading.Tasks.ParallelOptions
			$opts.MaxDegreeOfParallelism = [Math]::Min([Math]::Max(1, $MaxConcurrency), $all.Count)
			[void][System.Threading.Tasks.Parallel]::ForEach($all, $opts, [Action[string]] {
				param($t)
				$ping = New-Object System.Net.NetworkInformation.Ping
				try
				{
					$r = $ping.Send($t, $TimeoutMs)
					if ($r.Status -eq [System.Net.IPStatus]::Success) { $out[$t] = 'OK' } else { $out[$t] = 'FAIL' }
				}
				catch { $out[$t] = 'ERROR' }
				finally { $ping.Dispose() }
			})
			$done = $true
		}
		catch { $done = $false }
		
		if (-not $done)
		{
			# Serial fallback (previous behavior)
			foreach ($t in $all)
			{
				try
				{
					if (Test-Connection -ComputerName $t -Count 1 -Quiet -ErrorAction Stop) { $out[$t] = 'OK' }
					else { $out[$t] = 'FAIL' }
				}
				catch { $out[$t] = 'ERROR' }
			}
		}
		
		return $out
	}
}

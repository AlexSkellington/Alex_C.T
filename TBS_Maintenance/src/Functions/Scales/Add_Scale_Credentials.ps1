# ===================================================================================================
#                              FUNCTION: Add_Scale_Credentials
# ---------------------------------------------------------------------------------------------------
# Description:
#   Pre-seeds Windows credentials for every scale IP so \\<scaleIP>\c$ is accessible later.
#   For each unique scale IP (in parallel runspaces): skips if the share is already
#   accessible or TCP 445 is unreachable, otherwise tries the stored credentials in order
#   with "net use" and persists the working one with cmdkey. Silent - no logging.
#
# Parameters:
#   - ScaleCodeToIPInfo [hashtable]: scale code -> info with FullIP / IPNetwork+IPDevice (required)
#   - MaxParallel       [int]: max parallel credential tasks (default 16)
#
# Variables:
#   - $script:credBizerba / $script:credBiyerba: PSCredential objects tried in order
#   - $script:ScaleCredTasks: running task handles (kept so they are not GC'd)
#
# Writes:
#   - stored Windows credentials (cmdkey) for the working scale login
# ===================================================================================================

function Add_Scale_Credentials
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[hashtable]$ScaleCodeToIPInfo,
		[int]$MaxParallel = 16
	)

	if (-not $script:ScaleCredTasks -or -not ($script:ScaleCredTasks -is [hashtable])) { $script:ScaleCredTasks = @{ } }
	if ($MaxParallel -lt 1) { $MaxParallel = 1 }

	# Reap any completed tasks from prior calls before queuing more work.
	foreach ($k in @($script:ScaleCredTasks.Keys))
	{
		$st = $script:ScaleCredTasks[$k]
		$completed = $true
		try { $completed = (-not $st.Handle) -or $st.Handle.IsCompleted }
		catch { $completed = $true }

		if ($completed)
		{
			try { if ($st.Handle -and $st.PS) { $null = $st.PS.EndInvoke($st.Handle) } }
			catch { }
			try { if ($st.PS) { $st.PS.Dispose() } }
			catch { }
			[void]$script:ScaleCredTasks.Remove($k)
		}
	}

	# Build candidate credential list once so we don't duplicate literals in the worker.
	$credentialAttempts = @()
	$seenCreds = @{ }
	foreach ($cred in @($script:credBizerba, $script:credBiyerba))
	{
		if ($cred -is [System.Management.Automation.PSCredential])
		{
			$user = $null
			$pass = $null
			try { $user = [string]$cred.UserName }
			catch { $user = $null }
			try { $pass = [string]$cred.GetNetworkCredential().Password }
			catch { $pass = $null }

			if (-not [string]::IsNullOrWhiteSpace($user) -and -not [string]::IsNullOrEmpty($pass))
			{
				$key = "$user`n$pass"
				if (-not $seenCreds.ContainsKey($key))
				{
					$seenCreds[$key] = $true
					$credentialAttempts += [pscustomobject]@{
						User = $user
						Pass = $pass
					}
				}
			}
		}
	}

	# Safety fallback: preserve previous behavior even if the script-scoped credentials were not initialized.
	if (-not $credentialAttempts -or $credentialAttempts.Count -eq 0)
	{
		$credentialAttempts = @(
			[pscustomobject]@{ User = 'bizuser'; Pass = 'bizerba' },
			[pscustomobject]@{ User = 'bizuser'; Pass = 'biyerba' }
		)
	}

	# Collect unique, valid IPs only.
	$scaleIPs = @(
		$ScaleCodeToIPInfo.Values | ForEach-Object {
			$candidate = $null
			if ($_.FullIP) { $candidate = [string]$_.FullIP }
			elseif ($_.IPNetwork -and $_.IPDevice)
			{
				# "192.168.1" + "50" must not become "192.168.150" (TryParse reads that as 192.168.0.150)
				$candidate = ([string]$_.IPNetwork).Trim().TrimEnd('.') + '.' + ([string]$_.IPDevice).Trim()
			}

			if (-not [string]::IsNullOrWhiteSpace($candidate))
			{
				$candidate = $candidate.Trim()
				$parsedIp = $null
				if ([System.Net.IPAddress]::TryParse($candidate, [ref]$parsedIp))
				{
					$parsedIp.IPAddressToString
				}
			}
		} | Select-Object -Unique
	)

	if ($scaleIPs.Count -eq 0) { return }

	# Shared pool: reuse while active, recreate when idle and the requested width changed.
	$needNewPool = $false
	$hasOpenPool = $false
	try
	{
		if ($script:ScaleCredPool -and $script:ScaleCredPool.RunspacePoolStateInfo.State -eq 'Opened')
		{
			$hasOpenPool = $true
		}
	}
	catch
	{
		$hasOpenPool = $false
	}

	if (-not $hasOpenPool)
	{
		$needNewPool = $true
	}
	elseif (($script:ScaleCredPoolMaxParallel -ne $MaxParallel) -and $script:ScaleCredTasks.Count -eq 0)
	{
		try { $script:ScaleCredPool.Close() }
		catch { }
		try { $script:ScaleCredPool.Dispose() }
		catch { }
		$script:ScaleCredPool = $null
		$needNewPool = $true
	}

	if ($needNewPool)
	{
		$script:ScaleCredPool = [runspacefactory]::CreateRunspacePool(1, $MaxParallel)
		$script:ScaleCredPool.ApartmentState = 'MTA'
		$script:ScaleCredPool.Open()
		$script:ScaleCredPoolMaxParallel = $MaxParallel
	}

	# Auto-reap completed tasks so PS objects/pool get cleaned up even if this function is only called once.
	if (-not $script:ScaleCredReaper)
	{
		$script:ScaleCredReaper = New-Object System.Windows.Forms.Timer
		$script:ScaleCredReaper.Interval = 2000
		$script:ScaleCredReaper.Add_Tick({
				try
				{
					if ($script:ScaleCredTasks)
					{
						foreach ($key in @($script:ScaleCredTasks.Keys))
						{
							$task = $script:ScaleCredTasks[$key]
							$done = $true
							try { $done = (-not $task.Handle) -or $task.Handle.IsCompleted }
							catch { $done = $true }

							if ($done)
							{
								try { if ($task.Handle -and $task.PS) { $null = $task.PS.EndInvoke($task.Handle) } }
								catch { }
								try { if ($task.PS) { $task.PS.Dispose() } }
								catch { }
								[void]$script:ScaleCredTasks.Remove($key)
							}
						}
					}

					if ((-not $script:ScaleCredTasks) -or $script:ScaleCredTasks.Count -eq 0)
					{
						try
						{
							if ($script:ScaleCredPool)
							{
								$script:ScaleCredPool.Close()
								$script:ScaleCredPool.Dispose()
							}
						}
						catch { }
						$script:ScaleCredPool = $null
						$script:ScaleCredPoolMaxParallel = $null
						if ($script:ScaleCredReaper) { $script:ScaleCredReaper.Stop() }
					}
				}
				catch { }
			})
	}
	if ($script:ScaleCredReaper -and -not $script:ScaleCredReaper.Enabled) { $script:ScaleCredReaper.Start() }

	$worker = {
		param (
			[string]$ip,
			[object[]]$CredentialAttempts
		)

		$sharePath = "\\$ip\c$"

		# ---- quick SMB reachability (port 445) ----
		$smbOk = $false
		$client = $null
		$waitHandle = $null
		try
		{
			$client = New-Object System.Net.Sockets.TcpClient
			$ar = $client.BeginConnect($ip, 445, $null, $null)
			$waitHandle = $ar.AsyncWaitHandle
			if ($waitHandle.WaitOne(400))
			{
				try { $client.EndConnect($ar) }
				catch { }
				if ($client.Connected) { $smbOk = $true }
			}
		}
		catch { }
		finally
		{
			try { if ($waitHandle) { $waitHandle.Close() } }
			catch { }
			try { if ($client) { $client.Close() } }
			catch { }
		}

		if (-not $smbOk) { return }

		# ---- already accessible without creds? ----
		try
		{
			if (Test-Path -LiteralPath $sharePath) { return }
		}
		catch { }

		# ---- try credentials quickly via NET USE (no persistence) ----
		foreach ($c in $CredentialAttempts)
		{
			$user = $null
			$pass = $null
			try { $user = [string]$c.User }
			catch { $user = $null }
			try { $pass = [string]$c.Pass }
			catch { $pass = $null }

			if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrEmpty($pass)) { continue }

			# Clear any lingering connection before each attempt so we don't inherit a stale state.
			try { $null = & net.exe use $sharePath '/delete' '/y' 2>$null }
			catch { }

			$mapRc = -1
			try
			{
				$null = & net.exe use $sharePath "/user:$user" $pass '/persistent:no' 2>$null
				$mapRc = $LASTEXITCODE
			}
			catch { $mapRc = -1 }

			if ($mapRc -eq 0)
			{
				try { $null = & net.exe use $sharePath '/delete' '/y' 2>$null }
				catch { }

				$cmdKeyRc = -1
				try
				{
					$null = & cmdkey.exe "/add:$ip" "/user:$user" "/pass:$pass" 2>$null
					$cmdKeyRc = $LASTEXITCODE
				}
				catch { $cmdKeyRc = -1 }

				if ($cmdKeyRc -eq 0)
				{
					# Best-effort verification; even if Windows does not immediately reuse the credential,
					# a successful CMDKEY add is enough to stop retrying.
					try { $null = Test-Path -LiteralPath $sharePath }
					catch { }
					return
				}
			}

			try { $null = & net.exe use $sharePath '/delete' '/y' 2>$null }
			catch { }
		}
	}

	foreach ($ip in $scaleIPs)
	{
		$existingTask = $script:ScaleCredTasks[$ip]
		if ($existingTask)
		{
			$inFlight = $false
			try { $inFlight = $existingTask.Handle -and -not $existingTask.Handle.IsCompleted }
			catch { $inFlight = $false }

			if ($inFlight) { continue }

			try { if ($existingTask.Handle -and $existingTask.PS) { $null = $existingTask.PS.EndInvoke($existingTask.Handle) } }
			catch { }
			try { if ($existingTask.PS) { $existingTask.PS.Dispose() } }
			catch { }
			[void]$script:ScaleCredTasks.Remove($ip)
		}

		$ps = [powershell]::Create()
		$null = $ps.AddScript($worker.ToString()).AddArgument($ip).AddArgument($credentialAttempts)

		$ps.RunspacePool = $script:ScaleCredPool
		try
		{
			$handle = $ps.BeginInvoke()

			# Stash to prevent GC and allow cleanup/busy tracking.
			$script:ScaleCredTasks[$ip] = @{
				PS	    = $ps
				Handle  = $handle
				Started = Get-Date
			}
		}
		catch
		{
			try { $ps.Dispose() }
			catch { }
		}
	}

	# Note: We do not wait. This function returns immediately while tasks run in background.
}

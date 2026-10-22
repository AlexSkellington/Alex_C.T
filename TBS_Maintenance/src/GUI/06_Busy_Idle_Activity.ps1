# ------------------------------------------------------------------------------------------------
# --- Busy scanner (UI thread) -------------------------------------------------------------
#   (kept logic; additionally updates status strip for a modern "app" feel)
# ------------------------------------------------------------------------------------------------
$BusyTimer = New-Object System.Windows.Forms.Timer
$BusyTimer.Interval = 1500
$BusyTimer.add_Tick({
		# Reset busy flags each tick
		$script:IsBusy = $false
		$script:BusyReason = $null
		
		# 1) Lane protocol runspace tasks?
		try
		{
			if ($script:LaneProtocolJobs)
			{
				foreach ($kv in @($script:LaneProtocolJobs.GetEnumerator()))
				{
					$st = $kv.Value
					if ($st -and $st.Handle -and -not $st.Handle.IsCompleted)
					{
						$script:IsBusy = $true
						$script:BusyReason = "LaneProtocol runspace ($($kv.Key))"
						break
					}
				}
			}
		}
		catch { }
		
		# 2) Lane protocol pool in-flight work?
		if (-not $script:IsBusy)
		{
			try
			{
				if ($script:LaneProtocolPool)
				{
					$avail = $script:LaneProtocolPool.GetAvailableRunspaces()
					$max = $script:LaneProtocolPool.MaxRunspaces
					if ($avail -lt $max)
					{
						$script:IsBusy = $true
						$script:BusyReason = "LaneProtocol pool busy ($avail/$max available)"
					}
				}
			}
			catch { }
		}
		
		# 3) Scale credential helper tasks?
		if (-not $script:IsBusy)
		{
			try
			{
				if ($script:ScaleCredTasks)
				{
					foreach ($k in @($script:ScaleCredTasks.Keys))
					{
						$st = $script:ScaleCredTasks[$k]
						if ($st -and $st.Handle -and -not $st.Handle.IsCompleted)
						{
							$script:IsBusy = $true
							$script:BusyReason = "ScaleCred task ($k)"
							break
						}
					}
				}
			}
			catch { }
		}
		
		# 4) Known launch background jobs?
		if (-not $script:IsBusy)
		{
			try
			{
				if (
					$script:VncExportTask -and
					$script:VncExportTask.Job -and
					$script:VncExportTask.Job.State -in @('Running', 'NotStarted')
				)
				{
					$script:IsBusy = $true
					$script:BusyReason = "Exporting VNC files"
				}
			}
			catch { }
		}
		
		if (-not $script:IsBusy)
		{
			try
			{
				if (
					$script:RemoteMachineInfoTask -and
					$script:RemoteMachineInfoTask.Job -and
					$script:RemoteMachineInfoTask.Job.State -in @('Running', 'NotStarted')
				)
				{
					$script:IsBusy = $true
					$script:BusyReason = "Collecting remote machine info"
				}
			}
			catch { }
		}
		
		if (-not $script:IsBusy)
		{
			try
			{
				$insertTestJob = Get-Job -Name 'InsertTestItemJob' -State Running -ErrorAction SilentlyContinue | Select-Object -First 1
				if ($insertTestJob)
				{
					$script:IsBusy = $true
					$script:BusyReason = "Loading test item"
				}
			}
			catch { }
		}
		
		if (-not $script:IsBusy)
		{
			try
			{
				$laneSqlEnableJob = $null
				if ($script:LaneSqlProtocolEnableJobs)
				{
					$laneSqlEnableJob = @($script:LaneSqlProtocolEnableJobs.GetEnumerator() | Where-Object {
						$_.Value -and $_.Value.Process -and (-not $_.Value.Process.HasExited)
					} | Select-Object -First 1)
				}
				if ($laneSqlEnableJob)
				{
					$script:IsBusy = $true
					$laneBusyText = [string]$laneSqlEnableJob.Key
					$script:BusyReason = "Scheduling SQL enable ($laneBusyText)"
				}
			}
			catch { }
		}
		
		# 5) Any other background jobs (EXCEPT ignored names/patterns)
		if (-not $script:IsBusy)
		{
			try
			{
				$jobs = Get-Job -State Running -ErrorAction SilentlyContinue | Where-Object {
					$nm = $_.Name
					$ignore = $false
					foreach ($pat in $script:IdleIgnoreJobNames)
					{
						if ($nm -like $pat) { $ignore = $true; break }
					}
					-not $ignore
				}
				
				if ($jobs)
				{
					$script:IsBusy = $true
					$script:BusyReason = "Job: $($jobs[0].Name)"
				}
			}
			catch { }
		}
		
		# ---- Status strip update (modern) ----
		try
		{
			# Use the status helper function
			if (-not [string]::IsNullOrWhiteSpace($script:StartupWarmupStatus))
			{
				& $updateStatusBar -Status "Warmup" -StatusColor $uiStatusInfo -Operation $script:StartupWarmupStatus
			}
			elseif ($script:IsBusy)
			{
				& $updateStatusBar -Status "Running" -StatusColor $uiStatusWarning -Operation $script:BusyReason -BusyInfo "Working..."
			}
			else
			{
				& $updateStatusBar -Status "Ready" -StatusColor $uiStatusSuccess -Operation "Waiting for action..."
			}
		}
		catch { }
	})
$BusyTimer.Start()

# ------------------------------------------------------------------------------------------------
# --- Idle watchdog (UI thread) ------------------------------------------------------------
#   (kept logic; additionally shows countdown in status strip)
# ------------------------------------------------------------------------------------------------
$script:IdleTimer = New-Object System.Windows.Forms.Timer
$script:IdleTimer.Interval = 30000 # 30s
$script:IdleTimer.add_Tick({
		if ($script:IdleMinutesAllowed -le 0)
		{
			try { $stIdle.Text = "" }
			catch { }
			return
		}
		
		try
		{
			$minsIdle = (New-TimeSpan -Start $script:LastActivity -End (Get-Date)).TotalMinutes
			$hardCloseReady = ($script:IdleHardCloseAfterMinutes -gt 0 -and $minsIdle -ge $script:IdleHardCloseAfterMinutes)
			
			# Status strip idle countdown
			try
			{
				$left = [math]::Max(0, $script:IdleMinutesAllowed - $minsIdle)
				$stIdle.Text = ("Idle close in: {0:n1} min" -f $left)
			}
			catch { }
			
			if ($minsIdle -ge $script:IdleMinutesAllowed)
			{
				if (-not $script:IsBusy -or $hardCloseReady)
				{
					try
					{
						if ($hardCloseReady -and $script:IsBusy -and $script:BusyReason)
						{
							Write_Log "Idle $([math]::Round($minsIdle, 1)) min; overriding busy ($script:BusyReason) due to hard-close window." "yellow"
						}
						else
						{
							Write_Log "Idle for $([math]::Round($minsIdle, 1)) minute(s) - closing." "yellow"
						}
					}
					catch { }
					
					$script:SuppressClosePrompt = $true
					if ($form -and -not $form.IsDisposed) { $form.Close() }
				}
				else
				{
					# Busy detected at idle threshold; log once then reset window
					try
					{
						if ($script:BusyReason)
						{
							Write_Log "Idle $([math]::Round($minsIdle, 1)) min, but busy: $script:BusyReason (not closing)." "gray"
						}
						else
						{
							Write_Log "Idle $([math]::Round($minsIdle, 1)) min, but busy: <unknown> (not closing)." "gray"
						}
					}
					catch { }
					
					# Reset so we don't spam + so user gets another full idle window
					$script:LastActivity = Get-Date
				}
			}
		}
		catch { }
	})
$script:IdleTimer.Start()

# ------------------------------------------------------------------------------------------------
# --- Activity hooks (form + all existing child controls) ---------------------------------
#   (kept logic; plus hooks future-added controls to avoid "dead" activity areas)
# ------------------------------------------------------------------------------------------------
try { $form.add_KeyDown({ $script:LastActivity = Get-Date }) }
catch { }
try { $form.add_MouseMove({ $script:LastActivity = Get-Date }) }
catch { }

# Hook existing controls (your original stack approach)
try
{
	$stack = New-Object System.Collections.Stack
	$stack.Push($form)
	
	while ($stack.Count -gt 0)
	{
		$parent = [System.Windows.Forms.Control]$stack.Pop()
		foreach ($ctrl in @($parent.Controls))
		{
			try { $ctrl.add_KeyDown({ $script:LastActivity = Get-Date }) }
			catch { }
			try { $ctrl.add_Click({ $script:LastActivity = Get-Date }) }
			catch { }
			try { $ctrl.add_MouseMove({ $script:LastActivity = Get-Date }) }
			catch { }
			if ($ctrl -and $ctrl.HasChildren) { $stack.Push($ctrl) }
		}
	}
}
catch { }

# NEW: Hook controls added later (so new controls also reset idle timer)
try
{
	$form.add_ControlAdded({
			param ($s,
				$e)
			try
			{
				$c = $e.Control
				if (-not $c) { return }
				
				$stack2 = New-Object System.Collections.Stack
				$stack2.Push($c)
				
				while ($stack2.Count -gt 0)
				{
					$p = [System.Windows.Forms.Control]$stack2.Pop()
					try { $p.add_KeyDown({ $script:LastActivity = Get-Date }) }
					catch { }
					try { $p.add_Click({ $script:LastActivity = Get-Date }) }
					catch { }
					try { $p.add_MouseMove({ $script:LastActivity = Get-Date }) }
					catch { }
					if ($p.HasChildren) { foreach ($cc in @($p.Controls)) { $stack2.Push($cc) } }
				}
			}
			catch { }
		})
}
catch { }

# ------------------------------------------------------------------------------------------------
# Form Closing (X): close directly, then cleanup
# ------------------------------------------------------------------------------------------------
$form.add_FormClosing({
		$script:SuppressClosePrompt = $false
		
		# ---- Timers ----
		try { if ($script:IdleTimer) { $script:IdleTimer.Stop(); $script:IdleTimer.Dispose(); $script:IdleTimer = $null } }
		catch { }
		try { if ($BusyTimer) { $BusyTimer.Stop(); $BusyTimer.Dispose(); $BusyTimer = $null } }
		catch { }
		try { if ($refreshTimer) { $refreshTimer.Stop(); $refreshTimer.Dispose(); $refreshTimer = $null } }
		catch { }
		try { if ($script:VncExportReaper) { $script:VncExportReaper.Stop(); $script:VncExportReaper.Dispose(); $script:VncExportReaper = $null } }
		catch { }
		try { if ($script:RemoteMachineInfoReaper) { $script:RemoteMachineInfoReaper.Stop(); $script:RemoteMachineInfoReaper.Dispose(); $script:RemoteMachineInfoReaper = $null } }
		catch { }
		try { if ($script:StartupWarmupTimer) { $script:StartupWarmupTimer.Stop(); $script:StartupWarmupTimer.Dispose(); $script:StartupWarmupTimer = $null } }
		catch { }
		try { if ($script:PopupThemeTimer) { $script:PopupThemeTimer.Stop(); $script:PopupThemeTimer.Dispose(); $script:PopupThemeTimer = $null } }
		catch { }
		
		try
		{
			if ($global:ProtocolFormTimer)
			{
				$global:ProtocolFormTimer.Stop()
				$global:ProtocolFormTimer.Dispose()
				$global:ProtocolFormTimer = $null
			}
		}
		catch { }
		
		# ---- Popup form (optional) ----
		try
		{
			if ($global:ProtocolForm)
			{
				try { $global:ProtocolForm.Hide() }
				catch { }
				try { $global:ProtocolForm.Dispose() }
				catch { }
				$global:ProtocolForm = $null
			}
		}
		catch { }
		
		# Fast exit: the host process closes right after ShowDialog returns, so expensive
		# background teardown here only slows the X button down.
		try { $script:LaneProtocolStop = $true }
		catch { }
		try { $script:VncExportTask = $null; $script:RemoteMachineInfoTask = $null }
		catch { }
		try { $script:ScaleCredTasks = @{ } }
		catch { }
		try { Write_Log "Form is closing. Performing cleanup." "green" }
		catch { }
		try
		{
			Get-ChildItem -LiteralPath $TempDir -File -ErrorAction SilentlyContinue |
			Where-Object { $_.Name -like '*.sqi' -or $_.Name -like '*.sql' } |
			ForEach-Object {
				try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
				catch { }
			}
		}
		catch { }
		return
		
		# Stop XE watcher jobs (kept from your earlier close handler)
		try
		{
			$xeJobs = Get-Job -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'XEWatcherJob_*' }
			foreach ($j in $xeJobs)
			{
				try { Stop-Job -Job $j -Force -ErrorAction SilentlyContinue | Out-Null }
				catch { }
				try { Remove-Job -Job $j -Force -ErrorAction SilentlyContinue | Out-Null }
				catch { }
			}
		}
		catch { }
		
		# ---- Lane protocol updater cleanup (timer + pool) ----
		try { $script:LaneProtocolStop = $true }
		catch { }
		
		try
		{
			if ($script:protocolTimer)
			{
				try { $script:protocolTimer.Stop() }
				catch { }
				try { $script:protocolTimer.Dispose() }
				catch { }
				$script:protocolTimer = $null
			}
		}
		catch { }
		
		try
		{
			if ($script:LaneProtocolJobs)
			{
				foreach ($kv in @($script:LaneProtocolJobs.GetEnumerator()))
				{
					$st = $kv.Value
					if ($st)
					{
						try { if ($st.Handle -and (-not $st.Handle.IsCompleted) -and $st.PS) { $null = $st.PS.Stop() } }
						catch { }
						try { if ($st.PS) { $st.PS.Dispose() } }
						catch { }
					}
					[void]$script:LaneProtocolJobs.Remove($kv.Key)
				}
			}
		}
		catch { }
		
		try
		{
			if ($script:LaneDbSizeJobs)
			{
				foreach ($kv in @($script:LaneDbSizeJobs.GetEnumerator()))
				{
					$st = $kv.Value
					if ($st)
					{
						try { if ($st.Handle -and (-not $st.Handle.IsCompleted) -and $st.PS) { $null = $st.PS.Stop() } }
						catch { }
						try { if ($st.PS) { $st.PS.Dispose() } }
						catch { }
					}
					[void]$script:LaneDbSizeJobs.Remove($kv.Key)
				}
			}
		}
		catch { }
		
		try
		{
			if ($script:LaneProtocolPool)
			{
				try { $script:LaneProtocolPool.Close() }
				catch { }
				try { $script:LaneProtocolPool.Dispose() }
				catch { }
				$script:LaneProtocolPool = $null
			}
		}
		catch { }

		try
		{
			if ($script:LaneDbSizePool)
			{
				try { $script:LaneDbSizePool.Close() }
				catch { }
				try { $script:LaneDbSizePool.Dispose() }
				catch { }
				$script:LaneDbSizePool = $null
			}
		}
		catch { }
		
		try
		{
			if ($script:VncExportTask)
			{
				$task = $script:VncExportTask
				try
				{
					if ($task.Job -and $task.Job.State -in @('Running', 'NotStarted'))
					{
						Stop-Job -Job $task.Job -Force -ErrorAction SilentlyContinue | Out-Null
					}
					elseif ($task.Job)
					{
						Receive-Job -Job $task.Job -ErrorAction SilentlyContinue | Out-Null
					}
				}
				catch { }
				try { if ($task.Job) { Remove-Job -Job $task.Job -Force -ErrorAction SilentlyContinue | Out-Null } }
				catch { }
				$script:VncExportTask = $null
			}
		}
		catch { }
		
		try
		{
			if ($script:RemoteMachineInfoTask)
			{
				$task = $script:RemoteMachineInfoTask
				try
				{
					if ($task.Job -and $task.Job.State -in @('Running', 'NotStarted'))
					{
						Stop-Job -Job $task.Job -Force -ErrorAction SilentlyContinue | Out-Null
					}
					elseif ($task.Job)
					{
						Receive-Job -Job $task.Job -ErrorAction SilentlyContinue | Out-Null
					}
				}
				catch { }
				try { if ($task.Job) { Remove-Job -Job $task.Job -Force -ErrorAction SilentlyContinue | Out-Null } }
				catch { }
				$script:RemoteMachineInfoTask = $null
			}
		}
		catch { }
		
		# ---- Scale credential runspaces (if you used the background cred helper) ----
		try
		{
			if ($script:ScaleCredReaper)
			{
				try { $script:ScaleCredReaper.Stop() }
				catch { }
				try { $script:ScaleCredReaper.Dispose() }
				catch { }
				$script:ScaleCredReaper = $null
			}
		}
		catch { }
		
		try
		{
			if ($script:ScaleCredTasks)
			{
				foreach ($k in @($script:ScaleCredTasks.Keys))
				{
					$st = $script:ScaleCredTasks[$k]
					if ($st)
					{
						try { if ($st.Handle -and $st.PS) { $st.PS.EndInvoke($st.Handle) } }
						catch { }
						try { if ($st.PS) { $st.PS.Dispose() } }
						catch { }
					}
					[void]$script:ScaleCredTasks.Remove($k)
				}
			}
		}
		catch { }
		
		try
		{
			if ($script:ScaleCredPool)
			{
				try { $script:ScaleCredPool.Close() }
				catch { }
				try { $script:ScaleCredPool.Dispose() }
				catch { }
				$script:ScaleCredPool = $null
				$script:ScaleCredPoolMaxParallel = $null
			}
		}
		catch { }
		
		# --- Any lingering job from earlier versions (defensive) ---
		try
		{
			Get-Job -Name 'ClearXEFolderJob' -ErrorAction SilentlyContinue | ForEach-Object {
				try { Stop-Job $_ -Force -ErrorAction SilentlyContinue }
				catch { }
				try { Remove-Job $_ -Force -ErrorAction SilentlyContinue }
				catch { }
			}
		}
		catch { }
		
		# Final housekeeping
		try { Write_Log "Form is closing. Performing cleanup." "green" }
		catch { }
		try { Delete_Files -Path "$TempDir" -SpecifiedFiles "*.sqi", "*.sql" }
		catch { }
	})

# NOTE: version info now lives in the dedicated $stVersion status-bar item, so the
# busy scanner's status updates no longer wipe it out after the first tick.

# Final layout pass so everything starts aligned
& $layoutTopStrip
& $layoutLogArea

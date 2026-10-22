# ===================================================================================================
#                                  SECTION: GUI - Form Closing
# ---------------------------------------------------------------------------------------------------
# Description:
#   Closing the main window (X button or idle close) closes immediately: timers are stopped,
#   the protocol popup is disposed, background loops are told to stop, and the temporary
#   *.sqi / *.sql files in %TEMP% are deleted. Background runspaces, jobs and pools are NOT
#   torn down one by one - the host process exits right after ShowDialog returns, which ends
#   them anyway, and an orderly teardown only made the X button slow.
#   Also runs the initial layout pass for the top strip and the log area.
# ===================================================================================================

$form.add_FormClosing({
		# ---- Timers ----
		foreach ($timerName in 'IdleTimer', 'VncExportReaper', 'RemoteMachineInfoReaper', 'StartupWarmupTimer', 'PopupThemeTimer', 'protocolTimer')
		{
			try
			{
				$t = Get-Variable -Name $timerName -Scope Script -ValueOnly -ErrorAction SilentlyContinue
				if ($t) { $t.Stop(); $t.Dispose(); Set-Variable -Name $timerName -Scope Script -Value $null }
			}
			catch { }
		}
		try { if ($BusyTimer) { $BusyTimer.Stop(); $BusyTimer.Dispose() } }
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

		# ---- Protocol popup ----
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

		# ---- Background loops: stop scheduling new work ----
		$script:LaneProtocolStop = $true
		$script:VncExportTask = $null
		$script:RemoteMachineInfoTask = $null
		$script:ScaleCredTasks = @{ }

		try { Write_Log "Form is closing. Performing cleanup." "green" }
		catch { }

		# ---- Temporary SQL files written by the maintenance tasks ----
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
	})

# Initial layout pass so everything starts aligned (the bottom layout runs in 14)
& $layoutTopStrip
& $layoutLogArea

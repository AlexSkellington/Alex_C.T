# ===================================================================================================
#                                   FUNCTION: Wait_UI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Waits like Start-Sleep but keeps the window responsive: sleeps in 50 ms slices and lets the
#   window repaint / handle clicks in between (Application.DoEvents). Also counts as activity, so
#   the idle auto-close never fires while a task is waiting on purpose.
#
# Parameters:
#   - Milliseconds [int]: how long to wait (default 100)
#
# Notes:
#   - Use it for waits in code that runs on the UI thread (menu actions). Code shipped into
#     background jobs / runspaces keeps plain Start-Sleep (this function does not exist there).
# ===================================================================================================

function Wait_UI
{
	param (
		[int]$Milliseconds = 100
	)

	$end = [DateTime]::UtcNow.AddMilliseconds([Math]::Max(0, $Milliseconds))
	do
	{
		$script:LastActivity = Get-Date
		try { [System.Windows.Forms.Application]::DoEvents() }
		catch { }
		$left = ($end - [DateTime]::UtcNow).TotalMilliseconds
		if ($left -gt 0) { Start-Sleep -Milliseconds ([int][Math]::Min(50, [Math]::Ceiling($left))) }
	}
	while ([DateTime]::UtcNow -lt $end)
}

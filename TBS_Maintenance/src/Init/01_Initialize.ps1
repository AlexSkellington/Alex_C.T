# ===================================================================================================
#                                  SECTION: Runtime Initialization
# ---------------------------------------------------------------------------------------------------
# Description:
#   Runs once every function is defined, before the GUI is built:
#     - loads the MailslotSender interop type (used by the lane/scale command functions)
#     - headless worker mode: when started with -LaneSqlEnableWorker (by the scheduled task
#       Ensure_SQL_Enable_Scheduled_Task_For_File_Lane registers), runs that one job and exits
#       with 0 (success) / 1 (failure) - no window is created.
# ===================================================================================================

Ensure_MailslotSender_Type

if ($script:LaneSqlEnableWorkerMode)
{
	$__laneEnableWorkerOk = $false
	try
	{
		$__laneEnableWorkerOk = [bool](Ensure_SQL_Enable_Scheduled_Task_For_File_Lane `
			-LaneNumber $script:LaneSqlEnableWorkerLane `
			-MachineName $script:LaneSqlEnableWorkerMachine `
			-TcpPort $script:LaneSqlEnableWorkerTcpPort `
			-StartTime $script:LaneSqlEnableWorkerStartTime)
	}
	catch
	{
		$__laneEnableWorkerOk = $false
	}
	if ($__laneEnableWorkerOk) { exit 0 } else { exit 1 }
}

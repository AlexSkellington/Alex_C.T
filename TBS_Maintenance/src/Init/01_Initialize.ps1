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

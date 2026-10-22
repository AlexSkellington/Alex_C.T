# ===================================================================================================
#                              FUNCTION: Get_Server_Maintenance_Worker
# ---------------------------------------------------------------------------------------------------
# Description:
#   Writes the server maintenance worker (Resources\Workers\Server_Maintenance_Worker.ps1) to
#   Scripts\Server_Maintenance.ps1 when it is missing or different, so the scheduled tasks always
#   run the version of the running tool. A running worker is not disturbed: PowerShell has
#   already read its script.
#
# Variables:
#   - $script:ScriptsFolder: where the worker is written
#
# Returns:
#   The worker's full path, or $null when it could not be written.
# ===================================================================================================

function Get_Server_Maintenance_Worker
{
	$workerPath = Join-Path $script:ScriptsFolder 'Server_Maintenance.ps1'
	$workerText = @'
#TBS:EMBED Resources/Workers/Server_Maintenance_Worker.ps1
'@
	$workerText = $workerText -replace "`r?`n", "`r`n"
	try
	{
		if (-not (Test-Path -LiteralPath $script:ScriptsFolder)) { $null = New-Item -Path $script:ScriptsFolder -ItemType Directory -Force -ErrorAction Stop }
		$current = $null
		if (Test-Path -LiteralPath $workerPath) { $current = [System.IO.File]::ReadAllText($workerPath) }
		if ($current -ne $workerText) { [System.IO.File]::WriteAllText($workerPath, $workerText, [System.Text.Encoding]::ASCII) }
		return $workerPath
	}
	catch
	{
		try { Write_Log "Could not write the server maintenance script to ${workerPath}: $($_.Exception.Message)" 'red' }
		catch { }
		return $null
	}
}

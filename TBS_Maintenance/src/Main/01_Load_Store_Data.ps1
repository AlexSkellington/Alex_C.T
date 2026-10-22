# ===================================================================================================
#                                 SECTION: Main - Load Store Data
# ---------------------------------------------------------------------------------------------------
# Description:
#   Pre-window data load - only fast, local reads, so the window appears quickly; the network
#   refreshes run after Show as warmup steps (02).
#     - starts the NSSM download in the background (returns immediately)
#     - store + database info, read LIVE from the store's INI files (a few ms)
#     - store number normalisation (1 -> 001; 4+ digits keep 4+)
#     - node discovery (lanes, scales, backoffices) from cache
#
# Notes:
#   - tests\TBS_Maintenance_SelfTest.ps1 runs the store-number block verbatim, from the
#     '$script:StoreNumber = [string]...' line to the line before '$script:StoreName = ...'.
#   - At script scope $StoreNumber and $script:StoreNumber are the same variable; the menus
#     pass $StoreNumber to the functions.
# ===================================================================================================

Download_NSSM -Background | Out-Null

Get_Store_And_Database_Info -WinIniPath $WinIniPath -SmsStartIniPath $SmsStartIniPath -StartupIniPath $StartupIniPath -SystemIniPath $SystemIniPath
$script:StoreNumber = [string]$script:FunctionResults['StoreNumber']
if ($script:StoreNumber)
{
	$storeDigits = ($script:StoreNumber -replace '[^\d]', '')
	if ($storeDigits)
	{
		if ($storeDigits.Length -le 3) { $script:StoreNumber = ([int]$storeDigits).ToString('000') }
		else { $script:StoreNumber = ([int]$storeDigits).ToString('0000') }
	}
}
$script:StoreName = $script:FunctionResults['StoreName']
$script:SqlModuleName = $script:FunctionResults['SqlModuleName']

# Nodes from cache (the fresh discovery is the second warmup step, which also starts the
# lane protocol loop against the fresh node list)
$null = Retrieve_Nodes -StoreNumber $StoreNumber
$script:Nodes = $script:FunctionResults['Nodes']
$script:LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']

try { Write_Log ("Startup: pre-window data ready in {0:N1}s - store '{1}', {2} node(s)" -f ((Get-Date) - $script:StartupT0).TotalSeconds, $StoreNumber, @($Nodes).Count) 'blue' -FileOnly } catch { }

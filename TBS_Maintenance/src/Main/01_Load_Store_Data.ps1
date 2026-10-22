# ===================================================================================================
#                                       SECTION: Main Script Execution
# ---------------------------------------------------------------------------------------------------
# Description:
#   Orchestrates the execution flow of the script, initializing variables, processing items, and handling user interactions.
# ===================================================================================================

# Check for the precense of PsExec for later use
# $GetPsExec = Get_PsExec

# Starts download in background (returns immediately)
Download_NSSM -Background | Out-Null

# Refresh launch-owned cache files once, then reuse them for the rest of the session.
$script:UseCachedCoreStartup = $__isInlineLikeLaunch

# Get SQL Connection String
# v2.6.9: in EVERY launch mode this only reads the local cache before the
# window is shown (fast). The -RefreshCache network round trip runs as the
# first post-Show warmup step below, so the window always appears first.
# (No cache at all? The function does its full fetch here, exactly as before.)
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
$StoreNumber = $script:StoreNumber
$StoreName = $script:StoreName
$SqlModuleName = $script:SqlModuleName

# Count Nodes (cache first in every launch mode - same reasoning as the store
# info call above; the fresh refresh is the second warmup step)
$Nodes = Retrieve_Nodes -StoreNumber $StoreNumber
$script:Nodes = $script:FunctionResults['Nodes']
$Nodes = $script:Nodes

# Retrieve the list of machine names from the FunctionResults dictionary
$script:LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
$LaneNumToMachineName = $script:LaneNumToMachineName

# Lane protocol info: started by the 'Refreshing node discovery...' warmup
# step (post-Show, every launch mode) so it always runs against the fresh
# node list instead of blocking the first paint.
try { Write_Log ("Startup: pre-window data ready in {0:N1}s - store '{1}', {2} node(s)" -f ((Get-Date) - $script:StartupT0).TotalSeconds, $StoreNumber, @($Nodes).Count) 'blue' -FileOnly } catch { }

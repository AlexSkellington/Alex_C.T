# ===================================================================================================
#                              FUNCTION: Show_Table_Selection_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Dialog listing the distinct tables from Get_Table_Aliases rows (sorted) for Pump_Tables;
#   returns the checked table names (with _TAB). Thin wrapper over Show_CheckList_Form.
#
# Parameters:
#   - AliasResults  [System.Collections.ArrayList] (required): rows with a .Table property
#
# Returns:
#   Array of checked table names ($null when cancelled).
# ===================================================================================================

function Show_Table_Selection_Form
{
	param (
		[Parameter(Mandatory = $true)]
		[System.Collections.ArrayList]$AliasResults
	)

	$tables = @($AliasResults | Select-Object -ExpandProperty Table -Unique | Sort-Object)
	return Show_CheckList_Form -Title 'Select Tables to Process' -Prompt 'Please select the tables you want to pump:' -Items $tables -Width 410 -Height 400
}

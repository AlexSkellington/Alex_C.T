# ===================================================================================================
#                              FUNCTION: Show_Section_Selection_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   Dialog with a check list of SQL script sections (Process_Server / lane maintenance);
#   returns the checked section names. Thin wrapper over Show_CheckList_Form.
#
# Parameters:
#   - SectionNames  [string[]] (required): section names to list
#
# Returns:
#   Array of checked section names ($null when cancelled).
# ===================================================================================================

function Show_Section_Selection_Form
{
	param (
		[Parameter(Mandatory = $true)]
		[string[]]$SectionNames
	)

	return Show_CheckList_Form -Title 'Select SQL Sections' -Prompt 'Check the sections you want to run, then click OK.' -Items $SectionNames -Width 500 -Height 280
}

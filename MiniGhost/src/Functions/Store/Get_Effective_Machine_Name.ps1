# ===================================================================================================
#                              FUNCTION: Get_Effective_Machine_Name
# ---------------------------------------------------------------------------------------------------
# Description:
# Returns the pending/new machine name when one exists; otherwise returns the current computer name.
# ===================================================================================================

function Get_Effective_Machine_Name
{
	[CmdletBinding()]
	param ()
	
	if ($script:newMachineName -and -not [string]::IsNullOrWhiteSpace([string]$script:newMachineName))
	{
		return [string]$script:newMachineName
	}
	
	return $env:COMPUTERNAME
}

# ===================================================================================================
#                                FUNCTION: Set_Operation_Status
# ---------------------------------------------------------------------------------------------------
# Description:
# Central helper to update the shared operation-status table only when the requested key exists.
# ===================================================================================================

function Set_Operation_Status
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus,
		[Parameter(Mandatory = $true)]
		[string]$Key,
		[Parameter(Mandatory = $true)]
		[string]$Status,
		[Parameter(Mandatory = $false)]
		[string]$Message = "",
		[Parameter(Mandatory = $false)]
		[string]$Details = ""
	)
	
	if ($OperationStatus -and $OperationStatus.ContainsKey($Key))
	{
		$OperationStatus[$Key].Status = $Status
		$OperationStatus[$Key].Message = $Message
		$OperationStatus[$Key].Details = $Details
	}
}

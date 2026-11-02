# ===================================================================================================
#                              FUNCTION: Get_Backup_Copy_Paths
# ---------------------------------------------------------------------------------------------------
# Description:
#   Returns the saved extra copy folders for a backup type (Show_Backup_Copy_Paths_Dialog saves
#   them), or an empty list.
#
# Parameters:
#   - Kind [string] (required): ServerDB | StoremanZip | LaneDB
# ===================================================================================================

function Get_Backup_Copy_Paths
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$Kind
	)

	$settingsFile = Join-Path $script:AlexRoot ("Settings\Backup_Copy_Paths_" + $Kind + ".txt")
	if (-not (Test-Path -LiteralPath $settingsFile)) { return , ([string[]]@()) }
	try { return , ([string[]]@(Get-Content -LiteralPath $settingsFile -ErrorAction Stop | ForEach-Object { $_.Trim() } | Where-Object { $_ })) }
	catch { return , ([string[]]@()) }
}

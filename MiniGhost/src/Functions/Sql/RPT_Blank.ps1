# ===================================================================================================
#                                       FUNCTION: RPT_Blank
# ---------------------------------------------------------------------------------------------------
# Description:
#   Truncates the specified list of tables in the SQL database.
# ===================================================================================================

function RPT_Blank
{
	param (
		[string[]]$tables
	)
	
	# Initialize an array to store failed truncate commands
	$failedTruncateTables = @()
	
	# Prepare and execute Truncate Commands
	foreach ($table in $tables)
	{
		$command = "TRUNCATE TABLE $table;"
		if (-not (Execute_SQL_Commands -commandText $command))
		{
			$failedTruncateTables += $table # Add failed table to the array
		}
	}
	
	# Return the list of failed tables
	return $failedTruncateTables
}

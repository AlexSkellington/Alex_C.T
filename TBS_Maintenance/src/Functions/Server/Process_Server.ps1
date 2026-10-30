# ===================================================================================================
#                              FUNCTION: Process_Server
# ---------------------------------------------------------------------------------------------------
# Description:
#   Runs the server (store) SQL maintenance: takes the script from $script:ServerSQLScript
#   (generated on demand) or a file, optionally lets the user pick sections, runs each section
#   once with the built-in SqlClient (split on GO lines) in a background runspace so the window
#   stays responsive, and logs which sections failed.
#
# Parameters:
#   - StoresqlFilePath  [string]: SQL file to run instead of the script variable
#   - PromptForSections [switch]: show the SQL section editor first - pick sections, edit their SQL
#                                 for this run, add custom SQL (saved in Settings\Custom_Server_SQL.json)
#
# Variables:
#   - $script:ServerSQLScript: default SQL script text
#   - $script:FunctionResults['DBNAME'] / ['DBSERVER']: target database/server
# ===================================================================================================

function Process_Server
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StoresqlFilePath,
		[Parameter(Mandatory = $false)]
		[switch]$PromptForSections
	)

	# ------------------------------------------------------------------------------------------------
	# Banner: Start
	# ------------------------------------------------------------------------------------------------
	Write_Log "`r`n==================== Starting Server Database Maintenance ====================`r`n" "blue"

	# ------------------------------------------------------------------------------------------------
	# Variables and settings
	# ------------------------------------------------------------------------------------------------
	$sqlScript = $script:ServerSQLScript
	$dbName = $script:FunctionResults['DBNAME']
	$server = $script:FunctionResults['DBSERVER']

	# ------------------------------------------------------------------------------------------------
	# Load SQL script content. Primary source is the in-memory $script:ServerSQLScript
	# variable, set by Generate_SQL_Scripts during warmup. If it is not ready yet (the
	# user acted before the 'Generating SQL scripts' warmup step ran), generate it on
	# demand - pure CPU (~2 ms) and needs only the already-resolved store/DB info. The
	# file fallback below is only for an explicitly provided standalone .sqi.
	# ------------------------------------------------------------------------------------------------
	if ([string]::IsNullOrWhiteSpace($sqlScript))
	{
		try { Generate_SQL_Scripts -StoreNumber $script:StoreNumber | Out-Null } catch { }
		$sqlScript = $script:ServerSQLScript
	}

	if (-not [string]::IsNullOrWhiteSpace($sqlScript))
	{
		Write_Log "Executing SQL script from variable..." "blue"
	}
	elseif ($StoresqlFilePath)
	{
		if (-not (Test-Path $StoresqlFilePath))
		{
			Write_Log "SQL file not found: $StoresqlFilePath" "red"
			return
		}
		Write_Log "Executing SQL file: $StoresqlFilePath" "blue"
		try
		{
			$sqlScript = Get-Content -Path $StoresqlFilePath -Raw -ErrorAction Stop
		}
		catch
		{
			Write_Log "Failed to read SQL file: $_" "red"
			return
		}
	}
	else
	{
		Write_Log "No SQL script content or file path provided." "red"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# Parse SQL sections: each "/* Name */" comment starts a section
	# ------------------------------------------------------------------------------------------------
	$sectionPattern = '(?s)/\*\s*(?<SectionName>[^*/]+?)\s*\*/\s*(?<SQLCommands>(?:.(?!/\*)|.)*?)(?=(/\*|$))'
	$sectionMatches = [regex]::Matches($sqlScript, $sectionPattern)
	if ($sectionMatches.Count -eq 0)
	{
		Write_Log "No SQL sections found to execute." "red"
		return
	}

	# Sections in run order: Name + SQL
	$sectionsToRun = @($sectionMatches | ForEach-Object { [pscustomobject]@{ Name = $_.Groups['SectionName'].Value.Trim(); Sql = $_.Groups['SQLCommands'].Value.Trim(); Edited = $false } })
	if ($PromptForSections)
	{
		$customFile = Join-Path $script:AlexRoot 'Settings\Custom_Server_SQL.json'
		$sectionsToRun = @(Show_SQL_Section_Editor -Sections $sectionsToRun -Title 'Server Database Maintenance - SQL Sections' -CustomFile $customFile)
		if ($sectionsToRun.Count -eq 0 -or -not $sectionsToRun[0])
		{
			Write_Log "No sections selected or form was canceled. Aborting execution." "yellow"
			return
		}
	}

	# ------------------------------------------------------------------------------------------------
	# How to run SQL: the built-in SqlClient in a background runspace (Invoke_In_Background), split on
	# GO lines - index rebuilds / shrinks take minutes and used to freeze the window ("Not
	# Responding"). No SQL PowerShell module is loaded (that alone took seconds on the UI thread).
	# ------------------------------------------------------------------------------------------------
	$ConnectionString = $script:FunctionResults['ConnectionString']
	if (-not $ConnectionString -or -not $server -or -not $dbName)
	{
		Write_Log "DB server, DB name, or connection string not found. Cannot execute SQL script." "red"
		return
	}

	$sectionWorker = {
		param ([string]$ConnStr, [string]$Sql)
		$cn = New-Object System.Data.SqlClient.SqlConnection $ConnStr
		$cn.Open()
		try
		{
			foreach ($batch in ([regex]::Split($Sql, '(?im)^\s*GO\s*$')))
			{
				if ([string]::IsNullOrWhiteSpace($batch)) { continue }
				$cmd = $cn.CreateCommand()
				$cmd.CommandText = $batch
				$cmd.CommandTimeout = 0
				$null = $cmd.ExecuteNonQuery()
			}
		}
		finally { $cn.Dispose() }
	}

	# ------------------------------------------------------------------------------------------------
	# Execute each section once, in order
	# ------------------------------------------------------------------------------------------------
	$failedSections = @()
	foreach ($section in $sectionsToRun)
	{
		$sectionName = $section.Name
		$sqlCommands = ([string]$section.Sql).Trim()
		if ($section.Edited) { $sectionName += ' (edited)' }

		if ([string]::IsNullOrWhiteSpace($sqlCommands))
		{
			Write_Log "Section '$sectionName' has no commands. Skipping..." "yellow"
			continue
		}

		Write_Log "`r`nExecuting section: '$sectionName'" "blue"
		Write_Log "--------------------------------------------------------------------------------"
		Write_Log "$sqlCommands" "orange"
		Write_Log "--------------------------------------------------------------------------------"

		try
		{
			$null = Invoke_In_Background -ScriptBlock $sectionWorker -ArgumentList @($ConnectionString, $sqlCommands)
			Write_Log "Section '$sectionName' executed successfully." "green"
		}
		catch
		{
			Write_Log "Error executing section '$sectionName': $_" "red"
			$failedSections += $sectionName
		}
	}
	# ------------------------------------------------------------------------------------------------
	# Completion summary
	# ------------------------------------------------------------------------------------------------
	if ($failedSections.Count -eq 0)
	{
		Write_Log "`r`nAll SQL sections executed successfully." "green"
	}
	else
	{
		Write_Log "The following sections failed: $($failedSections -join ', ')" "red"
	}

	Write_Log "`r`n==================== Completed Server Database Maintenance ====================" "blue"
}

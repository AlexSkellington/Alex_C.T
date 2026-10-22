# ===================================================================================================
#                              FUNCTION: Process_Server
# ---------------------------------------------------------------------------------------------------
# Description:
#   Runs the server (store) SQL maintenance: takes the script from $script:ServerSQLScript
#   (generated on demand) or a file, optionally lets the user pick sections, runs each section
#   once - Invoke-Sqlcmd when a SQL module is available, else SqlClient split on GO lines - and
#   logs which sections failed.
#
# Parameters:
#   - StoresqlFilePath  [string]: SQL file to run instead of the script variable
#   - PromptForSections [switch]: show the section selection dialog first
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

	$SectionsToRun = $null
	if ($PromptForSections)
	{
		$allSectionNames = @($sectionMatches | ForEach-Object { $_.Groups['SectionName'].Value.Trim() })
		$SectionsToRun = Show_Section_Selection_Form -SectionNames $allSectionNames
		if (-not $SectionsToRun -or @($SectionsToRun).Count -eq 0)
		{
			Write_Log "No sections selected or form was canceled. Aborting execution." "yellow"
			return
		}
	}
	$useSpecificSections = $SectionsToRun -and @($SectionsToRun).Count -gt 0

	# ------------------------------------------------------------------------------------------------
	# How to run SQL: Invoke-Sqlcmd from the detected module, else the built-in SqlClient
	# (which does not understand GO - the sections are split into batches on GO lines)
	# ------------------------------------------------------------------------------------------------
	$ConnectionString = $script:FunctionResults['ConnectionString']
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	if (-not $ConnectionString -or -not $server -or -not $dbName)
	{
		Write_Log "DB server, DB name, or connection string not found. Cannot execute SQL script." "red"
		return
	}

	$InvokeSqlCmd = $null
	if ($SqlModuleName -and $SqlModuleName -ne "None")
	{
		try
		{
			Import-Module $SqlModuleName -ErrorAction Stop
			$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
		}
		catch
		{
			Write_Log "SQL module '$SqlModuleName' could not be loaded ($($_.Exception.Message)) - using the built-in SQL client." "yellow"
			$InvokeSqlCmd = $null
		}
	}
	else
	{
		Write_Log "No SQL PowerShell module - using the built-in SQL client." "yellow"
	}
	$supportsConnectionString = ($InvokeSqlCmd -and $InvokeSqlCmd.Parameters.ContainsKey('ConnectionString'))

	$runSection = {
		param ([string]$Sql)
		if ($InvokeSqlCmd)
		{
			if ($supportsConnectionString) { $null = & $InvokeSqlCmd -ConnectionString $ConnectionString -Query $Sql -ErrorAction Stop -QueryTimeout 0 }
			else { $null = & $InvokeSqlCmd -ServerInstance $server -Database $dbName -Query $Sql -ErrorAction Stop -QueryTimeout 0 }
			return
		}
		$cn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
		try
		{
			$cn.Open()
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
	foreach ($match in $sectionMatches)
	{
		$sectionName = $match.Groups['SectionName'].Value.Trim()
		$sqlCommands = $match.Groups['SQLCommands'].Value.Trim()

		if ($useSpecificSections -and ($SectionsToRun -notcontains $sectionName)) { continue }
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
			& $runSection $sqlCommands
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

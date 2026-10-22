# ===================================================================================================
#                              FUNCTION: Retrieve_Nodes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Finds the store's lanes, scales, server and backoffices and builds every lookup map the
#   rest of the script uses (lane number -> machine, scale code -> IP info, ...). Sources, in
#   order: the store database (TER_TAB + TBS_SCL_ver520), else Office\Load\Ter_Load.sql; TER
#   scales are enriched from XchScale\XchNodes.ini. Results are cached per store, the GUI
#   count labels are updated, and the nodes object is returned.
#
# Parameters:
#   - StoreNumber  [string]: store to count (required)
#   - RefreshCache [switch]: ignore the in-memory / file cache and rediscover
#
# Variables:
#   - $script:FunctionResults: receives Nodes, every map listed in $resultNames below and the
#     NumberOfLanes/Servers/Backoffices/Scales counts
#   - Reads $OfficePath, $LoadPath, $script:BasePath, FunctionResults DBSERVER / DBNAME /
#     ConnectionString / SqlModuleName / StoremanPath
#
# Writes:
#   - Setup_Files\Retrieve_Nodes.txt (JSON cache; CacheSchemaVersion must match to be used)
#
# Returns:
#   PSCustomObject with all counts and the lane/scale/backoffice maps.
#
# Notes:
#   - Node rules (TER_TAB / Ter_Load rows): 0xx (or a \\machine\ path) = lane, 8xx with an
#     XchScale.exe path = TER scale, 901 '@machine' = server, 902-998 '@machine' = backoffice.
#   - TBS scales keep their original ScaleCode; a TER scale with the same IP as a TBS scale
#     is dropped (the TBS row wins and inherits the TER path/port details).
#   - Run as a background job at startup (warmup): keep it free of GUI-only dependencies
#     (the label updates are skipped when the labels do not exist).
# ===================================================================================================

function Retrieve_Nodes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[switch]$RefreshCache
	)

	# Bump when the shape of the cached results changes (10: TBS scales keep their original
	# ScaleCode and are never normalised to 80X)
	$CacheSchemaVersion = 10
	$CachePath = Join-Path $script:SetupFiles 'Retrieve_Nodes.txt'

	# ------------------------------------------------------------------------------------------------
	# What this function publishes. FunctionResults keys = local variable names.
	# ------------------------------------------------------------------------------------------------
	# Nodes object properties, in order
	$nodeNames = @(
		'NumberOfLanes', 'NumberOfServers', 'NumberOfBackoffices', 'NumberOfScales',
		'LaneMachineNames', 'LaneNumToMachineName', 'LaneMachineLabels', 'LaneMachinePath', 'LaneMachineToServerPath',
		'ScaleCodes', 'ScaleLabels', 'ScaleExePaths', 'ScaleCodeToIPInfo', 'ScaleNameToCode', 'ScaleCodeToPath', 'ScaleNameToPath',
		'ScaleCodeToName', 'ScaleCodeToSection', 'ScaleCodeToBrand', 'ScaleCodeToModel', 'ScaleCodeToLocation',
		'ScaleCodeToIPNetwork', 'ScaleCodeToIPDevice', 'ScaleCodeToFullIP', 'ScaleCodeToPortRaw', 'ScaleCodeToTcpPort',
		'ScaleCodeToBaseMode', 'ScaleCodeToType', 'ScaleCodeToBaud', 'ScaleCodeToSub', 'ScaleCodeToPlumDir',
		'BackofficeNumToMachineName', 'BackofficeNumToLabel', 'BackofficeNumToPath',
		'ServerMachineName', 'ServerLabel', 'ServerPath', 'WindowsScales'
	)
	# Published to FunctionResults (the node properties plus the reverse maps)
	$resultNames = $nodeNames + @('MachineNameToLaneNum', 'MachineNameToBackofficeNum', 'MachineNameToBackofficePath')
	# Hashtable maps (cache: JSON object <-> hashtable)
	$mapNames = @($resultNames | Where-Object { $_ -notmatch '^(NumberOf|ServerMachineName$|ServerLabel$|ServerPath$|LaneMachineNames$|ScaleCodes$)' })
	# Scale fields that also get their own ScaleCodeTo<Field> map (only when set)
	$scaleFieldMaps = @(
		@('Section', 'ScaleCodeToSection'), @('ScaleBrand', 'ScaleCodeToBrand'), @('ScaleModel', 'ScaleCodeToModel'),
		@('ScaleLocation', 'ScaleCodeToLocation'), @('IPNetwork', 'ScaleCodeToIPNetwork'), @('IPDevice', 'ScaleCodeToIPDevice'),
		@('FullIP', 'ScaleCodeToFullIP'), @('PortRaw', 'ScaleCodeToPortRaw'), @('TcpPort', 'ScaleCodeToTcpPort'),
		@('BaseMode', 'ScaleCodeToBaseMode'), @('Type', 'ScaleCodeToType'), @('Baud', 'ScaleCodeToBaud'),
		@('Sub', 'ScaleCodeToSub'), @('PlumDir', 'ScaleCodeToPlumDir')
	)

	# Count labels (GUI only; absent in background jobs)
	$updateLabels = {
		param ($Servers, $Backoffices, $Lanes, $Scales)
		if ($null -ne $NodesBackoffices) { $NodesBackoffices.Text = "Number of Backoffices: $Backoffices" }
		if ($null -ne $NodesStore) { $NodesStore.Text = "Number of Lanes: $Lanes" }
		if ($null -ne $scalesLabel) { $scalesLabel.Text = "Number of Scales: $Scales" }
		if ($null -ne $form) { $form.Refresh() }
	}

	# Builds $Nodes from the local variables and publishes everything (dot-sourced: runs in
	# this function's scope)
	$publish = {
		$nodeProps = [ordered]@{ }
		foreach ($n in $nodeNames) { $nodeProps[$n] = Get-Variable -Name $n -ValueOnly }
		$Nodes = [PSCustomObject]$nodeProps
		foreach ($n in $resultNames) { $script:FunctionResults[$n] = Get-Variable -Name $n -ValueOnly }
		$script:FunctionResults['Nodes'] = $Nodes
		$script:FunctionResults['XchNodesIniPathsUsed'] = @($XchNodesIniPathsUsed | Select-Object -Unique)
		& $updateLabels $NumberOfServers $NumberOfBackoffices $NumberOfLanes $NumberOfScales
	}

	# Store number as used for the cache match (3 digits, 4+ digit stores keep theirs)
	$normalizeStore = {
		param ([string]$Value)
		$digits = ([string]$Value -replace '[^\d]', '')
		if (-not $digits) { return $Value }
		try
		{
			if ($digits.Length -le 3) { return ([int]$digits).ToString('000') }
			return ([int]$digits).ToString('0000')
		}
		catch { return $Value }
	}
	$StoreNumberNorm = & $normalizeStore $StoreNumber

	# ------------------------------------------------------------------------------------------------
	# 0a. In-memory results for this store (this launch already discovered them)
	# ------------------------------------------------------------------------------------------------
	if (-not $RefreshCache)
	{
		try
		{
			$existingStoreNorm = & $normalizeStore ([string]$script:FunctionResults['StoreNumber'])
			if ($script:FunctionResults['Nodes'] -and ($existingStoreNorm -eq $StoreNumberNorm) -and $script:FunctionResults['LaneNumToMachineName'])
			{
				& $updateLabels $script:FunctionResults['NumberOfServers'] $script:FunctionResults['NumberOfBackoffices'] $script:FunctionResults['NumberOfLanes'] $script:FunctionResults['NumberOfScales']
				return $script:FunctionResults['Nodes']
			}
		}
		catch { }
	}

	# ------------------------------------------------------------------------------------------------
	# 0b. File cache for this store
	# ------------------------------------------------------------------------------------------------
	if ((-not $RefreshCache) -and (Test-Path -LiteralPath $CachePath))
	{
		try
		{
			$raw = Get-Content -LiteralPath $CachePath -Raw -ErrorAction Stop
			$cacheObj = $null
			if (-not [string]::IsNullOrWhiteSpace($raw)) { $cacheObj = $raw | ConvertFrom-Json -ErrorAction Stop }
			if ($cacheObj -and $cacheObj.Results -and $cacheObj.CacheSchemaVersion -and
				([int]$cacheObj.CacheSchemaVersion -eq $CacheSchemaVersion) -and
				$cacheObj.StoreNumber -and ($cacheObj.StoreNumber.ToString() -eq $StoreNumberNorm))
			{
				$r = $cacheObj.Results
				foreach ($n in $mapNames)
				{
					$h = @{ }
					$o = $r.$n
					if ($o -and $o.PSObject -and $o.PSObject.Properties) { foreach ($p in $o.PSObject.Properties) { $h[$p.Name] = $p.Value } }
					Set-Variable -Name $n -Value $h
				}
				$NumberOfLanes = [int]$r.NumberOfLanes
				$NumberOfServers = [int]$r.NumberOfServers
				$NumberOfBackoffices = [int]$r.NumberOfBackoffices
				$NumberOfScales = [int]$r.NumberOfScales
				$LaneMachineNames = @($r.LaneMachineNames)
				$ScaleCodes = @($r.ScaleCodes)
				$ServerMachineName = $r.ServerMachineName
				$ServerLabel = $r.ServerLabel
				$ServerPath = $r.ServerPath
				$XchNodesIniPathsUsed = @()
				if ($r.PSObject.Properties.Name -contains 'XchNodesIniPathsUsed') { $XchNodesIniPathsUsed = @($r.XchNodesIniPathsUsed) }

				. $publish
				return $Nodes
			}
		}
		catch
		{
			# unreadable / old cache: discover again below
		}
	}

	# ====================================================================================
	# DISCOVERY - start empty
	# ====================================================================================
	foreach ($n in $mapNames) { Set-Variable -Name $n -Value @{ } }
	$LaneMachineNames = @()
	$ScaleCodes = @()
	$ServerMachineName = $null
	$ServerLabel = $null
	$ServerPath = $null
	$NumberOfLanes = 0
	$NumberOfServers = 0
	$NumberOfBackoffices = 0
	$NumberOfScales = 0

	# TER scales are grouped by their 80X terminal; TBS scales keep the ScaleCode stored in
	# TBS_SCL_ver520.
	$TerScaleBy80X = @{ }
	$TbsScaleRows = New-Object System.Collections.ArrayList

	# Classify one TER_TAB / Ter_Load row (dot-sourced: updates this function's variables)
	$addTerRow = {
		param ([string]$terminal, [string]$label, [string]$path, [string]$hostPath)

		if (($terminal -match '^0\d\d$' -or $path -match '^\\\\[^\\]+\\') -and $terminal -notmatch '^8' -and $terminal -notmatch '^9')
		{
			# Lane
			$machineName = if ($path -match '\\\\([^\\]+)\\') { $matches[1] } else { $terminal }
			$LaneMachineNames += $machineName
			if ($terminal -match '^0\d\d$') { $LaneNumToMachineName[$terminal] = $machineName }
			$LaneMachineLabels[$machineName] = $label
			$LaneMachinePath[$machineName] = $path
			$LaneMachineToServerPath[$machineName] = $hostPath
		}
		elseif ($terminal -match '^8\d\d$' -and $path -match '(?i)XchScale\\XchScale\.exe$')
		{
			# TER scale (one entry per 80X terminal; later rows only fill gaps)
			if (-not $TerScaleBy80X.ContainsKey($terminal))
			{
				$TerScaleBy80X[$terminal] = [PSCustomObject]@{
					Code          = $terminal
					ScaleCode     = $terminal
					BaseScaleCode = $terminal
					Label         = $label
					ScaleName     = $label
					Path          = $path
					HostPath      = $hostPath
					ScaleBrand    = $null
					Scalevendor   = $null
					ScaleModel    = $null
					ScaleLocation = $null
					IPNetwork     = $null
					IPDevice      = $null
					FullIP        = $null
					PortRaw       = $null
					TcpPort       = $null
					Section       = $null
					BaseMode      = $null
					Type          = $null
					Baud          = $null
					Sub           = $null
					PlumDir       = $null
					Active        = $null
					Source        = 'TER'
				}
			}
			else
			{
				$existingTer = $TerScaleBy80X[$terminal]
				if ([string]::IsNullOrWhiteSpace([string]$existingTer.Label) -and $label) { $existingTer.Label = $label }
				if ([string]::IsNullOrWhiteSpace([string]$existingTer.ScaleName) -and $label) { $existingTer.ScaleName = $label }
				if ([string]::IsNullOrWhiteSpace([string]$existingTer.Path) -and $path) { $existingTer.Path = $path }
				if ([string]::IsNullOrWhiteSpace([string]$existingTer.HostPath) -and $hostPath) { $existingTer.HostPath = $hostPath }
			}
		}
		elseif ($terminal -eq '901' -and $path -match '^@[^@]+$')
		{
			# Store server
			$ServerMachineName = $path -replace '^@', ''
			$ServerLabel = $label
			$ServerPath = $path
		}
		elseif ($terminal -match '^9(0[2-9]|[1-8]\d|9[0-8])$' -and $path -match '^@[^@]+$')
		{
			# Backoffice
			$BackofficeNumToMachineName[$terminal] = $path -replace '^@', ''
			$BackofficeNumToLabel[$terminal] = $label
			$BackofficeNumToPath[$terminal] = $path
		}
	}

	$TerLoadSqlPath = $null
	if (-not [string]::IsNullOrWhiteSpace($LoadPath)) { $TerLoadSqlPath = Join-Path $LoadPath 'Ter_Load.sql' }

	# ====================================================================================
	# XchNodes.ini candidates: <storeman>\XchScale (+ one or two folder levels below)
	# ====================================================================================
	$XchSearchRoots = @()
	foreach ($rootCandidate in @([string]$script:FunctionResults['StoremanPath'], [string]$script:BasePath))
	{
		if (-not [string]::IsNullOrWhiteSpace($rootCandidate)) { $XchSearchRoots += [System.IO.Path]::Combine($rootCandidate.Trim().TrimEnd('\'), 'XchScale') }
	}
	if (-not [string]::IsNullOrWhiteSpace($OfficePath))
	{
		try
		{
			$officeTrimmed = ([string]$OfficePath).Trim().TrimEnd('\')
			if ((Split-Path -Path $officeTrimmed -Leaf) -ieq 'Office')
			{
				$storemanRoot = Split-Path -Path $officeTrimmed -Parent
				if (-not [string]::IsNullOrWhiteSpace($storemanRoot)) { $XchSearchRoots += [System.IO.Path]::Combine($storemanRoot.TrimEnd('\'), 'XchScale') }
			}
			$XchSearchRoots += [System.IO.Path]::Combine($officeTrimmed, 'XchScale')
		}
		catch { }
	}

	$XchNodesIniCandidates = New-Object System.Collections.Generic.List[string]
	$seenIni = @{ }
	$addIni = {
		param ([string]$IniPath)
		if ([System.IO.File]::Exists($IniPath) -and -not $seenIni.ContainsKey($IniPath.Trim().ToLowerInvariant()))
		{
			$seenIni[$IniPath.Trim().ToLowerInvariant()] = $true
			$XchNodesIniCandidates.Add($IniPath)
		}
	}
	$seenRoots = @{ }
	foreach ($root in $XchSearchRoots)
	{
		$rootPath = ([string]$root).Trim().Trim('"')
		if ([string]::IsNullOrWhiteSpace($rootPath) -or $seenRoots.ContainsKey($rootPath.ToLowerInvariant())) { continue }
		$seenRoots[$rootPath.ToLowerInvariant()] = $true
		try
		{
			if (-not (Test-Path -LiteralPath $rootPath -PathType Container)) { continue }
			& $addIni ([System.IO.Path]::Combine($rootPath, 'XchNodes.ini'))

			# Bounded search (some stores have very large XchScale trees): children, then
			# grandchildren only if nothing was found yet
			$childDirs = @(Get-ChildItem -LiteralPath $rootPath -Directory -ErrorAction SilentlyContinue)
			foreach ($childDir in $childDirs) { & $addIni ([System.IO.Path]::Combine([string]$childDir.FullName, 'XchNodes.ini')) }
			if ($XchNodesIniCandidates.Count -eq 0)
			{
				foreach ($childDir in $childDirs)
				{
					foreach ($grandChildDir in @(Get-ChildItem -LiteralPath $childDir.FullName -Directory -ErrorAction SilentlyContinue))
					{
						& $addIni ([System.IO.Path]::Combine([string]$grandChildDir.FullName, 'XchNodes.ini'))
					}
				}
			}
		}
		catch { }
	}

	# ====================================================================================
	# 1. PRIMARY: STORE DATABASE (TER_TAB, TBS_SCL_ver520)
	# ====================================================================================
	$server = $script:FunctionResults['DBSERVER']
	$database = $script:FunctionResults['DBNAME']
	$ConnectionString = Normalize_Sql_ConnectionString -RawConnectionString $script:FunctionResults['ConnectionString'] -Server $server -Database $database -SqlUser $script:FunctionResults['SQL_UserID'] -SqlPwd $script:FunctionResults['SQL_UserPwd']
	$NodesFromDatabase = $false
	$SqlModule = $script:FunctionResults['SqlModuleName']

	if (-not $SqlModule -or $SqlModule -eq 'None')
	{
		Write_Log "SQL PowerShell module not installed - using the built-in SQL client for node info." "yellow"
	}

	# The lane query filters on this store (TER_TAB.F1056); without a known store number the
	# connection is a placeholder target and would only burn ~15 s on timeouts.
	$hasUsableStoreNumber = (([string]$StoreNumber -replace '[^\d]', '').Length -gt 0)
	if ($ConnectionString -and -not $hasUsableStoreNumber)
	{
		Write_Log "Store number unknown - skipping the store database lookup (no store to query against). Using local files instead." "yellow"
	}

	if ($ConnectionString -and $hasUsableStoreNumber)
	{
		try
		{
			if ($SqlModule -and $SqlModule -ne 'None') { Import-Module $SqlModule -ErrorAction SilentlyContinue | Out-Null }
		}
		catch { }
		$invokeSqlCmd = Get-Command Invoke-Sqlcmd -ErrorAction SilentlyContinue
		$supportsConnStr = $false
		$supportsConnTimeout = $false
		$startupQueryTimeoutSeconds = 10
		$startupConnectionTimeoutSeconds = 5
		$startupConnectionString = $ConnectionString
		if ($invokeSqlCmd)
		{
			$supportsConnStr = $invokeSqlCmd.Parameters.ContainsKey('ConnectionString')
			$supportsConnTimeout = $invokeSqlCmd.Parameters.ContainsKey('ConnectionTimeout')
		}
		elseif ($SqlModule -and $SqlModule -ne 'None')
		{
			Write_Log "SQL module '$SqlModule' not available on this machine - using the built-in SQL client." "yellow"
		}
		else
		{
			Write_Log "SQL module not available on this machine - using the built-in SQL client." "yellow"
		}

		# Invoke-Sqlcmd when available, else SqlClient (returns rows either way)
		$invokeStartupQuery = {
			param ([string]$Query, [int]$QueryTimeout = 10)

			if ($invokeSqlCmd)
			{
				if ($supportsConnStr) { return (& $invokeSqlCmd -ConnectionString $startupConnectionString -Query $Query -QueryTimeout $QueryTimeout -ErrorAction Stop) }
				if ($supportsConnTimeout) { return (& $invokeSqlCmd -ServerInstance $server -Database $database -Query $Query -QueryTimeout $QueryTimeout -ConnectionTimeout $startupConnectionTimeoutSeconds -ErrorAction Stop) }
				return (& $invokeSqlCmd -ServerInstance $server -Database $database -Query $Query -QueryTimeout $QueryTimeout -ErrorAction Stop)
			}

			$dt = New-Object System.Data.DataTable
			$cn = New-Object System.Data.SqlClient.SqlConnection $startupConnectionString
			try
			{
				$cn.Open()
				$cmd = $cn.CreateCommand()
				$cmd.CommandText = $Query
				$cmd.CommandTimeout = $QueryTimeout
				$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
				[void]$da.Fill($dt)
			}
			finally
			{
				try { if ($cn) { $cn.Close() } } catch { }
				try { if ($cn) { $cn.Dispose() } } catch { }
			}
			return $dt
		}

		# Short connect timeout for the startup queries
		if (-not [string]::IsNullOrWhiteSpace($startupConnectionString))
		{
			try
			{
				$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder $startupConnectionString
				$csb.PSBase.ConnectTimeout = $startupConnectionTimeoutSeconds
				$startupConnectionString = $csb.ConnectionString
			}
			catch
			{
				if ($startupConnectionString -notmatch '(?i)\bConnect\s*Timeout\s*=')
				{
					$startupConnectionString = $startupConnectionString.Trim().TrimEnd(';') + ";Connect Timeout=$startupConnectionTimeoutSeconds"
				}
			}
		}

		$NodesFromDatabase = $true
		$safeStoreNumber = ($StoreNumber -replace '[^\d]', '')
		if ($safeStoreNumber.Length -le 3) { $safeStoreNumber = ([int]$safeStoreNumber).ToString('000') }
		else { $safeStoreNumber = ([int]$safeStoreNumber).ToString('0000') }
		try
		{
			$terTabResult = & $invokeStartupQuery "SELECT F1057, F1058, F1125, F1169 FROM TER_TAB WHERE F1056 = '$safeStoreNumber'" $startupQueryTimeoutSeconds
			foreach ($row in $terTabResult)
			{
				. $addTerRow ([string]$row.F1057) ([string]$row.F1058) ([string]$row.F1125) ([string]$row.F1169)
			}

			$tbsSclScalesResult = $null
			try
			{
				$tbsSclScalesResult = & $invokeStartupQuery "SELECT ScaleCode, ScaleName, ScaleLocation, IPNetwork, IPDevice, Active, ScaleBrand, ScaleModel FROM TBS_SCL_ver520 WHERE Active = 'Y'" $startupQueryTimeoutSeconds
			}
			catch
			{
				# A store without the TBS scale table simply has no TBS scales
				if ($_.Exception.Message -notmatch "Invalid object name 'TBS_SCL_ver520'") { throw }
			}

			# Column value as string, $null when empty / DBNull
			$col = {
				param ($v)
				if ([string]::IsNullOrWhiteSpace([string]$v)) { return $null }
				return [string]$v
			}
			$tbsRowOrdinal = 0
			foreach ($row in @($tbsSclScalesResult))
			{
				if ($null -eq $row) { continue }
				$tbsRowOrdinal++
				$rawScaleCode = ([string]$row.ScaleCode).Trim()
				if ([string]::IsNullOrWhiteSpace($rawScaleCode)) { continue }

				$tbsScaleName = & $col $row.ScaleName
				$tbsIPNetwork = & $col $row.IPNetwork
				if ($tbsIPNetwork) { $tbsIPNetwork = $tbsIPNetwork.Trim() }
				$tbsIPDevice = & $col $row.IPDevice
				if ($tbsIPDevice) { $tbsIPDevice = $tbsIPDevice.Trim() }
				$tbsBrand = & $col $row.ScaleBrand
				$tbsLabel = if ($tbsScaleName) { $tbsScaleName } else { "Scale $rawScaleCode" }

				$null = $TbsScaleRows.Add([PSCustomObject]@{
						RawScaleCode  = $rawScaleCode
						TbsRowOrdinal = $tbsRowOrdinal
						ScaleCode     = $rawScaleCode
						BaseScaleCode = $rawScaleCode
						Label         = $tbsLabel
						ScaleName     = $tbsLabel
						Path          = $null
						HostPath      = $null
						ScaleBrand    = $tbsBrand
						Scalevendor   = $tbsBrand
						ScaleModel    = & $col $row.ScaleModel
						ScaleLocation = & $col $row.ScaleLocation
						IPNetwork     = $tbsIPNetwork
						IPDevice      = $tbsIPDevice
						FullIP        = if ($tbsIPNetwork -or $tbsIPDevice) { "$tbsIPNetwork$tbsIPDevice" } else { $null }
						PortRaw       = $null
						TcpPort       = $null
						Section       = $null
						BaseMode      = $null
						Type          = $null
						Baud          = $null
						Sub           = $null
						PlumDir       = $null
						Active        = & $col $row.Active
						Source        = 'TBS'
					})
			}
		}
		catch
		{
			Write_Log "Failed to retrieve counts from the database: $_" "yellow"
			$NodesFromDatabase = $false
		}
	}

	# ====================================================================================
	# 2. FALLBACK: Office\Load\Ter_Load.sql
	# ====================================================================================
	if ((-not $NodesFromDatabase) -and $TerLoadSqlPath -and (Test-Path -LiteralPath $TerLoadSqlPath))
	{
		Write_Log "Using Ter_Load.sql as backup for TER_TAB information." "yellow"
		$content = Get-Content -LiteralPath $TerLoadSqlPath -Raw
		if ($content -match 'INSERT INTO Ter_Load VALUES\s*(.*);' -or $content -match 'INSERT INTO Ter_Load VALUES\s*(.*)')
		{
			foreach ($value in ($matches[1] -split '\),\s*\(' | ForEach-Object { $_.Trim('() ') }))
			{
				$fields = @($value -split "',\s*'" | ForEach-Object { $_.Trim("'") })
				if ($fields.Count -ge 5 -and $fields[0] -eq $StoreNumber)
				{
					. $addTerRow ([string]$fields[1]) ([string]$fields[2]) ([string]$fields[3]) ([string]$fields[4])
				}
			}
		}
	}

	# ====================================================================================
	# 2.5 ENRICH TER 80X SCALES FROM XchNodes.ini (best section per scale wins)
	# ====================================================================================
	$XchNodesIniPathsUsed = @()
	if ($TerScaleBy80X.Count -gt 0 -and $XchNodesIniCandidates.Count -gt 0)
	{
		$AllowedScaleCodeSet = @{ }
		foreach ($allowedCode in @($TerScaleBy80X.Keys | Where-Object { $_ -match '^8\d\d$' })) { $AllowedScaleCodeSet[$allowedCode] = $true }

		# Section names are <store><terminal>, e.g. 001801 or 0001801
		$storeSectionPrefixes = @()
		$storeDigits = ($StoreNumber -replace '[^\d]', '')
		if ($storeDigits)
		{
			try { $storeSectionPrefixes += ([int]$storeDigits).ToString('000'); $storeSectionPrefixes += ([int]$storeDigits).ToString('0000') }
			catch { }
			$storeSectionPrefixes += $storeDigits
			$storeSectionPrefixes = @($storeSectionPrefixes | Select-Object -Unique)
		}

		$BestXchByBaseScaleCode = @{ }
		$sourceRank = 0
		foreach ($candidateIni in @($XchNodesIniCandidates | Select-Object -Unique))
		{
			$sourceRank++
			$XchNodesIniPathsUsed += $candidateIni
			try
			{
				$iniRaw = Get-Content -LiteralPath $candidateIni -Raw -ErrorAction Stop
				foreach ($sectionMatch in [regex]::Matches($iniRaw, '(?ms)^\[(?<name>[^\]]+)\]\s*(?<body>.*?)(?=^\[|\z)'))
				{
					$sectionKey = [string]$sectionMatch.Groups['name'].Value.Trim()
					if ([string]::IsNullOrWhiteSpace($sectionKey)) { continue }

					$currentMap = @{ }
					foreach ($bodyLine in ([string]$sectionMatch.Groups['body'].Value -split "`r?`n"))
					{
						$bodyTrim = $bodyLine.Trim()
						if ([string]::IsNullOrWhiteSpace($bodyTrim) -or $bodyTrim -match '^[;#]') { continue }
						if ($bodyTrim -match '^\s*([^=]+?)\s*=\s*(.*)\s*$') { $currentMap[$matches[1].Trim()] = $matches[2] }
					}
					$get = { param ($k) if ($currentMap.ContainsKey($k)) { [string]$currentMap[$k] } else { $null } }

					$sectionTerminal = $null
					if ($sectionKey -match '(\d{3})$' -and $matches[1] -match '^8\d\d$') { $sectionTerminal = $matches[1] }

					# NUMBER=1 -> 801, NUMBER=801 -> 801
					$numberRaw = & $get 'NUMBER'
					if ($numberRaw) { $numberRaw = $numberRaw.Trim() }
					$scaleCodeFromNumber = $null
					if ($numberRaw -match '^\d+$')
					{
						try
						{
							$n = [int]$numberRaw
							if ($n -ge 0 -and $n -le 99) { $scaleCodeFromNumber = ('8{0:00}' -f $n) }
							elseif ($n -ge 800 -and $n -le 899) { $scaleCodeFromNumber = ('{0:000}' -f $n) }
						}
						catch { }
					}

					# Match rank: 0 = <store><terminal> section, 1 = terminal suffix only, 2 = NUMBER
					$matchedScaleCode = $null
					$matchRank = 99
					if ($sectionTerminal -and $AllowedScaleCodeSet.ContainsKey($sectionTerminal))
					{
						$matchedScaleCode = $sectionTerminal
						$matchRank = 1
						foreach ($prefix in $storeSectionPrefixes) { if ($sectionKey -ieq "$prefix$sectionTerminal") { $matchRank = 0; break } }
					}
					elseif ($scaleCodeFromNumber -and $AllowedScaleCodeSet.ContainsKey($scaleCodeFromNumber))
					{
						$matchedScaleCode = $scaleCodeFromNumber
						$matchRank = 2
					}
					if (-not $matchedScaleCode) { continue }

					# Port=IP 192.168.1.50:3001 -> FullIP / TcpPort / IPNetwork + IPDevice
					$portRaw = & $get 'Port'
					$tcpPort = $null
					$fullIP = $null
					$ipNetwork = $null
					$ipDevice = $null
					if (-not [string]::IsNullOrWhiteSpace($portRaw))
					{
						$portValue = ($portRaw -replace '(?i)^\s*IP\s*', '').Trim()
						if ($portValue -match '^\s*([^:]+?)\s*:(\d+)\s*$')
						{
							$fullIP = $matches[1].Trim()
							$tcpPort = $matches[2].Trim()
						}
						else { $fullIP = $portValue }

						if ($fullIP -match '^\d{1,3}(?:\.\d{1,3}){3}$')
						{
							$lastDot = $fullIP.LastIndexOf('.')
							$ipNetwork = $fullIP.Substring(0, $lastDot + 1)
							$ipDevice = $fullIP.Substring($lastDot + 1)
						}
					}

					$completeness = 0
					if ($fullIP) { $completeness += 10 }
					if (-not [string]::IsNullOrWhiteSpace((& $get 'Scalevendor'))) { $completeness += 5 }
					if (-not [string]::IsNullOrWhiteSpace((& $get 'ScaleModel'))) { $completeness += 5 }
					if ($portRaw) { $completeness += 3 }
					foreach ($k in 'Baud', 'Type', 'BaseMode') { if (-not [string]::IsNullOrWhiteSpace((& $get $k))) { $completeness += 1 } }

					$entryObj = [PSCustomObject]@{
						SourceRank   = $sourceRank
						MatchRank    = $matchRank
						Completeness = $completeness
						Section      = $sectionKey
						BaseMode     = & $get 'BaseMode'
						Type         = & $get 'Type'
						ScaleBrand   = & $get 'Scalevendor'
						ScaleModel   = & $get 'ScaleModel'
						Baud         = & $get 'Baud'
						PortRaw      = $portRaw
						TcpPort      = $tcpPort
						IPNetwork    = $ipNetwork
						IPDevice     = $ipDevice
						FullIP       = $fullIP
						Sub          = & $get 'Sub'
						PlumDir      = & $get 'PlumDir'
					}

					# Keep the best section: lower MatchRank, then more complete, then earlier file
					$existingXch = $BestXchByBaseScaleCode[$matchedScaleCode]
					if (-not $existingXch -or
						$entryObj.MatchRank -lt $existingXch.MatchRank -or
						($entryObj.MatchRank -eq $existingXch.MatchRank -and $entryObj.Completeness -gt $existingXch.Completeness) -or
						($entryObj.MatchRank -eq $existingXch.MatchRank -and $entryObj.Completeness -eq $existingXch.Completeness -and $entryObj.SourceRank -lt $existingXch.SourceRank))
					{
						$BestXchByBaseScaleCode[$matchedScaleCode] = $entryObj
					}
				}
			}
			catch
			{
				Write_Log "Failed to read XchNodes.ini '$candidateIni': $_" "yellow"
			}
		}

		# Fill the TER scales' empty fields from their best XchNodes section
		foreach ($terCode in @($TerScaleBy80X.Keys))
		{
			$xchEntry = $BestXchByBaseScaleCode[$terCode]
			$terEntry = $TerScaleBy80X[$terCode]
			if (-not $xchEntry -or -not $terEntry) { continue }
			foreach ($field in 'IPNetwork', 'IPDevice', 'FullIP', 'PortRaw', 'TcpPort', 'Section', 'BaseMode', 'Type', 'Baud', 'Sub', 'PlumDir', 'ScaleModel')
			{
				if ([string]::IsNullOrWhiteSpace([string]$terEntry.$field) -and $xchEntry.$field) { $terEntry.$field = [string]$xchEntry.$field }
			}
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.ScaleBrand) -and $xchEntry.ScaleBrand)
			{
				$terEntry.ScaleBrand = [string]$xchEntry.ScaleBrand
				$terEntry.Scalevendor = [string]$xchEntry.ScaleBrand
			}
		}
	}

	# ====================================================================================
	# 2.75 DROP TER SCALES THAT ARE ALSO A TBS SCALE (same IP) - the TBS row wins
	# ====================================================================================
	$normalizedIp = {
		param ($e)
		if (-not [string]::IsNullOrWhiteSpace([string]$e.FullIP)) { return ([string]$e.FullIP).Trim().ToLowerInvariant() }
		if (-not [string]::IsNullOrWhiteSpace([string]$e.IPNetwork) -or -not [string]::IsNullOrWhiteSpace([string]$e.IPDevice))
		{
			return ("$([string]$e.IPNetwork)$([string]$e.IPDevice)").Trim().ToLowerInvariant()
		}
		return $null
	}
	$TbsByNormalizedIp = @{ }
	foreach ($tbsEntry in @($TbsScaleRows))
	{
		$ip = & $normalizedIp $tbsEntry
		if ($ip -and -not $TbsByNormalizedIp.ContainsKey($ip)) { $TbsByNormalizedIp[$ip] = $tbsEntry }
	}

	$FinalTerBy80X = @{ }
	foreach ($terCode in @($TerScaleBy80X.Keys | Sort-Object { [int]$_ }))
	{
		$terEntry = $TerScaleBy80X[$terCode]
		$ip = & $normalizedIp $terEntry
		if ($ip -and $TbsByNormalizedIp.ContainsKey($ip))
		{
			$matchingTbsRow = $TbsByNormalizedIp[$ip]
			foreach ($field in 'Path', 'HostPath', 'PortRaw', 'TcpPort', 'Section', 'BaseMode', 'Type', 'Baud', 'Sub', 'PlumDir')
			{
				if ([string]::IsNullOrWhiteSpace([string]$matchingTbsRow.$field) -and $terEntry.$field) { $matchingTbsRow.$field = [string]$terEntry.$field }
			}
			continue
		}
		$FinalTerBy80X[$terCode] = $terEntry
	}

	# ====================================================================================
	# 2.9 FINAL SCALE KEYS AND MAPS (TER by 80X first, then TBS by original ScaleCode)
	# ====================================================================================
	$addScale = {
		param ([string]$finalKey, $entry, [string]$source)

		$finalObj = [PSCustomObject]@{
			Code          = $finalKey
			ScaleCode     = $finalKey
			BaseScaleCode = $entry.BaseScaleCode
			Label         = $entry.Label
			ScaleName     = $entry.ScaleName
			Path          = $entry.Path
			HostPath      = $entry.HostPath
			ScaleBrand    = $entry.ScaleBrand
			Scalevendor   = if ($entry.Scalevendor) { $entry.Scalevendor } else { $entry.ScaleBrand }
			ScaleModel    = $entry.ScaleModel
			ScaleLocation = $entry.ScaleLocation
			IPNetwork     = $entry.IPNetwork
			IPDevice      = $entry.IPDevice
			FullIP        = $entry.FullIP
			PortRaw       = $entry.PortRaw
			TcpPort       = $entry.TcpPort
			Section       = $entry.Section
			BaseMode      = $entry.BaseMode
			Type          = $entry.Type
			Baud          = $entry.Baud
			Sub           = $entry.Sub
			PlumDir       = $entry.PlumDir
			Active        = $entry.Active
			Source        = $source
		}

		$ScaleCodes += $finalKey
		$ScaleCodeToIPInfo[$finalKey] = $finalObj
		$ScaleLabels[$finalKey] = $finalObj.Label
		$ScaleExePaths[$finalKey] = $finalObj.Path
		$ScaleCodeToPath[$finalKey] = $finalObj.Path
		if ($finalObj.ScaleName)
		{
			$ScaleNameToCode[$finalObj.ScaleName] = $finalKey
			$ScaleNameToPath[$finalObj.ScaleName] = $finalObj.Path
			$ScaleCodeToName[$finalKey] = $finalObj.ScaleName
		}
		foreach ($fm in $scaleFieldMaps)
		{
			$fv = $finalObj.($fm[0])
			if ($fv) { (Get-Variable -Name $fm[1] -ValueOnly)[$finalKey] = $fv }
		}
	}

	foreach ($terCode in @($FinalTerBy80X.Keys | Sort-Object { [int]$_ }))
	{
		. $addScale ([string]$terCode) $FinalTerBy80X[$terCode] 'TER'
	}

	# A TBS ScaleCode that collides with an existing key gets a _TBS<n> suffix
	$TbsKeyCounts = @{ }
	foreach ($tbsEntry in @($TbsScaleRows))
	{
		$rawKey = [string]$tbsEntry.RawScaleCode
		if ([string]::IsNullOrWhiteSpace($rawKey)) { continue }
		$finalKey = $rawKey
		if (-not $TbsKeyCounts.ContainsKey($rawKey)) { $TbsKeyCounts[$rawKey] = 1 }
		while ($ScaleCodeToIPInfo.ContainsKey($finalKey))
		{
			$TbsKeyCounts[$rawKey]++
			$finalKey = "{0}_TBS{1}" -f $rawKey, $TbsKeyCounts[$rawKey]
		}
		. $addScale $finalKey $tbsEntry 'TBS'
	}

	# ====================================================================================
	# 3. REVERSE MAPS, COUNTS, BIZERBA (WINDOWS) SCALES
	# ====================================================================================
	foreach ($kv in @($LaneNumToMachineName.GetEnumerator())) { $MachineNameToLaneNum[$kv.Value] = $kv.Key }
	foreach ($kv in @($BackofficeNumToMachineName.GetEnumerator()))
	{
		$MachineNameToBackofficeNum[$kv.Value] = $kv.Key
		if ($BackofficeNumToPath.ContainsKey($kv.Key)) { $MachineNameToBackofficePath[$kv.Value] = $BackofficeNumToPath[$kv.Key] }
	}

	$NumberOfLanes = $LaneMachineNames.Count
	$NumberOfServers = if ($ServerMachineName) { 1 } else { 0 }
	$NumberOfBackoffices = $BackofficeNumToMachineName.Count
	$NumberOfScales = $ScaleCodes.Count

	# Bizerba scales run Windows (C$ shares, credentials, file deployment)
	foreach ($code in @($ScaleCodeToIPInfo.Keys))
	{
		$scale = $ScaleCodeToIPInfo[$code]
		foreach ($field in 'ScaleBrand', 'Scalevendor', 'ScaleModel', 'ScaleName', 'Label')
		{
			if ($scale.$field -and ([string]$scale.$field -match '(?i)bizerba')) { $WindowsScales[$code] = $scale; break }
		}
	}

	. $publish

	# ====================================================================================
	# 4. CACHE WRITE (keys sorted so the file is readable and stable)
	# ====================================================================================
	try
	{
		$numKey = { [double]((([string]$_.Key) -replace '[^\d]', '0')) }
		$sortMap = {
			param ($Map, [object[]]$SortBy)
			$o = [ordered]@{ }
			foreach ($kv in @($Map.GetEnumerator() | Sort-Object $SortBy)) { $o[$kv.Key] = $kv.Value }
			return $o
		}

		$out = [ordered]@{ }
		foreach ($n in $mapNames) { $out[$n] = & $sortMap (Get-Variable -Name $n -ValueOnly) @($numKey, { $_.Key }) }

		# Lanes: only 3-digit lane numbers; the per-machine maps follow the lane order
		$sortedLaneNums = @($LaneNumToMachineName.Keys | Where-Object { $_ -match '^\d{3}$' } | Sort-Object { [int]$_ })
		$out['LaneNumToMachineName'] = [ordered]@{ }
		foreach ($k in 'LaneMachineLabels', 'LaneMachinePath', 'LaneMachineToServerPath') { $out[$k] = [ordered]@{ } }
		foreach ($ln in $sortedLaneNums)
		{
			$m = $LaneNumToMachineName[$ln]
			$out['LaneNumToMachineName'][$ln] = $m
			if (-not $m) { continue }
			foreach ($k in 'LaneMachineLabels', 'LaneMachinePath', 'LaneMachineToServerPath')
			{
				$src = Get-Variable -Name $k -ValueOnly
				if ($src.ContainsKey($m)) { $out[$k][$m] = $src[$m] }
			}
		}
		$byValueThenKey = @({ [double]((([string]$_.Value) -replace '[^\d]', '0')) }, { $_.Key })
		$out['MachineNameToLaneNum'] = & $sortMap $MachineNameToLaneNum $byValueThenKey
		$out['MachineNameToBackofficeNum'] = & $sortMap $MachineNameToBackofficeNum $byValueThenKey
		$out['ScaleNameToCode'] = & $sortMap $ScaleNameToCode @({ $_.Value }, { $_.Key })
		$out['ScaleNameToPath'] = & $sortMap $ScaleNameToPath @({ $_.Key })
		$out['MachineNameToBackofficePath'] = & $sortMap $MachineNameToBackofficePath @({ $_.Key })

		$outLaneMachineNames = @($LaneMachineNames)
		if ($outLaneMachineNames.Count -gt 0 -and $MachineNameToLaneNum.Count -gt 0)
		{
			$outLaneMachineNames = @($outLaneMachineNames | Sort-Object { [double](([string]$MachineNameToLaneNum[$_]) -replace '[^\d]', '0') }, { $_ })
		}

		$results = [ordered]@{
			NumberOfLanes       = $NumberOfLanes
			NumberOfServers     = $NumberOfServers
			NumberOfBackoffices = $NumberOfBackoffices
			NumberOfScales      = $NumberOfScales
			LaneMachineNames    = $outLaneMachineNames
			ScaleCodes          = @($ScaleCodes | Sort-Object { [double](([string]$_) -replace '[^\d]', '0') }, { $_ })
			ServerMachineName   = $ServerMachineName
			ServerLabel         = $ServerLabel
			ServerPath          = $ServerPath
		}
		foreach ($n in $mapNames) { $results[$n] = $out[$n] }
		$results['XchNodesIniPathsUsed'] = @($XchNodesIniPathsUsed | Select-Object -Unique)

		$cacheOut = [ordered]@{
			LastUpdated        = (Get-Date).ToString('o')
			CacheSchemaVersion = $CacheSchemaVersion
			StoreNumber        = $StoreNumberNorm
			Results            = $results
		}
		($cacheOut | ConvertTo-Json -Depth 16) | Set-Content -LiteralPath $CachePath -Encoding UTF8 -Force
	}
	catch { }

	return $Nodes
}

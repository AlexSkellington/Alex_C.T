# ===================================================================================================
#                              FUNCTION: Retrieve_Nodes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Counts the store's hosts, lanes, backoffices, servers and scales and builds the
#   lane-number-to-machine-name map. Prefers database queries (TER_TAB, TBS_SCL_ver520,
#   server check) with a file-system fallback, caches the result per store, updates the
#   GUI count labels and returns the full nodes object.
#
# Parameters:
#   - StoreNumber  [string]: store to count (required)
#   - RefreshCache [switch]: ignore in-memory/file cache and re-query
#
# Variables:
#   - $script:FunctionResults: receives Nodes, LaneNumToMachineName, ScaleCodeToIPInfo and
#     the NumberOfLanes/Servers/Backoffices/Scales counts
#
# Writes:
#   - Setup_Files\Retrieve_Nodes.txt (JSON cache with schema version)
#
# Returns:
#   PSCustomObject with all counts and the lane/scale maps.
# ===================================================================================================

function Retrieve_Nodes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[switch]$RefreshCache
	)

	# ------------------------------------------------------------------------------------------------
	# Per-function cache (small + fast)
	# ------------------------------------------------------------------------------------------------
	$CacheSchemaVersion = 10 # CHANGE: bumped because TBS scales now always keep their original ScaleCode and are never normalized to 80X.
	$CachePath = Join-Path $script:SetupFiles 'Retrieve_Nodes.txt'

	# Normalize store number for cache match only
	$StoreNumberNorm = $StoreNumber
	try
	{
		$sn = ($StoreNumber -replace '[^\d]', '')
		if ($sn.Length -gt 0)
		{
			if ($sn.Length -le 3) { $StoreNumberNorm = ([int]$sn).ToString('000') }
			else { $StoreNumberNorm = ([int]$sn).ToString('0000') }
		}
	}
	catch { $StoreNumberNorm = $StoreNumber }

	# Reuse in-memory results first
	if (-not $RefreshCache)
	{
		try
		{
			$existingStoreNorm = [string]$script:FunctionResults['StoreNumber']
			if ($existingStoreNorm)
			{
				$existingDigits = ($existingStoreNorm -replace '[^\d]', '')
				if ($existingDigits)
				{
					if ($existingDigits.Length -le 3) { $existingStoreNorm = ([int]$existingDigits).ToString('000') }
					else { $existingStoreNorm = ([int]$existingDigits).ToString('0000') }
				}
			}
			if (
				$script:FunctionResults['Nodes'] -and
				($existingStoreNorm -eq $StoreNumberNorm) -and
				$script:FunctionResults['LaneNumToMachineName']
			)
			{
				if ($NodesHost -ne $null) { $NodesHost.Text = "Number of Servers: $($script:FunctionResults['NumberOfServers'])" }
				if ($NodesBackoffices -ne $null) { $NodesBackoffices.Text = "Number of Backoffices: $($script:FunctionResults['NumberOfBackoffices'])" }
				if ($NodesStore -ne $null) { $NodesStore.Text = "Number of Lanes: $($script:FunctionResults['NumberOfLanes'])" }
				if ($scalesLabel -ne $null) { $scalesLabel.Text = "Number of Scales: $($script:FunctionResults['NumberOfScales'])" }
				if ($form -ne $null) { $form.Refresh() }
				return $script:FunctionResults['Nodes']
			}
		}
		catch { }
	}

	# Try load cache file
	$cacheFresh = $false
	$cacheObj = $null
	if ((-not $RefreshCache) -and $CachePath -and (Test-Path $CachePath))
	{
		try
		{
			$raw = Get-Content -LiteralPath $CachePath -Raw -ErrorAction Stop
			if (-not [string]::IsNullOrWhiteSpace($raw))
			{
				$cacheObj = $raw | ConvertFrom-Json -ErrorAction Stop
				if (
					$cacheObj -and
					$cacheObj.CacheSchemaVersion -and
					([int]$cacheObj.CacheSchemaVersion -eq $CacheSchemaVersion) -and
					$cacheObj.StoreNumber -and
					($cacheObj.StoreNumber.ToString() -eq $StoreNumberNorm)
				)
				{
					$cacheFresh = $true
				}
			}
		}
		catch
		{
			$cacheFresh = $false
			$cacheObj = $null
		}
	}

	# If cache is fresh, hydrate FunctionResults + return without doing DB/file work
	if ($cacheFresh -and $cacheObj -and $cacheObj.Results)
	{
		try
		{
			$r = $cacheObj.Results

			# Helper inline converter: PSCustomObject -> hashtable
			$toHash = {
				param ($o)
				$h = @{ }
				if ($o -and $o.PSObject -and $o.PSObject.Properties)
				{
					foreach ($p in $o.PSObject.Properties)
					{
						$h[$p.Name] = $p.Value
					}
				}
				$h
			}

			$NumberOfLanes = [int]$r.NumberOfLanes
			$NumberOfServers = [int]$r.NumberOfServers
			$NumberOfBackoffices = [int]$r.NumberOfBackoffices
			$NumberOfScales = [int]$r.NumberOfScales

			$LaneMachineNames = @($r.LaneMachineNames)
			$ScaleCodes = @($r.ScaleCodes)

			$LaneNumToMachineName = & $toHash $r.LaneNumToMachineName
			$MachineNameToLaneNum = & $toHash $r.MachineNameToLaneNum
			$LaneMachineLabels = & $toHash $r.LaneMachineLabels
			$LaneMachinePath = & $toHash $r.LaneMachinePath
			$LaneMachineToServerPath = & $toHash $r.LaneMachineToServerPath

			$ScaleLabels = & $toHash $r.ScaleLabels
			$ScaleExePaths = & $toHash $r.ScaleExePaths
			$ScaleCodeToIPInfo = & $toHash $r.ScaleCodeToIPInfo
			$ScaleNameToCode = & $toHash $r.ScaleNameToCode
			$ScaleCodeToPath = & $toHash $r.ScaleCodeToPath
			$ScaleNameToPath = & $toHash $r.ScaleNameToPath

			$ScaleCodeToName = & $toHash $r.ScaleCodeToName
			$ScaleCodeToSection = & $toHash $r.ScaleCodeToSection
			$ScaleCodeToBrand = & $toHash $r.ScaleCodeToBrand
			$ScaleCodeToModel = & $toHash $r.ScaleCodeToModel
			$ScaleCodeToLocation = & $toHash $r.ScaleCodeToLocation
			$ScaleCodeToIPNetwork = & $toHash $r.ScaleCodeToIPNetwork
			$ScaleCodeToIPDevice = & $toHash $r.ScaleCodeToIPDevice
			$ScaleCodeToFullIP = & $toHash $r.ScaleCodeToFullIP
			$ScaleCodeToPortRaw = & $toHash $r.ScaleCodeToPortRaw
			$ScaleCodeToTcpPort = & $toHash $r.ScaleCodeToTcpPort
			$ScaleCodeToBaseMode = & $toHash $r.ScaleCodeToBaseMode
			$ScaleCodeToType = & $toHash $r.ScaleCodeToType
			$ScaleCodeToBaud = & $toHash $r.ScaleCodeToBaud
			$ScaleCodeToSub = & $toHash $r.ScaleCodeToSub
			$ScaleCodeToPlumDir = & $toHash $r.ScaleCodeToPlumDir

			$BackofficeNumToMachineName = & $toHash $r.BackofficeNumToMachineName
			$MachineNameToBackofficeNum = & $toHash $r.MachineNameToBackofficeNum
			$BackofficeNumToLabel = & $toHash $r.BackofficeNumToLabel
			$BackofficeNumToPath = & $toHash $r.BackofficeNumToPath
			$MachineNameToBackofficePath = & $toHash $r.MachineNameToBackofficePath

			$ServerMachineName = $r.ServerMachineName
			$ServerLabel = $r.ServerLabel
			$ServerPath = $r.ServerPath

			$WindowsScales = & $toHash $r.WindowsScales

			$XchNodesIniPathsUsed = @()
			try
			{
				if ($r.PSObject.Properties.Name -contains 'XchNodesIniPathsUsed')
				{
					$XchNodesIniPathsUsed = @($r.XchNodesIniPathsUsed)
				}
			}
			catch { }

			$Nodes = [PSCustomObject]@{
				NumberOfLanes			   = $NumberOfLanes
				NumberOfServers		       = $NumberOfServers
				NumberOfBackoffices	       = $NumberOfBackoffices
				NumberOfScales			   = $NumberOfScales
				LaneMachineNames		   = $LaneMachineNames
				LaneNumToMachineName	   = $LaneNumToMachineName
				LaneMachineLabels		   = $LaneMachineLabels
				LaneMachinePath		       = $LaneMachinePath
				LaneMachineToServerPath    = $LaneMachineToServerPath
				ScaleCodes				   = $ScaleCodes
				ScaleLabels			       = $ScaleLabels
				ScaleExePaths			   = $ScaleExePaths
				ScaleCodeToIPInfo		   = $ScaleCodeToIPInfo
				ScaleNameToCode		       = $ScaleNameToCode
				ScaleCodeToPath		       = $ScaleCodeToPath
				ScaleNameToPath		       = $ScaleNameToPath
				ScaleCodeToName		       = $ScaleCodeToName
				ScaleCodeToSection		   = $ScaleCodeToSection
				ScaleCodeToBrand		   = $ScaleCodeToBrand
				ScaleCodeToModel		   = $ScaleCodeToModel
				ScaleCodeToLocation	       = $ScaleCodeToLocation
				ScaleCodeToIPNetwork	   = $ScaleCodeToIPNetwork
				ScaleCodeToIPDevice	       = $ScaleCodeToIPDevice
				ScaleCodeToFullIP		   = $ScaleCodeToFullIP
				ScaleCodeToPortRaw		   = $ScaleCodeToPortRaw
				ScaleCodeToTcpPort		   = $ScaleCodeToTcpPort
				ScaleCodeToBaseMode	       = $ScaleCodeToBaseMode
				ScaleCodeToType		       = $ScaleCodeToType
				ScaleCodeToBaud		       = $ScaleCodeToBaud
				ScaleCodeToSub			   = $ScaleCodeToSub
				ScaleCodeToPlumDir		   = $ScaleCodeToPlumDir
				BackofficeNumToMachineName = $BackofficeNumToMachineName
				BackofficeNumToLabel	   = $BackofficeNumToLabel
				BackofficeNumToPath	       = $BackofficeNumToPath
				ServerMachineName		   = $ServerMachineName
				ServerLabel			       = $ServerLabel
				ServerPath				   = $ServerPath
				WindowsScales			   = $WindowsScales
			}

			$script:FunctionResults['NumberOfLanes'] = $NumberOfLanes
			$script:FunctionResults['NumberOfServers'] = $NumberOfServers
			$script:FunctionResults['NumberOfBackoffices'] = $NumberOfBackoffices
			$script:FunctionResults['NumberOfScales'] = $NumberOfScales
			$script:FunctionResults['LaneMachineNames'] = $LaneMachineNames
			$script:FunctionResults['LaneNumToMachineName'] = $LaneNumToMachineName
			$script:FunctionResults['MachineNameToLaneNum'] = $MachineNameToLaneNum
			$script:FunctionResults['LaneMachineLabels'] = $LaneMachineLabels
			$script:FunctionResults['LaneMachinePath'] = $LaneMachinePath
			$script:FunctionResults['LaneMachineToServerPath'] = $LaneMachineToServerPath
			$script:FunctionResults['ScaleCodes'] = $ScaleCodes
			$script:FunctionResults['ScaleLabels'] = $ScaleLabels
			$script:FunctionResults['ScaleExePaths'] = $ScaleExePaths
			$script:FunctionResults['ScaleCodeToIPInfo'] = $ScaleCodeToIPInfo
			$script:FunctionResults['ScaleNameToCode'] = $ScaleNameToCode
			$script:FunctionResults['ScaleCodeToPath'] = $ScaleCodeToPath
			$script:FunctionResults['ScaleNameToPath'] = $ScaleNameToPath
			$script:FunctionResults['ScaleCodeToName'] = $ScaleCodeToName
			$script:FunctionResults['ScaleCodeToSection'] = $ScaleCodeToSection
			$script:FunctionResults['ScaleCodeToBrand'] = $ScaleCodeToBrand
			$script:FunctionResults['ScaleCodeToModel'] = $ScaleCodeToModel
			$script:FunctionResults['ScaleCodeToLocation'] = $ScaleCodeToLocation
			$script:FunctionResults['ScaleCodeToIPNetwork'] = $ScaleCodeToIPNetwork
			$script:FunctionResults['ScaleCodeToIPDevice'] = $ScaleCodeToIPDevice
			$script:FunctionResults['ScaleCodeToFullIP'] = $ScaleCodeToFullIP
			$script:FunctionResults['ScaleCodeToPortRaw'] = $ScaleCodeToPortRaw
			$script:FunctionResults['ScaleCodeToTcpPort'] = $ScaleCodeToTcpPort
			$script:FunctionResults['ScaleCodeToBaseMode'] = $ScaleCodeToBaseMode
			$script:FunctionResults['ScaleCodeToType'] = $ScaleCodeToType
			$script:FunctionResults['ScaleCodeToBaud'] = $ScaleCodeToBaud
			$script:FunctionResults['ScaleCodeToSub'] = $ScaleCodeToSub
			$script:FunctionResults['ScaleCodeToPlumDir'] = $ScaleCodeToPlumDir
			$script:FunctionResults['BackofficeNumToMachineName'] = $BackofficeNumToMachineName
			$script:FunctionResults['MachineNameToBackofficeNum'] = $MachineNameToBackofficeNum
			$script:FunctionResults['BackofficeNumToLabel'] = $BackofficeNumToLabel
			$script:FunctionResults['BackofficeNumToPath'] = $BackofficeNumToPath
			$script:FunctionResults['MachineNameToBackofficePath'] = $MachineNameToBackofficePath
			$script:FunctionResults['ServerMachineName'] = $ServerMachineName
			$script:FunctionResults['ServerLabel'] = $ServerLabel
			$script:FunctionResults['ServerPath'] = $ServerPath
			$script:FunctionResults['WindowsScales'] = $WindowsScales
			$script:FunctionResults['Nodes'] = $Nodes
			$script:FunctionResults['XchNodesIniPathsUsed'] = $XchNodesIniPathsUsed

			if ($NodesHost -ne $null) { $NodesHost.Text = "Number of Servers: $NumberOfServers" }
			if ($NodesBackoffices -ne $null) { $NodesBackoffices.Text = "Number of Backoffices: $NumberOfBackoffices" }
			if ($NodesStore -ne $null) { $NodesStore.Text = "Number of Lanes: $NumberOfLanes" }
			if ($scalesLabel -ne $null) { $scalesLabel.Text = "Number of Scales: $NumberOfScales" }
			if ($form -ne $null) { $form.Refresh() }

			return $Nodes
		}
		catch
		{
			# If cache hydrate fails for any reason, fall through to normal logic below
		}
	}

	# ====================================================================================
	# INITIALIZE ALL MAPPING STRUCTURES
	# ====================================================================================
	$HostPath = "$OfficePath"
	$LaneMachineNames = @()
	$LaneNumToMachineName = @{ }
	$LaneMachineLabels = @{ }
	$LaneMachinePath = @{ }
	$LaneMachineToServerPath = @{ }
	$ScaleCodes = @()
	$ScaleLabels = @{ }
	$ScaleExePaths = @{ }
	$ScaleCodeToIPInfo = @{ }

	# CHANGE:
	#   TER scales are grouped by 80X only.
	#   TBS scales keep original ScaleCode exactly as stored in TBS_SCL_ver520.
	$TerScaleBy80X = @{ }
	$TbsScaleRows = New-Object System.Collections.ArrayList

	$ScaleCodeToName = @{ }
	$ScaleCodeToSection = @{ }
	$ScaleCodeToBrand = @{ }
	$ScaleCodeToModel = @{ }
	$ScaleCodeToLocation = @{ }
	$ScaleCodeToIPNetwork = @{ }
	$ScaleCodeToIPDevice = @{ }
	$ScaleCodeToFullIP = @{ }
	$ScaleCodeToPortRaw = @{ }
	$ScaleCodeToTcpPort = @{ }
	$ScaleCodeToBaseMode = @{ }
	$ScaleCodeToType = @{ }
	$ScaleCodeToBaud = @{ }
	$ScaleCodeToSub = @{ }
	$ScaleCodeToPlumDir = @{ }

	$BackofficeNumToMachineName = @{ }
	$BackofficeNumToLabel = @{ }
	$BackofficeNumToPath = @{ }
	$ServerMachineName = $null
	$ServerLabel = $null
	$ServerPath = $null

	$MachineNameToLaneNum = @{ }
	$MachineNameToBackofficeNum = @{ }
	$MachineNameToBackofficePath = @{ }

	$TerLoadSqlPath = $null
	if (-not [string]::IsNullOrWhiteSpace($LoadPath))
	{
		$TerLoadSqlPath = Join-Path $LoadPath 'Ter_Load.sql'
	}

	# ====================================================================================
	# Build XchNodes.ini candidate list only.
	# ====================================================================================
	$XchScalePath = $null
	$XchNodesIniPath = $null
	$XchNodesIniCandidates = @()

	$XchSearchRoots = @()

	try
	{
		if ($script:FunctionResults.ContainsKey('StoremanPath') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['StoremanPath']))
		{
			$storemanRoot = ([string]$script:FunctionResults['StoremanPath']).Trim().TrimEnd('\')
			if (-not [string]::IsNullOrWhiteSpace($storemanRoot))
			{
				$XchSearchRoots += ([System.IO.Path]::Combine($storemanRoot, 'XchScale'))
			}
		}
	}
	catch { }

	try
	{
		if ($script:BasePath)
		{
			$storemanRoot = ([string]$script:BasePath).Trim().TrimEnd('\')
			if (-not [string]::IsNullOrWhiteSpace($storemanRoot))
			{
				$XchSearchRoots += ([System.IO.Path]::Combine($storemanRoot, 'XchScale'))
			}
		}
	}
	catch { }

	if (-not [string]::IsNullOrWhiteSpace($OfficePath))
	{
		try
		{
			$officeTrimmed = ([string]$OfficePath).Trim().TrimEnd('\')
			$officeLeaf = Split-Path -Path $officeTrimmed -Leaf
			if ($officeLeaf -ieq 'Office')
			{
				$storemanRoot = Split-Path -Path $officeTrimmed -Parent
				if (-not [string]::IsNullOrWhiteSpace($storemanRoot))
				{
					$XchSearchRoots += ([System.IO.Path]::Combine($storemanRoot.TrimEnd('\'), 'XchScale'))
				}
			}

			$XchSearchRoots += ([System.IO.Path]::Combine($officeTrimmed, 'XchScale'))
		}
		catch { }
	}

	$seenXchRoots = @{ }
	$UniqueXchSearchRoots = @()
	foreach ($root in $XchSearchRoots)
	{
		if ([string]::IsNullOrWhiteSpace([string]$root)) { continue }
		$rootPath = ([string]$root).Trim().Trim('"')
		if ([string]::IsNullOrWhiteSpace($rootPath)) { continue }

		$rootKey = $rootPath.ToLowerInvariant()
		if ($seenXchRoots.ContainsKey($rootKey)) { continue }
		$seenXchRoots[$rootKey] = $true
		$UniqueXchSearchRoots += $rootPath
	}
	$XchSearchRoots = $UniqueXchSearchRoots

	$seenIniPaths = @{ }
	foreach ($rootPath in $XchSearchRoots)
	{
		try
		{
			if (-not (Test-Path -LiteralPath $rootPath -PathType Container)) { continue }

			$directIni = [System.IO.Path]::Combine($rootPath, 'XchNodes.ini')
			if ([System.IO.File]::Exists($directIni))
			{
				$iniKey = $directIni.Trim().ToLowerInvariant()
				if (-not $seenIniPaths.ContainsKey($iniKey))
				{
					$seenIniPaths[$iniKey] = $true
					$XchNodesIniCandidates += $directIni
				}
			}

			# Avoid a full recursive crawl during startup. Some stores have very large/problematic
			# XchScale trees, and a bounded search is enough for the normal XchNodes.ini layouts.
			$childDirs = @()
			try { $childDirs = @(Get-ChildItem -LiteralPath $rootPath -Directory -ErrorAction SilentlyContinue) }
			catch { $childDirs = @() }

			foreach ($childDir in $childDirs)
			{
				if (-not $childDir) { continue }
				$iniFull = [System.IO.Path]::Combine([string]$childDir.FullName, 'XchNodes.ini')
				if (-not [System.IO.File]::Exists($iniFull)) { continue }

				$iniKey = $iniFull.Trim().ToLowerInvariant()
				if (-not $seenIniPaths.ContainsKey($iniKey))
				{
					$seenIniPaths[$iniKey] = $true
					$XchNodesIniCandidates += $iniFull
				}
			}

			if ($XchNodesIniCandidates.Count -eq 0 -and $childDirs.Count -gt 0)
			{
				foreach ($childDir in $childDirs)
				{
					if (-not $childDir) { continue }

					$grandChildDirs = @()
					try { $grandChildDirs = @(Get-ChildItem -LiteralPath $childDir.FullName -Directory -ErrorAction SilentlyContinue) }
					catch { $grandChildDirs = @() }

					foreach ($grandChildDir in $grandChildDirs)
					{
						if (-not $grandChildDir) { continue }
						$iniFull = [System.IO.Path]::Combine([string]$grandChildDir.FullName, 'XchNodes.ini')
						if (-not [System.IO.File]::Exists($iniFull)) { continue }

						$iniKey = $iniFull.Trim().ToLowerInvariant()
						if (-not $seenIniPaths.ContainsKey($iniKey))
						{
							$seenIniPaths[$iniKey] = $true
							$XchNodesIniCandidates += $iniFull
						}
					}
				}
			}
		}
		catch { }
	}

	if ($XchNodesIniCandidates.Count -gt 0)
	{
		$XchNodesIniPath = $XchNodesIniCandidates[0]
		try { $XchScalePath = Split-Path -Path $XchNodesIniPath -Parent }
		catch { }
	}

	$server = $script:FunctionResults['DBSERVER']
	$database = $script:FunctionResults['DBNAME']
	$ConnectionString = Normalize_Sql_ConnectionString -RawConnectionString $script:FunctionResults['ConnectionString'] -Server $server -Database $database -SqlUser $script:FunctionResults['SQL_UserID'] -SqlPwd $script:FunctionResults['SQL_UserPwd']
	$NodesFromDatabase = $false
	$SqlModule = $script:FunctionResults['SqlModuleName']

	# ====================================================================================
	# DETECT SQL MODULE
	# ====================================================================================
	if (-not $SqlModule -or $SqlModule -eq 'None')
	{
		Write_Log "SQL PowerShell module not installed - using the built-in SQL client for node info." "yellow"
	}

	# ====================================================================================
	# 1. PRIMARY: LOAD FROM DATABASE (TER_TAB, TBS_SCL_ver520)
	# ====================================================================================
	# The lane query targets this store's rows (TER_TAB WHERE F1056 = store number), so
	# a DB connection is only useful when the store number is actually known. Without the
	# storeman folder the store number is 'N/A' and the connection (if any) is a
	# localhost/placeholder target - opening it just burns ~15s on timeouts for a query
	# that can never return valid lanes. The local-file fallbacks below run instead.
	$hasUsableStoreNumber = $false
	try
	{
		$hasUsableStoreNumber = (([string]$StoreNumber -replace '[^\d]', '').Length -gt 0)
	}
	catch { }

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
		if ($invokeSqlCmd) { $supportsConnStr = $invokeSqlCmd.Parameters.ContainsKey('ConnectionString') }
		if ($invokeSqlCmd) { $supportsConnTimeout = $invokeSqlCmd.Parameters.ContainsKey('ConnectionTimeout') }
		else
		{
			if ($SqlModule -and $SqlModule -ne 'None')
			{
				Write_Log "SQL module '$SqlModule' not available on this machine - using the built-in SQL client." "yellow"
			}
			else
			{
				Write_Log "SQL module not available on this machine - using the built-in SQL client." "yellow"
			}
		}

		$invokeStartupQuery = {
			param(
				[string]$Query,
				[int]$QueryTimeout = 10
			)

			if ($invokeSqlCmd)
			{
				if ($supportsConnStr)
				{
					return (& $invokeSqlCmd -ConnectionString $startupConnectionString -Query $Query -QueryTimeout $QueryTimeout -ErrorAction Stop)
				}
				if ($supportsConnTimeout)
				{
					return (& $invokeSqlCmd -ServerInstance $server -Database $database -Query $Query -QueryTimeout $QueryTimeout -ConnectionTimeout $startupConnectionTimeoutSeconds -ErrorAction Stop)
				}
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

		if ($supportsConnStr -and -not [string]::IsNullOrWhiteSpace($startupConnectionString))
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
		if (-not [string]::IsNullOrWhiteSpace($safeStoreNumber))
		{
			if ($safeStoreNumber.Length -le 3) { $safeStoreNumber = ([int]$safeStoreNumber).ToString('000') }
			else { $safeStoreNumber = ([int]$safeStoreNumber).ToString('0000') }
		}
		else { $safeStoreNumber = '000' }
		try
		{
			$queryTerTab = @"
SELECT F1057, F1058, F1125, F1169
FROM TER_TAB
WHERE F1056 = '$safeStoreNumber'
"@
			$terTabResult = & $invokeStartupQuery $queryTerTab $startupQueryTimeoutSeconds

			foreach ($row in $terTabResult)
			{
				$terminal = [string]$row.F1057
				$label = [string]$row.F1058
				$path = [string]$row.F1125
				$hostPath = [string]$row.F1169

				if (
					($terminal -match '^0\d\d$' -or $path -match '^\\\\[^\\]+\\') -and
					$terminal -notmatch '^8' -and $terminal -notmatch '^9'
				)
				{
					$machineName = $null
					if ($path -match '\\\\([^\\]+)\\') { $machineName = $matches[1] }
					else { $machineName = $terminal }

					$LaneMachineNames += $machineName
					if ($terminal -match '^0\d\d$') { $LaneNumToMachineName[$terminal] = $machineName }
					$LaneMachineLabels[$machineName] = $label
					$LaneMachinePath[$machineName] = $path
					$LaneMachineToServerPath[$machineName] = $hostPath
				}
				elseif ($terminal -match '^8\d\d$' -and $terminal -notmatch '^0' -and $terminal -notmatch '^9' -and $path -match '(?i)XchScale\\XchScale\.exe$')
				{
					# CHANGE:
					#   TER stays grouped by 80X only.
					if (-not $TerScaleBy80X.ContainsKey($terminal))
					{
						$TerScaleBy80X[$terminal] = [PSCustomObject]@{
							Code		  = $terminal
							ScaleCode	  = $terminal
							BaseScaleCode = $terminal
							Label		  = $label
							ScaleName	  = $label
							Path		  = $path
							HostPath	  = $hostPath
							ScaleBrand    = $null
							Scalevendor   = $null
							ScaleModel    = $null
							ScaleLocation = $null
							IPNetwork	  = $null
							IPDevice	  = $null
							FullIP	      = $null
							PortRaw	      = $null
							TcpPort	      = $null
							Section	      = $null
							BaseMode	  = $null
							Type		  = $null
							Baud		  = $null
							Sub		      = $null
							PlumDir	      = $null
							Active	      = $null
							Source	      = 'TER'
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
					$ServerMachineName = $path -replace '^@', ''
					$ServerLabel = $label
					$ServerPath = $path
				}
				elseif ($terminal -match '^9(0[2-9]|[1-8]\d|9[0-8])$' -and $path -match '^@[^@]+$')
				{
					$machineName = $path -replace '^@', ''
					$BackofficeNumToMachineName[$terminal] = $machineName
					$BackofficeNumToLabel[$terminal] = $label
					$BackofficeNumToPath[$terminal] = $path
				}
			}

			$NumberOfLanes = $LaneMachineNames.Count
			$NumberOfServers = if ($ServerMachineName) { 1 }
			else { 0 }
			$NumberOfBackoffices = $BackofficeNumToMachineName.Count

			$queryTbsSclScales = @"
SELECT ScaleCode, ScaleName, ScaleLocation, IPNetwork, IPDevice, Active, ScaleBrand, ScaleModel
FROM TBS_SCL_ver520
WHERE Active = 'Y'
"@
			try
			{
				$tbsSclScalesResult = & $invokeStartupQuery $queryTbsSclScales $startupQueryTimeoutSeconds
			}
			catch
			{
				if ($_.Exception.Message -match "Invalid object name 'TBS_SCL_ver520'")
				{
					$tbsSclScalesResult = $null
				}
				else
				{
					throw
				}
			}

			if ($tbsSclScalesResult)
			{
				$tbsRowOrdinal = 0
				foreach ($row in $tbsSclScalesResult)
				{
					$tbsRowOrdinal++

					# CHANGE:
					#   Keep original TBS ScaleCode exactly as stored.
					$rawScaleCode = $null
					try { $rawScaleCode = ([string]$row.ScaleCode).Trim() }
					catch { $rawScaleCode = $null }
					if ([string]::IsNullOrWhiteSpace($rawScaleCode)) { continue }

					$tbsScaleName = $null
					if (-not [string]::IsNullOrWhiteSpace([string]$row.ScaleName)) { $tbsScaleName = [string]$row.ScaleName }

					$tbsScaleLocation = $null
					if (-not [string]::IsNullOrWhiteSpace([string]$row.ScaleLocation)) { $tbsScaleLocation = [string]$row.ScaleLocation }

					$tbsIPNetwork = $null
					if (-not [string]::IsNullOrWhiteSpace([string]$row.IPNetwork)) { $tbsIPNetwork = [string]$row.IPNetwork.Trim() }

					$tbsIPDevice = $null
					if (-not [string]::IsNullOrWhiteSpace([string]$row.IPDevice)) { $tbsIPDevice = [string]$row.IPDevice.Trim() }

					$tbsFullIP = $null
					if ($tbsIPNetwork -or $tbsIPDevice)
					{
						$tbsFullIP = "$tbsIPNetwork$tbsIPDevice"
					}

					$tbsActive = $null
					if (-not [string]::IsNullOrWhiteSpace([string]$row.Active)) { $tbsActive = [string]$row.Active }

					$tbsBrand = $null
					if (-not [string]::IsNullOrWhiteSpace([string]$row.ScaleBrand)) { $tbsBrand = [string]$row.ScaleBrand }

					$tbsModel = $null
					if (-not [string]::IsNullOrWhiteSpace([string]$row.ScaleModel)) { $tbsModel = [string]$row.ScaleModel }

					$null = $TbsScaleRows.Add(([PSCustomObject]@{
								RawScaleCode  = $rawScaleCode
								TbsRowOrdinal = $tbsRowOrdinal
								ScaleCode	  = $rawScaleCode
								BaseScaleCode = $rawScaleCode
								Label		  = if ($tbsScaleName) { $tbsScaleName } else { "Scale $rawScaleCode" }
								ScaleName	  = if ($tbsScaleName) { $tbsScaleName } else { "Scale $rawScaleCode" }
								Path		  = $null
								HostPath	  = $null
								ScaleBrand    = $tbsBrand
								Scalevendor   = $tbsBrand
								ScaleModel    = $tbsModel
								ScaleLocation = $tbsScaleLocation
								IPNetwork	  = $tbsIPNetwork
								IPDevice	  = $tbsIPDevice
								FullIP	      = $tbsFullIP
								PortRaw	      = $null
								TcpPort	      = $null
								Section	      = $null
								BaseMode	  = $null
								Type		  = $null
								Baud		  = $null
								Sub		      = $null
								PlumDir	      = $null
								Active	      = $tbsActive
								Source	      = 'TBS'
							}))
				}
			}
		}
		catch
		{
			Write_Log "Failed to retrieve counts from the database: $_" "yellow"
			$NodesFromDatabase = $false
		}
	}

	# ====================================================================================
	# 2. FALLBACK: LOAD FROM Ter_Load.sql IF DB IS NOT AVAILABLE
	# ====================================================================================
	$TerLoadUsed = $false
	if ((-not $NodesFromDatabase) -and $TerLoadSqlPath -and (Test-Path $TerLoadSqlPath))
	{
		Write_Log "Using Ter_Load.sql as backup for TER_TAB information." "yellow"
		$TerLoadUsed = $true
		$content = Get-Content $TerLoadSqlPath -Raw

		if ($content -match 'INSERT INTO Ter_Load VALUES\s*(.*);' -or $content -match 'INSERT INTO Ter_Load VALUES\s*(.*)')
		{
			$insertBlock = $matches[1]
			$values = $insertBlock -split '\),\s*\(' | ForEach-Object { $_.Trim('() ') }

			foreach ($value in $values)
			{
				$fields = $value -split "',\s*'" | ForEach-Object { $_.Trim("'") }
				if ($fields.Count -ge 5 -and $fields[0] -eq $StoreNumber)
				{
					$terminal = [string]$fields[1]
					$label = [string]$fields[2]
					$path = [string]$fields[3]
					$hostPath = [string]$fields[4]

					if (
						($terminal -match '^0\d\d$' -or $path -match '^\\\\[^\\]+\\') -and
						$terminal -notmatch '^8' -and $terminal -notmatch '^9'
					)
					{
						$machineName = $null
						if ($path -match '\\\\([^\\]+)\\') { $machineName = $matches[1] }
						else { $machineName = $terminal }

						$LaneMachineNames += $machineName
						if ($terminal -match '^0\d\d$') { $LaneNumToMachineName[$terminal] = $machineName }
						$LaneMachineLabels[$machineName] = $label
						$LaneMachinePath[$machineName] = $path
						$LaneMachineToServerPath[$machineName] = $hostPath
					}
					elseif ($terminal -match '^8\d\d$' -and $terminal -notmatch '^0' -and $terminal -notmatch '^9' -and $path -match '(?i)XchScale\\XchScale\.exe$')
					{
						if (-not $TerScaleBy80X.ContainsKey($terminal))
						{
							$TerScaleBy80X[$terminal] = [PSCustomObject]@{
								Code		  = $terminal
								ScaleCode	  = $terminal
								BaseScaleCode = $terminal
								Label		  = $label
								ScaleName	  = $label
								Path		  = $path
								HostPath	  = $hostPath
								ScaleBrand    = $null
								Scalevendor   = $null
								ScaleModel    = $null
								ScaleLocation = $null
								IPNetwork	  = $null
								IPDevice	  = $null
								FullIP	      = $null
								PortRaw	      = $null
								TcpPort	      = $null
								Section	      = $null
								BaseMode	  = $null
								Type		  = $null
								Baud		  = $null
								Sub		      = $null
								PlumDir	      = $null
								Active	      = $null
								Source	      = 'TER'
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
						$ServerMachineName = $path -replace '^@', ''
						$ServerLabel = $label
						$ServerPath = $path
					}
					elseif ($terminal -match '^9(0[2-9]|[1-8]\d|9[0-8])$' -and $path -match '^@[^@]+$')
					{
						$machineName = $path -replace '^@', ''
						$BackofficeNumToMachineName[$terminal] = $machineName
						$BackofficeNumToLabel[$terminal] = $label
						$BackofficeNumToPath[$terminal] = $path
					}
				}
			}
		}

		$NumberOfLanes = $LaneMachineNames.Count
		$NumberOfServers = if ($ServerMachineName) { 1 }
		else { 0 }
		$NumberOfBackoffices = $BackofficeNumToMachineName.Count
	}

	# ====================================================================================
	# 2.5 ENRICH TER 80X SCALES FROM ALL XchNodes.ini CANDIDATES
	# ====================================================================================
	$XchNodesIniPathsUsed = @()
	if ($TerScaleBy80X.Count -gt 0 -and $XchNodesIniCandidates.Count -gt 0)
	{
		$AllowedScaleCodes = @($TerScaleBy80X.Keys | Where-Object { $_ -match '^8\d\d$' } | Select-Object -Unique)
		$AllowedScaleCodeSet = @{ }
		foreach ($allowedCode in $AllowedScaleCodes) { $AllowedScaleCodeSet[$allowedCode] = $true }

		$storeDigits = ($StoreNumber -replace '[^\d]', '')
		$storeSectionPrefixes = @()
		if (-not [string]::IsNullOrWhiteSpace($storeDigits))
		{
			try
			{
				$storeSectionPrefixes += ([int]$storeDigits).ToString('000')
				$storeSectionPrefixes += ([int]$storeDigits).ToString('0000')
			}
			catch
			{
				$storeSectionPrefixes += $storeDigits
			}
			if (-not ($storeSectionPrefixes -contains $storeDigits)) { $storeSectionPrefixes += $storeDigits }
		}
		$storeSectionPrefixes = @($storeSectionPrefixes | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)

		$BestXchByBaseScaleCode = @{ }
		$sourceRank = 0

		foreach ($candidateIni in @($XchNodesIniCandidates | Select-Object -Unique))
		{
			$sourceRank++
			$XchNodesIniPathsUsed += $candidateIni

			try
			{
				$iniRaw = Get-Content -LiteralPath $candidateIni -Raw -ErrorAction Stop
				$sectionMatches = [regex]::Matches($iniRaw, '(?ms)^\[(?<name>[^\]]+)\]\s*(?<body>.*?)(?=^\[|\z)')

				foreach ($sectionMatch in $sectionMatches)
				{
					$sectionKey = [string]$sectionMatch.Groups['name'].Value.Trim()
					$sectionBody = [string]$sectionMatch.Groups['body'].Value
					if ([string]::IsNullOrWhiteSpace($sectionKey)) { continue }

					$currentMap = [ordered]@{ }
					foreach ($bodyLineRaw in ($sectionBody -split "`r?`n"))
					{
						$bodyLine = [string]$bodyLineRaw
						$bodyTrim = $bodyLine.Trim()
						if ([string]::IsNullOrWhiteSpace($bodyTrim) -or $bodyTrim -match '^[;#]') { continue }

						if ($bodyTrim -match '^\s*([^=]+?)\s*=\s*(.*)\s*$')
						{
							$key = $matches[1].Trim()
							$val = $matches[2]
							$currentMap[$key] = $val
						}
					}

					$sectionTerminal = $null
					if ($sectionKey -match '(\d{3})$')
					{
						if ($matches[1] -match '^8\d\d$') { $sectionTerminal = $matches[1] }
					}

					$numberRaw = $null
					if ($currentMap.Contains('NUMBER')) { $numberRaw = ([string]$currentMap['NUMBER']).Trim() }

					$scaleCodeFromNumber = $null
					if (-not [string]::IsNullOrWhiteSpace($numberRaw) -and $numberRaw -match '^\d+$')
					{
						try
						{
							$n = [int]$numberRaw
							if ($n -ge 0 -and $n -le 99)
							{
								$scaleCodeFromNumber = ('8{0:00}' -f $n)
							}
							elseif ($n -ge 800 -and $n -le 899)
							{
								$scaleCodeFromNumber = ('{0:000}' -f $n)
							}
						}
						catch { }
					}

					$matchedScaleCode = $null
					$matchRank = 99

					if ($sectionTerminal -and $AllowedScaleCodeSet.ContainsKey($sectionTerminal))
					{
						foreach ($prefix in $storeSectionPrefixes)
						{
							if ($sectionKey -ieq "$prefix$sectionTerminal")
							{
								$matchedScaleCode = $sectionTerminal
								$matchRank = 0
								break
							}
						}
					}

					if (-not $matchedScaleCode -and $sectionTerminal -and $AllowedScaleCodeSet.ContainsKey($sectionTerminal))
					{
						$matchedScaleCode = $sectionTerminal
						$matchRank = 1
					}

					if (-not $matchedScaleCode -and $scaleCodeFromNumber -and $AllowedScaleCodeSet.ContainsKey($scaleCodeFromNumber))
					{
						$matchedScaleCode = $scaleCodeFromNumber
						$matchRank = 2
					}

					if (-not $matchedScaleCode) { continue }

					$portRaw = $null
					if ($currentMap.Contains('Port')) { $portRaw = [string]$currentMap['Port'] }

					$portValue = $null
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
						else
						{
							$fullIP = $portValue
						}

						if ($fullIP -match '^\d{1,3}(?:\.\d{1,3}){3}$')
						{
							$lastDot = $fullIP.LastIndexOf('.')
							if ($lastDot -gt 0)
							{
								$ipNetwork = $fullIP.Substring(0, $lastDot + 1)
								$ipDevice = $fullIP.Substring($lastDot + 1)
							}
						}
					}

					$completeness = 0
					if ($fullIP) { $completeness += 10 }
					if ($currentMap.Contains('Scalevendor') -and -not [string]::IsNullOrWhiteSpace([string]$currentMap['Scalevendor'])) { $completeness += 5 }
					if ($currentMap.Contains('ScaleModel') -and -not [string]::IsNullOrWhiteSpace([string]$currentMap['ScaleModel'])) { $completeness += 5 }
					if ($portRaw) { $completeness += 3 }
					if ($currentMap.Contains('Baud') -and -not [string]::IsNullOrWhiteSpace([string]$currentMap['Baud'])) { $completeness += 1 }
					if ($currentMap.Contains('Type') -and -not [string]::IsNullOrWhiteSpace([string]$currentMap['Type'])) { $completeness += 1 }
					if ($currentMap.Contains('BaseMode') -and -not [string]::IsNullOrWhiteSpace([string]$currentMap['BaseMode'])) { $completeness += 1 }

					$entryObj = [PSCustomObject]@{
						SourceIniPath	    = $candidateIni
						SourceRank		    = $sourceRank
						MatchRank		    = $matchRank
						Completeness	    = $completeness
						Section			    = $sectionKey
						SectionTerminal	    = $sectionTerminal
						Number			    = $numberRaw
						ScaleCodeFromNumber = $scaleCodeFromNumber
						ScaleCode		    = $matchedScaleCode
						BaseMode		    = if ($currentMap.Contains('BaseMode')) { [string]$currentMap['BaseMode'] } else { $null }
						Type			    = if ($currentMap.Contains('Type')) { [string]$currentMap['Type'] } else { $null }
						ScaleBrand		    = if ($currentMap.Contains('Scalevendor')) { [string]$currentMap['Scalevendor'] } else { $null }
						Scalevendor		    = if ($currentMap.Contains('Scalevendor')) { [string]$currentMap['Scalevendor'] } else { $null }
						ScaleModel		    = if ($currentMap.Contains('ScaleModel')) { [string]$currentMap['ScaleModel'] } else { $null }
						Baud			    = if ($currentMap.Contains('Baud')) { [string]$currentMap['Baud'] } else { $null }
						PortRaw			    = $portRaw
						Port			    = $fullIP
						TcpPort			    = $tcpPort
						IPNetwork		    = $ipNetwork
						IPDevice		    = $ipDevice
						FullIP			    = $fullIP
						Sub				    = if ($currentMap.Contains('Sub')) { [string]$currentMap['Sub'] } else { $null }
						PlumDir			    = if ($currentMap.Contains('PlumDir')) { [string]$currentMap['PlumDir'] } else { $null }
						EnableLogging	    = if ($currentMap.Contains('EnableLogging')) { [string]$currentMap['EnableLogging'] } else { $null }
						ShowGUI			    = if ($currentMap.Contains('ShowGUI')) { [string]$currentMap['ShowGUI'] } else { $null }
					}

					$replaceXch = $false
					if (-not $BestXchByBaseScaleCode.ContainsKey($matchedScaleCode))
					{
						$replaceXch = $true
					}
					else
					{
						$existingXch = $BestXchByBaseScaleCode[$matchedScaleCode]
						if ($entryObj.MatchRank -lt $existingXch.MatchRank)
						{
							$replaceXch = $true
						}
						elseif ($entryObj.MatchRank -eq $existingXch.MatchRank -and $entryObj.Completeness -gt $existingXch.Completeness)
						{
							$replaceXch = $true
						}
						elseif ($entryObj.MatchRank -eq $existingXch.MatchRank -and $entryObj.Completeness -eq $existingXch.Completeness -and $entryObj.SourceRank -lt $existingXch.SourceRank)
						{
							$replaceXch = $true
						}
					}

					if ($replaceXch)
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

		foreach ($terCode in @($TerScaleBy80X.Keys))
		{
			if (-not $BestXchByBaseScaleCode.ContainsKey($terCode)) { continue }
			$xchEntry = $BestXchByBaseScaleCode[$terCode]
			$terEntry = $TerScaleBy80X[$terCode]
			if (-not $xchEntry -or -not $terEntry) { continue }

			if ([string]::IsNullOrWhiteSpace([string]$terEntry.IPNetwork) -and $xchEntry.IPNetwork) { $terEntry.IPNetwork = [string]$xchEntry.IPNetwork }
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.IPDevice) -and $xchEntry.IPDevice) { $terEntry.IPDevice = [string]$xchEntry.IPDevice }
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.FullIP) -and $xchEntry.FullIP) { $terEntry.FullIP = [string]$xchEntry.FullIP }
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.PortRaw) -and $xchEntry.PortRaw) { $terEntry.PortRaw = [string]$xchEntry.PortRaw }
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.TcpPort) -and $xchEntry.TcpPort) { $terEntry.TcpPort = [string]$xchEntry.TcpPort }
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.Section) -and $xchEntry.Section) { $terEntry.Section = [string]$xchEntry.Section }
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.BaseMode) -and $xchEntry.BaseMode) { $terEntry.BaseMode = [string]$xchEntry.BaseMode }
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.Type) -and $xchEntry.Type) { $terEntry.Type = [string]$xchEntry.Type }
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.Baud) -and $xchEntry.Baud) { $terEntry.Baud = [string]$xchEntry.Baud }
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.Sub) -and $xchEntry.Sub) { $terEntry.Sub = [string]$xchEntry.Sub }
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.PlumDir) -and $xchEntry.PlumDir) { $terEntry.PlumDir = [string]$xchEntry.PlumDir }
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.ScaleBrand) -and $xchEntry.ScaleBrand)
			{
				$terEntry.ScaleBrand = [string]$xchEntry.ScaleBrand
				$terEntry.Scalevendor = [string]$xchEntry.ScaleBrand
			}
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.ScaleModel) -and $xchEntry.ScaleModel) { $terEntry.ScaleModel = [string]$xchEntry.ScaleModel }
		}
	}

	# ====================================================================================
	# 2.75 REMOVE TER SCALES THAT MATCH A TBS SCALE BY SAME IP
	# ====================================================================================
	$TbsByNormalizedIp = @{ }
	foreach ($tbsEntry in @($TbsScaleRows))
	{
		if (-not $tbsEntry) { continue }

		$tbsIpNorm = $null
		if (-not [string]::IsNullOrWhiteSpace([string]$tbsEntry.FullIP))
		{
			$tbsIpNorm = ([string]$tbsEntry.FullIP).Trim().ToLowerInvariant()
		}
		elseif (-not [string]::IsNullOrWhiteSpace([string]$tbsEntry.IPNetwork) -or -not [string]::IsNullOrWhiteSpace([string]$tbsEntry.IPDevice))
		{
			$tbsIpNorm = ("$([string]$tbsEntry.IPNetwork)$([string]$tbsEntry.IPDevice)").Trim().ToLowerInvariant()
		}

		if ([string]::IsNullOrWhiteSpace($tbsIpNorm)) { continue }

		if (-not $TbsByNormalizedIp.ContainsKey($tbsIpNorm))
		{
			$TbsByNormalizedIp[$tbsIpNorm] = New-Object System.Collections.ArrayList
		}
		$null = $TbsByNormalizedIp[$tbsIpNorm].Add($tbsEntry)
	}

	$FinalTerBy80X = @{ }
	foreach ($terCode in @($TerScaleBy80X.Keys | Sort-Object { [int]$_ }))
	{
		$terEntry = $TerScaleBy80X[$terCode]
		if (-not $terEntry) { continue }

		$terIpNorm = $null
		if (-not [string]::IsNullOrWhiteSpace([string]$terEntry.FullIP))
		{
			$terIpNorm = ([string]$terEntry.FullIP).Trim().ToLowerInvariant()
		}
		elseif (-not [string]::IsNullOrWhiteSpace([string]$terEntry.IPNetwork) -or -not [string]::IsNullOrWhiteSpace([string]$terEntry.IPDevice))
		{
			$terIpNorm = ("$([string]$terEntry.IPNetwork)$([string]$terEntry.IPDevice)").Trim().ToLowerInvariant()
		}

		if (-not [string]::IsNullOrWhiteSpace($terIpNorm) -and $TbsByNormalizedIp.ContainsKey($terIpNorm))
		{
			# Same IP exists in TER and TBS -> keep TBS only
			$matchingTbsRow = @($TbsByNormalizedIp[$terIpNorm]) | Select-Object -First 1
			if ($matchingTbsRow)
			{
				if ([string]::IsNullOrWhiteSpace([string]$matchingTbsRow.Path) -and $terEntry.Path) { $matchingTbsRow.Path = [string]$terEntry.Path }
				if ([string]::IsNullOrWhiteSpace([string]$matchingTbsRow.HostPath) -and $terEntry.HostPath) { $matchingTbsRow.HostPath = [string]$terEntry.HostPath }
				if ([string]::IsNullOrWhiteSpace([string]$matchingTbsRow.PortRaw) -and $terEntry.PortRaw) { $matchingTbsRow.PortRaw = [string]$terEntry.PortRaw }
				if ([string]::IsNullOrWhiteSpace([string]$matchingTbsRow.TcpPort) -and $terEntry.TcpPort) { $matchingTbsRow.TcpPort = [string]$terEntry.TcpPort }
				if ([string]::IsNullOrWhiteSpace([string]$matchingTbsRow.Section) -and $terEntry.Section) { $matchingTbsRow.Section = [string]$terEntry.Section }
				if ([string]::IsNullOrWhiteSpace([string]$matchingTbsRow.BaseMode) -and $terEntry.BaseMode) { $matchingTbsRow.BaseMode = [string]$terEntry.BaseMode }
				if ([string]::IsNullOrWhiteSpace([string]$matchingTbsRow.Type) -and $terEntry.Type) { $matchingTbsRow.Type = [string]$terEntry.Type }
				if ([string]::IsNullOrWhiteSpace([string]$matchingTbsRow.Baud) -and $terEntry.Baud) { $matchingTbsRow.Baud = [string]$terEntry.Baud }
				if ([string]::IsNullOrWhiteSpace([string]$matchingTbsRow.Sub) -and $terEntry.Sub) { $matchingTbsRow.Sub = [string]$terEntry.Sub }
				if ([string]::IsNullOrWhiteSpace([string]$matchingTbsRow.PlumDir) -and $terEntry.PlumDir) { $matchingTbsRow.PlumDir = [string]$terEntry.PlumDir }
			}

			continue
		}

		$FinalTerBy80X[$terCode] = $terEntry
	}

	# ====================================================================================
	# 2.9 MATERIALIZE FINAL SCALE KEYS / MAPS
	# ====================================================================================
	$ScaleCodes = @()
	$ScaleLabels = @{ }
	$ScaleExePaths = @{ }
	$ScaleCodeToIPInfo = @{ }
	$ScaleNameToCode = @{ }
	$ScaleCodeToPath = @{ }
	$ScaleNameToPath = @{ }

	$ScaleCodeToName = @{ }
	$ScaleCodeToSection = @{ }
	$ScaleCodeToBrand = @{ }
	$ScaleCodeToModel = @{ }
	$ScaleCodeToLocation = @{ }
	$ScaleCodeToIPNetwork = @{ }
	$ScaleCodeToIPDevice = @{ }
	$ScaleCodeToFullIP = @{ }
	$ScaleCodeToPortRaw = @{ }
	$ScaleCodeToTcpPort = @{ }
	$ScaleCodeToBaseMode = @{ }
	$ScaleCodeToType = @{ }
	$ScaleCodeToBaud = @{ }
	$ScaleCodeToSub = @{ }
	$ScaleCodeToPlumDir = @{ }

	# Materialize TER first with true 80X keys
	foreach ($terCode in @($FinalTerBy80X.Keys | Sort-Object { [int]$_ }))
	{
		$terEntry = $FinalTerBy80X[$terCode]
		if (-not $terEntry) { continue }

		$finalKey = [string]$terCode

		$finalObj = [PSCustomObject]@{
			Code		  = $finalKey
			ScaleCode	  = $finalKey
			BaseScaleCode = $terEntry.BaseScaleCode
			Label		  = $terEntry.Label
			ScaleName	  = $terEntry.ScaleName
			Path		  = $terEntry.Path
			HostPath	  = $terEntry.HostPath
			ScaleBrand    = $terEntry.ScaleBrand
			Scalevendor   = if ($terEntry.Scalevendor) { $terEntry.Scalevendor } else { $terEntry.ScaleBrand }
			ScaleModel    = $terEntry.ScaleModel
			ScaleLocation = $terEntry.ScaleLocation
			IPNetwork	  = $terEntry.IPNetwork
			IPDevice	  = $terEntry.IPDevice
			FullIP	      = $terEntry.FullIP
			PortRaw	      = $terEntry.PortRaw
			TcpPort	      = $terEntry.TcpPort
			Section	      = $terEntry.Section
			BaseMode	  = $terEntry.BaseMode
			Type		  = $terEntry.Type
			Baud		  = $terEntry.Baud
			Sub		      = $terEntry.Sub
			PlumDir	      = $terEntry.PlumDir
			Active	      = $terEntry.Active
			Source	      = 'TER'
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
		if ($finalObj.Section) { $ScaleCodeToSection[$finalKey] = $finalObj.Section }
		if ($finalObj.ScaleBrand) { $ScaleCodeToBrand[$finalKey] = $finalObj.ScaleBrand }
		if ($finalObj.ScaleModel) { $ScaleCodeToModel[$finalKey] = $finalObj.ScaleModel }
		if ($finalObj.ScaleLocation) { $ScaleCodeToLocation[$finalKey] = $finalObj.ScaleLocation }
		if ($finalObj.IPNetwork) { $ScaleCodeToIPNetwork[$finalKey] = $finalObj.IPNetwork }
		if ($finalObj.IPDevice) { $ScaleCodeToIPDevice[$finalKey] = $finalObj.IPDevice }
		if ($finalObj.FullIP) { $ScaleCodeToFullIP[$finalKey] = $finalObj.FullIP }
		if ($finalObj.PortRaw) { $ScaleCodeToPortRaw[$finalKey] = $finalObj.PortRaw }
		if ($finalObj.TcpPort) { $ScaleCodeToTcpPort[$finalKey] = $finalObj.TcpPort }
		if ($finalObj.BaseMode) { $ScaleCodeToBaseMode[$finalKey] = $finalObj.BaseMode }
		if ($finalObj.Type) { $ScaleCodeToType[$finalKey] = $finalObj.Type }
		if ($finalObj.Baud) { $ScaleCodeToBaud[$finalKey] = $finalObj.Baud }
		if ($finalObj.Sub) { $ScaleCodeToSub[$finalKey] = $finalObj.Sub }
		if ($finalObj.PlumDir) { $ScaleCodeToPlumDir[$finalKey] = $finalObj.PlumDir }
	}

	# Materialize TBS next with original/raw ScaleCode keys
	$TbsKeyCounts = @{ }
	foreach ($tbsEntry in @($TbsScaleRows))
	{
		if (-not $tbsEntry) { continue }

		$rawKey = [string]$tbsEntry.RawScaleCode
		if ([string]::IsNullOrWhiteSpace($rawKey)) { continue }

		$finalKey = $rawKey

		# If raw/original TBS key collides with an existing key, keep original and add suffix only when needed
		if ($ScaleCodeToIPInfo.ContainsKey($finalKey))
		{
			if (-not $TbsKeyCounts.ContainsKey($rawKey)) { $TbsKeyCounts[$rawKey] = 1 }

			do
			{
				$TbsKeyCounts[$rawKey]++
				$finalKey = "{0}_TBS{1}" -f $rawKey, $TbsKeyCounts[$rawKey]
			}
			while ($ScaleCodeToIPInfo.ContainsKey($finalKey))
		}
		else
		{
			if (-not $TbsKeyCounts.ContainsKey($rawKey)) { $TbsKeyCounts[$rawKey] = 1 }
		}

		$finalObj = [PSCustomObject]@{
			Code		  = $finalKey
			ScaleCode	  = $finalKey
			BaseScaleCode = $tbsEntry.BaseScaleCode
			Label		  = $tbsEntry.Label
			ScaleName	  = $tbsEntry.ScaleName
			Path		  = $tbsEntry.Path
			HostPath	  = $tbsEntry.HostPath
			ScaleBrand    = $tbsEntry.ScaleBrand
			Scalevendor   = if ($tbsEntry.Scalevendor) { $tbsEntry.Scalevendor } else { $tbsEntry.ScaleBrand }
			ScaleModel    = $tbsEntry.ScaleModel
			ScaleLocation = $tbsEntry.ScaleLocation
			IPNetwork	  = $tbsEntry.IPNetwork
			IPDevice	  = $tbsEntry.IPDevice
			FullIP	      = $tbsEntry.FullIP
			PortRaw	      = $tbsEntry.PortRaw
			TcpPort	      = $tbsEntry.TcpPort
			Section	      = $tbsEntry.Section
			BaseMode	  = $tbsEntry.BaseMode
			Type		  = $tbsEntry.Type
			Baud		  = $tbsEntry.Baud
			Sub		      = $tbsEntry.Sub
			PlumDir	      = $tbsEntry.PlumDir
			Active	      = $tbsEntry.Active
			Source	      = 'TBS'
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
		if ($finalObj.Section) { $ScaleCodeToSection[$finalKey] = $finalObj.Section }
		if ($finalObj.ScaleBrand) { $ScaleCodeToBrand[$finalKey] = $finalObj.ScaleBrand }
		if ($finalObj.ScaleModel) { $ScaleCodeToModel[$finalKey] = $finalObj.ScaleModel }
		if ($finalObj.ScaleLocation) { $ScaleCodeToLocation[$finalKey] = $finalObj.ScaleLocation }
		if ($finalObj.IPNetwork) { $ScaleCodeToIPNetwork[$finalKey] = $finalObj.IPNetwork }
		if ($finalObj.IPDevice) { $ScaleCodeToIPDevice[$finalKey] = $finalObj.IPDevice }
		if ($finalObj.FullIP) { $ScaleCodeToFullIP[$finalKey] = $finalObj.FullIP }
		if ($finalObj.PortRaw) { $ScaleCodeToPortRaw[$finalKey] = $finalObj.PortRaw }
		if ($finalObj.TcpPort) { $ScaleCodeToTcpPort[$finalKey] = $finalObj.TcpPort }
		if ($finalObj.BaseMode) { $ScaleCodeToBaseMode[$finalKey] = $finalObj.BaseMode }
		if ($finalObj.Type) { $ScaleCodeToType[$finalKey] = $finalObj.Type }
		if ($finalObj.Baud) { $ScaleCodeToBaud[$finalKey] = $finalObj.Baud }
		if ($finalObj.Sub) { $ScaleCodeToSub[$finalKey] = $finalObj.Sub }
		if ($finalObj.PlumDir) { $ScaleCodeToPlumDir[$finalKey] = $finalObj.PlumDir }
	}

	$NumberOfLanes = $LaneMachineNames.Count
	$NumberOfServers = if ($ServerMachineName) { 1 }
	else { 0 }
	$NumberOfBackoffices = $BackofficeNumToMachineName.Count
	$NumberOfScales = $ScaleCodes.Count

	# ====================================================================================
	# 3. CLEAN 1:1 MAPPINGS FOR ALL NODES (NO ALIASES, INCLUDE PATHS)
	# ====================================================================================

	# ---- LANES ----
	$CleanLaneNumToMachineName = @{ }
	$CleanMachineNameToLaneNum = @{ }
	$CleanLaneNumToPath = @{ }
	$CleanMachineNameToPath = @{ }
	$CleanLaneNumToServerPath = @{ }
	$CleanMachineNameToServerPath = @{ }

	foreach ($kv in $LaneNumToMachineName.GetEnumerator())
	{
		$laneNum = $kv.Key
		$machine = $kv.Value
		$CleanLaneNumToMachineName[$laneNum] = $machine
		$CleanMachineNameToLaneNum[$machine] = $laneNum

		if ($LaneMachinePath.ContainsKey($machine))
		{
			$CleanLaneNumToPath[$laneNum] = $LaneMachinePath[$machine]
			$CleanMachineNameToPath[$machine] = $LaneMachinePath[$machine]
		}
		if ($LaneMachineToServerPath.ContainsKey($machine))
		{
			$CleanLaneNumToServerPath[$laneNum] = $LaneMachineToServerPath[$machine]
			$CleanMachineNameToServerPath[$machine] = $LaneMachineToServerPath[$machine]
		}
	}
	$LaneNumToMachineName = $CleanLaneNumToMachineName
	$MachineNameToLaneNum = $CleanMachineNameToLaneNum
	$LaneNumToPath = $CleanLaneNumToPath
	$MachineNameToPath = $CleanMachineNameToPath
	$LaneNumToServerPath = $CleanLaneNumToServerPath
	$MachineNameToServerPath = $CleanMachineNameToServerPath

	# ---- SCALES ----
	$CleanScaleCodeToIPInfo = @{ }
	$CleanScaleNameToCode = @{ }
	$CleanScaleCodeToPath = @{ }
	$CleanScaleNameToPath = @{ }

	$CleanScaleCodeToName = @{ }
	$CleanScaleCodeToSection = @{ }
	$CleanScaleCodeToBrand = @{ }
	$CleanScaleCodeToModel = @{ }
	$CleanScaleCodeToLocation = @{ }
	$CleanScaleCodeToIPNetwork = @{ }
	$CleanScaleCodeToIPDevice = @{ }
	$CleanScaleCodeToFullIP = @{ }
	$CleanScaleCodeToPortRaw = @{ }
	$CleanScaleCodeToTcpPort = @{ }
	$CleanScaleCodeToBaseMode = @{ }
	$CleanScaleCodeToType = @{ }
	$CleanScaleCodeToBaud = @{ }
	$CleanScaleCodeToSub = @{ }
	$CleanScaleCodeToPlumDir = @{ }

	foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
	{
		$scaleKey = $kv.Key
		$scale = $kv.Value

		$CleanScaleCodeToIPInfo[$scaleKey] = $scale

		if ($scale.ScaleName)
		{
			$CleanScaleNameToCode[$scale.ScaleName] = $scaleKey
			$CleanScaleNameToPath[$scale.ScaleName] = $scale.Path
		}

		$CleanScaleCodeToPath[$scaleKey] = $scale.Path

		if ($scale.ScaleName) { $CleanScaleCodeToName[$scaleKey] = $scale.ScaleName }
		if ($scale.Section) { $CleanScaleCodeToSection[$scaleKey] = $scale.Section }
		if ($scale.ScaleBrand) { $CleanScaleCodeToBrand[$scaleKey] = $scale.ScaleBrand }
		if ($scale.ScaleModel) { $CleanScaleCodeToModel[$scaleKey] = $scale.ScaleModel }
		if ($scale.ScaleLocation) { $CleanScaleCodeToLocation[$scaleKey] = $scale.ScaleLocation }
		if ($scale.IPNetwork) { $CleanScaleCodeToIPNetwork[$scaleKey] = $scale.IPNetwork }
		if ($scale.IPDevice) { $CleanScaleCodeToIPDevice[$scaleKey] = $scale.IPDevice }
		if ($scale.FullIP) { $CleanScaleCodeToFullIP[$scaleKey] = $scale.FullIP }
		if ($scale.PortRaw) { $CleanScaleCodeToPortRaw[$scaleKey] = $scale.PortRaw }
		if ($scale.TcpPort) { $CleanScaleCodeToTcpPort[$scaleKey] = $scale.TcpPort }
		if ($scale.BaseMode) { $CleanScaleCodeToBaseMode[$scaleKey] = $scale.BaseMode }
		if ($scale.Type) { $CleanScaleCodeToType[$scaleKey] = $scale.Type }
		if ($scale.Baud) { $CleanScaleCodeToBaud[$scaleKey] = $scale.Baud }
		if ($scale.Sub) { $CleanScaleCodeToSub[$scaleKey] = $scale.Sub }
		if ($scale.PlumDir) { $CleanScaleCodeToPlumDir[$scaleKey] = $scale.PlumDir }
	}

	$ScaleCodeToIPInfo = $CleanScaleCodeToIPInfo
	$ScaleNameToCode = $CleanScaleNameToCode
	$ScaleCodeToPath = $CleanScaleCodeToPath
	$ScaleNameToPath = $CleanScaleNameToPath

	$ScaleCodeToName = $CleanScaleCodeToName
	$ScaleCodeToSection = $CleanScaleCodeToSection
	$ScaleCodeToBrand = $CleanScaleCodeToBrand
	$ScaleCodeToModel = $CleanScaleCodeToModel
	$ScaleCodeToLocation = $CleanScaleCodeToLocation
	$ScaleCodeToIPNetwork = $CleanScaleCodeToIPNetwork
	$ScaleCodeToIPDevice = $CleanScaleCodeToIPDevice
	$ScaleCodeToFullIP = $CleanScaleCodeToFullIP
	$ScaleCodeToPortRaw = $CleanScaleCodeToPortRaw
	$ScaleCodeToTcpPort = $CleanScaleCodeToTcpPort
	$ScaleCodeToBaseMode = $CleanScaleCodeToBaseMode
	$ScaleCodeToType = $CleanScaleCodeToType
	$ScaleCodeToBaud = $CleanScaleCodeToBaud
	$ScaleCodeToSub = $CleanScaleCodeToSub
	$ScaleCodeToPlumDir = $CleanScaleCodeToPlumDir

	# ---- BACKOFFICES ----
	$CleanBackofficeNumToMachineName = @{ }
	$CleanMachineNameToBackofficeNum = @{ }
	$CleanBackofficeNumToPath = @{ }
	$CleanMachineNameToBOPath = @{ }

	foreach ($kv in $BackofficeNumToMachineName.GetEnumerator())
	{
		$boNum = $kv.Key
		$machine = $kv.Value
		$CleanBackofficeNumToMachineName[$boNum] = $machine
		$CleanMachineNameToBackofficeNum[$machine] = $boNum

		if ($BackofficeNumToPath.ContainsKey($boNum))
		{
			$CleanBackofficeNumToPath[$boNum] = $BackofficeNumToPath[$boNum]
			$CleanMachineNameToBOPath[$machine] = $BackofficeNumToPath[$boNum]
		}
	}
	$BackofficeNumToMachineName = $CleanBackofficeNumToMachineName
	$MachineNameToBackofficeNum = $CleanMachineNameToBackofficeNum
	$BackofficeNumToPath = $CleanBackofficeNumToPath
	$MachineNameToBackofficePath = $CleanMachineNameToBOPath

	# ====================================================================================
	# 5. BUILD WINDOWS SCALES ONLY
	# ====================================================================================
	$WindowsScales = @{ }
	foreach ($code in $ScaleCodeToIPInfo.Keys)
	{
		$scale = $ScaleCodeToIPInfo[$code]
		if (-not $scale) { continue }

		$isBizerba = $false

		if ($scale.ScaleBrand -and ([string]$scale.ScaleBrand -match '(?i)bizerba')) { $isBizerba = $true }
		elseif ($scale.Scalevendor -and ([string]$scale.Scalevendor -match '(?i)bizerba')) { $isBizerba = $true }
		elseif ($scale.ScaleModel -and ([string]$scale.ScaleModel -match '(?i)bizerba')) { $isBizerba = $true }
		elseif ($scale.ScaleName -and ([string]$scale.ScaleName -match '(?i)bizerba')) { $isBizerba = $true }
		elseif ($scale.Label -and ([string]$scale.Label -match '(?i)bizerba')) { $isBizerba = $true }

		if ($isBizerba)
		{
			$WindowsScales[$code] = $scale
		}
	}

	# ====================================================================================
	# 6. BUILD RETURN OBJECT & STORE TO GLOBAL FUNCTIONRESULTS
	# ====================================================================================
	$Nodes = [PSCustomObject]@{
		NumberOfLanes			   = $NumberOfLanes
		NumberOfServers		       = $NumberOfServers
		NumberOfBackoffices	       = $NumberOfBackoffices
		NumberOfScales			   = $NumberOfScales
		LaneMachineNames		   = $LaneMachineNames
		LaneNumToMachineName	   = $LaneNumToMachineName
		LaneMachineLabels		   = $LaneMachineLabels
		LaneMachinePath		       = $LaneMachinePath
		LaneMachineToServerPath    = $LaneMachineToServerPath
		ScaleCodes				   = $ScaleCodes
		ScaleLabels			       = $ScaleLabels
		ScaleExePaths			   = $ScaleExePaths
		ScaleCodeToIPInfo		   = $ScaleCodeToIPInfo
		ScaleNameToCode		       = $ScaleNameToCode
		ScaleCodeToPath		       = $ScaleCodeToPath
		ScaleNameToPath		       = $ScaleNameToPath
		ScaleCodeToName		       = $ScaleCodeToName
		ScaleCodeToSection		   = $ScaleCodeToSection
		ScaleCodeToBrand		   = $ScaleCodeToBrand
		ScaleCodeToModel		   = $ScaleCodeToModel
		ScaleCodeToLocation	       = $ScaleCodeToLocation
		ScaleCodeToIPNetwork	   = $ScaleCodeToIPNetwork
		ScaleCodeToIPDevice	       = $ScaleCodeToIPDevice
		ScaleCodeToFullIP		   = $ScaleCodeToFullIP
		ScaleCodeToPortRaw		   = $ScaleCodeToPortRaw
		ScaleCodeToTcpPort		   = $ScaleCodeToTcpPort
		ScaleCodeToBaseMode	       = $ScaleCodeToBaseMode
		ScaleCodeToType		       = $ScaleCodeToType
		ScaleCodeToBaud		       = $ScaleCodeToBaud
		ScaleCodeToSub			   = $ScaleCodeToSub
		ScaleCodeToPlumDir		   = $ScaleCodeToPlumDir
		BackofficeNumToMachineName = $BackofficeNumToMachineName
		BackofficeNumToLabel	   = $BackofficeNumToLabel
		BackofficeNumToPath	       = $BackofficeNumToPath
		ServerMachineName		   = $ServerMachineName
		ServerLabel			       = $ServerLabel
		ServerPath				   = $ServerPath
		WindowsScales			   = $WindowsScales
	}

	$script:FunctionResults['NumberOfLanes'] = $NumberOfLanes
	$script:FunctionResults['NumberOfServers'] = $NumberOfServers
	$script:FunctionResults['NumberOfBackoffices'] = $NumberOfBackoffices
	$script:FunctionResults['NumberOfScales'] = $NumberOfScales
	$script:FunctionResults['LaneMachineNames'] = $LaneMachineNames
	$script:FunctionResults['LaneNumToMachineName'] = $LaneNumToMachineName
	$script:FunctionResults['MachineNameToLaneNum'] = $MachineNameToLaneNum
	$script:FunctionResults['LaneMachineLabels'] = $LaneMachineLabels
	$script:FunctionResults['LaneMachinePath'] = $LaneMachinePath
	$script:FunctionResults['LaneMachineToServerPath'] = $LaneMachineToServerPath
	$script:FunctionResults['ScaleCodes'] = $ScaleCodes
	$script:FunctionResults['ScaleLabels'] = $ScaleLabels
	$script:FunctionResults['ScaleExePaths'] = $ScaleExePaths
	$script:FunctionResults['ScaleCodeToIPInfo'] = $ScaleCodeToIPInfo
	$script:FunctionResults['ScaleNameToCode'] = $ScaleNameToCode
	$script:FunctionResults['ScaleCodeToPath'] = $ScaleCodeToPath
	$script:FunctionResults['ScaleNameToPath'] = $ScaleNameToPath
	$script:FunctionResults['ScaleCodeToName'] = $ScaleCodeToName
	$script:FunctionResults['ScaleCodeToSection'] = $ScaleCodeToSection
	$script:FunctionResults['ScaleCodeToBrand'] = $ScaleCodeToBrand
	$script:FunctionResults['ScaleCodeToModel'] = $ScaleCodeToModel
	$script:FunctionResults['ScaleCodeToLocation'] = $ScaleCodeToLocation
	$script:FunctionResults['ScaleCodeToIPNetwork'] = $ScaleCodeToIPNetwork
	$script:FunctionResults['ScaleCodeToIPDevice'] = $ScaleCodeToIPDevice
	$script:FunctionResults['ScaleCodeToFullIP'] = $ScaleCodeToFullIP
	$script:FunctionResults['ScaleCodeToPortRaw'] = $ScaleCodeToPortRaw
	$script:FunctionResults['ScaleCodeToTcpPort'] = $ScaleCodeToTcpPort
	$script:FunctionResults['ScaleCodeToBaseMode'] = $ScaleCodeToBaseMode
	$script:FunctionResults['ScaleCodeToType'] = $ScaleCodeToType
	$script:FunctionResults['ScaleCodeToBaud'] = $ScaleCodeToBaud
	$script:FunctionResults['ScaleCodeToSub'] = $ScaleCodeToSub
	$script:FunctionResults['ScaleCodeToPlumDir'] = $ScaleCodeToPlumDir
	$script:FunctionResults['BackofficeNumToMachineName'] = $BackofficeNumToMachineName
	$script:FunctionResults['MachineNameToBackofficeNum'] = $MachineNameToBackofficeNum
	$script:FunctionResults['BackofficeNumToLabel'] = $BackofficeNumToLabel
	$script:FunctionResults['BackofficeNumToPath'] = $BackofficeNumToPath
	$script:FunctionResults['MachineNameToBackofficePath'] = $MachineNameToBackofficePath
	$script:FunctionResults['ServerMachineName'] = $ServerMachineName
	$script:FunctionResults['ServerLabel'] = $ServerLabel
	$script:FunctionResults['ServerPath'] = $ServerPath
	$script:FunctionResults['WindowsScales'] = $WindowsScales
	$script:FunctionResults['Nodes'] = $Nodes
	$script:FunctionResults['XchNodesIniPathsUsed'] = @($XchNodesIniPathsUsed | Select-Object -Unique)

	# ====================================================================================
	# 7. OPTIONAL: UPDATE GUI LABELS IF PRESENT
	# ====================================================================================
	if ($NodesHost -ne $null) { $NodesHost.Text = "Number of Servers: $NumberOfServers" }
	if ($NodesBackoffices -ne $null) { $NodesBackoffices.Text = "Number of Backoffices: $NumberOfBackoffices" }
	if ($NodesStore -ne $null) { $NodesStore.Text = "Number of Lanes: $NumberOfLanes" }
	if ($scalesLabel -ne $null) { $scalesLabel.Text = "Number of Scales: $NumberOfScales" }
	if ($form -ne $null) { $form.Refresh() }

	# ====================================================================================
	# 8. CACHE WRITE
	# ====================================================================================
	if ($CachePath)
	{
		try
		{
			$sortedLaneNums = @($LaneNumToMachineName.Keys) | Where-Object { $_ -match '^\d{3}$' } | Sort-Object { [int]$_ }

			$outLaneNumToMachineName = [ordered]@{ }
			foreach ($k in $sortedLaneNums) { $outLaneNumToMachineName[$k] = $LaneNumToMachineName[$k] }

			$outMachineNameToLaneNum = [ordered]@{ }
			foreach ($kv in ($MachineNameToLaneNum.GetEnumerator() | Sort-Object { [int]($_.Value) }, { $_.Key }))
			{
				$outMachineNameToLaneNum[$kv.Key] = $kv.Value
			}

			$outLaneMachineLabels = [ordered]@{ }
			$outLaneMachinePath = [ordered]@{ }
			$outLaneMachineToServerPath = [ordered]@{ }
			foreach ($ln in $sortedLaneNums)
			{
				$m = $LaneNumToMachineName[$ln]
				if ($m)
				{
					if ($LaneMachineLabels.ContainsKey($m)) { $outLaneMachineLabels[$m] = $LaneMachineLabels[$m] }
					if ($LaneMachinePath.ContainsKey($m)) { $outLaneMachinePath[$m] = $LaneMachinePath[$m] }
					if ($LaneMachineToServerPath.ContainsKey($m)) { $outLaneMachineToServerPath[$m] = $LaneMachineToServerPath[$m] }
				}
			}

			$outScaleLabels = [ordered]@{ }
			foreach ($k in (@($ScaleLabels.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleLabels[$k] = $ScaleLabels[$k] }

			$outScaleExePaths = [ordered]@{ }
			foreach ($k in (@($ScaleExePaths.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleExePaths[$k] = $ScaleExePaths[$k] }

			$outScaleCodeToIPInfo = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToIPInfo.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToIPInfo[$k] = $ScaleCodeToIPInfo[$k] }

			$outScaleNameToCode = [ordered]@{ }
			foreach ($kv in ($ScaleNameToCode.GetEnumerator() | Sort-Object { $_.Value }, { $_.Key })) { $outScaleNameToCode[$kv.Key] = $kv.Value }

			$outScaleCodeToPath = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToPath.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToPath[$k] = $ScaleCodeToPath[$k] }

			$outScaleNameToPath = [ordered]@{ }
			foreach ($kv in ($ScaleNameToPath.GetEnumerator() | Sort-Object { $_.Key })) { $outScaleNameToPath[$kv.Key] = $kv.Value }

			$outScaleCodeToName = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToName.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToName[$k] = $ScaleCodeToName[$k] }

			$outScaleCodeToSection = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToSection.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToSection[$k] = $ScaleCodeToSection[$k] }

			$outScaleCodeToBrand = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToBrand.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToBrand[$k] = $ScaleCodeToBrand[$k] }

			$outScaleCodeToModel = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToModel.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToModel[$k] = $ScaleCodeToModel[$k] }

			$outScaleCodeToLocation = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToLocation.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToLocation[$k] = $ScaleCodeToLocation[$k] }

			$outScaleCodeToIPNetwork = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToIPNetwork.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToIPNetwork[$k] = $ScaleCodeToIPNetwork[$k] }

			$outScaleCodeToIPDevice = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToIPDevice.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToIPDevice[$k] = $ScaleCodeToIPDevice[$k] }

			$outScaleCodeToFullIP = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToFullIP.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToFullIP[$k] = $ScaleCodeToFullIP[$k] }

			$outScaleCodeToPortRaw = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToPortRaw.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToPortRaw[$k] = $ScaleCodeToPortRaw[$k] }

			$outScaleCodeToTcpPort = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToTcpPort.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToTcpPort[$k] = $ScaleCodeToTcpPort[$k] }

			$outScaleCodeToBaseMode = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToBaseMode.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToBaseMode[$k] = $ScaleCodeToBaseMode[$k] }

			$outScaleCodeToType = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToType.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToType[$k] = $ScaleCodeToType[$k] }

			$outScaleCodeToBaud = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToBaud.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToBaud[$k] = $ScaleCodeToBaud[$k] }

			$outScaleCodeToSub = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToSub.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToSub[$k] = $ScaleCodeToSub[$k] }

			$outScaleCodeToPlumDir = [ordered]@{ }
			foreach ($k in (@($ScaleCodeToPlumDir.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outScaleCodeToPlumDir[$k] = $ScaleCodeToPlumDir[$k] }

			$outBackofficeNumToMachineName = [ordered]@{ }
			foreach ($k in (@($BackofficeNumToMachineName.Keys) | Sort-Object { [int]$_ })) { $outBackofficeNumToMachineName[$k] = $BackofficeNumToMachineName[$k] }

			$outMachineNameToBackofficeNum = [ordered]@{ }
			foreach ($kv in ($MachineNameToBackofficeNum.GetEnumerator() | Sort-Object { [int]($_.Value) }, { $_.Key })) { $outMachineNameToBackofficeNum[$kv.Key] = $kv.Value }

			$outBackofficeNumToLabel = [ordered]@{ }
			foreach ($k in (@($BackofficeNumToLabel.Keys) | Sort-Object { [int]$_ })) { $outBackofficeNumToLabel[$k] = $BackofficeNumToLabel[$k] }

			$outBackofficeNumToPath = [ordered]@{ }
			foreach ($k in (@($BackofficeNumToPath.Keys) | Sort-Object { [int]$_ })) { $outBackofficeNumToPath[$k] = $BackofficeNumToPath[$k] }

			$outMachineNameToBackofficePath = [ordered]@{ }
			foreach ($kv in ($MachineNameToBackofficePath.GetEnumerator() | Sort-Object { $_.Key })) { $outMachineNameToBackofficePath[$kv.Key] = $kv.Value }

			$outWindowsScales = [ordered]@{ }
			foreach ($k in (@($WindowsScales.Keys) | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ })) { $outWindowsScales[$k] = $WindowsScales[$k] }

			$outXchNodesIniPathsUsed = @($XchNodesIniPathsUsed | Select-Object -Unique)

			$outLaneMachineNames = @($LaneMachineNames)
			try
			{
				if ($outLaneMachineNames.Count -gt 0 -and $MachineNameToLaneNum.Count -gt 0)
				{
					$outLaneMachineNames = $outLaneMachineNames | Sort-Object {
						try { [int]($MachineNameToLaneNum[$_] -replace '[^\d]', '0') }
						catch { 999999 }
					}, { $_ }
				}
			}
			catch { }

			$outScaleCodes = @($ScaleCodes)
			try
			{
				$outScaleCodes = $outScaleCodes | Sort-Object { [int](($_ -replace '[^\d]', '0')) }, { $_ }
			}
			catch { }

			$cacheOut = [pscustomobject]@{
				LastUpdated = (Get-Date).ToString('o')
				CacheSchemaVersion = $CacheSchemaVersion
				StoreNumber = $StoreNumberNorm
				Results	    = [pscustomobject]@{
					NumberOfLanes			    = $NumberOfLanes
					NumberOfServers			    = $NumberOfServers
					NumberOfBackoffices		    = $NumberOfBackoffices
					NumberOfScales			    = $NumberOfScales
					LaneMachineNames		    = $outLaneMachineNames
					LaneNumToMachineName	    = $outLaneNumToMachineName
					MachineNameToLaneNum	    = $outMachineNameToLaneNum
					LaneMachineLabels		    = $outLaneMachineLabels
					LaneMachinePath			    = $outLaneMachinePath
					LaneMachineToServerPath	    = $outLaneMachineToServerPath
					ScaleCodes				    = $outScaleCodes
					ScaleLabels				    = $outScaleLabels
					ScaleExePaths			    = $outScaleExePaths
					ScaleCodeToIPInfo		    = $outScaleCodeToIPInfo
					ScaleNameToCode			    = $outScaleNameToCode
					ScaleCodeToPath			    = $outScaleCodeToPath
					ScaleNameToPath			    = $outScaleNameToPath
					ScaleCodeToName			    = $outScaleCodeToName
					ScaleCodeToSection		    = $outScaleCodeToSection
					ScaleCodeToBrand		    = $outScaleCodeToBrand
					ScaleCodeToModel		    = $outScaleCodeToModel
					ScaleCodeToLocation		    = $outScaleCodeToLocation
					ScaleCodeToIPNetwork	    = $outScaleCodeToIPNetwork
					ScaleCodeToIPDevice		    = $outScaleCodeToIPDevice
					ScaleCodeToFullIP		    = $outScaleCodeToFullIP
					ScaleCodeToPortRaw		    = $outScaleCodeToPortRaw
					ScaleCodeToTcpPort		    = $outScaleCodeToTcpPort
					ScaleCodeToBaseMode		    = $outScaleCodeToBaseMode
					ScaleCodeToType			    = $outScaleCodeToType
					ScaleCodeToBaud			    = $outScaleCodeToBaud
					ScaleCodeToSub			    = $outScaleCodeToSub
					ScaleCodeToPlumDir		    = $outScaleCodeToPlumDir
					BackofficeNumToMachineName  = $outBackofficeNumToMachineName
					MachineNameToBackofficeNum  = $outMachineNameToBackofficeNum
					BackofficeNumToLabel	    = $outBackofficeNumToLabel
					BackofficeNumToPath		    = $outBackofficeNumToPath
					MachineNameToBackofficePath = $outMachineNameToBackofficePath
					ServerMachineName		    = $ServerMachineName
					ServerLabel				    = $ServerLabel
					ServerPath				    = $ServerPath
					WindowsScales			    = $outWindowsScales
					XchNodesIniPathsUsed	    = $outXchNodesIniPathsUsed
				}
			}

			($cacheOut | ConvertTo-Json -Depth 16) | Set-Content -LiteralPath $CachePath -Encoding UTF8 -Force
		}
		catch { }
	}

	# ====================================================================================
	# 9. RETURN THE NODES OBJECT FOR SCRIPT CALLERS
	# ====================================================================================
	return $Nodes
}

# ===================================================================================================
#                              FUNCTION: Get_Table_Aliases
# ---------------------------------------------------------------------------------------------------
# Description:
#   Returns a hashtable mapping SQL table aliases to table names. Uses a hardcoded mapping
#   when the SMS version is in the supported range, otherwise scans the Office Load
#   *_Load.sql files dynamically. Cached in memory for the launch and in a small file.
#
# Parameters:
#   - RefreshCache  [switch]: ignore in-memory/file cache and rebuild
#
# Variables:
#   - $script:FunctionResults: stores the alias/table hashes (Get_Table_Aliases bucket)
#   - $LoadPath / $OfficePath: location of the Load SQL files (dynamic mode)
#
# Returns:
#   Hashtable alias -> table name.
#
# Writes:
#   - Setup_Files\Get_Table_Aliases.txt (JSON cache: AliasHash + TableHash only)
# ===================================================================================================

function Get_Table_Aliases
{
	param (
		[switch]$RefreshCache
	)
	
	# ------------------------------------------------------------------------------------------------
	# Per-function cache
	# ------------------------------------------------------------------------------------------------
	$CachePath = Join-Path $script:SetupFiles 'Get_Table_Aliases.txt'
	
	# Reuse in-memory alias map first
	if (-not $RefreshCache)
	{
		try
		{
			if ($script:FunctionResults.ContainsKey('Get_Table_Aliases') -and $script:FunctionResults['Get_Table_Aliases'])
			{
				$aliasData = $script:FunctionResults['Get_Table_Aliases']
				if ($aliasData.AliasHash -is [hashtable]) { return $aliasData.AliasHash }
				if ($aliasData.AliasHash -and $aliasData.AliasHash.PSObject -and $aliasData.AliasHash.PSObject.Properties)
				{
					$aliasHash = @{ }
					foreach ($p in $aliasData.AliasHash.PSObject.Properties) { $aliasHash[$p.Name] = $p.Value }
					$tableHash = @{ }
					if ($aliasData.TableHash -and $aliasData.TableHash.PSObject -and $aliasData.TableHash.PSObject.Properties)
					{
						foreach ($p in $aliasData.TableHash.PSObject.Properties) { $tableHash[$p.Name] = $p.Value }
					}
					$script:FunctionResults['Get_Table_Aliases'] = @{
						Aliases   = @()
						AliasHash = $aliasHash
						TableHash = $tableHash
					}
					return $aliasHash
				}
			}
		}
		catch { }
	}
	
	# ------------------------------------------------------------------------------------------------
	# Version detection (relies on Get_Store_And_Database_Info setting SMSVersionFull)
	# ------------------------------------------------------------------------------------------------
	$MinSupportedSMSVersion = "3.3.0.0"
	$MaxSupportedSMSVersion = "3.6.1.8"
	
	$SMSVersionRaw = $script:FunctionResults['SMSVersionFull']
	$SMSVersion = $SMSVersionRaw
	if ($SMSVersion -match '([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)')
	{
		$SMSVersion = $Matches[1]
	}
	else
	{
		$SMSVersion = "0.0.0.0"
	}
	
	# Version range check
	$versionInRange = $false
	try
	{
		$vTest = [version]$SMSVersion
		$vMin = [version]$MinSupportedSMSVersion
		$vMax = [version]$MaxSupportedSMSVersion
		if ($vTest -ge $vMin -and $vTest -le $vMax) { $versionInRange = $true }
	}
	catch { $versionInRange = $false }
	
	# ------------------------------------------------------------------------------------------------
	# Ensure LoadPath exists (dynamic mode only)
	# ------------------------------------------------------------------------------------------------
	if (-not $LoadPath -or [string]::IsNullOrWhiteSpace($LoadPath))
	{
		try
		{
			if ($OfficePath -and (Test-Path $OfficePath)) { $LoadPath = Join-Path $OfficePath 'Load' }
		}
		catch { }
	}
	
	# ------------------------------------------------------------------------------------------------
	# ------------------------------------------------------------------------------------------------
	# Try load cache file (small)
	# ------------------------------------------------------------------------------------------------
	if ((-not $RefreshCache) -and $CachePath -and (Test-Path $CachePath))
	{
		try
		{
			$raw = Get-Content -LiteralPath $CachePath -Raw -ErrorAction Stop
			if (-not [string]::IsNullOrWhiteSpace($raw))
			{
				$cacheObj = $raw | ConvertFrom-Json -ErrorAction Stop
				if ($cacheObj -and $cacheObj.Results)
				{
					# Rebuild AliasHash hashtable (so .ContainsKey() works)
					$AliasHash = @{ }
					if ($cacheObj.Results.AliasHash -and $cacheObj.Results.AliasHash.PSObject -and $cacheObj.Results.AliasHash.PSObject.Properties)
					{
						foreach ($p in $cacheObj.Results.AliasHash.PSObject.Properties) { $AliasHash[$p.Name] = $p.Value }
					}
					
					# Rebuild TableHash hashtable
					$TableHash = @{ }
					if ($cacheObj.Results.TableHash -and $cacheObj.Results.TableHash.PSObject -and $cacheObj.Results.TableHash.PSObject.Properties)
					{
						foreach ($p in $cacheObj.Results.TableHash.PSObject.Properties) { $TableHash[$p.Name] = $p.Value }
					}
					
					$script:FunctionResults['Get_Table_Aliases'] = @{
						Aliases   = @()
						AliasHash = $AliasHash
						TableHash = $TableHash
					}
					
					return $AliasHash
				}
			}
		}
		catch
		{
			# Cache errors fall through to rebuild
		}
	}
	
	# ------------------------------------------------------------------------------------------------
	# Compute signature for *_Load.sql files (used only when dynamic rebuild is needed)
	# ------------------------------------------------------------------------------------------------
	$loadSignature = $null
	try
	{
		if ($LoadPath -and (Test-Path $LoadPath))
		{
			$files = Get-ChildItem -LiteralPath $LoadPath -Recurse -Filter '*_Load.sql' -File -ErrorAction SilentlyContinue
			$fileCount = 0
			$newestUtc = $null
			$totalBytes = 0
			
			if ($files)
			{
				$fileCount = $files.Count
				$newestUtc = ($files | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1).LastWriteTimeUtc
				$totalBytes = [int64](($files | Measure-Object Length -Sum).Sum)
			}
			
			$loadSignature = [pscustomobject]@{
				Root = [string]$LoadPath
				FileCount = [int]$fileCount
				NewestUtc = if ($newestUtc) { $newestUtc.ToString('o') } else { "" }
				TotalBytes = [int64]$totalBytes
			}
		}
	}
	catch { $loadSignature = $null }
	
	# ------------------------------------------------------------------------------------------------
	# Hardcoded alias [?][?][?] table mapping
	# ------------------------------------------------------------------------------------------------
	$AliasToTable = @{
		'ALT'	   = 'ALT_TAB'
		'BIO'	   = 'BIO_TAB'
		'BTL'	   = 'BTL_TAB'
		'CAT'	   = 'CAT_TAB'
		'CFG'	   = 'CFG_TAB'
		'CLF'	   = 'CLF_TAB'
		'CLF_SDP'  = 'CLF_SDP_TAB'
		'CLG'	   = 'CLG_TAB'
		'CLK'	   = 'CLK_TAB'
		'CLL'	   = 'CLL_TAB'
		'CLR'	   = 'CLR_TAB'
		'CLT'	   = 'CLT_TAB'
		'CLT_ITM'  = 'CLT_ITM_TAB'
		'COST'	   = 'COST_TAB'
		'CPN'	   = 'CPN_TAB'
		'DELV'	   = 'DELV_TAB'
		'DEPT'	   = 'DEPT_TAB'
		'DSD'	   = 'DSD_TAB'
		'ECL'	   = 'ECL_TAB'
		'FAM'	   = 'FAM_TAB'
		'KIT'	   = 'KIT_TAB'
		'LOC'	   = 'LOC_TAB'
		'LVL'	   = 'LVL_TAB'
		'MIX'	   = 'MIX_TAB'
		'MOD'	   = 'MOD_TAB'
		'OBJ'	   = 'OBJ_TAB'
		'POS'	   = 'POS_TAB'
		'PRICE'    = 'PRICE_TAB'
		'PUB'	   = 'PUB_TAB'
		'RES'	   = 'RES_TAB'
		'ROUTE'    = 'ROUTE_TAB'
		'RPC'	   = 'RPC_TAB'
		'SCAL_ITM' = 'SCL_TAB'
		'SCAL_NUT' = 'SCL_NUT_TAB'
		'SCAL_TXT' = 'SCL_TXT_TAB'
		'SDP'	   = 'SDP_TAB'
		'STD'	   = 'STD_TAB'
		'TAR'	   = 'TAR_TAB'
		'UNT'	   = 'UNT_TAB'
		'VENDOR'   = 'VENDOR_TAB'
	}
	
	# Build reverse (table [?][?][?] alias)
	$TableToAlias = @{ }
	foreach ($k in $AliasToTable.Keys)
	{
		$TableToAlias[$AliasToTable[$k]] = $k
	}
	
	# ------------------------------------------------------------------------------------------------
	# Static mode (supported range)
	# ------------------------------------------------------------------------------------------------
	if ($versionInRange)
	{
		# Store only what matters
		$script:FunctionResults['Get_Table_Aliases'] = @{
			Aliases   = @() # intentionally empty (no @CREATE clutter)
			AliasHash = $AliasToTable
			TableHash = $TableToAlias
		}
		
		# Cache write (Static, small)
		if ($CachePath)
		{
			try
			{
				$cacheOut = [pscustomobject]@{
					LastUpdated = (Get-Date).ToString('o')
					SMSVersion  = $SMSVersion
					Mode	    = 'Static'
					LoadSignature = $null
					Results	    = [pscustomobject]@{
						AliasHash = [pscustomobject]$AliasToTable
						TableHash = [pscustomobject]$TableToAlias
					}
				}
				
				($cacheOut | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $CachePath -Encoding UTF8 -Force
			}
			catch { }
		}
		
		return $AliasToTable
	}
	
	# ------------------------------------------------------------------------------------------------
	# Dynamic scan mode (out of supported range)
	# ------------------------------------------------------------------------------------------------
	Write-Host "SMS Version $SMSVersion is outside supported range ($MinSupportedSMSVersion - $MaxSupportedSMSVersion). Scanning SQL files for table/alias map..." -ForegroundColor Yellow
	
	$BaseTables = @(
		'OBJ', 'POS', 'PRICE', 'COST', 'DSD', 'KIT', 'LOC', 'ALT', 'ECL',
		'SCL', 'SCL_TXT', 'SCL_NUT', 'DEPT', 'SDP', 'CAT', 'RPC', 'FAM',
		'CPN', 'PUB', 'BIO', 'CLK', 'LVL', 'MIX', 'BTL', 'TAR', 'UNT',
		'RES', 'ROUTE', 'VENDOR', 'DELV', 'CLT', 'CLG', 'CLF', 'CLR',
		'CLL', 'CLT_ITM', 'CLF_SDP', 'STD', 'CFG', 'MOD'
	)
	
	$escapedTables = $BaseTables | Sort-Object Length -Descending | ForEach-Object { [regex]::Escape($_) }
	$tablesPattern = $escapedTables -join '|'
	$pattern = "^\s*@CREATE\s*\(\s*['""]?(?<Table>($tablesPattern)(_[A-Z]+))?['""]?\s*,\s*['""]?(?<Alias>[\w-]+)['""]?\s*\);"
	$regex = [regex]::new($pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
	
	if (-not $LoadPath -or -not (Test-Path $LoadPath))
	{
		Write-Warning "LoadPath is not available; cannot scan *_Load.sql files. Returning empty hashtable."
		$empty = @{ }
		$script:FunctionResults['Get_Table_Aliases'] = @{ Aliases = @(); AliasHash = $empty; TableHash = @{ } }
		return $empty
	}
	
	$allSqlFiles = Get-ChildItem -LiteralPath $LoadPath -Recurse -Filter '*_Load.sql' -File -ErrorAction SilentlyContinue
	
	$AliasToTableLive = @{ }
	$TableToAliasLive = @{ }
	
	foreach ($file in $allSqlFiles)
	{
		foreach ($baseTable in $BaseTables)
		{
			if ($file.Name -ieq "$baseTable`_Load.sql")
			{
				$content = Get-Content -LiteralPath $file.FullName -ErrorAction SilentlyContinue
				if (-not $content) { break }
				
				foreach ($line in $content)
				{
					$lineClean = $line -replace '--.*', '' -replace '/\*.*?\*/', ''
					if ($lineClean -match '@CREATE')
					{
						$match = $regex.Match($lineClean)
						if ($match.Success)
						{
							$table = $match.Groups['Table'].Value
							$alias = $match.Groups['Alias'].Value
							if ($table -and $alias)
							{
								$AliasToTableLive[$alias] = $table
								$TableToAliasLive[$table] = $alias
							}
						}
					}
				}
				break
			}
		}
	}
	
	if ($AliasToTableLive.Count -eq 0)
	{
		Write-Warning "No table-alias pairs detected. Using empty hashtable."
	}
	
	$script:FunctionResults['Get_Table_Aliases'] = @{
		Aliases   = @() # intentionally empty (keep cache small)
		AliasHash = $AliasToTableLive
		TableHash = $TableToAliasLive
	}
	
	# Cache write (Dynamic, small + signature)
	if ($CachePath)
	{
		try
		{
			$cacheOut = [pscustomobject]@{
				LastUpdated = (Get-Date).ToString('o')
				SMSVersion  = $SMSVersion
				Mode	    = 'Dynamic'
				LoadSignature = $loadSignature
				Results	    = [pscustomobject]@{
					AliasHash = [pscustomobject]$AliasToTableLive
					TableHash = [pscustomobject]$TableToAliasLive
				}
			}
			
			($cacheOut | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $CachePath -Encoding UTF8 -Force
		}
		catch { }
	}
	
	return $AliasToTableLive
}

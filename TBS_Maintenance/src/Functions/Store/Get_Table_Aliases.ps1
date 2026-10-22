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
#   - $script:FunctionResults['Get_Table_Aliases']: @{ Aliases; AliasHash; TableHash }
#   - $script:FunctionResults['SMSVersionFull']: decides static vs dynamic mode
#   - $LoadPath / $OfficePath: location of the Load SQL files (dynamic mode)
#
# Returns:
#   Hashtable alias -> table name.
#
# Writes:
#   - Setup_Files\Get_Table_Aliases.txt (JSON cache: mode, SMS version, Load-file signature,
#     AliasHash + TableHash). The file cache is reused only while the SMS version and (dynamic
#     mode) the *_Load.sql signature still match.
# ===================================================================================================

function Get_Table_Aliases
{
	param (
		[switch]$RefreshCache
	)

	$CachePath = Join-Path $script:SetupFiles 'Get_Table_Aliases.txt'

	# PSCustomObject (from JSON) -> hashtable, so .ContainsKey() works for the callers
	$toHash = {
		param ($o)
		$h = @{ }
		if ($o -is [hashtable]) { return $o }
		if ($o -and $o.PSObject -and $o.PSObject.Properties) { foreach ($p in $o.PSObject.Properties) { $h[$p.Name] = $p.Value } }
		return $h
	}
	$publish = {
		param ($AliasHash, $TableHash)
		$script:FunctionResults['Get_Table_Aliases'] = @{ Aliases = @(); AliasHash = $AliasHash; TableHash = $TableHash }
		return $AliasHash
	}

	# ------------------------------------------------------------------------------------------------
	# 1. In-memory (this launch)
	# ------------------------------------------------------------------------------------------------
	if (-not $RefreshCache)
	{
		try
		{
			$aliasData = $script:FunctionResults['Get_Table_Aliases']
			if ($aliasData -and $aliasData.AliasHash)
			{
				if ($aliasData.AliasHash -is [hashtable]) { return $aliasData.AliasHash }
				return (& $publish (& $toHash $aliasData.AliasHash) (& $toHash $aliasData.TableHash))
			}
		}
		catch { }
	}

	# ------------------------------------------------------------------------------------------------
	# SMS version -> static (supported range) or dynamic mode
	# ------------------------------------------------------------------------------------------------
	$MinSupportedSMSVersion = "3.3.0.0"
	$MaxSupportedSMSVersion = "3.6.1.8"

	$SMSVersion = "0.0.0.0"
	if ([string]$script:FunctionResults['SMSVersionFull'] -match '([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)') { $SMSVersion = $Matches[1] }
	$versionKnown = ($SMSVersion -ne "0.0.0.0")

	$versionInRange = $false
	try
	{
		$vTest = [version]$SMSVersion
		$versionInRange = ($vTest -ge [version]$MinSupportedSMSVersion -and $vTest -le [version]$MaxSupportedSMSVersion)
	}
	catch { $versionInRange = $false }

	if ([string]::IsNullOrWhiteSpace($LoadPath))
	{
		try { if ($OfficePath -and (Test-Path -LiteralPath $OfficePath)) { $LoadPath = Join-Path $OfficePath 'Load' } }
		catch { }
	}

	# Signature of the *_Load.sql files (dynamic mode cache validation)
	$getLoadSignature = {
		try
		{
			if (-not $LoadPath -or -not (Test-Path -LiteralPath $LoadPath)) { return $null }
			$files = @(Get-ChildItem -LiteralPath $LoadPath -Recurse -Filter '*_Load.sql' -File -ErrorAction SilentlyContinue)
			$newest = $null
			$total = [int64]0
			foreach ($f in $files)
			{
				if (-not $newest -or $f.LastWriteTimeUtc -gt $newest) { $newest = $f.LastWriteTimeUtc }
				$total += $f.Length
			}
			return [pscustomobject]@{
				Root       = [string]$LoadPath
				FileCount  = [int]$files.Count
				NewestUtc  = if ($newest) { $newest.ToString('o') } else { "" }
				TotalBytes = [int64]$total
			}
		}
		catch { return $null }
	}

	# ------------------------------------------------------------------------------------------------
	# 2. File cache (only while it still describes this store's SMS version / Load files)
	# ------------------------------------------------------------------------------------------------
	if ((-not $RefreshCache) -and (Test-Path -LiteralPath $CachePath))
	{
		try
		{
			$raw = Get-Content -LiteralPath $CachePath -Raw -ErrorAction Stop
			$cacheObj = $null
			if (-not [string]::IsNullOrWhiteSpace($raw)) { $cacheObj = $raw | ConvertFrom-Json -ErrorAction Stop }
			if ($cacheObj -and $cacheObj.Results)
			{
				$cacheValid = $true
				# A known SMS version that differs from the cached one means an upgrade: rebuild
				if ($versionKnown -and $cacheObj.SMSVersion -and ([string]$cacheObj.SMSVersion -ne $SMSVersion)) { $cacheValid = $false }
				if ($cacheValid -and [string]$cacheObj.Mode -eq 'Dynamic' -and $cacheObj.LoadSignature)
				{
					$sig = & $getLoadSignature
					if ($sig -and (
							[int]$cacheObj.LoadSignature.FileCount -ne $sig.FileCount -or
							[string]$cacheObj.LoadSignature.NewestUtc -ne $sig.NewestUtc -or
							[int64]$cacheObj.LoadSignature.TotalBytes -ne $sig.TotalBytes))
					{
						$cacheValid = $false
					}
				}
				if ($cacheValid)
				{
					return (& $publish (& $toHash $cacheObj.Results.AliasHash) (& $toHash $cacheObj.Results.TableHash))
				}
			}
		}
		catch
		{
			# unreadable cache: rebuild
		}
	}

	$writeCache = {
		param ([string]$Mode, $Signature, $AliasHash, $TableHash)
		try
		{
			$cacheOut = [pscustomobject]@{
				LastUpdated   = (Get-Date).ToString('o')
				SMSVersion    = $SMSVersion
				Mode          = $Mode
				LoadSignature = $Signature
				Results       = [pscustomobject]@{
					AliasHash = [pscustomobject]$AliasHash
					TableHash = [pscustomobject]$TableHash
				}
			}
			($cacheOut | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $CachePath -Encoding UTF8 -Force
		}
		catch { }
	}

	# ------------------------------------------------------------------------------------------------
	# 3a. Static mode (supported SMS range): hardcoded alias -> table map
	# ------------------------------------------------------------------------------------------------
	if ($versionInRange)
	{
		$AliasToTable = @{
			'ALT'      = 'ALT_TAB'
			'BIO'      = 'BIO_TAB'
			'BTL'      = 'BTL_TAB'
			'CAT'      = 'CAT_TAB'
			'CFG'      = 'CFG_TAB'
			'CLF'      = 'CLF_TAB'
			'CLF_SDP'  = 'CLF_SDP_TAB'
			'CLG'      = 'CLG_TAB'
			'CLK'      = 'CLK_TAB'
			'CLL'      = 'CLL_TAB'
			'CLR'      = 'CLR_TAB'
			'CLT'      = 'CLT_TAB'
			'CLT_ITM'  = 'CLT_ITM_TAB'
			'COST'     = 'COST_TAB'
			'CPN'      = 'CPN_TAB'
			'DELV'     = 'DELV_TAB'
			'DEPT'     = 'DEPT_TAB'
			'DSD'      = 'DSD_TAB'
			'ECL'      = 'ECL_TAB'
			'FAM'      = 'FAM_TAB'
			'KIT'      = 'KIT_TAB'
			'LOC'      = 'LOC_TAB'
			'LVL'      = 'LVL_TAB'
			'MIX'      = 'MIX_TAB'
			'MOD'      = 'MOD_TAB'
			'OBJ'      = 'OBJ_TAB'
			'POS'      = 'POS_TAB'
			'PRICE'    = 'PRICE_TAB'
			'PUB'      = 'PUB_TAB'
			'RES'      = 'RES_TAB'
			'ROUTE'    = 'ROUTE_TAB'
			'RPC'      = 'RPC_TAB'
			'SCAL_ITM' = 'SCL_TAB'
			'SCAL_NUT' = 'SCL_NUT_TAB'
			'SCAL_TXT' = 'SCL_TXT_TAB'
			'SDP'      = 'SDP_TAB'
			'STD'      = 'STD_TAB'
			'TAR'      = 'TAR_TAB'
			'UNT'      = 'UNT_TAB'
			'VENDOR'   = 'VENDOR_TAB'
		}
		$TableToAlias = @{ }
		foreach ($k in $AliasToTable.Keys) { $TableToAlias[$AliasToTable[$k]] = $k }

		& $writeCache 'Static' $null $AliasToTable $TableToAlias
		return (& $publish $AliasToTable $TableToAlias)
	}

	# ------------------------------------------------------------------------------------------------
	# 3b. Dynamic mode (SMS version outside the supported range): read @CREATE lines from the
	#     <TABLE>_Load.sql files, e.g.  @CREATE('OBJ_TAB', 'OBJ');
	# ------------------------------------------------------------------------------------------------
	Write-Host "SMS Version $SMSVersion is outside supported range ($MinSupportedSMSVersion - $MaxSupportedSMSVersion). Scanning SQL files for table/alias map..." -ForegroundColor Yellow

	if (-not $LoadPath -or -not (Test-Path -LiteralPath $LoadPath))
	{
		Write-Warning "LoadPath is not available; cannot scan *_Load.sql files. Returning empty hashtable."
		return (& $publish @{ } @{ })
	}

	$BaseTables = @(
		'OBJ', 'POS', 'PRICE', 'COST', 'DSD', 'KIT', 'LOC', 'ALT', 'ECL',
		'SCL', 'SCL_TXT', 'SCL_NUT', 'DEPT', 'SDP', 'CAT', 'RPC', 'FAM',
		'CPN', 'PUB', 'BIO', 'CLK', 'LVL', 'MIX', 'BTL', 'TAR', 'UNT',
		'RES', 'ROUTE', 'VENDOR', 'DELV', 'CLT', 'CLG', 'CLF', 'CLR',
		'CLL', 'CLT_ITM', 'CLF_SDP', 'STD', 'CFG', 'MOD'
	)
	$loadFileNames = @{ }
	foreach ($bt in $BaseTables) { $loadFileNames["${bt}_Load.sql"] = $true }

	$tablesPattern = ($BaseTables | Sort-Object Length -Descending | ForEach-Object { [regex]::Escape($_) }) -join '|'
	$regex = [regex]::new("^\s*@CREATE\s*\(\s*['""]?(?<Table>($tablesPattern)(_[A-Z]+))?['""]?\s*,\s*['""]?(?<Alias>[\w-]+)['""]?\s*\);", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

	$AliasToTableLive = @{ }
	$TableToAliasLive = @{ }
	foreach ($file in @(Get-ChildItem -LiteralPath $LoadPath -Recurse -Filter '*_Load.sql' -File -ErrorAction SilentlyContinue))
	{
		if (-not $loadFileNames.ContainsKey($file.Name)) { continue }
		foreach ($line in @(Get-Content -LiteralPath $file.FullName -ErrorAction SilentlyContinue))
		{
			$lineClean = $line -replace '--.*', '' -replace '/\*.*?\*/', ''
			if ($lineClean -notmatch '@CREATE') { continue }
			$match = $regex.Match($lineClean)
			if ($match.Success -and $match.Groups['Table'].Value -and $match.Groups['Alias'].Value)
			{
				$AliasToTableLive[$match.Groups['Alias'].Value] = $match.Groups['Table'].Value
				$TableToAliasLive[$match.Groups['Table'].Value] = $match.Groups['Alias'].Value
			}
		}
	}
	if ($AliasToTableLive.Count -eq 0) { Write-Warning "No table-alias pairs detected. Using empty hashtable." }

	& $writeCache 'Dynamic' (& $getLoadSignature) $AliasToTableLive $TableToAliasLive
	return (& $publish $AliasToTableLive $TableToAliasLive)
}

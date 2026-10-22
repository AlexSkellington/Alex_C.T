# ===================================================================================================
#                              FUNCTION: Get_Store_And_Database_Info
# ---------------------------------------------------------------------------------------------------
# Description:
#   Reads store identity and SQL connection details (store name, address, SMS version, key
#   number, DB server, database, connection string) from the Storeman INI files, preferring
#   in-memory results, then the cache file; re-parses only when missing or -RefreshCache.
#
# Parameters:
#   - WinIniPath      [string]: path to INFO_*901_WIN.INI
#   - SmsStartIniPath [string]: path to INFO_*901_SMSStart.ini
#   - StartupIniPath  [string]: path to the Storeman Startup.ini
#   - SystemIniPath   [string]: path to the Storeman system.ini
#   - RefreshCache    [switch]: force re-parse and cache rewrite
#
# Variables:
#   - $script:FunctionResults: receives StoreNumber, DBSERVER, DBNAME, ConnectionString,
#     SMSVersionFull, SqlModuleName and friends
#
# Writes:
#   - Setup_Files\Get_Store_And_Database_Info.txt (JSON cache)
#
# Notes:
#   - Always re-detects the SQL module; updates GUI labels even on a cache hit (finally block).
# ===================================================================================================

function Get_Store_And_Database_Info
{
	param (
		[string]$WinIniPath,
		[string]$SmsStartIniPath,
		[string]$StartupIniPath,
		[string]$SystemIniPath,
		[switch]$RefreshCache
	)
	
	# ------------------------------------------------------------------------------------------------
	# Cache (per-function file) - only re-run if memory/file data is missing/unreadable or -RefreshCache is used
	# ------------------------------------------------------------------------------------------------
	$CachePath = Join-Path $script:SetupFiles 'Get_Store_And_Database_Info.txt'
	$fields = @(
		'StoreNumber', 'StoreName', 'CompanyName', 'Terminal', 'Address', 'City', 'State',
		'SMSVersionFull', 'StampDate', 'StampTime', 'KeyNumber',
		'DBSERVER', 'DBNAME', 'ConnectionString', 'ResolvedWinIniPath'
	)
	
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
	
	# CHANGE: Wrap the whole function in try/finally so GUI labels update even if we "return" on cache hit
	try
	{
		$availableSqlModule = $null
		try
		{
			if (
				$script:FunctionResults.ContainsKey('SqlModuleName') -and
				-not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['SqlModuleName']) -and
				([string]$script:FunctionResults['SqlModuleName'] -ne 'N/A')
			)
			{
				$availableSqlModule = [string]$script:FunctionResults['SqlModuleName']
			}
		}
		catch { $availableSqlModule = $null }
		
		if (-not $RefreshCache)
		{
			# Reuse in-memory values if they are already populated for this launch
			$hasInMemoryResults = $true
			foreach ($field in @('StoreNumber', 'StoreName', 'DBSERVER', 'DBNAME', 'ConnectionString', 'SMSVersionFull'))
			{
				if (
					-not $script:FunctionResults.ContainsKey($field) -or
					[string]::IsNullOrWhiteSpace([string]$script:FunctionResults[$field]) -or
					([string]$script:FunctionResults[$field] -eq 'N/A')
				)
				{
					$hasInMemoryResults = $false
					break
				}
			}
			if ($hasInMemoryResults)
			{
				return
			}
			
			# Reuse file cache if present
			if ($CachePath -and (Test-Path $CachePath))
			{
				try
				{
					$raw = Get-Content -LiteralPath $CachePath -Raw -ErrorAction Stop
					if (-not [string]::IsNullOrWhiteSpace($raw))
					{
						$cache = $raw | ConvertFrom-Json -ErrorAction Stop
						if ($cache -and $cache.Results)
						{
							foreach ($p in $cache.Results.PSObject.Properties)
							{
								$script:FunctionResults[$p.Name] = $p.Value
							}
							return
						}
					}
				}
				catch { }
			}
		}
		
		if ([string]::IsNullOrWhiteSpace($availableSqlModule) -or $availableSqlModule -eq 'N/A')
		{
			try
			{
				$loadedSqlModule = Get-Module SqlServer, SQLPS | Select-Object -First 1
				if ($loadedSqlModule) { $availableSqlModule = [string]$loadedSqlModule.Name }
			}
			catch { }
		}
		if ([string]::IsNullOrWhiteSpace($availableSqlModule) -or $availableSqlModule -eq 'N/A')
		{
			try
			{
				if (Get-Module -ListAvailable -Name SqlServer | Select-Object -First 1) { $availableSqlModule = "SqlServer" }
				elseif (Get-Module -ListAvailable -Name SQLPS | Select-Object -First 1) { $availableSqlModule = "SQLPS" }
				else { $availableSqlModule = "None" }
			}
			catch
			{
				$availableSqlModule = if ($availableSqlModule) { $availableSqlModule }
				else { "None" }
			}
		}
		$script:FunctionResults['SqlModuleName'] = $availableSqlModule
		
		# ------------------------------------------------------------------------------------------------
		# Initialize results with N/A for every expected property (+ resolved WIN.INI path)
		# ------------------------------------------------------------------------------------------------
		foreach ($f in $fields) { $script:FunctionResults[$f] = 'N/A' }
		
		# ==================================================================================================
		# Parse Startup.ini first to learn STORE and TER and to derive likely DBs folder(s)
		# ==================================================================================================
		$startupStore = $null
		$startupTer = $null
		$probableDbsRoots = @()
		
		if ($StartupIniPath -and (Test-Path $StartupIniPath))
		{
			try
			{
				foreach ($line in Get-Content -LiteralPath $StartupIniPath)
				{
					$t = $line.Trim()
					if ($t.StartsWith(';')) { continue }
					
					if ($t -match '^\s*STORE\s*=\s*(\d+)\s*$')
					{
						$startupStore = ($Matches[1] -as [int]).ToString('000')
					}
					elseif ($t -match '^\s*TER\s*=\s*(\d+)\s*$')
					{
						$startupTer = ($Matches[1] -as [int]).ToString('000')
					}
				}
				
				if ($startupStore)
				{
					$script:FunctionResults['StoreNumber'] = $startupStore
					if ($startupTer) { $script:FunctionResults['Terminal'] = $startupTer }
				}
				else
				{
					Write_Log "STORE= not found in Startup.ini. Will still attempt WIN.INI path resolution." "yellow"
				}
				
				$storemanRoot = Split-Path -Path $StartupIniPath -Parent
				if ($storemanRoot)
				{
					$derivedOffice = Join-Path $storemanRoot 'office'
					$derivedDbs = Join-Path $derivedOffice 'dbs'
					if (Test-Path $derivedDbs) { $probableDbsRoots += $derivedDbs }
				}
			}
			catch
			{
				Write_Log "Error parsing Startup.ini: $($_.Exception.Message)" "red"
			}
		}
		else
		{
			Write_Log "Startup.ini path not provided or not found. Skipping STORE/TER pre-resolution." "yellow"
		}
		
		foreach ($fallback in $script:StoremanOfficeDbsFallbacks)
		{
			if (Test-Path -LiteralPath $fallback) { $probableDbsRoots += $fallback }
		}
		$probableDbsRoots = $probableDbsRoots | Sort-Object -Unique
		
		# ==================================================================================================
		# Resolve correct INFO_*_WIN.INI matching Startup.ini STORE/TER
		# ==================================================================================================
		$resolvedWinIni = $null
		$needResolve = $true
		
		if ($WinIniPath -and (Test-Path $WinIniPath))
		{
			$name = [System.IO.Path]::GetFileName($WinIniPath)
			if ($name -match '^INFO_(\d{3})(\d{3})_WIN\.INI$')
			{
				$suppliedStore = $Matches[1]
				$suppliedTer = $Matches[2]
				$storeOk = -not $startupStore -or ($suppliedStore -eq $startupStore)
				$terOk = -not $startupTer -or ($suppliedTer -eq $startupTer)
				if ($storeOk -and $terOk)
				{
					$resolvedWinIni = $WinIniPath
					$needResolve = $false
				}
				else
				{
					Write_Log "Provided WIN.INI '$name' does not coincide with Startup.ini (STORE=$startupStore TER=$startupTer). Will search for the correct one." "yellow"
				}
			}
			else
			{
				Write_Log "Provided WIN.INI name '$name' does not match expected pattern 'INFO_SSSLLL_WIN.INI'. Will search for the correct one." "yellow"
			}
		}
		
		if ($needResolve -and $probableDbsRoots.Count -gt 0)
		{
			try
			{
				$candidates = foreach ($root in $probableDbsRoots)
				{
					Get-ChildItem -LiteralPath $root -Filter 'INFO_*_WIN.INI' -File -ErrorAction SilentlyContinue
				}
				
				$scored = @(
					foreach ($f in $candidates)
					{
						$m = [System.Text.RegularExpressions.Regex]::Match($f.Name, '^INFO_(\d{3})(\d{3})_WIN\.INI$')
						if (-not $m.Success) { continue }
						
						$s = $m.Groups[1].Value
						$t = $m.Groups[2].Value
						
						$score = 0
						if ($startupStore -and $s -eq $startupStore) { $score += 10 }
						if ($startupTer -and $t -eq $startupTer) { $score += 5 }
						if (-not $startupTer -and $t -eq '901') { $score += 3 }
						
						[PSCustomObject]@{
							File  = $f
							Store = $s
							Ter   = $t
							Score = $score
							Age   = $f.LastWriteTimeUtc
						}
					}
				) | Sort-Object -Property @{ Expression = 'Score'; Descending = $true }, @{ Expression = 'Age'; Descending = $true }
				
				if ($scored -and $scored[0].Score -gt 0)
				{
					$resolvedWinIni = $scored[0].File.FullName
				}
				elseif ($scored)
				{
					$resolvedWinIni = ($scored | Sort-Object -Property Age -Descending | Select-Object -First 1).File.FullName
				}
				else
				{
					Write_Log "No INFO_*_WIN.INI candidates found under: $($probableDbsRoots -join '; ')" "red"
				}
			}
			catch
			{
				Write_Log "Error while resolving WIN.INI: $($_.Exception.Message)" "red"
			}
		}
		
		if ($resolvedWinIni)
		{
			$WinIniPath = $resolvedWinIni
			$script:FunctionResults['ResolvedWinIniPath'] = $resolvedWinIni
		}
		else
		{
			if (-not ($WinIniPath -and (Test-Path $WinIniPath)))
			{
				Write_Log "WIN.INI is unresolved; continuing without store metadata from WIN.INI." "red"
			}
			else
			{
				$script:FunctionResults['ResolvedWinIniPath'] = $WinIniPath
				Write_Log "Proceeding with provided WIN.INI (unvalidated): $WinIniPath" "yellow"
			}
		}
		
		# ------------------------------------------------------------------------------------------------
		# Extract Store Info from WIN.INI
		# ------------------------------------------------------------------------------------------------
		if ($WinIniPath -and (Test-Path $WinIniPath))
		{
			$currentSection = ""
			foreach ($line in Get-Content -LiteralPath $WinIniPath)
			{
				$trimmed = $line.Trim()
				if ($trimmed -match '^\[(.+)\]$') { $currentSection = $Matches[1]; continue }
				if ($trimmed -notmatch "=" -or $trimmed.StartsWith(";")) { continue }
				
				$parts = $trimmed -split "=", 2
				$key = $parts[0].Trim()
				$value = $parts[1].Trim()
				
				switch ($currentSection)
				{
					"ORIGIN" {
						if ($key -ieq "StampDate") { $script:FunctionResults['StampDate'] = $value }
						if ($key -ieq "StampTime") { $script:FunctionResults['StampTime'] = $value }
					}
					"SYSTEM" {
						if ($key -ieq "CompanyName") { $script:FunctionResults['CompanyName'] = $value }
						if ($key -ieq "Store") { $script:FunctionResults['StoreNumber'] = ($value -as [int]).ToString('000') }
						if ($key -ieq "Terminal") { $script:FunctionResults['Terminal'] = ($value -as [int]).ToString('000') }
					}
					"STOREDETAIL" {
						if ($key -ieq "Name") { $script:FunctionResults['StoreName'] = $value }
						if ($key -ieq "Address") { $script:FunctionResults['Address'] = $value }
						if ($key -ieq "City") { $script:FunctionResults['City'] = $value }
						if ($key -ieq "State") { $script:FunctionResults['State'] = $value }
					}
					"KEY" {
						if ($key -ieq "KeyNumber") { $script:FunctionResults['KeyNumber'] = $value }
					}
					"Versions" {
						if ($key -ieq "VersionIni") { $script:FunctionResults['SMSVersionFull'] = $value }
					}
				}
			}
		}
		else
		{
			if ([string]::IsNullOrWhiteSpace($WinIniPath))
			{
				Write_Log "No store info file (INFO_*_WIN.INI) found - no storeman folder detected on this machine." "red"
			}
			else
			{
				Write_Log "No store info file (INFO_*_WIN.INI) found at $WinIniPath" "red"
			}
		}
		
		# ------------------------------------------------------------------------------------------------
		# Fallback: StoreName from system.ini ([SMS] Name=...)
		# ------------------------------------------------------------------------------------------------
		if (
			($script:FunctionResults['StoreName'] -eq 'N/A' -or [string]::IsNullOrWhiteSpace($script:FunctionResults['StoreName'])) -and
			$SystemIniPath -and (Test-Path $SystemIniPath)
		)
		{
			$inSMSSection = $false
			foreach ($line in Get-Content -LiteralPath $SystemIniPath)
			{
				$trimmed = $line.Trim()
				if ($trimmed -match '^\[SMS\]$') { $inSMSSection = $true; continue }
				if ($inSMSSection)
				{
					if ($trimmed -match '^\[.+\]$') { break }
					if ($trimmed -match '^\s*Name\s*=(.*)$')
					{
						$storeNameBackup = $Matches[1].Trim()
						if ($storeNameBackup) { $script:FunctionResults['StoreName'] = $storeNameBackup; break }
					}
				}
			}
		}
		
		# ------------------------------------------------------------------------------------------------
		# Extract SQL Server/Database from SMSStart INI
		# ------------------------------------------------------------------------------------------------
		$dbServer = $null
		$dbName = $null
		
		if ($SmsStartIniPath -and (Test-Path $SmsStartIniPath))
		{
			foreach ($line in Get-Content -LiteralPath $SmsStartIniPath)
			{
				$trimmed = $line.Trim()
				if ($trimmed -notmatch "=" -or $trimmed.StartsWith(";")) { continue }
				$parts = $trimmed -split "=", 2
				$key = $parts[0].Trim()
				$value = $parts[1].Trim()
				if (-not $dbServer -and $key -match 'ServerName') { $dbServer = $value }
				if (-not $dbName -and $key -match 'DatabaseName') { $dbName = $value }
				if ($dbServer -and $dbName) { break }
			}
		}
		
		# ------------------------------------------------------------------------------------------------
		# Fallback: Startup.ini for DBSERVER/DBNAME
		# ------------------------------------------------------------------------------------------------
		if ((!$dbServer -or !$dbName) -and $StartupIniPath -and (Test-Path $StartupIniPath))
		{
			foreach ($line in Get-Content -LiteralPath $StartupIniPath)
			{
				$trimmed = $line.Trim()
				if ($trimmed -notmatch "=" -or $trimmed.StartsWith(";")) { continue }
				$parts = $trimmed -split "=", 2
				$key = $parts[0].Trim()
				$value = $parts[1].Trim()
				if (-not $dbServer -and $key -match '^DBSERVER$') { $dbServer = $value }
				if (-not $dbName -and $key -match '^DBNAME$') { $dbName = $value }
				if ($dbServer -and $dbName) { break }
			}
			if (-not $dbServer) { $dbServer = "localhost" }
			if (-not $dbName) { $dbName = "STORESQL" }
		}
		
		# ------------------------------------------------------------------------------------------------
		# Finalize connection string
		# ------------------------------------------------------------------------------------------------
		if ($dbServer -and $dbName)
		{
			$script:FunctionResults['DBSERVER'] = $dbServer
			$script:FunctionResults['DBNAME'] = $dbName
			$script:FunctionResults['ConnectionString'] =
			"Server=$dbServer;Database=$dbName;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connect Timeout=5;Application Name=TBS_Maintenance_Script"
		}
		
		# ------------------------------------------------------------------------------------------------
		# Cache write (store EVERYTHING this function collects)
		# ------------------------------------------------------------------------------------------------
		if ($CachePath)
		{
			try
			{
				$results = [ordered]@{ }
				foreach ($k in ($fields + 'SqlModuleName'))
				{
					$v = $script:FunctionResults[$k]
					if ($null -eq $v -or ($v -is [string] -and [string]::IsNullOrWhiteSpace($v))) { $v = 'N/A' }
					$results[$k] = $v
				}
				
				$cacheOut = [ordered]@{
					LastUpdated = (Get-Date).ToString('o')
					StoreNumber = $script:FunctionResults['StoreNumber']
					Results	    = $results
				}
				
				($cacheOut | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $CachePath -Encoding UTF8 -Force
			}
			catch { }
		}
	}
	finally
	{
		# ------------------------------------------------------------------------------------------------
		# GUI label updates (optional)
		# ------------------------------------------------------------------------------------------------
		# CHANGE: This runs even when we returned early due to a cache hit (fixes "labels not updating").
		# CHANGE: Safer Refresh/DoEvents calls (won't throw if $form isn't set or WinForms isn't loaded).
		try
		{
			$anyLabelChanged = $false
			
			if ($storeNumberLabel -ne $null)
			{
				$storeNumberLabel.Text = "Store Number: $($script:FunctionResults['StoreNumber'])"
				$anyLabelChanged = $true
			}
			if ($storeNameLabel -ne $null)
			{
				$storeNameLabel.Text = "Store Name: $($script:FunctionResults['StoreName'])"
				$anyLabelChanged = $true
			}
			if ($bannerDateLabel -ne $null)
			{
				try
				{
					$bn = $script:FunctionResults['StoreNumber']
					$bnm = $script:FunctionResults['StoreName']
					if ($bn -and "$bn" -ne 'N/A' -and -not [string]::IsNullOrWhiteSpace("$bn"))
					{
						if ($bnm -and "$bnm" -ne 'N/A' -and -not [string]::IsNullOrWhiteSpace("$bnm"))
						{
						$bnmShort = "$bnm"
						if ($bnmShort.Length -gt 26) { $bnmShort = $bnmShort.Substring(0, 25).TrimEnd() + '...' }
						$bannerDateLabel.Text = "Store $bn  -  $bnmShort"
						}
						else { $bannerDateLabel.Text = "Store $bn" }
					}
					else
					{
					$bannerDateLabel.Text = "Store Maintenance Console"
					}
					$anyLabelChanged = $true
					try { if ($layoutBanner) { & $layoutBanner } } catch { }
				}
				catch { }
			}
			if ($smsVersionLabel -ne $null)
			{
				$smsVersionLabel.Text = "SMS Version: $($script:FunctionResults['SMSVersionFull'])"
				$anyLabelChanged = $true
			}
			
			if ($anyLabelChanged)
			{
				if ($form -ne $null) { $form.Refresh() }
				try { [System.Windows.Forms.Application]::DoEvents() }
				catch { }
			}
		}
		catch { }
	}
}

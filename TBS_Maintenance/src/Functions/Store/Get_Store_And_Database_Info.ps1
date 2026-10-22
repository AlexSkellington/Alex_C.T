# ===================================================================================================
#                              FUNCTION: Get_Store_And_Database_Info
# ---------------------------------------------------------------------------------------------------
# Description:
#   Reads store identity and SQL connection details (store name, address, SMS version, key
#   number, DB server, database, connection string) from the Storeman INI files. The files are
#   local and read in a few milliseconds, so they are parsed LIVE on every launch (no file
#   cache: a cached copy showed stale names / servers until a background refresh replaced it).
#   Within one launch, a repeat call returns the values already in memory.
#
# Parameters:
#   - WinIniPath      [string]: path to INFO_*901_WIN.INI
#   - SmsStartIniPath [string]: path to INFO_*901_SMSStart.ini
#   - StartupIniPath  [string]: path to the Storeman Startup.ini
#   - SystemIniPath   [string]: path to the Storeman system.ini
#   - RefreshCache    [switch]: re-parse even if this launch already has the values, and
#                                re-detect the SQL module
#
# Variables:
#   - $script:FunctionResults: receives StoreNumber, DBSERVER, DBNAME, ConnectionString,
#     SMSVersionFull, SqlModuleName and friends
#
# Writes:
#   - Setup_Files\SqlModule.txt - the one remembered value: which SQL PowerShell module is
#     installed (finding out scans every module folder - seconds on some store PCs). Re-checked
#     once a day, when the remembered module folder is gone, or with -RefreshCache.
#
# Notes:
#   - Updates the GUI labels on every call, including the in-memory return (finally block).
#   - StoreNumber / Terminal are 3-digit padded when numeric (1 -> 001).
#   - StoreName is resolved here only: system.ini [SMS] Name= first, INFO_*_WIN.INI
#     [STOREDETAIL] Name= as fallback ("Company Name" placeholder = empty); StoreNameSource
#     says which. Other functions read $script:FunctionResults['StoreName'].
# ===================================================================================================

function Get_Store_And_Database_Info
{
	param (
		# Default to the store paths resolved at startup (Bootstrap\06_Store_Paths)
		[string]$WinIniPath = $script:WinIniPath,
		[string]$SmsStartIniPath = $script:SmsStartIniPath,
		[string]$StartupIniPath = $script:StartupIniPath,
		[string]$SystemIniPath = $script:SystemIniPath,
		[switch]$RefreshCache
	)

	$fields = @(
		'StoreNumber', 'StoreName', 'StoreNameSource', 'CompanyName', 'Terminal', 'Address', 'City', 'State',
		'SMSVersionFull', 'StampDate', 'StampTime', 'KeyNumber',
		'DBSERVER', 'DBNAME', 'ConnectionString', 'ResolvedWinIniPath'
	)

	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }

	# try/finally: the GUI labels update even on the in-memory early return
	try
	{
		# Already parsed during this launch -> nothing to re-read
		if (-not $RefreshCache -and $script:FunctionResults.ContainsKey('StoreInfoParsedAt')) { return }

		# The cache file of v2.7.3 and earlier is no longer used
		try
		{
			$oldCache = Join-Path $script:SetupFiles 'Get_Store_And_Database_Info.txt'
			if (Test-Path -LiteralPath $oldCache) { Remove-Item -LiteralPath $oldCache -Force -ErrorAction SilentlyContinue }
		}
		catch { }

		# ------------------------------------------------------------------------------------------------
		# SQL module (SqlServer / SQLPS / None) - the only remembered value, see header
		# ------------------------------------------------------------------------------------------------
		$availableSqlModule = $null
		$moduleMemo = $null
		if ($script:SetupFiles) { $moduleMemo = Join-Path $script:SetupFiles 'SqlModule.txt' }
		if (-not $RefreshCache -and $moduleMemo -and (Test-Path -LiteralPath $moduleMemo))
		{
			try
			{
				# line 1: module name, line 2: its folder ('' for None)
				$memo = @(Get-Content -LiteralPath $moduleMemo -ErrorAction Stop)
				$memoAge = (Get-Date) - (Get-Item -LiteralPath $moduleMemo).LastWriteTime
				$memoName = [string]$memo[0]
				$memoBase = if ($memo.Count -gt 1) { [string]$memo[1] } else { '' }
				if ($memoName -and $memoAge.TotalHours -lt 24 -and ($memoName -eq 'None' -or ($memoBase -and (Test-Path -LiteralPath $memoBase))))
				{
					$availableSqlModule = $memoName
				}
			}
			catch { $availableSqlModule = $null }
		}
		if (-not $availableSqlModule)
		{
			$moduleBase = ''
			try
			{
				$found = Get-Module SqlServer, SQLPS | Select-Object -First 1
				if (-not $found) { $found = Get-Module -ListAvailable -Name SqlServer | Select-Object -First 1 }
				if (-not $found) { $found = Get-Module -ListAvailable -Name SQLPS | Select-Object -First 1 }
				if ($found) { $availableSqlModule = [string]$found.Name; $moduleBase = [string]$found.ModuleBase }
				else { $availableSqlModule = 'None' }
			}
			catch { $availableSqlModule = 'None' }
			if ($moduleMemo)
			{
				try { Set-Content -LiteralPath $moduleMemo -Value @($availableSqlModule, $moduleBase) -Encoding ASCII -Force }
				catch { }
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
						$startupStore = ([int]$Matches[1]).ToString('000')
					}
					elseif ($t -match '^\s*TER\s*=\s*(\d+)\s*$')
					{
						$startupTer = ([int]$Matches[1]).ToString('000')
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
		$winIniStoreName = $null
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
						if ($key -ieq "Store") { $script:FunctionResults['StoreNumber'] = if ($value -match '^\d+$') { ([int]$value).ToString('000') } else { $value } }
						if ($key -ieq "Terminal") { $script:FunctionResults['Terminal'] = if ($value -match '^\d+$') { ([int]$value).ToString('000') } else { $value } }
					}
					"STOREDETAIL" {
						if ($key -ieq "Name") { $winIniStoreName = $value } # fallback only - see "Store name" below
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
		# Store name - resolved ONCE here; everything else reads $script:FunctionResults['StoreName']
		#   1. Office\system.ini: [SMS] Name=, else the first NAME= line in any section
		#   2. INFO_*_WIN.INI [STOREDETAIL] Name=  (fallback)
		#   SMS writes the placeholder "Company Name" when the store details were never filled in;
		#   it counts as empty. StoreNameSource records where the name came from.
		# ------------------------------------------------------------------------------------------------
		$cleanStoreName = {
			param ([string]$Value)
			if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
			$v = $Value.Trim()
			if ($v.StartsWith('"'))
			{
				# quoted: everything inside the outer quotes
				$lastQuote = $v.LastIndexOf('"')
				$v = if ($lastQuote -gt 0) { $v.Substring(1, $lastQuote - 1).Trim() } else { $v.Trim('"').Trim() }
			}
			else
			{
				# ' ;' starts an inline comment; '#' stays ("Supermarket #2")
				$mSemi = [regex]::Match($v, '^(.*?)\s;.*$')
				if ($mSemi.Success) { $v = $mSemi.Groups[1].Value.Trim() }
			}
			if ([string]::IsNullOrWhiteSpace($v) -or $v -ieq 'Company Name' -or $v -ieq 'N/A') { return $null }
			return $v
		}

		$systemIniStoreName = $null
		if ($SystemIniPath -and (Test-Path -LiteralPath $SystemIniPath))
		{
			$smsSectionName = $null
			$firstName = $null
			$section = ''
			foreach ($line in Get-Content -LiteralPath $SystemIniPath -ErrorAction SilentlyContinue)
			{
				$trimmed = $line.Trim()
				if ($trimmed -match '^\[(.+)\]$') { $section = $Matches[1].Trim(); continue }
				if ($trimmed -match '^(?i)NAME\s*=(.*)$')
				{
					$candidate = & $cleanStoreName $Matches[1]
					if (-not $candidate) { continue }
					if ($section -ieq 'SMS' -and -not $smsSectionName) { $smsSectionName = $candidate }
					if (-not $firstName) { $firstName = $candidate }
				}
			}
			$systemIniStoreName = if ($smsSectionName) { $smsSectionName } else { $firstName }
		}

		$winIniName = & $cleanStoreName $winIniStoreName
		if ($systemIniStoreName)
		{
			$script:FunctionResults['StoreName'] = $systemIniStoreName
			$script:FunctionResults['StoreNameSource'] = "system.ini ($SystemIniPath)"
		}
		elseif ($winIniName)
		{
			$script:FunctionResults['StoreName'] = $winIniName
			$script:FunctionResults['StoreNameSource'] = "INFO WIN.INI ($WinIniPath)"
		}
		else
		{
			$script:FunctionResults['StoreName'] = 'N/A'
			$script:FunctionResults['StoreNameSource'] = 'not found (system.ini / INFO WIN.INI have no usable Name=)'
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

		$script:FunctionResults['StoreInfoParsedAt'] = Get-Date
	}
	finally
	{
		# ------------------------------------------------------------------------------------------------
		# GUI label updates (optional)
		# ------------------------------------------------------------------------------------------------
		# Runs even after an early return on a cache hit; safe without a form (background jobs).
		try
		{
			$anyLabelChanged = $false

			if ($null -ne $storeNumberLabel)
			{
				$storeNumberLabel.Text = "Store Number: $($script:FunctionResults['StoreNumber'])"
				$anyLabelChanged = $true
			}
			if ($null -ne $storeNameLabel)
			{
				$storeNameLabel.Text = "Store Name: $($script:FunctionResults['StoreName'])"
				# Hover shows where the name came from (system.ini / INFO WIN.INI)
				try { if ($toolTip -and $script:FunctionResults['StoreNameSource']) { $toolTip.SetToolTip($storeNameLabel, "Store name from " + $script:FunctionResults['StoreNameSource'] + ". Click to ping all nodes (Lanes > Scales > Backoffices).") } }
				catch { }
				$anyLabelChanged = $true
			}
			if ($null -ne $bannerDateLabel)
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
			if ($null -ne $smsVersionLabel)
			{
				$smsVersionLabel.Text = "SMS Version: $($script:FunctionResults['SMSVersionFull'])"
				$anyLabelChanged = $true
			}

			if ($anyLabelChanged)
			{
				if ($null -ne $script:MainForm) { $script:MainForm.Refresh() }
				try { [System.Windows.Forms.Application]::DoEvents() }
				catch { }
			}
		}
		catch { }
	}
}

# ===================================================================================================
#                              FUNCTION: Troubleshoot_ScaleCommApp
# ---------------------------------------------------------------------------------------------------
# Description:
#   Audits and optionally repairs the ScaleCommApp *.exe.config files: sets StoreName (the
#   store name resolved by Get_Store_And_Database_Info), FistScaleIP (first active scale IP),
#   SQL_InstanceName, SQL_databaseName and Recs_BatchSendFull; restores the deploy line in
#   Office\DEPLOY_CHG.sql and removes the "Update_Scales_Specials_Task_Minutes" task if present
#   (both through Fix_Deploy_CHG).
#
# Parameters:
#   - ConfigPath          [string]: specific config file to target
#   - ScaleCommRoots      [string[]]: candidate ScaleCommApp install roots (default $script:ScaleCommRoots)
#   - AllMatches          [switch]: fix every matching config (default: first only)
#   - NoAutoFix           [switch]: audit only, do not save
#   - Force               [switch]: overwrite without confirmation
#   - ExpectedStoreName   [string]: expected StoreName (default: $script:FunctionResults['StoreName'])
#   - ExpectedSqlInstance / ExpectedDatabase [string]: expected SQL instance (DBSERVER) / database (DBNAME)
#   - OfficePath          [string]: Office folder for DEPLOY_CHG.sql (default $script:OfficePath)
#
# Variables:
#   - $script:FunctionResults: DBNAME, DBSERVER, ConnectionString, SqlModuleName, StoreName; $script:ScaleCommRoots
#
# Writes:
#   - ScaleCommApp *.exe.config files (UTF-8 no BOM, .bak backups) and Office\DEPLOY_CHG.sql
# ===================================================================================================

function Troubleshoot_ScaleCommApp
{
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
		[string]$ConfigPath,
		[string[]]$ScaleCommRoots = $script:ScaleCommRoots,
		[switch]$AllMatches,
		[switch]$NoAutoFix,
		[switch]$Force,
		[string]$ExpectedStoreName,
		[string]$ExpectedSqlInstance,
		[string]$ExpectedDatabase,
		# NEW: where Office\DEPLOY_CHG.sql lives. If not passed, uses $script:OfficePath when available.
		[string]$OfficePath
	)

	Write_Log "`r`n==================== Starting Troubleshoot_ScaleCommApp ====================`r`n" "blue"

	# ------------------------------------------------------------------------------------------------
	# Resolve expectations & connection (same behavior; no ternary)
	# ------------------------------------------------------------------------------------------------
	$dbName = $script:FunctionResults['DBNAME']
	$server = $script:FunctionResults['DBSERVER']
	$connStr = $script:FunctionResults['ConnectionString']
	$sqlModule = $script:FunctionResults['SqlModuleName']

	if (-not $ExpectedDatabase -and $dbName) { $ExpectedDatabase = $dbName }
	if (-not $ExpectedSqlInstance -and $server) { $ExpectedSqlInstance = $server }
	# Store name: resolved once by Get_Store_And_Database_Info (system.ini first, INFO WIN.INI
	# fallback) - no second parse here.

	if (-not $connStr -and $server -and $dbName)
	{
		# Build connection string inline (no helpers)
		$sqlUser = $script:FunctionResults['SQL_UserID']
		$sqlPwd = $script:FunctionResults['SQL_UserPwd']
		if ($sqlUser -and $sqlPwd)
		{
			$connStr = "Server=$server;Database=$dbName;User ID=$sqlUser;Password=$sqlPwd;TrustServerCertificate=True;"
		}
		else
		{
			$connStr = "Server=$server;Database=$dbName;Integrated Security=SSPI;TrustServerCertificate=True;"
		}
		$script:FunctionResults['ConnectionString'] = $connStr
	}

	# ------------------------------------------------------------------------------------------------
	# Active NIC prefix (fallback) - logged ONCE (inline)
	# ------------------------------------------------------------------------------------------------
	$ActiveNicPrefix = $null
	try
	{
		# Collect server IPv4s (for subnet check)
		$serverIPv4s = @()
		try
		{
			[System.Net.Dns]::GetHostAddresses((($server -replace '^(?i)tcp:', '') -split '[\\,]')[0]) | ForEach-Object {
				if ($_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { $serverIPv4s += $_.ToString() }
			}
		}
		catch { }

		# Get all UP NICs
		$allNics = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
		Where-Object { $_.OperationalStatus -eq [System.Net.NetworkInformation.OperationalStatus]::Up }

		# Try to pick an interface on the same subnet as SQL server (inline mask compare)
		$pickedIP = $null
		foreach ($ni in $allNics)
		{
			$ipProps = $ni.GetIPProperties()
			foreach ($ua in $ipProps.UnicastAddresses)
			{
				if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
				if (-not $ua.IPv4Mask) { continue }

				foreach ($srv in $serverIPv4s)
				{
					$ok = $false
					try
					{
						$a = [System.Net.IPAddress]::Parse($ua.Address.ToString()).GetAddressBytes()
						$b = [System.Net.IPAddress]::Parse($srv).GetAddressBytes()
						$m = [System.Net.IPAddress]::Parse($ua.IPv4Mask.ToString()).GetAddressBytes()
						$ok = $true
						for ($i = 0; $i -lt 4; $i++) { if (($a[$i] -band $m[$i]) -ne ($b[$i] -band $m[$i])) { $ok = $false; break } }
					}
					catch { $ok = $false }
					if ($ok) { $pickedIP = $ua.Address.ToString(); break }
				}
				if ($pickedIP) { break }
			}
			if ($pickedIP) { break }
		}

		# Gateway-backed interface fallback
		if (-not $pickedIP)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				$hasGw = $false
				foreach ($gw in $ipProps.GatewayAddresses)
				{
					if ($gw.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { $hasGw = $true; break }
				}
				if (-not $hasGw) { continue }
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					$pickedIP = $ua.Address.ToString(); break
				}
				if ($pickedIP) { break }
			}
		}

		# RFC1918 preference fallback
		if (-not $pickedIP)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					$ip = $ua.Address.ToString()
					if ($ip.StartsWith('10.') -or $ip.StartsWith('192.168.')) { $pickedIP = $ip; break }
					if ($ip.StartsWith('172.'))
					{
						$parts = $ip.Split('.')
						if ($parts.Length -ge 2)
						{
							$sec = 0
							if ([int]::TryParse($parts[1], [ref]$sec)) { if ($sec -ge 16 -and $sec -le 31) { $pickedIP = $ip; break } }
						}
					}
				}
				if ($pickedIP) { break }
			}
		}

		# Derive /24 prefix
		if ($pickedIP) { $p = $pickedIP.Split('.'); if ($p.Length -ge 3) { $ActiveNicPrefix = "$($p[0]).$($p[1]).$($p[2])" } }
	}
	catch { }

	if ($ActiveNicPrefix) { Write_Log "Active NIC /24 prefix: $ActiveNicPrefix" "cyan" }
	else { Write_Log "Active NIC /24 prefix not detected (fallback may be unavailable)." "yellow" }

	# ------------------------------------------------------------------------------------------------
	# FIRST active scale full IP (robust) - logged ONCE (inline)
	# ------------------------------------------------------------------------------------------------
	$firstScaleIP = $null
	if ($server -and $dbName -and $connStr)
	{
		# Decide Invoke-Sqlcmd availability and argument style
		$InvokeSqlCmd = $null
		if ($sqlModule -and $sqlModule -ne 'None')
		{
			try { Import-Module $sqlModule -ErrorAction Stop; $InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $sqlModule -ErrorAction Stop }
			catch { $InvokeSqlCmd = $null }
		}
		$supportsConnStr = $false
		if ($InvokeSqlCmd) { $supportsConnStr = $InvokeSqlCmd.Parameters.Keys -contains 'ConnectionString' }

		# Query
		$q = @"
SELECT IPNetwork, IPDevice, ScaleCode, Active
FROM dbo.TBS_SCL_ver520
ORDER BY
    CASE WHEN TRY_CAST(ScaleCode AS INT) IS NULL THEN 1 ELSE 0 END,
    TRY_CAST(ScaleCode AS INT),
    ScaleCode;
"@

		# Execute query (inline; no helpers)
		$rows = $null
		try
		{
			if ($InvokeSqlCmd)
			{
				if ($supportsConnStr) { $rows = Invoke-Sqlcmd -ConnectionString $connStr -Query $q -ErrorAction Stop }
				else { $rows = Invoke-Sqlcmd -ServerInstance $server -Database $dbName -Query $q -ErrorAction Stop }
			}
			else
			{
				$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
				$csb.PSBase.ConnectionString = $connStr
				$conn = New-Object System.Data.SqlClient.SqlConnection $csb.ConnectionString
				$conn.Open()
				$cmd = $conn.CreateCommand(); $cmd.CommandText = $q; $cmd.CommandTimeout = 30
				$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
				$dt = New-Object System.Data.DataTable
				[void]$da.Fill($dt)
				$conn.Close()
				$rows = $dt
			}
		}
		catch { $rows = $null }

		# Process rows
		$hasRows = $false
		if ($rows -is [System.Data.DataTable]) { if ($rows.Rows.Count -gt 0) { $hasRows = $true } }
		elseif ($rows) { if ($rows.Count -gt 0) { $hasRows = $true } }

		if ($hasRows)
		{
			Write_Log "Loaded rows from TBS_SCL_ver520 for first-scale detection." "cyan"

			# Normalize enumeration
			$enum = @()
			if ($rows -is [System.Data.DataTable]) { $enum = $rows.Rows }
			else { $enum = $rows }

			# Prefix frequency map (inline)
			$pfxCounts = @{ }
			foreach ($r in $enum)
			{
				$ipn = $null
				if ($null -ne $r.IPNetwork) { $ipn = [string]$r.IPNetwork }
				if ($ipn)
				{
					$ipn = $ipn.Trim()
					$pref = $null
					$m = [regex]::Match($ipn, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$')
					if ($m.Success) { $pref = "$($m.Groups[1].Value).$($m.Groups[2].Value).$($m.Groups[3].Value)" }
					else
					{
						$m4 = [regex]::Match($ipn, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\s*$')
						if ($m4.Success) { $pref = "$($m4.Groups[1].Value).$($m4.Groups[2].Value).$($m4.Groups[3].Value)" }
					}
					if ($pref)
					{
						if (-not $pfxCounts.ContainsKey($pref)) { $pfxCounts[$pref] = 0 }
						$pfxCounts[$pref] = $pfxCounts[$pref] + 1
					}
				}
			}

			$MostCommonPrefix = $null
			if ($pfxCounts.Count -gt 0)
			{
				$MostCommonPrefix = ($pfxCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
				Write_Log "Most common IPNetwork prefix in table: $MostCommonPrefix" "cyan"
			}

			# Filter "active" rows (inline)
			$candidates = @()
			foreach ($r in $enum)
			{
				$act = $null
				if ($null -ne $r.Active) { $act = ([string]$r.Active).Trim().ToUpper() }
				else { $act = '' }
				$isActive = $false
				if ($act -eq '') { $isActive = $true }
				elseif ($act -eq 'Y' -or $act -eq '1' -or $act -eq 'TRUE') { $isActive = $true }
				if ($isActive) { $candidates += $r }
			}

			# Try 1: exact per-row prefix + device
			foreach ($r in $candidates)
			{
				$pref = $null
				$ipn = $null
				if ($null -ne $r.IPNetwork) { $ipn = ([string]$r.IPNetwork).Trim() }
				if ($ipn)
				{
					$m = [regex]::Match($ipn, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$')
					if ($m.Success) { $pref = "$($m.Groups[1].Value).$($m.Groups[2].Value).$($m.Groups[3].Value)" }
					if (-not $pref)
					{
						$m4 = [regex]::Match($ipn, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\s*$')
						if ($m4.Success)
						{
							$pref = "$($m4.Groups[1].Value).$($m4.Groups[2].Value).$($m4.Groups[3].Value)"
							if (-not $r.IPDevice) { $r.IPDevice = $m4.Groups[4].Value }
						}
					}
				}

				$devOct = $null
				if ($null -ne $r.IPDevice)
				{
					$tmp = 0
					if ([int]::TryParse(([string]$r.IPDevice).Trim(), [ref]$tmp))
					{
						if ($tmp -ge 0 -and $tmp -le 255) { $devOct = $tmp }
					}
				}

				if ($pref -and ($null -ne $devOct)) { $firstScaleIP = "$pref.$devOct"; break }
			}

			# Try 2: table-most-common prefix + device
			if (-not $firstScaleIP -and $MostCommonPrefix)
			{
				foreach ($r in $candidates)
				{
					$devOct = $null
					if ($null -ne $r.IPDevice)
					{
						$tmp = 0
						if ([int]::TryParse(([string]$r.IPDevice).Trim(), [ref]$tmp))
						{
							if ($tmp -ge 0 -and $tmp -le 255) { $devOct = $tmp }
						}
					}
					if ($null -ne $devOct) { $firstScaleIP = "$MostCommonPrefix.$devOct"; break }
				}
			}

			# Try 3: active NIC /24 prefix + device
			if (-not $firstScaleIP -and $ActiveNicPrefix)
			{
				foreach ($r in $candidates)
				{
					$devOct = $null
					if ($null -ne $r.IPDevice)
					{
						$tmp = 0
						if ([int]::TryParse(([string]$r.IPDevice).Trim(), [ref]$tmp))
						{
							if ($tmp -ge 0 -and $tmp -le 255) { $devOct = $tmp }
						}
					}
					if ($null -ne $devOct) { $firstScaleIP = "$ActiveNicPrefix.$devOct"; break }
				}
			}

			if ($firstScaleIP) { Write_Log "First scale IP resolved: $firstScaleIP" "cyan" }
			else { Write_Log "First scale IP NOT resolved (need prefix + device)." "yellow" }
		}
		else
		{
			Write_Log "Could not read rows from TBS_SCL_ver520 (DB issue or empty table)." "yellow"
		}
	}
	else
	{
		Write_Log "DB connection info incomplete; first-scale detection skipped." "yellow"
	}

	# ------------------------------------------------------------------------------------------------
	# Locate target configs
	# ------------------------------------------------------------------------------------------------
	$targets = @()
	if ($ConfigPath)
	{
		if (Test-Path -LiteralPath $ConfigPath) { $targets += (Get-Item -LiteralPath $ConfigPath) }
		else { Write_Log "Config path not found: $ConfigPath" "red"; return }
	}
	else
	{
		foreach ($root in $ScaleCommRoots)
		{
			if (Test-Path -LiteralPath $root)
			{
				$found = Get-ChildItem -LiteralPath $root -File -Filter '*.exe.config' -ErrorAction SilentlyContinue
				if ($found) { $targets += $found }
			}
		}
	}
	$targets = $targets | Sort-Object FullName -Unique
	if (-not $targets -or $targets.Count -eq 0) { Write_Log "No *.exe.config found under ScaleCommApp roots." "red"; return }
	if (-not $AllMatches -and $targets.Count -gt 1)
	{
		Write_Log ("Multiple configs found; using first. Use -AllMatches to update all." + "`r`n - " + (($targets | ForEach-Object FullName) -join "`r`n - ")) "yellow"
		$targets = @($targets[0])
	}

	# ------------------------------------------------------------------------------------------------
	# ONE-TIME "Desired Settings" + "Existing (first file)" snapshot (inline getters)
	# ------------------------------------------------------------------------------------------------
	$previewPath = $targets[0].FullName
	$previewXml = $null
	try { [xml]$previewXml = Get-Content -LiteralPath $previewPath -Raw -ErrorAction Stop }
	catch { }

	$curStorePreview = $null; $curFirstIPPreview = $null; $curSqlInstPreview = $null; $curDbNamePreview = $null; $curBatchPreview = $null
	if ($previewXml)
	{
		$cfgPrev = $previewXml.configuration; if (-not $cfgPrev) { $cfgPrev = $previewXml.DocumentElement }
		if ($cfgPrev)
		{
			$asPrev = $cfgPrev.appSettings
			if ($asPrev)
			{
				$n = $asPrev.SelectSingleNode("add[@key='StoreName']"); if ($n) { $curStorePreview = [string]$n.GetAttribute('value') }
				$n = $asPrev.SelectSingleNode("add[@key='FistScaleIP']"); if ($n) { $curFirstIPPreview = [string]$n.GetAttribute('value') }
				$n = $asPrev.SelectSingleNode("add[@key='SQL_InstanceName']"); if ($n) { $curSqlInstPreview = [string]$n.GetAttribute('value') }
				$n = $asPrev.SelectSingleNode("add[@key='SQL_databaseName']"); if ($n) { $curDbNamePreview = [string]$n.GetAttribute('value') }
				$n = $asPrev.SelectSingleNode("add[@key='Recs_BatchSendFull']"); if ($n) { $curBatchPreview = [string]$n.GetAttribute('value') }
			}
		}
	}

	$dStore = '<unchanged/skip>'; if ($ExpectedStoreName) { $dStore = $ExpectedStoreName }
	$dIP = '<skip>'; if ($firstScaleIP) { $dIP = $firstScaleIP }
	$dInst = '<unchanged/skip>'; if ($ExpectedSqlInstance) { $dInst = $ExpectedSqlInstance }
	$dDb = '<unchanged/skip>'; if ($ExpectedDatabase) { $dDb = $ExpectedDatabase }

	Write_Log "Desired Settings:" "blue"
	Write_Log ("  StoreName          : {0}" -f $dStore) "green"
	Write_Log ("  FistScaleIP        : {0}" -f $dIP) "green"
	Write_Log ("  SQL_InstanceName   : {0}" -f $dInst) "green"
	Write_Log ("  SQL_databaseName   : {0}" -f $dDb) "green"
	Write_Log ("  Recs_BatchSendFull : 10000") "green"

	$leafPreview = Split-Path -Leaf $previewPath
	Write_Log ("Existing (first file: {0}):" -f $leafPreview) "blue"
	if (-not $curStorePreview) { $curStorePreview = '<missing>' }
	if (-not $curFirstIPPreview) { $curFirstIPPreview = '<missing>' }
	if (-not $curSqlInstPreview) { $curSqlInstPreview = '<missing>' }
	if (-not $curDbNamePreview) { $curDbNamePreview = '<missing>' }
	if (-not $curBatchPreview) { $curBatchPreview = '<missing>' }
	Write_Log ("  StoreName          : {0}" -f $curStorePreview) "yellow"
	Write_Log ("  FistScaleIP        : {0}" -f $curFirstIPPreview) "yellow"
	Write_Log ("  SQL_InstanceName   : {0}" -f $curSqlInstPreview) "yellow"
	Write_Log ("  SQL_databaseName   : {0}" -f $curDbNamePreview) "yellow"
	Write_Log ("  Recs_BatchSendFull : {0}" -f $curBatchPreview) "yellow"

	# ------------------------------------------------------------------------------------------------
	# Apply to all targets (quiet per-file; track counts + saved list)
	# ------------------------------------------------------------------------------------------------
	$autoFix = (-not $NoAutoFix.IsPresent)
	[int]$audited = 0
	[int]$updated = 0
	[int]$skipped = 0
	$savedFiles = New-Object System.Collections.Generic.List[string]

	foreach ($t in $targets)
	{
		$audited++
		$path = $t.FullName

		# Load XML
		$xml = $null
		try { [xml]$xml = Get-Content -LiteralPath $path -Raw -ErrorAction Stop }
		catch { $skipped++; continue }
		$cfg = $xml.configuration; if (-not $cfg) { $cfg = $xml.DocumentElement }
		if (-not $cfg) { $skipped++; continue }
		$as = $cfg.appSettings
		if (-not $as) { $skipped++; continue }

		# READ current values (inline)
		$curStore = $null; $curFirstIP = $null; $curInst = $null; $curDb = $null; $curBatch = $null
		$n = $as.SelectSingleNode("add[@key='StoreName']"); if ($n) { $curStore = [string]$n.GetAttribute('value') }
		$n = $as.SelectSingleNode("add[@key='FistScaleIP']"); if ($n) { $curFirstIP = [string]$n.GetAttribute('value') }
		$n = $as.SelectSingleNode("add[@key='SQL_InstanceName']"); if ($n) { $curInst = [string]$n.GetAttribute('value') }
		$n = $as.SelectSingleNode("add[@key='SQL_databaseName']"); if ($n) { $curDb = [string]$n.GetAttribute('value') }
		$n = $as.SelectSingleNode("add[@key='Recs_BatchSendFull']"); if ($n) { $curBatch = [string]$n.GetAttribute('value') }

		$changedHere = $false

		# SET (inline create-if-missing)
		if ($ExpectedStoreName -and $curStore -ne $ExpectedStoreName)
		{
			if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set StoreName'))
			{
				$n = $as.SelectSingleNode("add[@key='StoreName']")
				if (-not $n) { $n = $xml.CreateElement('add'); $null = $n.SetAttribute('key', 'StoreName'); $null = $as.AppendChild($n) }
				$null = $n.SetAttribute('value', [string]$ExpectedStoreName)
				$changedHere = $true
			}
		}

		if ($firstScaleIP -and $curFirstIP -ne $firstScaleIP)
		{
			if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set FistScaleIP'))
			{
				$n = $as.SelectSingleNode("add[@key='FistScaleIP']")
				if (-not $n) { $n = $xml.CreateElement('add'); $null = $n.SetAttribute('key', 'FistScaleIP'); $null = $as.AppendChild($n) }
				$null = $n.SetAttribute('value', [string]$firstScaleIP)
				$changedHere = $true
			}
		}

		if ($ExpectedSqlInstance -and $curInst -ne $ExpectedSqlInstance)
		{
			if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set SQL_InstanceName'))
			{
				$n = $as.SelectSingleNode("add[@key='SQL_InstanceName']")
				if (-not $n) { $n = $xml.CreateElement('add'); $null = $n.SetAttribute('key', 'SQL_InstanceName'); $null = $as.AppendChild($n) }
				$null = $n.SetAttribute('value', [string]$ExpectedSqlInstance)
				$changedHere = $true
			}
		}

		if ($ExpectedDatabase -and $curDb -ne $ExpectedDatabase)
		{
			if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set SQL_databaseName'))
			{
				$n = $as.SelectSingleNode("add[@key='SQL_databaseName']")
				if (-not $n) { $n = $xml.CreateElement('add'); $null = $n.SetAttribute('key', 'SQL_databaseName'); $null = $as.AppendChild($n) }
				$null = $n.SetAttribute('value', [string]$ExpectedDatabase)
				$changedHere = $true
			}
		}

		if ($curBatch -ne '10000')
		{
			if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set Recs_BatchSendFull=10000'))
			{
				$n = $as.SelectSingleNode("add[@key='Recs_BatchSendFull']")
				if (-not $n) { $n = $xml.CreateElement('add'); $null = $n.SetAttribute('key', 'Recs_BatchSendFull'); $null = $as.AppendChild($n) }
				$null = $n.SetAttribute('value', '10000')
				$changedHere = $true
			}
		}

		# Save if changed
		if ($changedHere -and $autoFix)
		{
			try
			{
				$item = Get-Item -LiteralPath $path -ErrorAction Stop
				if ($item.Attributes -band [IO.FileAttributes]::ReadOnly)
				{
					if ($Force) { $item.Attributes = ($item.Attributes -bxor [IO.FileAttributes]::ReadOnly) }
					else { $skipped++; continue }
				}
			}
			catch { $skipped++; continue }

			try
			{
				$dir = Split-Path -LiteralPath $path -Parent
				$leaf = Split-Path -LiteralPath $path -Leaf
				$stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
				$bak = Join-Path $dir "$leaf.$stamp.bak"
				Copy-Item -LiteralPath $path -Destination $bak -Force
			}
			catch { }

			try
			{
				$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
				$sw = New-Object System.IO.StreamWriter($path, $false, $utf8NoBom)
				try { $xml.Save($sw) }
				finally { $sw.Close() }
				$updated++; $savedFiles.Add((Split-Path -Leaf $path)) | Out-Null
			}
			catch { $skipped++ }
		}
	}

	# ==================================================================================================
	# NEW: Restore DEPLOY_CHG.sql deploy line (integrated Fix_Deploy_CHG)
	# ==================================================================================================
	if (-not $OfficePath -and $script:OfficePath) { $OfficePath = $script:OfficePath }
	if (-not $OfficePath)
	{
		Write_Log "Office path not provided and \$script:OfficePath not available; skipping DEPLOY_CHG.sql restore." "yellow"
	}
	elseif ($PSCmdlet.ShouldProcess((Join-Path $OfficePath 'DEPLOY_CHG.sql'), 'Restore DEPLOY_CHG @EXEC line'))
	{
		[void](Fix_Deploy_CHG -OfficePath $OfficePath -Quiet)
	}
	# ------------------------------------------------------------------------------------------------
	# Final summary
	# ------------------------------------------------------------------------------------------------
	$sumIP = 'n/a'
	if ($firstScaleIP) { $sumIP = $firstScaleIP }

	Write_Log ("`r`nTroubleshoot_ScaleCommApp: Audited={0}, Updated={1}, Skipped={2}, FirstScaleIP={3}" -f $audited, $updated, $skipped, $sumIP) "blue"
	if ($updated -gt 0) { Write_Log ("Saved files: {0}" -f ($savedFiles -join ', ')) "green" }

	Write_Log "`r`n==================== Troubleshoot_ScaleCommApp Completed ====================" "blue"
}

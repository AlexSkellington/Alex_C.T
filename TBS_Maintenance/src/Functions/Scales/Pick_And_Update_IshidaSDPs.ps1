# ===================================================================================================
#                              FUNCTION: Pick_And_Update_IshidaSDPs
# ---------------------------------------------------------------------------------------------------
# Description:
#   One-shot picker + writer for the IshidaSDPs config entry based on SDP_TAB (F04 = SDP
#   number, F1022 = name): a GUI to pick which SDPs go into the config, pre-checking those
#   already present, then saves <add key="IshidaSDPs" value="..."> as a comma-separated list
#   of numbers, or "Null" when none are selected.
#
# Parameters:
#   - ConfigPath      [string]: Ishida config file (resolved from roots if empty)
#   - ScaleCommRoots  [string[]]: config roots to search (default $script:ScaleCommRoots)
#   - AllMatches      [switch]: accept all matching config files
#   - IncludeInactive [switch]: include inactive SDPs in the picker
#   - Force           [switch]: reserved for future use
#
# Variables:
#   - $script:FunctionResults: DBNAME, DBSERVER, ConnectionString, SqlModuleName
#                              (falls back to Get_Store_And_Database_Info)
#
# Notes:
#   - Builds a ConnectionString when missing (SQL Auth preferred, else Integrated).
#   - Saves UTF-8 without BOM and writes a timestamped .bak next to the config.
# ===================================================================================================

function Pick_And_Update_IshidaSDPs
{
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
		[string]$ConfigPath,
		[string[]]$ScaleCommRoots = $script:ScaleCommRoots,
		[switch]$AllMatches,
		[switch]$IncludeInactive,
		# reserved for future use
		[switch]$Force
	)
	
	# ---------------- Logging shim (Write_Log if present, else host) ----------------
	$logInfo = {
		param ($m) if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log $m 'cyan' }
		else { Write-Host "[INFO] $m" -ForegroundColor Cyan }
	}
	$logWarn = {
		param ($m) if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log $m 'yellow' }
		else { Write-Host "[WARN] $m" -ForegroundColor Yellow }
	}
	$logErr = {
		param ($m) if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log $m 'red' }
		else { Write-Host "[ERR ] $m" -ForegroundColor Red }
	}
	$logOk = {
		param ($m) if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log $m 'green' }
		else { Write-Host "[ OK ] $m" -ForegroundColor Green }
	}
	$logBlue = {
		param ($m) if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log $m 'blue' }
		else { Write-Host $m -ForegroundColor Blue }
	}
	
	& $logBlue "`r`n==================== Starting Pick_And_Update_IshidaSDPs ====================`r`n"
	
	# ---------------- Resolve DB info from your cache ----------------
	$dbName = $script:FunctionResults['DBNAME']
	$server = $script:FunctionResults['DBSERVER']
	$ConnectionString = $script:FunctionResults['ConnectionString']
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	
	if (-not $dbName -or -not $server -or -not $ConnectionString)
	{
		if (Get-Command Get_Store_And_Database_Info -ErrorAction SilentlyContinue)
		{
			& $logWarn "DB cache incomplete; calling Get_Store_And_Database_Info..."
			try { Get_Store_And_Database_Info | Out-Null }
			catch { & $logErr "Failed to resolve DB info: $($_.Exception.Message)" }
			$dbName = $script:FunctionResults['DBNAME']
			$server = $script:FunctionResults['DBSERVER']
			$ConnectionString = $script:FunctionResults['ConnectionString']
			$SqlModuleName = $script:FunctionResults['SqlModuleName']
		}
	}
	
	# ---------------- Build a ConnectionString when missing ----------------
	if (-not $ConnectionString -and $server -and $dbName)
	{
		$sqlUser = $script:FunctionResults['SQL_UserID']
		$sqlPwd = $script:FunctionResults['SQL_UserPwd']
		if ($sqlUser -and $sqlPwd)
		{
			$ConnectionString = "Server=$server;Database=$dbName;User ID=$sqlUser;Password=$sqlPwd;TrustServerCertificate=True;"
			& $logInfo "Built SQL Authentication connection string from cached credentials."
		}
		else
		{
			$ConnectionString = "Server=$server;Database=$dbName;Integrated Security=SSPI;TrustServerCertificate=True;"
			& $logInfo "Built Integrated Security connection string (no SQL creds found)."
		}
		$script:FunctionResults['ConnectionString'] = $ConnectionString
	}
	
	if (-not $server -or -not $dbName)
	{
		& $logErr "DB server or DB name unresolved. Cannot read SDP_TAB."
		& $logBlue "`r`n==================== Pick_And_Update_IshidaSDPs Completed ====================`r`n"
		return
	}
	
	# ---------------- Prepare SQL execution (Invoke-Sqlcmd or .NET) ----------------
	$InvokeSqlCmd = $null
	if ($SqlModuleName -and $SqlModuleName -ne 'None')
	{
		try
		{
			Import-Module $SqlModuleName -ErrorAction Stop
			$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
		}
		catch
		{
			& $logWarn "SQL module '$SqlModuleName' could not be loaded. Will use .NET SqlClient fallback."
		}
	}
	else
	{
		& $logWarn "No SQL module name provided. Will use .NET SqlClient fallback."
	}
	$supportsConnectionString = $false
	if ($InvokeSqlCmd) { $supportsConnectionString = $InvokeSqlCmd.Parameters.Keys -contains 'ConnectionString' }
	
	# ---------------- Read SDP_TAB (F04 number, F1022 name) ----------------
	$sql = @"
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='SDP_TAB')
    RAISERROR('SDP_TAB not found',16,1);

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SDP_TAB' AND COLUMN_NAME='F04')
    RAISERROR('Column F04 not found in SDP_TAB',16,1);

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SDP_TAB' AND COLUMN_NAME='F1022')
    RAISERROR('Column F1022 not found in SDP_TAB',16,1);

SELECT
    CAST([F04]   AS int)            AS SdpNumber,
    CAST([F1022] AS nvarchar(200))  AS SdpName
FROM SDP_TAB
ORDER BY CAST([F04] AS int);
"@
	
	try
	{
		if ($InvokeSqlCmd)
		{
			if ($supportsConnectionString -and $ConnectionString)
			{
				$rows = & $InvokeSqlCmd -ConnectionString $ConnectionString -Query $sql -QueryTimeout 60 -ErrorAction Stop
			}
			else
			{
				$rows = & $InvokeSqlCmd -ServerInstance $server -Database $dbName -Query $sql -QueryTimeout 60 -ErrorAction Stop
			}
		}
		else
		{
			$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
			if ($ConnectionString) { $csb.ConnectionString = $ConnectionString }
			else { $csb.DataSource = $server; $csb.InitialCatalog = $dbName; $csb.IntegratedSecurity = $true }
			$conn = New-Object System.Data.SqlClient.SqlConnection $csb.ConnectionString
			$cmd = $conn.CreateCommand(); $cmd.CommandText = $sql
			$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
			$dt = New-Object System.Data.DataTable
			$conn.Open(); [void]$da.Fill($dt); $conn.Close()
			$rows = $dt
		}
	}
	catch
	{
		& $logErr "Failed to read SDP_TAB (F04/F1022): $($_.Exception.Message)"
		& $logBlue "`r`n==================== Pick_And_Update_IshidaSDPs Completed ====================`r`n"
		return
	}
	
	if (-not $rows -or $rows.Rows.Count -eq 0)
	{
		& $logWarn "SDP_TAB returned 0 rows."
		& $logBlue "`r`n==================== Pick_And_Update_IshidaSDPs Completed ====================`r`n"
		return
	}
	
	# Normalize for UI
	$sdpList = @()
	foreach ($r in $rows)
	{
		$sdpList += [pscustomobject]@{
			SdpNumber = [int]$r.SdpNumber
			SdpName   = [string]$r.SdpName
		}
	}
	& $logInfo ("Loaded SDPs: " + ($sdpList.Count))
	
	# ---------------- Find target config files ----------------
	$targets = @()
	if ($ConfigPath)
	{
		if (Test-Path -LiteralPath $ConfigPath) { $targets += (Get-Item -LiteralPath $ConfigPath) }
		else { & $logErr "Config path not found: $ConfigPath" }
	}
	else
	{
		foreach ($root in $ScaleCommRoots)
		{
			if (-not (Test-Path -LiteralPath $root)) { continue }
			$found = Get-ChildItem -LiteralPath $root -File -Filter '*.exe.config' -ErrorAction SilentlyContinue
			if ($found) { $targets += $found }
		}
	}
	$targets = $targets | Sort-Object FullName -Unique
	if (-not $AllMatches -and $targets.Count -gt 1)
	{
		& $logWarn ("Multiple configs found; using first. Use -AllMatches to update all." +
			"`n - " + (($targets | ForEach-Object FullName) -join "`n - "))
		$targets = @($targets[0])
	}
	if (-not $targets -or $targets.Count -eq 0)
	{
		& $logErr "No target *.exe.config found."
		& $logBlue "`r`n==================== Pick_And_Update_IshidaSDPs Completed ====================`r`n"
		return
	}
	
	# ---------------- Read IshidaSDPs from first config (for pre-checks) ----------------
	$firstConfig = $targets[0].FullName
	[xml]$xmlFirst = Get-Content -LiteralPath $firstConfig -Raw
	$configuration = $xmlFirst.configuration; if (-not $configuration) { $configuration = $xmlFirst.DocumentElement }
	if (-not $configuration) { & $logErr "Invalid config (no <configuration>): $firstConfig"; return }
	$appSettings = $configuration.appSettings
	if (-not $appSettings)
	{
		$appSettings = $xmlFirst.CreateElement('appSettings')
		[void]$configuration.AppendChild($appSettings)
	}
	$keyName = 'IshidaSDPs'
	$node = $appSettings.SelectSingleNode("add[@key='$keyName']")
	$existingRaw = ''
	$existingNums = @()
	if ($node)
	{
		$existingRaw = [string]$node.GetAttribute('value')
		foreach ($part in ($existingRaw -split ',')) { $p = $part.Trim(); if ($p -match '^\s*(\d+)') { $existingNums += [int]$Matches[1] } }
	}
	& $logInfo ("Current IshidaSDPs in first config: " + ($existingRaw -replace '\s+', ' '))
	
	# ---------- build the "top bar" title text with the already-checked list ----------
	# CHANGE: Put already-in-config SDPs on the form title (top bar). Truncate if too long.
	$existingListText = ($existingNums | Sort-Object -Unique) -join ','
	if ([string]::IsNullOrWhiteSpace($existingListText)) { $existingListText = 'None' }
	$existingShort = $existingListText
	if ($existingShort.Length -gt 90) { $existingShort = $existingShort.Substring(0, 90) + '... (' + $existingNums.Count + ' total)' } # CHANGE
	
	# ---------------- Build WinForms picker ----------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select Ishida SDPs  |  Already in config: $existingShort" # CHANGE: show on top bar
	$form.StartPosition = 'CenterScreen'
	$form.Size = New-Object System.Drawing.Size(640, 640)
	$form.MinimizeBox = $false
	$form.MaximizeBox = $false
	$form.TopMost = $true
	
	$search = New-Object System.Windows.Forms.TextBox
	$search.Location = New-Object System.Drawing.Point(10, 10)
	$search.Size = New-Object System.Drawing.Size(505, 24)
	$form.Controls.Add($search)
	
	$btnAll = New-Object System.Windows.Forms.Button
	$btnAll.Text = "Select All"
	$btnAll.Location = New-Object System.Drawing.Point(520, 10)
	$btnAll.Size = New-Object System.Drawing.Size(100, 28)
	$form.Controls.Add($btnAll)
	
	$btnNone = New-Object System.Windows.Forms.Button
	$btnNone.Text = "Clear"
	$btnNone.Location = New-Object System.Drawing.Point(520, 44)
	$btnNone.Size = New-Object System.Drawing.Size(100, 28)
	$form.Controls.Add($btnNone)
	
	$list = New-Object System.Windows.Forms.CheckedListBox
	$list.Location = New-Object System.Drawing.Point(10, 44)
	$list.Size = New-Object System.Drawing.Size(505, 520)
	$list.CheckOnClick = $true
	$form.Controls.Add($list)
	
	$legend = New-Object System.Windows.Forms.Label
	$legend.AutoSize = $true
	$legend.Location = New-Object System.Drawing.Point(10, 570)
	$legend.Text = "Checked = selected. Pre-checked = already in config. Click Save to write."
	$form.Controls.Add($legend)
	
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "Save"
	$btnOK.Location = New-Object System.Drawing.Point(520, 536)
	$btnOK.Size = New-Object System.Drawing.Size(100, 28)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(520, 570)
	$btnCancel.Size = New-Object System.Drawing.Size(100, 28)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)
	
	$allEntries = @()
	foreach ($s in $sdpList)
	{
		$display = "{0} - {1}" -f $s.SdpNumber, $s.SdpName
		$null = $list.Items.Add($display)
		$allEntries += [pscustomobject]@{ Display = $display; Number = $s.SdpNumber }
		if ($existingNums -contains [int]$s.SdpNumber) { $list.SetItemChecked(($list.Items.Count - 1), $true) }
	}
	
	$search.Add_TextChanged({
			$needle = $search.Text.Trim()
			$list.Items.Clear()
			if ([string]::IsNullOrWhiteSpace($needle)) { foreach ($e in $allEntries) { [void]$list.Items.Add($e.Display) } }
			else { foreach ($e in $allEntries) { if ($e.Display -like "*$needle*") { [void]$list.Items.Add($e.Display) } } }
			for ($i = 0; $i -lt $list.Items.Count; $i++)
			{
				$text = [string]$list.Items[$i]
				if ($text -match '^\s*(\d+)') { if ($existingNums -contains [int]$Matches[1]) { $list.SetItemChecked($i, $true) } }
			}
		})
	
	$btnAll.Add_Click({ for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $true) } })
	$btnNone.Add_Click({ for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $false) } })
	
	$dialogResult = $form.ShowDialog()
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		& $logWarn "User cancelled. No changes saved."
		& $logBlue "`r`n==================== Pick_And_Update_IshidaSDPs Completed ====================`r`n"
		return
	}
	
	$selectedNums = @()
	foreach ($item in $list.CheckedItems)
	{
		$text = [string]$item
		if ($text -match '^\s*(\d+)') { $selectedNums += [int]$Matches[1] }
	}
	$selectedNums = $selectedNums | Sort-Object -Unique
	
	# ---------- write "Null" when nothing is selected ----------
	# CHANGE: If user leaves everything unchecked, write "Null" to the config value.
	$valueToWrite = if ($selectedNums.Count -gt 0) { ($selectedNums -join ',') }
	else { 'Null' }
	
	& $logInfo ("User selected SDPs -> " + $valueToWrite)
	
	# ---------------- Write back to each target config ----------------
	foreach ($t in $targets)
	{
		$path = $t.FullName
		if (-not $PSCmdlet.ShouldProcess($path, "Set appSettings 'IshidaSDPs'")) { continue }
		
		try
		{
			$item = Get-Item -LiteralPath $path -ErrorAction Stop
			if ($item.Attributes -band [IO.FileAttributes]::ReadOnly)
			{
				if ($Force) { & $logWarn "Removing ReadOnly attribute: $path"; $item.Attributes = ($item.Attributes -bxor [IO.FileAttributes]::ReadOnly) }
				else { & $logErr "File is read-only. Re-run with -Force to clear attribute."; continue }
			}
		}
		catch { & $logErr "Cannot access file attributes for ${path}: $($_.Exception.Message)"; continue }
		
		try
		{
			$dir = Split-Path -LiteralPath $path -Parent
			$base = Split-Path -LiteralPath $path -Leaf
			$stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
			$bak = Join-Path $dir "$base.$stamp.bak"
			Copy-Item -LiteralPath $path -Destination $bak -Force
			& $logOk "Backup created: $bak"
		}
		catch { & $logWarn "Backup failed for ${path}: $($_.Exception.Message)" }
		
		try
		{
			[xml]$xml = Get-Content -LiteralPath $path -Raw
			$cfg = $xml.configuration; if (-not $cfg) { $cfg = $xml.DocumentElement }
			if (-not $cfg) { & $logErr "Invalid config: no <configuration> root in $path"; continue }
			
			$as = $cfg.appSettings
			if (-not $as) { $as = $xml.CreateElement('appSettings'); [void]$cfg.AppendChild($as); & $logWarn "Created <appSettings> section in $path" }
			
			$key = 'IshidaSDPs'
			$node = $as.SelectSingleNode("add[@key='$key']")
			if ($node)
			{
				$old = $node.GetAttribute('value')
				if ($old -ne $valueToWrite) { $node.SetAttribute('value', $valueToWrite); & $logOk "Updated ${key}: '$old' -> '$valueToWrite'" }
				else { & $logInfo "No change needed in $path ($key already '$valueToWrite')." }
			}
			else
			{
				$add = $xml.CreateElement('add')
				[void]$add.SetAttribute('key', $key)
				[void]$add.SetAttribute('value', $valueToWrite)
				[void]$as.AppendChild($add)
				& $logOk "Added $key with value '$valueToWrite'"
			}
			
			$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
			$sw = New-Object System.IO.StreamWriter($path, $false, $utf8NoBom)
			$xml.Save($sw); $sw.Close()
			& $logOk "Saved config: $path"
		}
		catch
		{
			& $logErr "XML update failed for ${path}: $($_.Exception.Message)"
		}
	}
	
	& $logBlue "`r`n==================== Pick_And_Update_IshidaSDPs Completed ====================`r`n"
}

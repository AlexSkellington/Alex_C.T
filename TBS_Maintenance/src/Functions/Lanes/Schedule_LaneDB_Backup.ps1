# ===================================================================================================
#                              FUNCTION: Schedule_LaneDB_Backup
# ---------------------------------------------------------------------------------------------------
# Description:
#   Lane DB backup wizard: pick lanes (only those with a valid TCP/Named Pipes protocol),
#   prompt for schedule details, then per lane create Run_<machine>DB_Backup.ps1 under the
#   per-machine scripts folder and schedule a task that backs up to
#   $BackupRoot\<machine>\STORESQL_<timestamp>.bak using the detected protocol.
#
# Parameters:
#   (none)
#
# Variables:
#   - $script:FunctionResults: StoreNumber, node mappings
#   - $script:LaneProtocols: per-lane protocol (TCP / Named Pipes / File)
#   - $ScriptsFolder / $BackupRoot: per-machine script and backup locations
# ===================================================================================================

function Schedule_LaneDB_Backup
{
	Write_Log "`r`n==================== Starting Schedule_LaneDB_Backup ====================`r`n" "blue"

	# 1. Node selection (lanes only)
	$StoreNumber = $script:FunctionResults['StoreNumber']
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select Lanes for DB Backup"
	if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected for backup." "yellow"
		return
	}

	# 2. Check ProtocolResults for connectivity
	$goodLanes = @()
	foreach ($lane in $selection.Lanes)
	{
		$protocol = if ($script:LaneProtocols) { $script:LaneProtocols[$lane] }
		if ($protocol -in @("TCP", "Named Pipes")) { $goodLanes += $lane }
		else { Write_Log "Lane $lane not available in ProtocolResults or protocol not valid. Skipping." "yellow" }
	}
	if (-not $goodLanes -or $goodLanes.Count -eq 0)
	{
		Write_Log "No selected lanes are available for DB backup." "red"
		return
	}

	# 3. Prompt for backup options, user and password (all in one form, more spacing)
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object Windows.Forms.Form
	$form.Text = "Configure Lane DB Backup Scheduler"
	$form.Size = [Drawing.Size]::new(500, 425) # Larger form
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = 'FixedDialog'
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false

	# -- Backup Settings GroupBox --
	$grpBackup = New-Object Windows.Forms.GroupBox
	$grpBackup.Text = "Backup Settings"
	$grpBackup.Location = [Drawing.Point]::new(20, 15)
	$grpBackup.Size = [Drawing.Size]::new(450, 180)
	$form.Controls.Add($grpBackup)

	# Time
	$lblTime = New-Object Windows.Forms.Label
	$lblTime.Text = "Time to run backup (24h, HH:mm):"
	$lblTime.Location = [Drawing.Point]::new(25, 35)
	$lblTime.Size = [Drawing.Size]::new(200, 20)
	$grpBackup.Controls.Add($lblTime)
	$txtTime = New-Object Windows.Forms.MaskedTextBox
	$txtTime.Mask = "00:00"
	$txtTime.Text = "01:00"
	$txtTime.Location = [Drawing.Point]::new(260, 33)
	$txtTime.Size = [Drawing.Size]::new(80, 24)
	$grpBackup.Controls.Add($txtTime)

	# Frequency
	$lblFreq = New-Object Windows.Forms.Label
	$lblFreq.Text = "Frequency (every X days):"
	$lblFreq.Location = [Drawing.Point]::new(25, 70)
	$lblFreq.Size = [Drawing.Size]::new(200, 20)
	$grpBackup.Controls.Add($lblFreq)
	$numFreq = New-Object Windows.Forms.NumericUpDown
	$numFreq.Minimum = 1
	$numFreq.Maximum = 31
	$numFreq.Value = 1
	$numFreq.Location = [Drawing.Point]::new(260, 68)
	$numFreq.Size = [Drawing.Size]::new(80, 24)
	$grpBackup.Controls.Add($numFreq)

	# Retention
	$lblKeep = New-Object Windows.Forms.Label
	$lblKeep.Text = "How many backups to keep:"
	$lblKeep.Location = [Drawing.Point]::new(25, 105)
	$lblKeep.Size = [Drawing.Size]::new(200, 20)
	$grpBackup.Controls.Add($lblKeep)
	$numKeep = New-Object Windows.Forms.NumericUpDown
	$numKeep.Minimum = 1
	$numKeep.Maximum = 99
	$numKeep.Value = 1
	$numKeep.Location = [Drawing.Point]::new(260, 103)
	$numKeep.Size = [Drawing.Size]::new(80, 24)
	$grpBackup.Controls.Add($numKeep)

	# Backup folder (local copy) - picker opens at the default; accept it or switch
	$lblDir = New-Object Windows.Forms.Label
	$lblDir.Text = "Backup folder (local copy):"
	$lblDir.Location = [Drawing.Point]::new(25, 140)
	$lblDir.Size = [Drawing.Size]::new(145, 18)
	$grpBackup.Controls.Add($lblDir)
	$txtDir = New-Object Windows.Forms.TextBox
	$txtDir.ReadOnly = $true
	$txtDir.Text = $script:BackupRoot.TrimEnd('\')
	$txtDir.Location = [Drawing.Point]::new(175, 138)
	$txtDir.Size = [Drawing.Size]::new(165, 22)
	$grpBackup.Controls.Add($txtDir)
	$btnBrowse = New-Object Windows.Forms.Button
	$btnBrowse.Text = "Browse..."
	$btnBrowse.Location = [Drawing.Point]::new(346, 136)
	$btnBrowse.Size = [Drawing.Size]::new(78, 24)
	$btnBrowse.Add_Click({
		$p = Select-FolderOrDefault -initialPath $txtDir.Text -description "Choose where to keep the LOCAL copy of the lane DB backup (a <machine> subfolder is created inside it)."
		if ($p) { $txtDir.Text = $p }
	})
	$grpBackup.Controls.Add($btnBrowse)

	# -- Task User/Password GroupBox --
	$grpUser = New-Object Windows.Forms.GroupBox
	$grpUser.Text = "Scheduled Task Credentials"
	$grpUser.Location = [Drawing.Point]::new(20, 205)
	$grpUser.Size = [Drawing.Size]::new(450, 100)
	$form.Controls.Add($grpUser)

	# Username (default Administrator)
	$lblUser = New-Object Windows.Forms.Label
	$lblUser.Text = "Task user (default: Administrator):"
	$lblUser.Location = [Drawing.Point]::new(25, 35)
	$lblUser.Size = [Drawing.Size]::new(200, 18)
	$grpUser.Controls.Add($lblUser)
	$txtUser = New-Object Windows.Forms.TextBox
	$txtUser.Text = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
	$txtUser.Location = [Drawing.Point]::new(260, 33)
	$txtUser.Size = [Drawing.Size]::new(130, 24)
	$grpUser.Controls.Add($txtUser)

	# Password (masked)
	$lblPwd = New-Object Windows.Forms.Label
	$lblPwd.Text = "Task password (blank = only when user logged in):"
	$lblPwd.Location = [Drawing.Point]::new(25, 65)
	$lblPwd.Size = [Drawing.Size]::new(270, 18)
	$grpUser.Controls.Add($lblPwd)
	$txtPwd = New-Object Windows.Forms.TextBox
	$txtPwd.Text = ""
	$txtPwd.Location = [Drawing.Point]::new(305, 63)
	$txtPwd.Size = [Drawing.Size]::new(120, 24)
	$txtPwd.UseSystemPasswordChar = $true
	$grpUser.Controls.Add($txtPwd)

	# OK/Cancel
	$btnOK = New-Object Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = [Drawing.Point]::new(120, 325)
	$btnOK.Size = [Drawing.Size]::new(110, 38)
	$btnOK.Font = [System.Drawing.Font]::new("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)

	$btnCancel = New-Object Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = [Drawing.Point]::new(260, 325)
	$btnCancel.Size = [Drawing.Size]::new(110, 38)
	$btnCancel.Font = [System.Drawing.Font]::new("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)

	# Show dialog and handle input
	$result = $form.ShowDialog()
	if ($result -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "Lane backup scheduling cancelled by user." "yellow"
		return
	}
	# Local-copy location: the folder the user picked (default if unchanged)
	$chosenParent = $txtDir.Text.Trim()
	if (-not $chosenParent) { $chosenParent = $script:BackupRoot.TrimEnd('\') }
	Write_Log "Lane DB local-copy destination root: $chosenParent" "gray"
	$timeInput = $txtTime.Text
	$freqInput = [int]$numFreq.Value
	$keepInput = [int]$numKeep.Value
	$taskUser = $txtUser.Text.Trim()
	if (-not $taskUser) { $taskUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name }
	$taskPwd = $txtPwd.Text

	# Validate time
	if ($timeInput -notmatch '^\d{2}:\d{2}$')
	{
		[void][System.Windows.Forms.MessageBox]::Show("Time must be in HH:mm (24h) format.", "Input Error", 0, 16)
		Write_Log "Invalid time format entered." "red"
		return
	}
	$hours, $minutes = $timeInput.Split(":")
	if ([int]$hours -gt 23 -or [int]$minutes -gt 59)
	{
		[void][System.Windows.Forms.MessageBox]::Show("Invalid time entered.", "Input Error", 0, 16)
		Write_Log "Invalid time (out of range)." "red"
		return
	}

	# 4. Loop over each lane and schedule backup
	foreach ($lane in $goodLanes)
	{
		# Memory first, then the cache file, then the lane itself
		$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $lane
		if (-not $laneInfo)
		{
			Write_Log "Could not get DB info for lane $lane. Skipping." "yellow"
			continue
		}
		$machine = $laneInfo['MachineName']
		$dbName = $laneInfo['DBName']

		# sqlcmd -S string for the lane's protocol (only TCP / Named Pipes lanes got this far)
		$sqlcmdS = if ($script:LaneProtocols[$lane] -eq 'Named Pipes') { $laneInfo['NamedPipes'] } else { $laneInfo['TcpServer'] }

		# Set backup and script directories (per-lane); local-copy root from the user's choice
		$scriptsDir = $script:ScriptsFolder
		$machineFolder = Join-Path $chosenParent $machine

		# ----> KEY: name backup script exactly as Run_${MachineName}_DB_Backup.ps1
		$scriptName = "Run_${machine}_DB_Backup.ps1"
		$backupScriptPath = Join-Path $scriptsDir $scriptName

		# Compose the backup script for this lane (matches server style)
		$backupScript = @"
# Auto-generated by Schedule_${machine}_DB_Backup (Alex_C.T)
`$ErrorActionPreference = 'Stop'
`$BackupRoot = `"$machineFolder`"
`$Database = `"$dbName`"
`$SqlcmdS = `"$sqlcmdS`"
`$RemoteBackupFolder = "\\$machine\C$\Tecnica_Systems\Alex_C.T\Backups\$machine"
`$MaxBackups = $keepInput
`$MaxRemoteBackups = 3

# Compose backup file name
`$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
`$bakName = ("{0}_{1}.bak" -f `$Database, `$timestamp)
`$remoteBakFile = Join-Path `$RemoteBackupFolder `$bakName
`$localBakFile  = Join-Path `$BackupRoot `$bakName

# Ensure backup root exists (server)
if (-not (Test-Path `$BackupRoot)) { New-Item -Path `$BackupRoot -ItemType Directory -Force | Out-Null }
# Ensure backup root exists (remote/lane)
if (-not (Test-Path `$RemoteBackupFolder)) { New-Item -Path `$RemoteBackupFolder -ItemType Directory -Force | Out-Null }

# Keep only the newest N backups in a folder (called only after a good backup)
function Remove-OldBackups([string]`$Folder, [int]`$Keep) {
    @(Get-ChildItem -Path `$Folder -Filter "`$Database`_*.bak" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -Skip `$Keep) |
        ForEach-Object { Remove-Item -LiteralPath `$_.FullName -Force -ErrorAction SilentlyContinue }
}

# Run the backup on the lane via UNC SQL backup path (old backups are only rotated out AFTER a good one)
`$tsql = "BACKUP DATABASE [`$Database] TO DISK = N'`$remoteBakFile' WITH NOFORMAT, NOINIT, NAME = N'`$Database-FullBackup-`$timestamp', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
`$sqlOut = & sqlcmd -S `"$SqlcmdS`" -Q `$tsql -b 2>&1
`$exitCode = `$LASTEXITCODE

if (`$exitCode -eq 0) {
    # Wait for backup to finish (file to appear and not be locked)
    `$waitSec = 0
    while (-not (Test-Path `$remoteBakFile) -and `$waitSec -lt 120) {
        Start-Sleep -Seconds 1; `$waitSec++
    }
    # Optionally, wait until file is not locked (ready for copy)
    `$ready = `$false; `$tries = 0
    while (-not `$ready -and `$tries -lt 60) {
        try {
            `$s = [System.IO.File]::Open(`$remoteBakFile, 'Open', 'Read', 'None')
            `$s.Close(); `$ready = `$true
        } catch { Start-Sleep -Milliseconds 500; `$tries++ }
    }

    if ((Test-Path `$remoteBakFile) -and `$ready) {
        try {
            Copy-Item -Path `$remoteBakFile -Destination `$localBakFile -Force
            "[`$(Get-Date)] Backup complete: `$localBakFile" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
            Remove-OldBackups `$BackupRoot `$MaxBackups
        } catch {
            "[`$(Get-Date)] Backup succeeded but copy failed: `$(`$_.Exception.Message)" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
        }
        Remove-OldBackups `$RemoteBackupFolder `$MaxRemoteBackups
    } else {
        "[`$(Get-Date)] Backup file not found or not ready for copy after backup." | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
    }
} else {
    # Remove failed/partial file from lane backup folder if exists
    if (Test-Path `$remoteBakFile) { Remove-Item -Path `$remoteBakFile -Force }
    `$errMsg = (`$sqlOut | Out-String).Trim()
    if (-not `$errMsg) { `$errMsg = "(sqlcmd produced no output)" }
    "[`$(Get-Date)] Backup FAILED for [`$Database] on [`$SqlcmdS] (exit code: `$exitCode)" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
    "[`$(Get-Date)]   sqlcmd output: `$errMsg" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
}
"@

		# Write the script to scripts folder under machine folder
		$parentDir = Split-Path $backupScriptPath -Parent
		if (-not (Test-Path $parentDir)) { New-Item -Path $parentDir -ItemType Directory -Force | Out-Null }
		Set-Content -Path $backupScriptPath -Value $backupScript -Encoding UTF8

		# Build Task Scheduler arguments
		$hour = "{0:D2}" -f [int]$hours
		$min = "{0:D2}" -f [int]$minutes
		$startTime = "$hour`:$min"
		$taskName = "${machine}_DB_Backup_Schedule"
		# \" survives 5.1's native argument passing, so a path with spaces stays one -File argument
		$action = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \`"$backupScriptPath\`""
		$createArgs = @('/Create', '/TN', $taskName, '/TR', $action, '/SC', 'DAILY', '/ST', $startTime, '/RL', 'HIGHEST', '/F')
		if ($freqInput -gt 1) { $createArgs += @('/MO', [string]$freqInput) }

		# SYSTEM; a user with no password runs interactive only (/IT, must be logged in);
		# with a password the task also runs when nobody is logged in
		if ($taskUser -eq "SYSTEM") { $createArgs += @('/RU', 'SYSTEM'); $runAs = 'SYSTEM' }
		elseif ($taskPwd -eq "") { $createArgs += @('/RU', $taskUser, '/IT'); $runAs = "$taskUser (interactive)" }
		else { $createArgs += @('/RU', $taskUser, '/RP', $taskPwd); $runAs = "$taskUser (with password)" }

		Write_Log "Scheduling lane $lane ($machine) as $runAs`: $action" "cyan"
		$schtasksOut = & schtasks.exe @createArgs 2>&1
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Backup task scheduled for $machine ($dbName) at $startTime every $freqInput day(s)." "green"
		}
		else
		{
			Write_Log "Failed to schedule backup for $machine`: $(($schtasksOut | Out-String).Trim())" "red"
		}
	}

	Write_Log "`r`n==================== Schedule_LaneDB_Backup Completed ====================" "blue"
}

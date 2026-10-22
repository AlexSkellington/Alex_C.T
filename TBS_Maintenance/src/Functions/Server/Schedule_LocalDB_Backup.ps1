# ===================================================================================================
#                              FUNCTION: Schedule_LocalDB_Backup
# ---------------------------------------------------------------------------------------------------
# Description:
#   GUI to configure, schedule, and maintain the automated backup of the local store SQL
#   database: prompts for time, frequency, and retention, then generates and schedules a
#   SYSTEM task running a PowerShell script that writes a dated .bak and deletes oldest
#   backups, keeping only the configured count.
#
# Parameters:
#   (none - uses $script:FunctionResults for database info)
#
# Variables:
#   - $script:FunctionResults: DBNAME, DBSERVER
#   - $script:BackupRoot: backup destination (per-host subfolder)
#   - $script:ScriptsFolder: where the backup script is written
# ===================================================================================================

function Schedule_LocalDB_Backup
{
	Write_Log "`r`n==================== Starting Schedule_LocalDB_Backup ====================`r`n" "blue"

	try
	{
		# --- Validate FunctionResults for DB and Paths
		$dbName = $script:FunctionResults['DBNAME']
		$dbServer = $script:FunctionResults['DBSERVER']
		$backupDir = $script:BackupRoot
		$scriptsDir = $script:ScriptsFolder
		$serverFolder = Join-Path $backupDir $LocalHost

		if (-not $dbName -or -not $dbServer)
		{
			Write_Log "DBNAME or DBSERVER not found in FunctionResults. Aborting." "red"
			return
		}
		# --- Prompt User: Backup Time, Freq, Retention
		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing

		$form = New-Object System.Windows.Forms.Form
		$form.Text = "Configure Local DB Backup Scheduler"
		$form.Size = New-Object System.Drawing.Size(375, 280)
		$form.StartPosition = "CenterScreen"
		$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
		$form.MaximizeBox = $false
		$form.MinimizeBox = $false

		$lblTime = New-Object System.Windows.Forms.Label
		$lblTime.Text = "Time to run backup (24h, HH:mm):"
		$lblTime.Location = New-Object System.Drawing.Point(10, 18)
		$lblTime.Size = New-Object System.Drawing.Size(210, 20)
		$form.Controls.Add($lblTime)

		$txtTime = New-Object System.Windows.Forms.MaskedTextBox
		$txtTime.Mask = "00:00"
		$txtTime.Text = "01:00"
		$txtTime.Location = New-Object System.Drawing.Point(220, 16)
		$txtTime.Size = New-Object System.Drawing.Size(60, 20)
		$form.Controls.Add($txtTime)

		$lblFreq = New-Object System.Windows.Forms.Label
		$lblFreq.Text = "Frequency (every X days):"
		$lblFreq.Location = New-Object System.Drawing.Point(10, 58)
		$lblFreq.Size = New-Object System.Drawing.Size(210, 20)
		$form.Controls.Add($lblFreq)

		$numFreq = New-Object System.Windows.Forms.NumericUpDown
		$numFreq.Minimum = 1
		$numFreq.Maximum = 31
		$numFreq.Value = 1
		$numFreq.Location = New-Object System.Drawing.Point(220, 56)
		$numFreq.Size = New-Object System.Drawing.Size(60, 20)
		$form.Controls.Add($numFreq)

		$lblKeep = New-Object System.Windows.Forms.Label
		$lblKeep.Text = "How many backups to keep:"
		$lblKeep.Location = New-Object System.Drawing.Point(10, 98)
		$lblKeep.Size = New-Object System.Drawing.Size(210, 20)
		$form.Controls.Add($lblKeep)

		$numKeep = New-Object System.Windows.Forms.NumericUpDown
		$numKeep.Minimum = 1
		$numKeep.Maximum = 99
		$numKeep.Value = 3
		$numKeep.Location = New-Object System.Drawing.Point(220, 96)
		$numKeep.Size = New-Object System.Drawing.Size(60, 20)
		$form.Controls.Add($numKeep)

		$lblDir = New-Object System.Windows.Forms.Label
		$lblDir.Text = "Backup folder:"
		$lblDir.Location = New-Object System.Drawing.Point(10, 138)
		$lblDir.Size = New-Object System.Drawing.Size(95, 20)
		$form.Controls.Add($lblDir)

		$txtDir = New-Object System.Windows.Forms.TextBox
		$txtDir.ReadOnly = $true
		$txtDir.Text = $backupDir.TrimEnd('\')
		$txtDir.Location = New-Object System.Drawing.Point(108, 136)
		$txtDir.Size = New-Object System.Drawing.Size(150, 20)
		$form.Controls.Add($txtDir)

		$btnBrowse = New-Object System.Windows.Forms.Button
		$btnBrowse.Text = "Browse..."
		$btnBrowse.Location = New-Object System.Drawing.Point(262, 134)
		$btnBrowse.Size = New-Object System.Drawing.Size(75, 24)
		$btnBrowse.Add_Click({
			$p = Select-FolderOrDefault -initialPath $txtDir.Text -description "Choose where to store the local DB backup (a <host> subfolder is created inside it)."
			if ($p) { $txtDir.Text = $p }
		})
		$form.Controls.Add($btnBrowse)

		$btnOK = New-Object System.Windows.Forms.Button
		$btnOK.Text = "OK"
		$btnOK.Location = New-Object System.Drawing.Point(70, 190)
		$btnOK.Size = New-Object System.Drawing.Size(90, 30)
		$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$form.AcceptButton = $btnOK
		$form.Controls.Add($btnOK)

		$btnCancel = New-Object System.Windows.Forms.Button
		$btnCancel.Text = "Cancel"
		$btnCancel.Location = New-Object System.Drawing.Point(185, 190)
		$btnCancel.Size = New-Object System.Drawing.Size(90, 30)
		$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$form.CancelButton = $btnCancel
		$form.Controls.Add($btnCancel)

		$result = $form.ShowDialog()
		if ($result -ne [System.Windows.Forms.DialogResult]::OK)
		{
			Write_Log "Backup scheduling cancelled by user." "yellow"
			Write_Log "`r`n==================== Schedule_LocalDB_Backup Completed ====================" "blue"
			return
		}

		# Backup location: the folder the user picked (default if unchanged) + <host> subfolder
		$chosenParent = $txtDir.Text.Trim()
		if (-not $chosenParent) { $chosenParent = $backupDir.TrimEnd('\') }
		$serverFolder = Join-Path $chosenParent $LocalHost
		Write_Log "Local DB backup destination: $serverFolder" "gray"

		$timeInput = $txtTime.Text
		$freqInput = [int]$numFreq.Value
		$keepInput = [int]$numKeep.Value

		if ($timeInput -notmatch '^\d{2}:\d{2}$')
		{
			[System.Windows.Forms.MessageBox]::Show("Time must be in HH:mm (24h) format.", "Input Error", 0, 16)
			Write_Log "Invalid time format entered." "red"
			return
		}
		$hours, $minutes = $timeInput.Split(":")
		if ([int]$hours -gt 23 -or [int]$minutes -gt 59)
		{
			[System.Windows.Forms.MessageBox]::Show("Invalid time entered.", "Input Error", 0, 16)
			Write_Log "Invalid time (out of range)." "red"
			return
		}

		# --- Compose the backup script (no nested here-string)
		$backupScript = @"
# Auto-generated by Schedule_${LocalHost}_DB_Backup (Alex_C.T)
`$ErrorActionPreference = 'Stop'
`$BackupRoot = `"$serverFolder`"
`$Database = `"$dbName`"
`$SqlInstance = `"$dbServer`"
`$MaxBackups = $keepInput

# Compose backup file path with timestamp
`$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
`$backupFile = Join-Path `$BackupRoot ("{0}_{1}.bak" -f `$Database, `$timestamp)

# Ensure folder exists
if (-not (Test-Path `$BackupRoot)) {
    New-Item -Path `$BackupRoot -ItemType Directory -Force | Out-Null
}

# Delete oldest backups if needed
`$bakFiles = Get-ChildItem -Path `$BackupRoot -Filter "`$Database`_*.bak" | Sort-Object LastWriteTime
while (`$bakFiles.Count -ge `$MaxBackups) {
    Remove-Item -Path `$bakFiles[0].FullName -Force
    `$bakFiles = Get-ChildItem -Path `$BackupRoot -Filter "`$Database`_*.bak" | Sort-Object LastWriteTime
}

# Run the backup
`$tsql = "BACKUP DATABASE [`$Database] TO DISK = N'`$backupFile' WITH NOFORMAT, NOINIT, NAME = N'`$Database-FullBackup-`$timestamp', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
`$sqlOut = & sqlcmd -S `"$dbServer`" -Q `$tsql -b 2>&1
`$exitCode = `$LASTEXITCODE

if (`$exitCode -eq 0) {
    # Success log
    "[`$(Get-Date)] Backup complete: `$backupFile" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
}
else {
    # Failure log - capture the real sqlcmd error so it is diagnosable on the store
    `$errMsg = (`$sqlOut | Out-String).Trim()
    if (-not `$errMsg) { `$errMsg = "(sqlcmd produced no output)" }
    "[`$(Get-Date)] Backup FAILED for [`$Database] on [`$SqlInstance] (exit code: `$exitCode)" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
    "[`$(Get-Date)]   sqlcmd output: `$errMsg" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
    if (`$errMsg -match 'Msg 916|Msg 3013|Msg 3201|not able to access the database') {
        "[`$(Get-Date)]   ACCESS problem: backup ran as an account with no access to [`$Database]. Re-run 'Schedule DB Backup' (it now runs as the logged-in user) or set the task to an account that can open the database." | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
    }
}
"@

		# --- Write the backup script to the scripts folder
		$scriptName = "Run_${LocalHost}_DB_Backup.ps1"
		$backupScriptPath = Join-Path $scriptsDir $scriptName
		if (-not (Test-Path $scriptsDir))
		{
			New-Item -Path $scriptsDir -ItemType Directory -Force | Out-Null
		}
		Set-Content -Path $backupScriptPath -Value $backupScript -Encoding UTF8

		# --- Build Task Scheduler arguments ---
		$hour = "{0:D2}" -f [int]$hours
		$min = "{0:D2}" -f [int]$minutes
		$startTime = "$hour`:$min"

		$freqArg = "/SC DAILY"
		if ($freqInput -gt 1)
		{
			$freqArg = "/SC DAILY /MO $freqInput"
		}
		$taskName = "${LocalHost}_DB_Backup_Schedule"
		$action = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"`"$backupScriptPath`"`""

		# --- Remove existing task if present ---
		schtasks.exe /Delete /TN "$taskName" /F 2>$null | Out-Null


		# --- Create scheduled task ---
		# Run as the CURRENT (logged-in) user, not SYSTEM: on the store the SQL
		# login that can open the database is the interactive account, so the
		# backup must run under it (interactive, /IT). SYSTEM has no DB access.
		$backupTaskUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
		$cmd = "schtasks.exe /Create /RU `"$backupTaskUser`" /IT /RL HIGHEST /TN `"$taskName`" /TR `"$action`" $freqArg /ST $startTime /F"
		Write_Log "Creating backup task as '$backupTaskUser' (interactive /IT) with command:" "yellow"
		Write_Log $cmd "gray"
		Invoke-Expression $cmd

		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Backup task scheduled successfully (`"$taskName`") as '$backupTaskUser' at $startTime every $freqInput day(s)." "green"
			[System.Windows.Forms.MessageBox]::Show("Backup task scheduled successfully as '$backupTaskUser'.`nScript: $backupScriptPath", "Scheduled!", 0, 64)
		}
		else
		{
			Write_Log "Failed to schedule backup task. Check permissions or path." "red"
			[System.Windows.Forms.MessageBox]::Show("Failed to schedule backup task.`nCheck permissions or path.", "Error", 0, 16)
		}
	}
	catch
	{
		Write_Log "Fatal error in Schedule_LocalDB_Backup: $($_.Exception.Message)" "red"
		[System.Windows.Forms.MessageBox]::Show("Fatal error: $($_.Exception.Message)", "Error", 0, 16)
	}
	Write_Log "`r`n==================== Schedule_LocalDB_Backup Completed ====================" "blue"
}

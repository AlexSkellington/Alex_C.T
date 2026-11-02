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
		$form.Size = New-Object System.Drawing.Size(375, 316)
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
		$txtDir.ReadOnly = $false
		$txtDir.Text = $backupDir.TrimEnd('\')
		$txtDir.Location = New-Object System.Drawing.Point(108, 136)
		$txtDir.Size = New-Object System.Drawing.Size(150, 20)
		$form.Controls.Add($txtDir)
		try { $toolTip.SetToolTip($txtDir, "Type or paste a folder - a network path like \\server\share\Backups works - or use Browse. The SQL Server service writes the .bak, so for a share its account needs write access too.") } catch { }

		$btnBrowse = New-Object System.Windows.Forms.Button
		$btnBrowse.Text = "Browse..."
		$btnBrowse.Location = New-Object System.Drawing.Point(262, 134)
		$btnBrowse.Size = New-Object System.Drawing.Size(75, 24)
		$btnBrowse.Add_Click({
			$p = Select-FolderOrDefault -initialPath $txtDir.Text -description "Choose where to store the local DB backup (a <host> subfolder is created inside it)."
			if ($p) { $txtDir.Text = $p }
		})
		$form.Controls.Add($btnBrowse)

		# Extra copies ("Also copy to") - the list is remembered per backup type
		$btnCopies = New-Object System.Windows.Forms.Button
		$btnCopies.Location = New-Object System.Drawing.Point(108, 166)
		$btnCopies.Size = New-Object System.Drawing.Size(229, 26)
		$btnCopies.Tag = Get_Backup_Copy_Paths -Kind 'ServerDB'
		$btnCopies.Text = $(if (@($btnCopies.Tag).Count -gt 0) { "Also copy to " + @($btnCopies.Tag).Count + " more folder(s)..." } else { "Also copy to other folders..." })
		$btnCopies.Add_Click({
				param ($s, $e)
				$r = Show_Backup_Copy_Paths_Dialog -Kind 'ServerDB' -Current ([string[]]@($s.Tag))
				if ($r -is [System.Array]) { $s.Tag = $r; $s.Text = $(if (@($r).Count -gt 0) { "Also copy to " + @($r).Count + " more folder(s)..." } else { "Also copy to other folders..." }) }
			})
		try { $toolTip.SetToolTip($btnCopies, "Extra folders (local or network) that get a copy of every successful backup.") } catch { }
		$form.Controls.Add($btnCopies)

		$btnOK = New-Object System.Windows.Forms.Button
		$btnOK.Text = "OK"
		$btnOK.Location = New-Object System.Drawing.Point(70, 226)
		$btnOK.Size = New-Object System.Drawing.Size(90, 30)
		$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$form.AcceptButton = $btnOK
		$form.Controls.Add($btnOK)

		$btnCancel = New-Object System.Windows.Forms.Button
		$btnCancel.Text = "Cancel"
		$btnCancel.Location = New-Object System.Drawing.Point(185, 226)
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
		$folderCheck = Resolve_Backup_Folder -Path $chosenParent
		if (-not $folderCheck.Ok)
		{
			Write_Log ("Backup folder problem: " + $folderCheck.Error) "red"
			[void][System.Windows.Forms.MessageBox]::Show($folderCheck.Error, "Backup Folder", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			return
		}
		$chosenParent = $folderCheck.Path
		if ($folderCheck.IsNetwork) { Write_Log ("Network folder: the scheduled backup runs as " + "the logged-in user (and SQL Server writes the .bak with its service account)" + " - that account needs write access to " + $chosenParent + ".") "yellow" }
		$copyCheck = Resolve_Backup_Copy_Paths -Paths ([string[]]@($btnCopies.Tag)) -Primary $chosenParent
		if (-not $copyCheck.Ok)
		{
			Write_Log ("Extra backup folder problem: " + $copyCheck.Error) "red"
			[void][System.Windows.Forms.MessageBox]::Show($copyCheck.Error, "Also Copy To", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			return
		}
		$copyPaths = $copyCheck.Paths
		if ($copyPaths.Count -gt 0) { Write_Log ("Each backup is also copied to: " + ($copyPaths -join '; ')) "gray" }
		$serverFolder = Join-Path $chosenParent $LocalHost
		Write_Log "Local DB backup destination: $serverFolder" "gray"

		$timeInput = $txtTime.Text
		$freqInput = [int]$numFreq.Value
		$keepInput = [int]$numKeep.Value

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

		# Lines that copy each good backup to the extra folders (empty when there are none)
		$copySnippet = New_Backup_Copy_Snippet -Paths $copyPaths -SubFolder $LocalHost -Filter ($dbName + '_*.bak') -Keep $keepInput -SourceExpression '$backupFile' -LogExpression '(Join-Path $BackupRoot ''backup.log'')'

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

# Run the backup (old backups are only rotated out AFTER a successful one)
`$tsql = "BACKUP DATABASE [`$Database] TO DISK = N'`$backupFile' WITH NOFORMAT, NOINIT, NAME = N'`$Database-FullBackup-`$timestamp', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
`$sqlOut = & sqlcmd -S `"$dbServer`" -Q `$tsql -b 2>&1
`$exitCode = `$LASTEXITCODE

if (`$exitCode -eq 0) {
    # Success log, then keep only the newest MaxBackups
    "[`$(Get-Date)] Backup complete: `$backupFile" | Out-File -FilePath (Join-Path `$BackupRoot `"backup.log`") -Append -Encoding utf8
    @(Get-ChildItem -Path `$BackupRoot -Filter "`$Database`_*.bak" | Sort-Object LastWriteTime -Descending | Select-Object -Skip `$MaxBackups) |
        ForEach-Object { Remove-Item -LiteralPath `$_.FullName -Force -ErrorAction SilentlyContinue }
$copySnippet
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

		$freqArgs = @('/SC', 'DAILY')
		if ($freqInput -gt 1) { $freqArgs += @('/MO', [string]$freqInput) }
		$taskName = "${LocalHost}_DB_Backup_Schedule"
		$action = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \`"$backupScriptPath\`""

		# --- Remove existing task if present ---
		$null = & schtasks.exe /Delete /TN $taskName /F 2>$null

		# --- Create scheduled task ---
		# Run as the CURRENT (logged-in) user, not SYSTEM: on the store the SQL
		# login that can open the database is the interactive account, so the
		# backup must run under it (interactive, /IT). SYSTEM has no DB access.
		# (argument array, not Invoke-Expression of one string: that broke on '$' and quotes)
		$backupTaskUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
		$createArgs = @('/Create', '/RU', $backupTaskUser, '/IT', '/RL', 'HIGHEST', '/TN', $taskName, '/TR', $action) + $freqArgs + @('/ST', $startTime, '/F')
		Write_Log "Creating backup task as '$backupTaskUser' (interactive /IT): schtasks $($createArgs -join ' ')" "gray"
		$createOut = & schtasks.exe @createArgs 2>&1
		if ($LASTEXITCODE -ne 0) { Write_Log ("schtasks: " + (($createOut | Out-String).Trim())) "red" }

		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Backup task scheduled successfully (`"$taskName`") as '$backupTaskUser' at $startTime every $freqInput day(s)." "green"
			[void][System.Windows.Forms.MessageBox]::Show("Backup task scheduled successfully as '$backupTaskUser'.`nScript: $backupScriptPath", "Scheduled!", 0, 64)
		}
		else
		{
			Write_Log "Failed to schedule backup task. Check permissions or path." "red"
			[void][System.Windows.Forms.MessageBox]::Show("Failed to schedule backup task.`nCheck permissions or path.", "Error", 0, 16)
		}
	}
	catch
	{
		Write_Log "Fatal error in Schedule_LocalDB_Backup: $($_.Exception.Message)" "red"
		[void][System.Windows.Forms.MessageBox]::Show("Fatal error: $($_.Exception.Message)", "Error", 0, 16)
	}
	Write_Log "`r`n==================== Schedule_LocalDB_Backup Completed ====================" "blue"
}

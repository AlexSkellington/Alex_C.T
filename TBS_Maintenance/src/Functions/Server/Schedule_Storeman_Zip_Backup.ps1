# ===================================================================================================
#                              FUNCTION: Schedule_Storeman_Zip_Backup
# ---------------------------------------------------------------------------------------------------
# Description:
#   GUI to configure, schedule, and maintain the automated ZIP backup of the Storeman folder
#   (weekly by default): prompts for time, frequency, and retention, writes
#   Run_<host>_Storeman_Zip_Backup.ps1, and schedules the <host>_Storeman_Zip_Backup SYSTEM
#   task; the script zips to a dated .zip and keeps only the configured number of backups.
#
# Parameters:
#   (none - uses $BasePath as the detected Storeman folder)
#
# Variables:
#   - $BasePath: detected Storeman folder
#   - $script:BackupRoot: backup destination (per-host subfolder)
#   - $script:ScriptsFolder: where the backup script is written
# ===================================================================================================

function Schedule_Storeman_Zip_Backup
{
	Write_Log "`r`n==================== Starting Schedule_Storeman_Zip_Backup ====================`r`n" "blue"
	
	try
	{
		# --- Use detected Storeman folder ---
		$storemanPath = $BasePath
		if (-not (Test-Path $storemanPath))
		{
			Write_Log "Storeman path ($storemanPath) not found. Aborting." "red"
			return
		}
		$LocalHost = $env:COMPUTERNAME
		$backupRoot = "${script:BackupRoot}${LocalHost}"
		$scriptsDir = $script:ScriptsFolder
		$backupScriptName = "Run_${LocalHost}_Storeman_Zip_Backup.ps1"
		$taskName = "${LocalHost}_Storeman_Zip_Backup"
		
		# --- Prompt User: Backup Time, Frequency, Retention (defaults set for weekly) ---
		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing
		
		$form = New-Object System.Windows.Forms.Form
		$form.Text = "Configure Storeman ZIP Backup Scheduler"
		$form.Size = New-Object System.Drawing.Size(385, 280)
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
		$txtTime.Text = "02:00"
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
		$numFreq.Value = 7 # Default = once a week
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
		$numKeep.Value = 1
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
		$txtDir.Text = $script:BackupRoot.TrimEnd('\')
		$txtDir.Location = New-Object System.Drawing.Point(108, 136)
		$txtDir.Size = New-Object System.Drawing.Size(150, 20)
		$form.Controls.Add($txtDir)
		
		$btnBrowse = New-Object System.Windows.Forms.Button
		$btnBrowse.Text = "Browse..."
		$btnBrowse.Location = New-Object System.Drawing.Point(262, 134)
		$btnBrowse.Size = New-Object System.Drawing.Size(75, 24)
		$btnBrowse.Add_Click({
			$p = Select-FolderOrDefault -initialPath $txtDir.Text -description "Choose where to store the Storeman ZIP backup (a <host> subfolder is created inside it)."
			if ($p) { $txtDir.Text = $p }
		})
		$form.Controls.Add($btnBrowse)
		
		$btnOK = New-Object System.Windows.Forms.Button
		$btnOK.Text = "OK"
		$btnOK.Location = New-Object System.Drawing.Point(80, 190)
		$btnOK.Size = New-Object System.Drawing.Size(90, 30)
		$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$form.AcceptButton = $btnOK
		$form.Controls.Add($btnOK)
		
		$btnCancel = New-Object System.Windows.Forms.Button
		$btnCancel.Text = "Cancel"
		$btnCancel.Location = New-Object System.Drawing.Point(195, 190)
		$btnCancel.Size = New-Object System.Drawing.Size(90, 30)
		$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$form.CancelButton = $btnCancel
		$form.Controls.Add($btnCancel)
		
		$result = $form.ShowDialog()
		if ($result -ne [System.Windows.Forms.DialogResult]::OK)
		{
			Write_Log "Storeman ZIP backup scheduling cancelled by user." "yellow"
			Write_Log "`r`n==================== Schedule_Storeman_Zip_Backup Completed ====================" "blue"
			return
		}
		
		# Backup location: the folder the user picked (default if unchanged) + <host> subfolder
		$chosenParent = $txtDir.Text.Trim()
		if (-not $chosenParent) { $chosenParent = $script:BackupRoot.TrimEnd('\') }
		$backupRoot = Join-Path $chosenParent $LocalHost
		Write_Log "Storeman ZIP backup destination: $backupRoot" "gray"
		
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
		
		# --- Compose the backup script ---
		$backupScript = @"
# Auto-generated by Schedule_${LocalHost}_Storeman_Zip_Backup (Alex_C.T)
`$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

`$storemanPath = `"$storemanPath`"
`$backupRoot = `"$backupRoot`"
if (-not (Test-Path `$storemanPath)) {
    Write-Host "Storeman path (`$storemanPath) not found. Aborting." -ForegroundColor Red
    exit 1
}
if (-not (Test-Path `$backupRoot)) {
    New-Item -Path `$backupRoot -ItemType Directory -Force | Out-Null
}

# Compose backup ZIP filename with timestamp
`$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
`$backupZipPath = Join-Path `$backupRoot ("Storeman_Backup_`$timestamp.zip")

# Delete oldest backups if needed
`$oldZips = Get-ChildItem -Path `$backupRoot -Filter "Storeman_Backup_*.zip" | Sort-Object LastWriteTime -Descending
if (`$oldZips.Count -ge $keepInput) {
    `$oldZips | Select-Object -Skip ($keepInput - 1) | ForEach-Object {
        try {
            Remove-Item `$_.FullName -Force
        } catch {
            Write-Host "Failed to delete old backup: `$_.FullName" -ForegroundColor Yellow
        }
    }
}

# Run the backup
try {
    `$zip = [System.IO.Compression.ZipFile]::Open(`$backupZipPath, 1)  # 1 = Create

    # Gather all root directories starting with "install" (case-insensitive)
    `$rootDirs = Get-ChildItem -Path `$storemanPath -Directory | Select-Object -ExpandProperty FullName
    `$installDirs = `$rootDirs | Where-Object { `$_ -match '(?i)\\install' }

    `$backupFullPathU = [System.IO.Path]::GetFullPath((Join-Path `$storemanPath "BACKUP"))
    `$backupFullPathL = [System.IO.Path]::GetFullPath((Join-Path `$storemanPath "backup"))
    `$logFullPath     = [System.IO.Path]::GetFullPath((Join-Path `$storemanPath "log"))

    `$files = Get-ChildItem -Path `$storemanPath -Recurse -File
    `$countAdded = 0

    foreach (`$file in `$files) {
        `$filePath = [System.IO.Path]::GetFullPath(`$file.FullName)

        # Skip files under any install* folder
        `$skipInstall = `$false
        foreach (`$dir in `$installDirs) {
            if (`$filePath -like "`$dir*") {
                `$skipInstall = `$true
                break
            }
        }
        if (`$skipInstall) {
            continue
        }

        # Skip other excluded folders and log files
        if (
            (`$filePath -like "`$backupFullPathU*") -or
            (`$filePath -like "`$backupFullPathL*") -or
            (`$filePath -like "`$logFullPath*") -or
            (`$file.Extension -eq ".log")
        ) {
            continue
        }

        `$relativePath = `$filePath.Substring(`$storemanPath.Length).TrimStart('\','/')
        try {
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(`$zip, `$file.FullName, `$relativePath) | Out-Null
            `$countAdded++
        }
        catch {
            continue
        }
    }
    `$zip.Dispose()
    "`$([DateTime]::Now) Backup complete: `$backupZipPath (`$countAdded files)" | Out-File -FilePath (Join-Path `$backupRoot "backup.log") -Append -Encoding utf8
    Write-Host "Backup complete: `$backupZipPath (`$countAdded files)" -ForegroundColor Green
}
catch {
    if (`$zip) { `$zip.Dispose() }
    "`$([DateTime]::Now) Backup FAILED for `$storemanPath - `$(`$_.Exception.Message)" | Out-File -FilePath (Join-Path `$backupRoot "backup.log") -Append -Encoding utf8
    Write-Host "Backup FAILED: `$(`$_.Exception.Message)" -ForegroundColor Red
}
"@
		
		# --- Write the backup script to the scripts folder
		$backupScriptPath = Join-Path $scriptsDir $backupScriptName
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
		$action = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"`"$backupScriptPath`"`""
		
		# --- Remove existing task if present ---
		schtasks.exe /Delete /TN "$taskName" /F 2>$null | Out-Null
		
		# --- Create scheduled task ---
		$cmd = "schtasks.exe /Create /RU SYSTEM /RL HIGHEST /TN `"$taskName`" /TR `"$action`" $freqArg /ST $startTime /F"
		Write_Log "Creating scheduled task with command:" "yellow"
		Write_Log $cmd "gray"
		Invoke-Expression $cmd
		
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Storeman ZIP backup task scheduled successfully (`"$taskName`") at $startTime every $freqInput day(s)." "green"
			[System.Windows.Forms.MessageBox]::Show("Storeman ZIP backup scheduled successfully.`nScript: $backupScriptPath", "Scheduled!", 0, 64)
		}
		else
		{
			Write_Log "Failed to schedule Storeman ZIP backup task. Check permissions or path." "red"
			[System.Windows.Forms.MessageBox]::Show("Failed to schedule Storeman ZIP backup task.`nCheck permissions or path.", "Error", 0, 16)
		}
	}
	catch
	{
		Write_Log "Fatal error in Schedule_Storeman_Zip_Backup: $($_.Exception.Message)" "red"
		[System.Windows.Forms.MessageBox]::Show("Fatal error: $($_.Exception.Message)", "Error", 0, 16)
	}
	Write_Log "`r`n==================== Schedule_Storeman_Zip_Backup Completed ====================" "blue"
}

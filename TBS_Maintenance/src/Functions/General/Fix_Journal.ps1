# ===================================================================================================
#                              FUNCTION: Fix_Journal
# ---------------------------------------------------------------------------------------------------
# Description:
#   Repairs journal (EJ) files in the store ZX folder: prompts for a date in a dialog, finds
#   the EJ files for that date/store and removes each block from a '<trs F10' line up to (not
#   including) the next '<trs F1068' line. Each changed file is backed up first - outside the
#   ZX folder, where SMS would read a backup as another journal - then rewritten in ANSI.
#
# Parameters:
#   - OfficePath   [string]: Storeman Office base path (required)
#   - StoreNumber  [string]: store number, e.g. 001 (required)
#
# Writes:
#   - ZX folder EJ files (fixed lines); originals in <BackupRoot>\Journal_Fix\<yyyyMMdd_HHmmss>
# ===================================================================================================

function Fix_Journal
{
	[CmdletBinding()]
	param (
		# The base "OfficePath", e.g. "C:\storeman\office"
		[Parameter(Mandatory = $true)]
		[string]$OfficePath,
		# The store number (e.g., "001")
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)

	Write_Log "`r`n==================== Starting Fix_Journal Function ====================`r`n" "blue"

	# ---------------------------------------------------------------------------------------------
	# 1) Load Windows Forms assembly
	# ---------------------------------------------------------------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	# ---------------------------------------------------------------------------------------------
	# 2) Create and configure the Windows Form
	# ---------------------------------------------------------------------------------------------
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Select Date"
	$form.Size = New-Object System.Drawing.Size(300, 200)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = "FixedDialog"
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false

	# Create Label
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Please select the date:"
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(10, 20)
	$form.Controls.Add($label)

	# Create DateTimePicker
	$dateTimePicker = New-Object System.Windows.Forms.DateTimePicker
	$dateTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
	$dateTimePicker.Location = New-Object System.Drawing.Point(10, 50)
	$dateTimePicker.Width = 260
	$form.Controls.Add($dateTimePicker)

	# Create OK Button
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "OK"
	$okButton.Location = New-Object System.Drawing.Point(110, 100)
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $okButton
	$form.Controls.Add($okButton)

	# Show the form and capture the result
	$dialogResult = $form.ShowDialog()

	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log -Message "Date selection canceled by user. Exiting function." "yellow"
		return
	}

	# ---------------------------------------------------------------------------------------------
	# 3) Retrieve and format the selected date
	# ---------------------------------------------------------------------------------------------
	$snippetDate = $dateTimePicker.Value
	$formattedDate = $snippetDate.ToString('MMddyyyy') # MMDDYYYY format

	Write_Log -Message "Selected date: $formattedDate" "magenta"

	# ---------------------------------------------------------------------------------------------
	# 4) Construct ZX folder path from $OfficePath + $StoreNumber
	#    ZX folder: $OfficePath\ZX${StoreNumber}901
	# ---------------------------------------------------------------------------------------------
	$zxFolderPath = Join-Path $OfficePath "ZX${StoreNumber}901"

	# ---------------------------------------------------------------------------------------------
	# 5) Confirm the ZX folder exists
	# ---------------------------------------------------------------------------------------------
	if (-not (Test-Path -Path $zxFolderPath))
	{
		Write_Log -Message "ZX folder not found: $zxFolderPath." "red"
		return
	}

	# ---------------------------------------------------------------------------------------------
	# 6) Build the file prefix: YMMDDSSS (ignoring lane)
	#     Y = last digit of year
	#     MM = 2-digit month
	#     DD = 2-digit day
	#     SSS = store number (3 digits, e.g., "001")
	# ---------------------------------------------------------------------------------------------
	$yearLastDigit = ($snippetDate.Year % 10)
	$mm = $snippetDate.ToString('MM')
	$dd = $snippetDate.ToString('dd')
	$filePrefix = "$yearLastDigit$mm$dd$StoreNumber" # e.g., 41227001

	Write_Log -Message "Looking for files named '$filePrefix.*' in $zxFolderPath..." "blue"

	# ---------------------------------------------------------------------------------------------
	# 7) Find matching EJ files in ZX folder: e.g., 41227001.*
	# ---------------------------------------------------------------------------------------------
	$searchPattern = "$filePrefix.*"
	$matchingFiles = Get-ChildItem -Path $zxFolderPath -Filter $searchPattern -File -ErrorAction SilentlyContinue

	if (-not $matchingFiles)
	{
		Write_Log -Message "No files matching '$searchPattern' found in $zxFolderPath." "yellow"
		return
	}

	Write_Log -Message "Found $(@($matchingFiles).Count) file(s) to check." "green"
	$backupDir = $null

	# ---------------------------------------------------------------------------------------------
	# 8) For each matching EJ file, remove lines from <trs F10... up to <trs F1068...
	# ---------------------------------------------------------------------------------------------
	foreach ($file in $matchingFiles)
	{
		# [Optional] Skip files that have ".bak" anywhere in their name
		# to avoid infinite backup loops:
		if ($file.Extension -eq ".bak")
		{
			Write_Log -Message "Skipping backup file: $($file.Name)" "yellow"
			continue
		}

		Write_Log -Message "Fixing lines in: $($file.FullName)" "yellow"

		# Read the file lines
		try
		{
			$originalLines = Get-Content -Path $file.FullName -ErrorAction Stop
		}
		catch
		{
			Write_Log -Message "Failed to read EJ file: $($file.FullName). Skipping." "red"
			continue
		}

		# Prepare a list for the fixed lines
		$fixedLines = New-Object System.Collections.Generic.List[string]

		$skip = $false
		$skipped = New-Object System.Collections.Generic.List[string]   # the block being removed

		foreach ($line in $originalLines)
		{
			# 1) Start skipping at '<trs F10'
			if ($line -match '^\s*<trs\s+F10\b')
			{
				if (-not $skip) { $skipped.Clear() }
				$skipped.Add($line)
				$skip = $true
				continue
			}

			# 2) Stop skipping at '<trs F1068'
			if ($skip -and ($line -match '^\s*<trs\s+F1068\b'))
			{
				$skip = $false
				$skipped.Clear()
				# We *do* want to keep this line
				$fixedLines.Add($line)
				continue
			}

			# Keep the line if we're not skipping
			if ($skip) { $skipped.Add($line) }
			else { $fixedLines.Add($line) }
		}
		# A block that never reached its <trs F1068 line is kept: removing to the end of the file
		# truncated the journal
		if ($skip) { $fixedLines.AddRange($skipped) }

		if ($fixedLines.Count -eq @($originalLines).Count)
		{
			Write_Log -Message "Nothing to fix in: $($file.Name)" "gray"
			continue
		}

		# -----------------------------------------------------------------------------------------
		# 10) Back up the original OUTSIDE the ZX folder (a .bak inside ZX reads as a journal)
		# -----------------------------------------------------------------------------------------
		if (-not $backupDir)
		{
			$backupDir = Join-Path $script:BackupRoot ("Journal_Fix\" + (Get-Date -Format 'yyyyMMdd_HHmmss'))
			try { New-Item -Path $backupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }
			catch
			{
				Write_Log -Message "Cannot create backup folder $backupDir - no journal was changed: $($_.Exception.Message)" "red"
				return
			}
		}
		$backupPath = Join-Path $backupDir $file.Name
		try { Copy-Item -LiteralPath $file.FullName -Destination $backupPath -Force -ErrorAction Stop }
		catch
		{
			Write_Log -Message "Failed to back up $($file.Name) - file left unchanged." "red"
			continue
		}

		# -----------------------------------------------------------------------------------------
		# 11) Overwrite the original file with the fixed lines in ANSI encoding
		# -----------------------------------------------------------------------------------------
		try
		{
			$fixedLines | Set-Content -LiteralPath $file.FullName -Encoding Default -ErrorAction Stop
			Write_Log -Message "Fixed $($file.Name): removed $(@($originalLines).Count - $fixedLines.Count) line(s). Backup: $backupPath" "green"
		}
		catch
		{
			Write_Log -Message "Failed to write fixed content to: $($file.FullName)." "red"
			continue
		}
	}
	Write_Log "`r`n==================== Fix_Journal Function Completed ====================" "blue"
}

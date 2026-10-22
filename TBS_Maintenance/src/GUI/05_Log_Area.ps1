# ===================================================================================================
# ENHANCED LOG AREA - Modern card with header bar
# ===================================================================================================
$logFrame = New-Object System.Windows.Forms.Panel
$logFrame.Name = "LogFrame"
$logFrame.BackColor = [System.Drawing.Color]::White
$logFrame.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$logFrame.Padding = New-Object System.Windows.Forms.Padding(0)
$logFrame.Margin = New-Object System.Windows.Forms.Padding(0)
$form.Controls.Add($logFrame) | Out-Null

# Log frame header (colored top bar with "LOG" text)
$logFrameHeader = New-Object System.Windows.Forms.Panel
$logFrameHeader.Name = "LogFrameHeader"
$logFrameHeader.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$logFrameHeader.Height = 28
$logFrameHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$logFrame.Controls.Add($logFrameHeader)

# Header text
$logHeaderLabel = New-Object System.Windows.Forms.Label
$logHeaderLabel.Name = "LogHeaderLabel"
$logHeaderLabel.Text = "LOG OUTPUT"
$logHeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI", [float]9, [System.Drawing.FontStyle]::Regular)
$logHeaderLabel.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$logHeaderLabel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$logHeaderLabel.AutoSize = $false
$logHeaderLabel.Width = 120
$logHeaderLabel.Height = 28
$logHeaderLabel.Location = New-Object System.Drawing.Point(12, 4)
$logFrameHeader.Controls.Add($logHeaderLabel)

# Clear button in header
$logClearButton = New-Object System.Windows.Forms.Button
$logClearButton.Name = "LogClearButton"
$logClearButton.Text = "X"
$logClearButton.Font = New-Object System.Drawing.Font("Segoe UI", [float]8, [System.Drawing.FontStyle]::Regular)
$logClearButton.Size = New-Object System.Drawing.Size(24, 20)
$logClearButton.Location = New-Object System.Drawing.Point(($logFrameHeader.Width - 32), 4)
$logClearButton.BackColor = [System.Drawing.Color]::Transparent
$logClearButton.ForeColor = [System.Drawing.Color]::FromArgb(150, 150, 150)
$logClearButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$logClearButton.FlatAppearance.BorderSize = 0
$logClearButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$logClearButton.UseVisualStyleBackColor = $false
$logClearButton.Margin = New-Object System.Windows.Forms.Padding(0)
# Follow the right edge of the log header when the frame resizes
$logClearButton.Anchor = 'Top,Right'
$toolTip.SetToolTip($logClearButton, "Clear log")
$logClearButton.Add_MouseEnter({ $logClearButton.ForeColor = [System.Drawing.Color]::Red })
$logClearButton.Add_MouseLeave({ $logClearButton.ForeColor = [System.Drawing.Color]::FromArgb(150, 150, 150) })
$logClearButton.Add_Click({
		$logBox.Clear()
		if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log "Log Cleared" }
	})
$logFrameHeader.Controls.Add($logClearButton)

# Actual log box (RichTextBox)
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Name = "LogBox"
$logBox.ReadOnly = $true
$logBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$logBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$logBox.BackColor = [System.Drawing.Color]::White
$logBox.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
$logBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$logBox.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 12)
$logBox.Margin = New-Object System.Windows.Forms.Padding(0)
$logFrame.Controls.Add($logBox) | Out-Null

# Reference binding for logger
try { $script:logBox = $logBox } catch { }
try { $script:LogBox = $logBox } catch { }
try { $global:logBox = $logBox } catch { }
try { $global:LogBox = $logBox } catch { }

# Publish to FunctionResults
try
{
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
	$script:FunctionResults['LogBox'] = $logBox
}
catch { }

# Announce the storeman environment in the log.
# Detection happens before the GUI, so the original message was console-only.
# It must be written AFTER the form is shown: before that the RichTextBox has
# no window handle and the colored append is silently dropped. Shown fires
# exactly once, right when the handle exists.
$form.add_Shown({
		try
		{
			# Startup summary block (v2.6.8): one calm, scannable overview of
			# the environment, placed ABOVE the individual startup lines.
			# Detection happens before the GUI, so this is where the user
			# gets the plain-language explanation of what this machine has
			# (and doesn't have) available.
			$rule = '--------------------------------------------------'
			$modeTag = if ($script:StoremanKeyStale) { '[ KEY  ]' } elseif ($script:IsTestEnvironment) { '[ TEST ]' } else { '[ LIVE ]' }
			
			# Snapshot the lines already in the log (text + color) so the
			# summary can be written on top of them, in order.
			$oldLines = @()
			try
			{
				$t = $logBox.Text
				$i = 0
				while ($i -lt $t.Length)
				{
					$nl = $t.IndexOf("`n", $i)
					if ($nl -lt 0) { $nl = $t.Length }
					$lineText = $t.Substring($i, $nl - $i)
					$lineColor = [System.Drawing.Color]::Black
					try { $lineColor = $logBox.GetCharColor($i) } catch { }
					$oldLines += ,@($lineText, $lineColor)
					$i = $nl + 1
				}
			}
			catch { $oldLines = @() }
			
			try { $logBox.Clear() } catch { }
			
			$storeNum = ''
			$storeName = ''
			$smsVer = ''
			try
			{
				if ($script:FunctionResults)
				{
					$storeNum  = [string]$script:FunctionResults['StoreNumber']
					$storeName = [string]$script:FunctionResults['StoreName']
					$smsVer    = [string]$script:FunctionResults['SMSVersionFull']
				}
			}
			catch { }
			
			Write_Log $rule -FileOnly -IncludeTimestamp:$false -Raw
			Write_Log (" TBS MAINTENANCE SCRIPT   v{0} ({1})   {2}" -f $VersionNumber, $VersionDate, $modeTag) -FileOnly -IncludeTimestamp:$false -Raw
			Write_Log $rule -FileOnly -IncludeTimestamp:$false -Raw
			
			if ($script:StoremanKeyStale)
			{
				$staleKeyDate = if ($script:StoremanKeyDate) { $script:StoremanKeyDate.ToString('MM/dd/yyyy') } else { 'unknown' }
				$storeDataUsable = $false
				try { $storeDataUsable = Test-Path (Join-Path $BasePath 'Office\Dbs') } catch { }
				Write_Log ("[WARN] KEY EXPIRED - storeman key is older than 14 days (KeyValidDate: {0}). Re-key required." -f $staleKeyDate) "orange" -FileOnly -IncludeTimestamp:$false -Raw
				if ($storeDataUsable)
				{
					Write_Log "       Store data (storeman folder + store database) IS available and in use: $BasePath" "orange" -FileOnly -IncludeTimestamp:$false -Raw
					Write_Log "       Only the key needs refreshing - every other task works normally." "orange" -FileOnly -IncludeTimestamp:$false -Raw
				}
				else
				{
					Write_Log "       Tasks that need the storeman folder or the store database are unavailable." "orange" -FileOnly -IncludeTimestamp:$false -Raw
					Write_Log ("       To enable them: run on a store PC, or create the local test folder ({0}\Office\Dbs)." -f $BasePath) -FileOnly -IncludeTimestamp:$false -Raw
				}
			}
			elseif ($script:IsTestEnvironment)
			{
				Write_Log "[WARN] TEST ENVIRONMENT - $($script:StoremanSourceReason)" "orange" -FileOnly -IncludeTimestamp:$false -Raw
				Write_Log "       Tasks that need the storeman folder or the store database are unavailable." -FileOnly -IncludeTimestamp:$false -Raw
				Write_Log "       Everything else works normally." -FileOnly -IncludeTimestamp:$false -Raw
				Write_Log ("       To enable them: run on a store PC, or create the local test folder ({0}\Office\Dbs)." -f $BasePath) -FileOnly -IncludeTimestamp:$false -Raw
			}
			else
			{
				Write_Log "[ OK ] PRODUCTION - storeman share 'storeman' in use" "green" -FileOnly -IncludeTimestamp:$false -Raw
				Write_Log "       Path: $BasePath" -FileOnly -IncludeTimestamp:$false -Raw
				if ($storeNum -and $storeNum -ne 'N/A')
				{
					$storeLine = "       Store: $storeNum"
					if ($storeName -and $storeName -ne 'N/A') { $storeLine += " `"$storeName`"" }
					if ($smsVer -and $smsVer -ne 'N/A') { $storeLine += "   (SMS $smsVer)" }
					Write_Log $storeLine -FileOnly -IncludeTimestamp:$false -Raw
				}
			}
			
			Write_Log $rule -FileOnly -IncludeTimestamp:$false -Raw
			
			# Re-append the original startup lines with their colors
			foreach ($old in $oldLines)
			{
				try
				{
					if ($old[0] -and $old[0].Trim())
					{
						$logBox.SelectionStart = $logBox.TextLength
						$logBox.SelectionLength = 0
						$logBox.SelectionColor = $old[1]
						$logBox.AppendText($old[0] + "`r`n")
						$logBox.SelectionColor = $logBox.ForeColor
					}
				}
				catch { }
			}
			try
			{
				$logBox.SelectionStart = $logBox.TextLength
				[void]$logBox.ScrollToCaret()
			}
			catch { }
		}
		catch { }
	})

$toolTip.SetToolTip($logBox, "Log output. Right-click to clear.")
$logBox.Add_MouseUp({
		param ($sender, $eventArgs)
		if ($eventArgs.Button -eq [System.Windows.Forms.MouseButtons]::Right)
		{
			$logBox.Clear()
			if (Get-Command Write_Log -ErrorAction SilentlyContinue) { Write_Log "Log Cleared" }
		}
	})

# Layout that also positions the log frame + log box
$layoutLogArea = {
	if (-not $form -or $form.IsDisposed) { return }
	
	try
	{
		# Width aligned to top strip card
		$rowW = [math]::Max(200, $form.ClientSize.Width - (2 * $padSide))
		
		# Log top starts just below the top strip card
		$logTop = $topStripCard.Bottom + $gapCards
		
		# Bottom bound uses status strip height
		$bottomReserve = 0
		if ($statusStrip -and -not $statusStrip.IsDisposed) { $bottomReserve = $statusStrip.Height }
		
		# Compute available height for the log area
		$availH = $form.ClientSize.Height - $logTop - $bottomReserve - $gapCards
		$logH = [math]::Max($minLogH, $availH)
		
		# Frame
		$logFrame.Left = $padSide
		$logFrame.Top = $logTop
		$logFrame.Width = $rowW
		$logFrame.Height = $logH
		
		# Log box inset (padding inside frame) - account for header
		$logBox.Left = $logFrame.Left + $logInnerPad
		$logBox.Top = 28 + $logInnerPad # Account for header height
		$logBox.Width = [math]::Max(50, $logFrame.Width - (2 * $logInnerPad))
		$logBox.Height = [math]::Max(50, $logFrame.Height - 28 - (2 * $logInnerPad))
		
		# Keep top strip aligned
		& $layoutTopStrip
	}
	catch { }
}

# Run once + re-run on resize
& $layoutLogArea
$form.add_Resize({ & $layoutLogArea })

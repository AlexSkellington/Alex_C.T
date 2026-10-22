# ===================================================================================================
#                              FUNCTION: Deploy_Scale_Currency_Files
# ---------------------------------------------------------------------------------------------------
# Description:
#   Currency deployment for Bizerba scales: pick scales in the node selector, enter the
#   currency symbol (default "$", common symbols in a combo), then copy the .txt/.properties
#   currency files (all contents written inline, no external dependencies) to each scale:
#     \\<ScaleIP>\c$\bizstorecard\bizerba\_fileIO\generic_data\in
#
# Parameters:
#   (none)
#
# Variables:
#   - $script:FunctionResults: StoreNumber (node lookup)
#
# Notes:
#   - Files are written UTF-8 without BOM so no BOM appears before the currency symbol.
#   - Offers to reboot the successfully deployed scales (via Reboot_Nodes).
# ===================================================================================================

function Deploy_Scale_Currency_Files
{
	Write_Log "`r`n==================== Starting Deploy_Scale_Currency_Files ====================`r`n" "blue"
	
	# ---- Node selection for scales (show tabbed GUI) ----
	$StoreNumber = $script:FunctionResults['StoreNumber']
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Scale" -OnlyBizerbaScales
	if (-not $selection -or -not $selection.Scales -or $selection.Scales.Count -eq 0)
	{
		Write_Log "No scales selected for deployment." "yellow"
		return
	}
	
	# ---- Load Assemblies ----
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# ---- Create Form ----
	$form = New-Object Windows.Forms.Form
	$form.Text = "Set Currency Symbol"
	$form.Size = [System.Drawing.Size]::new(420, 250)
	$form.FormBorderStyle = 'FixedDialog'
	$form.StartPosition = 'CenterScreen'
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
	$form.Icon = [System.Drawing.SystemIcons]::Information
	
	# ---- Centered, bold, wrapped label ----
	$lbl = New-Object Windows.Forms.Label
	$lbl.Text = "Please enter the currency symbol to use`nfor all scale price files:"
	$lbl.Font = New-Object System.Drawing.Font("Segoe UI", [float]11, [System.Drawing.FontStyle]::Bold)
	$lbl.Width = $form.ClientSize.Width - 40
	$lbl.Height = 50
	$lbl.Location = New-Object System.Drawing.Point(20, 20)
	$lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$lbl.AutoSize = $false
	$lbl.ForeColor = [System.Drawing.Color]::FromArgb(33, 33, 33)
	$form.Controls.Add($lbl)
	
	# ---- ComboBox for common symbols ----
	$cmbCurrency = New-Object Windows.Forms.ComboBox
	$cmbCurrency.Items.AddRange(@('$', [char]0x20AC, [char]0x00A3, [char]0x00A5, [char]0x20B9, [char]0x20BD, [char]0x20A9, [char]0x20BA, [char]0x20AA))
	$cmbCurrency.Text = '$'
	$cmbCurrency.Font = New-Object System.Drawing.Font("Segoe UI", [float]13, [System.Drawing.FontStyle]::Regular)
	$cmbCurrency.DropDownStyle = 'DropDown'
	$cmbCurrency.MaxLength = 3
	$cmbCurrency.Width = 100
	$cmbCurrency.Height = 30
	$cmbCurrency.Location = [System.Drawing.Point]::new([Math]::Floor(($form.ClientSize.Width - $cmbCurrency.Width)/2), 80)
	$cmbCurrency.FlatStyle = 'Flat'
	$form.Controls.Add($cmbCurrency)
	
	# ---- Tooltip ----
	$toolTip = New-Object System.Windows.Forms.ToolTip
	$toolTip.SetToolTip($cmbCurrency, "Select a common symbol or type a custom one (up to 3 characters).")
	
	# ---- Preview Label ----
	$previewLabel = New-Object Windows.Forms.Label
	$previewLabel.Text = "Preview: $1.99"
	$previewLabel.Font = New-Object System.Drawing.Font("Segoe UI", [float]12, [System.Drawing.FontStyle]::Italic)
	$previewLabel.Width = $form.ClientSize.Width - 40
	$previewLabel.Height = 30
	$previewLabel.Location = New-Object System.Drawing.Point(20, 120)
	$previewLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$previewLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
	$form.Controls.Add($previewLabel)
	
	# ---- Update preview on text change ----
	$cmbCurrency.Add_TextChanged({
			$symbol = ($this.Text).Trim()
			if ([string]::IsNullOrWhiteSpace($symbol)) { $symbol = '$' }
			$previewLabel.Text = "Preview: $($symbol)1.99"
		})
	$cmbCurrency.Add_Leave({
			if ([string]::IsNullOrWhiteSpace($cmbCurrency.Text)) { $cmbCurrency.Text = '$' }
		})
	$form.Add_Shown({
			if ([string]::IsNullOrWhiteSpace($cmbCurrency.Text)) { $cmbCurrency.Text = '$' }
			$previewLabel.Text = "Preview: $($cmbCurrency.Text.Trim())1.99"
		})
	
	# ---- OK / Cancel ----
	$btnWidth = 100; $btnHeight = 36; $btnSpacing = 24; $btnY = 170
	$totalBtnWidth = $btnWidth * 2 + $btnSpacing
	$startX = [Math]::Floor(($form.ClientSize.Width - $totalBtnWidth)/2)
	
	$btnOK = New-Object Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Size = [System.Drawing.Size]::new($btnWidth, $btnHeight)
	$btnOK.Location = [System.Drawing.Point]::new($startX, $btnY)
	$btnOK.Font = New-Object System.Drawing.Font("Segoe UI", [float]10, [System.Drawing.FontStyle]::Bold)
	$btnOK.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)
	$btnOK.ForeColor = [System.Drawing.Color]::White
	$btnOK.FlatStyle = 'Flat'
	$btnOK.FlatAppearance.BorderSize = 0
	$btnOK.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 103, 173) })
	$btnOK.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204) })
	$btnOK.Add_Click({
			if ([string]::IsNullOrWhiteSpace($cmbCurrency.Text))
			{
				[System.Windows.Forms.MessageBox]::Show("Please enter a currency symbol.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
				return
			}
			$form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close()
		})
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)
	
	$btnCancel = New-Object Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Size = [System.Drawing.Size]::new($btnWidth, $btnHeight)
	$btnCancel.Location = [System.Drawing.Point]::new($startX + $btnWidth + $btnSpacing, $btnY)
	$btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
	$btnCancel.BackColor = [System.Drawing.Color]::FromArgb(232, 72, 85)
	$btnCancel.ForeColor = [System.Drawing.Color]::White
	$btnCancel.FlatStyle = 'Flat'
	$btnCancel.FlatAppearance.BorderSize = 0
	$btnCancel.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(196, 61, 72) })
	$btnCancel.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(232, 72, 85) })
	$btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)
	
	$form.Add_Resize({
			$lbl.Width = $form.ClientSize.Width - 40
			$cmbCurrency.Location = [System.Drawing.Point]::new([Math]::Floor(($form.ClientSize.Width - $cmbCurrency.Width)/2), 80)
			$previewLabel.Width = $form.ClientSize.Width - 40
			$totalBtnWidth = $btnWidth * 2 + $btnSpacing
			$startX = [Math]::Floor(($form.ClientSize.Width - $totalBtnWidth)/2)
			$btnOK.Location = [System.Drawing.Point]::new($startX, $btnY)
			$btnCancel.Location = [System.Drawing.Point]::new($startX + $btnWidth + $btnSpacing, $btnY)
		})
	
	# ---- Show the form ----
	$res = $form.ShowDialog()
	if ($res -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "Cancelled by user." "yellow"
		return
	}
	$currency = $cmbCurrency.Text.Trim()
	if (-not $currency) { $currency = '$' }
	
	# ---- Inline file content templates ----
	# NOTE: These are what get *inlined* by the scale. If a BOM is present at the start of the file,
	#       the BOM shows up right before the first visible token (your currency). Hence No-BOM writes.
	$totalprice_txt = "=$%$ BT20 =`"$currency *#.##`""
	$unitprice_txt = "=$%$ BT10 =`"$currency *#.##`""
	
	$totalprice_properties = @"
<?xml version="1.0" encoding="UTF-8"?>
<properties>
    <source type="FILE">
        <path>$($script:BizStoreCardGenericInputPath)\na_f_totalprice.txt</path>
        <result>VALUE</result>
    </source>
</properties>
"@
	
	$unitprice_properties = @"
<?xml version="1.0" encoding="UTF-8"?>
<properties>
    <source type="FILE">
        <path>$($script:BizStoreCardGenericInputPath)\na_f_unitprice.txt</path>
        <result>VALUE</result>
    </source>
</properties>
"@
	
	# ---- Push files to each selected scale ----
	$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
	$results = @()
	
	foreach ($scaleCode in $selection.Scales)
	{
		$scaleObj = $ScaleCodeToIPInfo[$scaleCode]
		$scaleIP = $scaleObj.FullIP
		$scaleLabel = $scaleObj.ScaleName
		$targetPath = "\\$scaleIP\c$\bizstorecard\bizerba\_fileIO\generic_data\in"
		
		$result = [PSCustomObject]@{
			Scale = "$scaleLabel [$scaleIP]"
			ScaleCode = $scaleCode
			ScaleIP = $scaleIP
			Result = "Success"
			Details = @()
		}
		
		$isAccessible = $false
		try
		{
			if (Test-Path $targetPath -ErrorAction Stop) { $isAccessible = $true }
		}
		catch
		{
			Write_Log "Scale $scaleLabel [$scaleIP] is offline or share inaccessible. Skipping." "yellow"
			$result.Result = "Failed"
			$result.Details = @("Share not reachable")
			$results += ,$result
			continue
		}
		
		if (-not $isAccessible)
		{
			try
			{
				New-Item -Path $targetPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
				$isAccessible = $true
			}
			catch
			{
				Write_Log "Could not create remote folder on $scaleLabel [$scaleIP]. Skipping." "yellow"
				$result.Result = "Failed"
				$result.Details = @("Failed to create share folder")
				$results += ,$result
				continue
			}
		}
		
		# ==========================================================================================
		# CHANGED: Write files as UTF-8 **without BOM** so the first bytes are the actual content.
		#          Using .NET API here because PS5.1 `-Encoding UTF8` includes a BOM.
		# ==========================================================================================
		try
		{
			$utf8NoBom = New-Object System.Text.UTF8Encoding($false) # $false => NO BOM
			
			[System.IO.File]::WriteAllText((Join-Path $targetPath 'na_f_totalprice.txt'), $totalprice_txt, $utf8NoBom) # No BOM
			[System.IO.File]::WriteAllText((Join-Path $targetPath 'na_f_unitprice.txt'), $unitprice_txt, $utf8NoBom) # No BOM
			[System.IO.File]::WriteAllText((Join-Path $targetPath 'na_f_totalprice.properties'), $totalprice_properties, $utf8NoBom) # No BOM
			[System.IO.File]::WriteAllText((Join-Path $targetPath 'na_f_unitprice.properties'), $unitprice_properties, $utf8NoBom) # No BOM
			
			# Optional: ensure CRLF line endings (Windows) on the XML (usually fine as-is).
			# If your environment requires explicit CRLF normalization, uncomment:
			# $crlf = "`r`n"
			# [System.IO.File]::WriteAllText((Join-Path $targetPath 'na_f_totalprice.properties'),
			#     ($totalprice_properties -replace "`r?`n",$crlf), $utf8NoBom)
			# [System.IO.File]::WriteAllText((Join-Path $targetPath 'na_f_unitprice.properties'),
			#     ($unitprice_properties -replace "`r?`n",$crlf), $utf8NoBom)
			
			$result.Details += "Files copied successfully (UTF-8 without BOM)."
			Write_Log "Deployed price files to $($scaleLabel) [$scaleIP] with Currency ($currency)" "green"
		}
		catch
		{
			$result.Result = "Failed"
			$result.Details = @("File write failed (network/permission).")
			Write_Log "Could not deploy price files to $($scaleLabel) [$scaleIP]. File write failed or network/permission denied." "yellow"
		}
		
		$results += ,$result
	}
	
	# ---- Show summary ----
	Write_Log "Deployment summary:" "Magenta"
	foreach ($r in $results)
	{
		$msg = "$($r.Scale): $($r.Result) - $($r.Details -join '; ')"
		if ($r.Result -eq "Success") { Write_Log $msg "green" }
		else { Write_Log $msg "red" }
	}
	
	# ---- Prompt for reboot of only SUCCESSFUL scales ----
	$successScales = $results | Where-Object { $_.Result -eq "Success" -and $_.PSObject.Properties.Match('ScaleCode') } | Select-Object -ExpandProperty ScaleCode
	$successScalesLabels = $results | Where-Object { $_.Result -eq "Success" } | ForEach-Object { $_.Scale }
	
	if ($successScales.Count -gt 0)
	{
		$scaleListText = ($successScalesLabels -join "`n")
		$rebootMsg = "Do you want to reboot the following successfully deployed scales now to apply changes?`n`n$scaleListText"
		$dialogResult = [System.Windows.Forms.MessageBox]::Show($rebootMsg, "Reboot Scales?", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
		
		if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			Write_Log "User chose to reboot all successfully deployed scales." "cyan"
			$rebootSelection = [PSCustomObject]@{ Lanes = @(); Scales = $successScales; Backoffices = @() }
			Reboot_Nodes -StoreNumber $StoreNumber -NodeTypes Scale -Selection $rebootSelection
		}
		else
		{
			Write_Log "User chose not to reboot. The following scales will need to be rebooted to apply changes:" "yellow"
			foreach ($s in $successScalesLabels) { Write_Log "$s" "yellow" }
		}
	}
	
	Write_Log "`r`n==================== Deploy_Scale_Currency_Files Completed ====================" "blue"
}

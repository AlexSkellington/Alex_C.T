# ===================================================================================================
# 									FUNCTION: Get_NEW_Store_Number
# ---------------------------------------------------------------------------------------------------
# Description:
# Prompts the user via a GUI to enter a valid store number.
# Accepts exactly 3 or 4 digits (no leading zeros required, not all zeros).
# Returns the entered value as-is (no padding) or $null if cancelled.
# ===================================================================================================

function Get_NEW_Store_Number
{
	while ($true)
	{
		# Decide required store length based on CURRENT store number length (3 vs 4)
		$requiredLen = 4
		
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('StoreNumber') -and $script:FunctionResults['StoreNumber'])
		{
			$cur = ($script:FunctionResults['StoreNumber'].ToString()).Trim()
			if ($cur -match '^\d{3,4}$') { $requiredLen = $cur.Length }
		}
		else
		{
			$iniStore = Get_Store_Number_From_INI # returns as-is (3 or 4) if your function is set that way
			if ($iniStore)
			{
				$iniStore = ($iniStore.ToString()).Trim()
				if ($iniStore -match '^\d{3,4}$') { $requiredLen = $iniStore.Length }
			}
		}
		
		$storeNumberForm = New-Object System.Windows.Forms.Form
		$storeNumberForm.Text = "Enter New Store Number"
		$storeNumberForm.ClientSize = New-Object System.Drawing.Size(340, 135)
		$storeNumberForm.StartPosition = "CenterParent"
		$storeNumberForm.FormBorderStyle = 'FixedDialog'
		$storeNumberForm.MaximizeBox = $false
		$storeNumberForm.MinimizeBox = $false
		$storeNumberForm.ShowInTaskbar = $false
		
		$label = New-Object System.Windows.Forms.Label
		$label.Text = "New Store Number (exactly $requiredLen digits, e.g. " + ($(if ($requiredLen -eq 3) { "123" }
				else { "4123" })) + "):"
		$label.Location = New-Object System.Drawing.Point(10, 20)
		$label.Size = New-Object System.Drawing.Size(315, 40)
		$label.AutoSize = $false
		
		$textBox = New-Object System.Windows.Forms.TextBox
		$textBox.Location = New-Object System.Drawing.Point(10, 65)
		$textBox.Size = New-Object System.Drawing.Size(320, 20)
		$textBox.MaxLength = $requiredLen
		
		$okButton = New-Object System.Windows.Forms.Button
		$okButton.Text = "OK"
		$okButton.Location = New-Object System.Drawing.Point(90, 98)
		$okButton.Size = New-Object System.Drawing.Size(75, 23)
		
		$cancelButton = New-Object System.Windows.Forms.Button
		$cancelButton.Text = "Cancel"
		$cancelButton.Location = New-Object System.Drawing.Point(175, 98)
		$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
		
		$dialogSelection = $null
		$okButton.Add_Click({
				$script:MiniGhostTempDialogSelection = 'OK'
				$storeNumberForm.Close()
			})
		$cancelButton.Add_Click({
				$script:MiniGhostTempDialogSelection = 'Cancel'
				$storeNumberForm.Close()
			})
		
		$storeNumberForm.Controls.AddRange(@($label, $textBox, $okButton, $cancelButton))
		$storeNumberForm.AcceptButton = $okButton
		$storeNumberForm.CancelButton = $cancelButton
		
		$script:MiniGhostTempDialogSelection = $null
		[void]$storeNumberForm.ShowDialog()
		
		# IMPORTANT: read before Dispose (PS WinForms reliability)
		$userInput = [string]$textBox.Text
		$dialogSelection = $script:MiniGhostTempDialogSelection
		$script:MiniGhostTempDialogSelection = $null
		$storeNumberForm.Dispose()
		
		if ($dialogSelection -ne 'OK')
		{
			return $null
		}
		
		if ($null -eq $userInput) { $userInput = "" }
		$userInput = $userInput.Trim()
		
		if ([string]::IsNullOrEmpty($userInput))
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Store number cannot be empty.",
				"Error",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			continue
		}
		
		# Build regex based on required length; reject all zeros for that length
		$re = '^(?!0{' + $requiredLen + '})\d{' + $requiredLen + '}$'
		
		if ($userInput -match $re)
		{
			if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
			$script:FunctionResults['StoreNumber'] = $userInput
			return $userInput # return as-is, no padding
		}
		else
		{
			[System.Windows.Forms.MessageBox]::Show(
				"Invalid store number.`n`nMust be exactly $requiredLen digits (numeric).`nNot allowed: all zeros, too short, too long.`nExample: " + ($(if ($requiredLen -eq 3) { "123" }
						else { "4123" })),
				"Invalid Input",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
		}
	}
}

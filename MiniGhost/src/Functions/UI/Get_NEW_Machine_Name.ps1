# ===================================================================================================
# 									FUNCTION: Get_NEW_Machine_Name
# ---------------------------------------------------------------------------------------------------
# Description:
# Prompts user for a new machine name via GUI.
# Supports:
#   - Classic: PREFIX + 3 digits, with optional hyphen before digits (e.g. LANE003, POS-001, SERVER-001)
#   - Prefixed: optional 1-4 digit store + PREFIX + optional hyphen + 3 digits (e.g. 0231LANE006, 0231POS-001)
# If prefixed format used -> offers to sync store number
# Returns uppercase validated name or $null if cancelled
# ===================================================================================================

function Get_NEW_Machine_Name
{
	while ($true)
	{
		$nodeRole = "Unknown"
		if (-not [string]::IsNullOrWhiteSpace([string]$script:NodeRole))
		{
			$nodeRole = ([string]$script:NodeRole).Trim()
		}
		
		$machineNamePromptText = if ($nodeRole -match '^(?i)(StoreServer|HostServer)$')
		{
			"Server examples:`nSERVER001  -  SERVER-001  -  0231SERVER001`nRe-enter the current machine name to force a full INI + SQL sync."
		}
		elseif ($nodeRole -match '^(?i)Lane$')
		{
			"Lane examples:`nPOS001  -  POS-001  -  LN026001  -  0384LANE001  -  001`nRe-enter the current machine name to force a full INI + SQL sync."
		}
		else
		{
			"Lane examples:`nPOS001  -  POS-001  -  LN026001  -  0384LANE001  -  001`nServer examples:`nSERVER001  -  SERVER-001  -  0231SERVER001`nRe-enter the current machine name to force a full INI + SQL sync."
		}
		$promptLineCount = ($machineNamePromptText -split "`n").Count
		$labelHeight = [Math]::Max(54, (18 * $promptLineCount) + 6)
		$formClientWidth = 430
		$formClientHeight = $labelHeight + 112
		
		$form = New-Object System.Windows.Forms.Form
		$form.Text = "Enter New Machine Name"
		$form.ClientSize = New-Object System.Drawing.Size($formClientWidth, $formClientHeight)
		$form.StartPosition = "CenterParent"
		$form.FormBorderStyle = 'FixedDialog'
		$form.MaximizeBox = $false
		$form.MinimizeBox = $false
		$form.ShowInTaskbar = $false
		
		$label = New-Object System.Windows.Forms.Label
		$label.Text = $machineNamePromptText
		$label.Location = New-Object System.Drawing.Point(10, 15)
		$label.Size = New-Object System.Drawing.Size(($formClientWidth - 20), $labelHeight)
		$label.AutoSize = $false
		
		$textBox = New-Object System.Windows.Forms.TextBox
		$textBox.Location = New-Object System.Drawing.Point(10, ($label.Bottom + 10))
		$textBox.Size = New-Object System.Drawing.Size(($formClientWidth - 20), 22)
		
		$okBtn = New-Object System.Windows.Forms.Button
		$okBtn.Text = "OK"
		$okBtn.Location = New-Object System.Drawing.Point(([int](($formClientWidth / 2) - 75)), ($textBox.Bottom + 14))
		$okBtn.Size = New-Object System.Drawing.Size(70, 28)
		$okBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
		
		$cancelBtn = New-Object System.Windows.Forms.Button
		$cancelBtn.Text = "Cancel"
		$cancelBtn.Location = New-Object System.Drawing.Point(($okBtn.Right + 15), $okBtn.Top)
		$cancelBtn.Size = New-Object System.Drawing.Size(70, 28)
		$cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		
		$form.Controls.AddRange(@($label, $textBox, $okBtn, $cancelBtn))
		$form.AcceptButton = $okBtn
		$form.CancelButton = $cancelBtn
		
		$dialogResult = $form.ShowDialog()
		
		# PS 5.1-safe: read textbox BEFORE disposing
		$userInputRaw = $textBox.Text
		$form.Dispose()
		
		if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
		{
			return $null
		}
		
		$userInput = $userInputRaw
		if ($null -eq $userInput) { $userInput = "" }
		$userInput = $userInput.Trim().ToUpper()
		
		if ([string]::IsNullOrEmpty($userInput))
		{
			[void][System.Windows.Forms.MessageBox]::Show("Machine name cannot be empty.", "Error")
			continue
		}
		
		if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
		
		$currentMachineForStyle = Get_Effective_Machine_Name
		$currentStore = $null
		if ($script:FunctionResults.ContainsKey('StoreNumber') -and $script:FunctionResults['StoreNumber'])
		{
			$currentStore = ($script:FunctionResults['StoreNumber'].ToString()).Trim()
		}
		else
		{
			$currentStore = Get_Store_Number_From_INI
			if ($currentStore) { $currentStore = ($currentStore.ToString()).Trim() }
		}
		
		$machineContext = Get_Machine_Name_Context -MachineName $userInput -CurrentMachineName $currentMachineForStyle
		if (-not $machineContext.Success)
		{
			[void][System.Windows.Forms.MessageBox]::Show(
				$machineContext.ErrorMessage,
				"Invalid Machine Name",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Error
			)
			continue
		}
		
		$normalizedName = $machineContext.MachineName
		$storePrefixFromName = $machineContext.StorePrefix
		$currentStoreComparable = $null
		if ($currentStore -match '^\d{3,4}$')
		{
			$currentStoreComparable = [int]$currentStore
		}
		
		if ($storePrefixFromName -and ($storePrefixFromName -match '^\d{3,4}$'))
		{
			$storePrefixComparable = [int]$storePrefixFromName
			if (($null -eq $currentStoreComparable) -or ($storePrefixComparable -ne $currentStoreComparable))
			{
				$displayStore = $currentStore
				if ([string]::IsNullOrEmpty($displayStore)) { $displayStore = "UNKNOWN" }
				
				$sync = [System.Windows.Forms.MessageBox]::Show(
					"Machine name has store prefix '$storePrefixFromName',`nbut current store is '$displayStore'.`n`nUpdate store number to '$storePrefixFromName'?",
					"Store Number Mismatch",
					[System.Windows.Forms.MessageBoxButtons]::YesNo,
					[System.Windows.Forms.MessageBoxIcon]::Question
				)
				
				if ($sync -eq [System.Windows.Forms.DialogResult]::Yes)
				{
					$iniParams = @{
						newStoreNumber = $storePrefixFromName
						MachineName    = $normalizedName
						OldMachineName = $currentMachineForStyle
					}
					if ($currentStore -match '^\d{3,4}$') { $iniParams["OldStoreNumber"] = $currentStore }
					
					$success = Update_INIs @iniParams
					
					if ($success)
					{
						$script:newStoreNumber = $storePrefixFromName
						$script:FunctionResults['StoreNumber'] = $storePrefixFromName
						
						if ($storeNumberLabel)
						{
							$storeNumberLabel.Text = "Store Number: $storePrefixFromName (updated)"
							$storeNumberLabel.Refresh()
						}
					}
					else
					{
						[void][System.Windows.Forms.MessageBox]::Show("Failed to update store number in INI.", "Error")
						continue
					}
				}
			}
		}
		
		return $normalizedName
	}
}

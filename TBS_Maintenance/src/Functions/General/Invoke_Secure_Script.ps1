# ===================================================================================================
#                              FUNCTION: Invoke_Secure_Script
# ---------------------------------------------------------------------------------------------------
# Description:
#   DISABLED as of v2.6.9: this used to prompt for a hard-coded password and execute a remote
#   script via Invoke-Expression (irm <url>). The remote execution and the password were
#   removed for security; the function now logs the disabled state and returns. The
#   "Activate Windows" menu item stays so the reason is visible in the log.
#
# Notes:
#   - The old password prompt and URL download code below the disabled check is inert.
# ===================================================================================================

function Invoke_Secure_Script
{
	[CmdletBinding()]
	param ()
	
	# --- Security hardening: remote script execution disabled ---
	Write_Log "`r`n==================== Starting Invoke_Secure_Script Function ====================`r`n" "blue"
	Write_Log "Invoke_Secure_Script is disabled: remote Invoke-Expression/irm execution was removed for security." "red"
	Write_Log "If this workflow is required, replace it with a signed local script and hash verification." "yellow"
	Write_Log "`r`n==================== Invoke_Secure_Script Function Completed ====================" "blue"
	return
	
	# --- Password Prompt (GUI) ---
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Authentication Required"
	$form.Size = New-Object System.Drawing.Size(350, 150)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = 'FixedDialog'
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Please enter the password to proceed:"
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(10, 20)
	$form.Controls.Add($label)
	
	$textbox = New-Object System.Windows.Forms.TextBox
	$textbox.Location = New-Object System.Drawing.Point(10, 50)
	$textbox.Size = New-Object System.Drawing.Size(310, 20)
	$textbox.UseSystemPasswordChar = $true
	$form.Controls.Add($textbox)
	
	$buttonOK = New-Object System.Windows.Forms.Button
	$buttonOK.Text = "OK"
	$buttonOK.Location = New-Object System.Drawing.Point(160, 80)
	$buttonOK.Size = New-Object System.Drawing.Size(80, 30)
	$buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $buttonOK
	$form.Controls.Add($buttonOK)
	
	$buttonCancel = New-Object System.Windows.Forms.Button
	$buttonCancel.Text = "Cancel"
	$buttonCancel.Location = New-Object System.Drawing.Point(240, 80)
	$buttonCancel.Size = New-Object System.Drawing.Size(80, 30)
	$buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $buttonCancel
	$form.Controls.Add($buttonCancel)
	
	$dialogResult = $form.ShowDialog()
	$password = if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) { $textbox.Text }
	else { $null }
	
	if (-not $password)
	{
		Write_Log "User canceled the authentication prompt." "yellow"
		Write_Log "`r`n==================== Invoke_Secure_Script Function Completed ====================" "blue"
		return
	}
	
	# --- Verify Password ---
	if ($password -ne $storedPassword)
	{
		Write_Log "Authentication failed. Incorrect password." "red"
		Write_Log "`r`n==================== Invoke_Secure_Script Function Completed ====================" "blue"
		return
	}
	
	Write_Log "Authentication successful. Proceeding with script execution." "green"
	
	# --- Execute Script from URL ---
	# Remote execution intentionally removed. This block is inert.
	Write_Log "Remote script execution remains disabled." "gray"
	
	# --- Log End ---
	Write_Log "`r`n==================== Invoke_Secure_Script Function Completed ====================" "blue"
}

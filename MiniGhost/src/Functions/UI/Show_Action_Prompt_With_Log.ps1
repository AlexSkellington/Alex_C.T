# ===================================================================================================
#                              FUNCTION: Show_Action_Prompt_With_Log
# ---------------------------------------------------------------------------------------------------
# Description:
#   Message dialog with one or two action buttons and an optional View Log button that opens the
#   given log text in Show_Text_Report_Form.
#
# Returns:
#   "Primary", "Secondary" or "Closed" (window closed with X).
#
# Notes:
#   - Click handlers run in their own scope: a plain $selection = ... inside them would set a new
#     local variable and the function would always return "Closed". The choice is carried on the
#     form's .Tag instead.
# ===================================================================================================

function Show_Action_Prompt_With_Log
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$Text,
		[Parameter(Mandatory = $true)]
		[string]$Title,
		[Parameter(Mandatory = $false)]
		[string]$PrimaryButtonText = "OK",
		[Parameter(Mandatory = $false)]
		[string]$SecondaryButtonText,
		[Parameter(Mandatory = $false)]
		[string]$LogText,
		[Parameter(Mandatory = $false)]
		[string]$LogTitle = "Details"
	)
	
	$promptForm = New-Object System.Windows.Forms.Form
	$promptForm.Text = $Title
	$promptForm.Size = New-Object System.Drawing.Size(620, 230)
	$promptForm.StartPosition = "CenterParent"
	$promptForm.FormBorderStyle = "FixedDialog"
	$promptForm.MaximizeBox = $false
	$promptForm.MinimizeBox = $false
	$promptForm.ShowInTaskbar = $false
	
	$messageTextBox = New-Object System.Windows.Forms.TextBox
	$messageTextBox.Multiline = $true
	$messageTextBox.ReadOnly = $true
	$messageTextBox.BorderStyle = "None"
	$messageTextBox.BackColor = $promptForm.BackColor
	$messageTextBox.TabStop = $false
	$messageTextBox.Text = $Text
	$messageTextBox.SetBounds(16, 16, 570, 110)
	
	$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
	$buttonPanel.Dock = "Bottom"
	$buttonPanel.Height = 48
	$buttonPanel.FlowDirection = "RightToLeft"
	$buttonPanel.WrapContents = $false
	$buttonPanel.Padding = New-Object System.Windows.Forms.Padding(0, 6, 12, 6)
	
	$promptForm.Tag = "Closed"
	
	$primaryButton = New-Object System.Windows.Forms.Button
	$primaryButton.Text = $PrimaryButtonText
	$primaryButton.AutoSize = $true
	$primaryButton.Add_Click({
			$promptForm.Tag = "Primary"
			$promptForm.Close()
		})
	$buttonPanel.Controls.Add($primaryButton)
	$promptForm.AcceptButton = $primaryButton
	$promptForm.CancelButton = $primaryButton
	
	if (-not [string]::IsNullOrWhiteSpace($SecondaryButtonText))
	{
		$secondaryButton = New-Object System.Windows.Forms.Button
		$secondaryButton.Text = $SecondaryButtonText
		$secondaryButton.AutoSize = $true
		$secondaryButton.Add_Click({
				$promptForm.Tag = "Secondary"
				$promptForm.Close()
			})
		$buttonPanel.Controls.Add($secondaryButton)
		$promptForm.CancelButton = $secondaryButton
	}
	
	if (-not [string]::IsNullOrWhiteSpace($LogText))
	{
		$viewLogButton = New-Object System.Windows.Forms.Button
		$viewLogButton.Text = "View Log"
		$viewLogButton.AutoSize = $true
		$viewLogButton.Add_Click({
				Show_Text_Report_Form -Title $LogTitle -Text $LogText
			})
		$buttonPanel.Controls.Add($viewLogButton)
	}
	
	$promptForm.Controls.Add($messageTextBox)
	$promptForm.Controls.Add($buttonPanel)
	Set_Modern_Dialog_Style -Form $promptForm -PrimaryButton $primaryButton
	[void]$promptForm.ShowDialog()
	$selection = [string]$promptForm.Tag
	$promptForm.Dispose()
	return $selection
}

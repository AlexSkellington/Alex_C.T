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
	
	$selection = "Closed"
	
	$primaryButton = New-Object System.Windows.Forms.Button
	$primaryButton.Text = $PrimaryButtonText
	$primaryButton.AutoSize = $true
	$primaryButton.Add_Click({
			$selection = "Primary"
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
				$selection = "Secondary"
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
	[void]$promptForm.ShowDialog()
	return $selection
}

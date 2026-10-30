function Show_Text_Report_Form
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$Title,
		[Parameter(Mandatory = $true)]
		[string]$Text,
		[Parameter(Mandatory = $false)]
		[int]$Width = 720,
		[Parameter(Mandatory = $false)]
		[int]$Height = 500
	)
	
	$reportForm = New-Object System.Windows.Forms.Form
	$reportForm.Text = $Title
	$reportForm.Size = New-Object System.Drawing.Size($Width, $Height)
	$reportForm.StartPosition = "CenterParent"
	$reportForm.ShowInTaskbar = $false
	
	$closeButton = New-Object System.Windows.Forms.Button
	$closeButton.Text = "Close"
	$closeButton.Dock = "Bottom"
	$closeButton.Height = 34
	$closeButton.Add_Click({ $reportForm.Close() })
	
	$textBox = New-Object System.Windows.Forms.TextBox
	$textBox.Multiline = $true
	$textBox.ReadOnly = $true
	$textBox.ScrollBars = "Vertical"
	$textBox.Dock = "Fill"
	$textBox.Font = New-Object System.Drawing.Font("Consolas", 9)
	$textBox.Text = $Text
	
	$reportForm.AcceptButton = $closeButton
	$reportForm.Controls.Add($textBox)
	$reportForm.Controls.Add($closeButton)
	[void]$reportForm.ShowDialog()
}

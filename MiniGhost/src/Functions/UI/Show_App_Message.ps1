# ===================================================================================================
#                                 FUNCTION: Show_App_Message
# ---------------------------------------------------------------------------------------------------
# Description:
# Small wrapper around WinForms MessageBox to reduce repeated enum boilerplate.
# ===================================================================================================

function Show_App_Message
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$Text,
		[Parameter(Mandatory = $true)]
		[string]$Title,
		[Parameter(Mandatory = $false)]
		[ValidateSet('OK', 'OKCancel', 'YesNo', 'YesNoCancel')]
		[string]$Buttons = 'OK',
		[Parameter(Mandatory = $false)]
		[ValidateSet('None', 'Information', 'Warning', 'Error', 'Question', 'Asterisk', 'Exclamation', 'Hand', 'Stop')]
		[string]$Icon = 'Information'
	)
	
	switch -Regex ($Icon)
	{
		'^(?i)Asterisk$'    { $Icon = 'Information'; break }
		'^(?i)Exclamation$' { $Icon = 'Warning'; break }
		'^(?i)Hand$'        { $Icon = 'Error'; break }
		'^(?i)Stop$'        { $Icon = 'Error'; break }
	}
	
	$buttonValue = [System.Enum]::Parse([System.Windows.Forms.MessageBoxButtons], $Buttons)
	$iconValue = [System.Enum]::Parse([System.Windows.Forms.MessageBoxIcon], $Icon)
	return [System.Windows.Forms.MessageBox]::Show($Text, $Title, $buttonValue, $iconValue)
}

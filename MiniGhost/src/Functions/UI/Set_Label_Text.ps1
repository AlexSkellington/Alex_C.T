# ===================================================================================================
#                                  FUNCTION: Set_Label_Text
# ---------------------------------------------------------------------------------------------------
# Description:
# Updates a WinForms label and optionally refreshes its parent layout.
# ===================================================================================================

function Set_Label_Text
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		$Label,
		[Parameter(Mandatory = $true)]
		[string]$Text,
		[Parameter(Mandatory = $false)]
		[switch]$RefreshParent,
		[Parameter(Mandatory = $false)]
		[switch]$ProcessEvents
	)
	
	if (-not $Label) { return }
	
	$Label.Text = $Text
	$Label.Refresh()
	
	if ($RefreshParent -and $Label.Parent)
	{
		$Label.Parent.PerformLayout()
		$Label.Parent.Refresh()
	}
	
	if ($ProcessEvents)
	{
		[System.Windows.Forms.Application]::DoEvents()
	}
}

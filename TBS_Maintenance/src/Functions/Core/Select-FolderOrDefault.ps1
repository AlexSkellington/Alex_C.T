# ===================================================================================================
#                              FUNCTION: Select-FolderOrDefault
# ---------------------------------------------------------------------------------------------------
# Description:
#   Opens a folder picker (FolderBrowserDialog) pre-set to a starting folder and returns the
#   folder the user chose; if the user cancels it returns $null so the caller keeps its default.
#   Shared helper for the Local DB / Storeman ZIP / Lane DB backup schedulers' "Backup folder"
#   option, so the backup can be saved to the default location or a location of choice.
#
# Parameters:
#   - initialPath: folder the picker opens at (usually the default backup root)
#   - description: prompt text shown at the top of the picker
#
# Returns:
#   The chosen folder path, or $null if the user cancelled.
# ===================================================================================================

function Select-FolderOrDefault
{
	param
	(
		[string]$initialPath = "",
		[string]$description = "Choose a folder"
	)

	Add-Type -AssemblyName System.Windows.Forms
	$fd = New-Object System.Windows.Forms.FolderBrowserDialog
	$fd.Description = $description
	$fd.ShowNewFolderButton = $true
	if ($initialPath -and (Test-Path -LiteralPath $initialPath)) { $fd.SelectedPath = $initialPath }

	$res = $fd.ShowDialog()
	if ($res -eq [System.Windows.Forms.DialogResult]::OK -and $fd.SelectedPath)
	{
		return $fd.SelectedPath
	}
	return $null
}

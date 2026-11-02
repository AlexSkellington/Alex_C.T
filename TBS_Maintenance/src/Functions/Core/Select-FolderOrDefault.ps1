# ===================================================================================================
#                              FUNCTION: Select-FolderOrDefault
# ---------------------------------------------------------------------------------------------------
# Description:
#   Opens an Explorer-style folder picker pre-set to a starting folder and returns the folder the
#   user chose; if the user cancels it returns $null so the caller keeps its default. Shared helper
#   for the Local DB / Storeman ZIP / Lane DB backup schedulers' "Backup folder" option.
#
# Parameters:
#   - initialPath: folder the picker opens at (usually the default backup root)
#   - description: prompt text shown in the picker's title
#
# Returns:
#   The chosen folder path, or $null if the user cancelled.
#
# Notes:
#   - The Explorer dialog's address bar accepts typed or pasted paths, network ones included
#     (\\server\share\folder) - the old FolderBrowserDialog tree did not. It is an OpenFileDialog
#     that returns the folder it was in when "Open" was pressed ("Select this folder" as the name).
#   - Falls back to the FolderBrowserDialog tree if the Explorer dialog cannot be shown.
# ===================================================================================================

function Select-FolderOrDefault
{
	param
	(
		[string]$initialPath = "",
		[string]$description = "Choose a folder"
	)

	Add-Type -AssemblyName System.Windows.Forms
	$placeholder = 'Select this folder'
	try
	{
		$ofd = New-Object System.Windows.Forms.OpenFileDialog
		try
		{
			$ofd.Title = $description + "  (type or paste a path - \\server\share works - in the address bar, then Open)"
			$ofd.ValidateNames = $false
			$ofd.CheckFileExists = $false
			$ofd.CheckPathExists = $true
			$ofd.Multiselect = $false
			$ofd.FileName = $placeholder
			$ofd.Filter = 'Folders|*.this-shows-folders-only'
			if ($initialPath -and (Test-Path -LiteralPath $initialPath)) { $ofd.InitialDirectory = $initialPath }
			if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
			$picked = $ofd.FileName
			if ([System.IO.Path]::GetFileName($picked) -eq $placeholder -or -not (Test-Path -LiteralPath $picked -PathType Container))
			{
				$picked = [System.IO.Path]::GetDirectoryName($picked)
			}
			if ([string]::IsNullOrWhiteSpace($picked)) { return $null }
			return $picked.TrimEnd('\')
		}
		finally { $ofd.Dispose() }
	}
	catch
	{
		$fd = New-Object System.Windows.Forms.FolderBrowserDialog
		$fd.Description = $description
		$fd.ShowNewFolderButton = $true
		if ($initialPath -and (Test-Path -LiteralPath $initialPath)) { $fd.SelectedPath = $initialPath }
		try
		{
			$res = $fd.ShowDialog()
			if ($res -eq [System.Windows.Forms.DialogResult]::OK -and $fd.SelectedPath) { return $fd.SelectedPath }
			return $null
		}
		finally { $fd.Dispose() }
	}
}

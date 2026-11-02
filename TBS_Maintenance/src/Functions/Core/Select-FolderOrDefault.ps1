# ===================================================================================================
#                              FUNCTION: Select-FolderOrDefault
# ---------------------------------------------------------------------------------------------------
# Description:
#   Opens the Windows folder picker (the Explorer-style one with a "Select Folder" button) pre-set to
#   a starting folder and returns the folder the user chose; if the user cancels it returns $null
#   so the caller keeps its default. Shared helper for the backup schedulers' folder options.
#
# Parameters:
#   - initialPath: folder the picker opens at (usually the default backup root)
#   - description: window title
#   - ProbeOnly  : build the picker without showing it and return $true when it can be shown
#                  (used by the tests)
#
# Returns:
#   The chosen folder path, or $null if the user cancelled.
#
# Notes:
#   - The address bar and the "Folder:" box accept typed or pasted paths, network ones included
#     (\\server\share\folder); "Select Folder" returns it. .NET Framework only exposes this picker
#     (IFileDialog with FOS_PICKFOLDERS) through OpenFileDialog's internal Vista dialog, reached
#     here by reflection - no C# is compiled.
#   - When that is not available it falls back to the older FolderBrowserDialog tree.
#   - v2.7.9 used an OpenFileDialog with a "Select this folder" file name: typing a folder path
#     there only navigated into it, so a typed network path could not be selected.
# ===================================================================================================

function Select-FolderOrDefault
{
	param
	(
		[string]$initialPath = "",
		[string]$description = "Choose a folder",
		[switch]$ProbeOnly
	)

	Add-Type -AssemblyName System.Windows.Forms
	$flags = [System.Reflection.BindingFlags]'Instance, Static, Public, NonPublic'
	try
	{
		$asm = [System.Windows.Forms.FileDialog].Assembly
		$iFileDialog = $asm.GetType('System.Windows.Forms.FileDialogNative+IFileDialog', $true)
		$fosType = $asm.GetType('System.Windows.Forms.FileDialogNative+FOS', $true)
		$eventsType = $asm.GetType('System.Windows.Forms.FileDialog+VistaDialogEvents', $true)
		$createVista = [System.Windows.Forms.OpenFileDialog].GetMethod('CreateVistaDialog', $flags)
		$beforeVista = [System.Windows.Forms.FileDialog].GetMethod('OnBeforeVistaDialog', $flags)
		$eventsCtor = $eventsType.GetConstructor($flags, $null, [type[]]@([System.Windows.Forms.FileDialog]), $null)
		if (-not $createVista -or -not $beforeVista -or -not $eventsCtor) { throw "folder picker internals not found" }

		$ofd = New-Object System.Windows.Forms.OpenFileDialog
		try
		{
			$ofd.Title = $description
			$ofd.AddExtension = $false
			$ofd.CheckFileExists = $false
			$ofd.CheckPathExists = $true
			$ofd.DereferenceLinks = $true
			$ofd.Multiselect = $false
			if ($initialPath -and (Test-Path -LiteralPath $initialPath)) { $ofd.InitialDirectory = $initialPath }

			$dialog = $createVista.Invoke($ofd, $null)
			$null = $beforeVista.Invoke($ofd, @($dialog))
			$optArgs = @($null)
			$null = $iFileDialog.GetMethod('GetOptions').Invoke($dialog, $optArgs)
			$withFolders = [Enum]::ToObject($fosType, ([uint32]$optArgs[0] -bor [uint32]0x20)) # FOS_PICKFOLDERS
			$null = $iFileDialog.GetMethod('SetOptions').Invoke($dialog, @($withFolders))
			if ($ProbeOnly) { return $true }

			# The events object copies the picked folder into $ofd.FileName when "Select Folder" is pressed
			$sink = $eventsCtor.Invoke(@($ofd))
			$adviseArgs = @($sink, [uint32]0)
			$null = $iFileDialog.GetMethod('Advise').Invoke($dialog, $adviseArgs)
			$owner = [IntPtr]::Zero
			# owned by the window that is open now (the backup dialog), so the picker opens on top of it
			try
			{
				$active = [System.Windows.Forms.Form]::ActiveForm
				if (-not $active) { $active = $script:MainForm }
				if ($active -and -not $active.IsDisposed) { $owner = $active.Handle }
			}
			catch { }
			try { $hr = [int]$iFileDialog.GetMethod('Show').Invoke($dialog, @($owner)) }
			finally { try { $null = $iFileDialog.GetMethod('Unadvise').Invoke($dialog, @($adviseArgs[1])) } catch { } }
			if ($hr -ne 0 -or [string]::IsNullOrWhiteSpace($ofd.FileName)) { return $null }
			return $ofd.FileName.TrimEnd('\')
		}
		finally { $ofd.Dispose() }
	}
	catch
	{
		if ($ProbeOnly) { return $false }
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

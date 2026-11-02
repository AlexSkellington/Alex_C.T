# ===================================================================================================
#                           FUNCTION: Show_Backup_Copy_Paths_Dialog
# ---------------------------------------------------------------------------------------------------
# Description:
#   Dialog for the extra places a scheduled backup is copied to after it succeeds: a list of
#   folders, a box to type or paste one (network paths like \\server\share\Backups work), Browse,
#   Add and Remove. The list is remembered per backup type on this PC.
#
# Parameters:
#   - Kind    [string] (required): ServerDB | StoremanZip | LaneDB - which list to load / save
#   - Current [string[]]: the list to start from (default: the saved one)
#
# Returns:
#   The folders (string[], may be empty) when OK is clicked, $null when cancelled.
#
# Notes:
#   - Saved in <AlexRoot>\Settings\Backup_Copy_Paths_<Kind>.txt, one folder per line.
#   - Paths are checked (created / write-tested) by Resolve_Backup_Folder when the backup is
#     scheduled, not here.
# ===================================================================================================

function Show_Backup_Copy_Paths_Dialog
{
	param (
		[Parameter(Mandatory = $true)]
		[ValidateSet('ServerDB', 'StoremanZip', 'LaneDB')]
		[string]$Kind,
		[string[]]$Current
	)
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$settingsFile = Join-Path $script:AlexRoot ("Settings\Backup_Copy_Paths_" + $Kind + ".txt")
	if ($null -eq $Current) { $Current = Get_Backup_Copy_Paths -Kind $Kind }

	$dlg = New-Object System.Windows.Forms.Form
	$dlg.Text = "Also Copy the Backup To"
	$dlg.ClientSize = New-Object System.Drawing.Size(520, 330)
	$dlg.StartPosition = "CenterParent"
	$dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$dlg.MaximizeBox = $false
	$dlg.MinimizeBox = $false
	$dlg.ShowInTaskbar = $false

	$info = New-Object System.Windows.Forms.Label
	$info.Text = "After each successful backup a copy goes to every folder below (into the same <machine> subfolder), keeping the same number of backups there. A failed copy never fails the backup."
	$info.Location = New-Object System.Drawing.Point(12, 10)
	$info.Size = New-Object System.Drawing.Size(496, 36)
	$dlg.Controls.Add($info)

	$list = New-Object System.Windows.Forms.ListBox
	$list.Location = New-Object System.Drawing.Point(12, 52)
	$list.Size = New-Object System.Drawing.Size(496, 160)
	$list.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
	$list.HorizontalScrollbar = $true
	foreach ($p in @($Current)) { if (-not [string]::IsNullOrWhiteSpace($p)) { [void]$list.Items.Add([string]$p) } }
	$dlg.Controls.Add($list)

	$txt = New-Object System.Windows.Forms.TextBox
	$txt.Location = New-Object System.Drawing.Point(12, 224)
	$txt.Size = New-Object System.Drawing.Size(316, 24)
	$dlg.Controls.Add($txt)

	$mk = {
		param ([string]$Text, [int]$X, [int]$Y, [int]$W)
		$b = New-Object System.Windows.Forms.Button
		$b.Text = $Text
		$b.Location = New-Object System.Drawing.Point($X, $Y)
		$b.Size = New-Object System.Drawing.Size($W, 26)
		$dlg.Controls.Add($b)
		return $b
	}
	$btnAdd = & $mk 'Add' 334 223 80
	$btnBrowse = & $mk 'Browse...' 420 223 88
	$btnRemove = & $mk 'Remove Selected' 12 256 130
	$btnOk = & $mk 'OK' 334 290 84
	$btnCancel = & $mk 'Cancel' 424 290 84
	$btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$dlg.CancelButton = $btnCancel
	try { $toolTip.SetToolTip($txt, "Type or paste a folder (D:\Backups, \\server\share\Backups), then Add.") } catch { }

	# Adds the typed path unless it is empty or already listed (no state outside the controls)
	$btnAdd.Add_Click({
			$p = $txt.Text.Trim().Trim('"').TrimEnd('\')
			if (-not $p) { return }
			foreach ($it in $list.Items) { if ([string]::Equals([string]$it, $p, [System.StringComparison]::OrdinalIgnoreCase)) { $txt.Clear(); return } }
			[void]$list.Items.Add($p)
			$txt.Clear()
		})
	$btnBrowse.Add_Click({
			$p = Select-FolderOrDefault -initialPath $txt.Text -description "Choose a folder to also copy the backup to"
			if ($p) { $txt.Text = $p; $btnAdd.PerformClick() }
		})
	$btnRemove.Add_Click({
			foreach ($it in @($list.SelectedItems)) { $list.Items.Remove($it) }
		})
	$txt.Add_KeyDown({
			param ($s, $e)
			if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $btnAdd.PerformClick(); $e.SuppressKeyPress = $true }
		})

	try
	{
		if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
		# A path still in the box counts too
		if ($txt.Text.Trim()) { $btnAdd.PerformClick() }
		$paths = [string[]]@($list.Items | ForEach-Object { [string]$_ })
		try
		{
			$dir = Split-Path -Parent $settingsFile
			if (-not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
			[System.IO.File]::WriteAllLines($settingsFile, $paths)
		}
		catch { Write_Log ("Extra backup folders could not be saved: " + $_.Exception.Message) "yellow" }
		return , $paths
	}
	finally { $dlg.Dispose() }
}

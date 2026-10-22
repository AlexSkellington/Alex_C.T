# ===================================================================================================
#                              FUNCTION: Organize_Desktop_Items
# ---------------------------------------------------------------------------------------------------
# Description:
#   Moves non-system, non-excluded items from the current user's Desktop into a single
#   "Unorganized Items" folder. Keeps common system shortcuts, the TBS / MiniGhost launchers
#   (any version) and the Lanes / Scales / BackOffices folders; never overwrites an item that
#   is already in the target folder (the moved one gets a ' (2)' style suffix).
#
# Parameters:
#   - UnorganizedFolderName [string]: target folder on the Desktop (default "Unorganized Items")
#   - AdditionalExclusions  [string[]]: extra item names to leave on the Desktop
#
# Writes:
#   - moves files/folders under the user Desktop into the Unorganized folder
# ===================================================================================================

function Organize_Desktop_Items
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$UnorganizedFolderName = "Unorganized Items",
		[Parameter(Mandatory = $false)]
		[string[]]$AdditionalExclusions
	)

	Write_Log "`r`n==================== Starting Organize_Desktop_Items ====================`r`n" "blue"

	try
	{
		# -- Resolve Desktop path
		$DesktopPath = [Environment]::GetFolderPath("Desktop")

		# -- Compute the destination folder path
		$UnorganizedFolder = Join-Path -Path $DesktopPath -ChildPath $UnorganizedFolderName

		# -- The running tool itself (.ps1 / .exe), when it lives on the Desktop
		$scriptName = if ($script:LaunchEntryPath) { [System.IO.Path]::GetFileName([string]$script:LaunchEntryPath) } else { "TBS_Maintenance_Script.exe" }
		# Launchers of any version (TBS_Maintenance_Script_V2.6.9.exe, MiniGhost_V38.ps1, ...)
		$keepPatterns = @('TBS_Maintenance_Script*', 'TBS_Maintenace_Script*', 'Execute(TBS*', 'MiniGhost*', 'Execute(MiniGhost*')

		# -- Known system/launcher items to preserve on Desktop (case-insensitive match by name)
		$systemIcons = @(
			"This PC.lnk", "Network.lnk", "Control Panel.lnk", "Recycle Bin.lnk", "User's Files.lnk",
			"Execute(TBS_Maintenance_Script).bat", "Execute(MiniGhost).bat",
			"TBS_Maintenance_Script.exe", "MiniGhost.exe",
			$scriptName,
			"desktop.ini"
		)

		# -- Default excluded folders that must remain on Desktop
		$defaultExcludedFolders = @("Lanes", "Scales", "BackOffices", $UnorganizedFolderName)

		# -- Merge additional exclusions if provided
		$excludedSet = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
		foreach ($n in $systemIcons) { [void]$excludedSet.Add($n) }
		foreach ($n in $defaultExcludedFolders) { [void]$excludedSet.Add($n) }
		if ($AdditionalExclusions)
		{
			foreach ($n in $AdditionalExclusions) { if ($n) { [void]$excludedSet.Add($n) } }
		}

		# -- Ensure the target "Unorganized" folder exists
		if (-not (Test-Path -LiteralPath $UnorganizedFolder))
		{
			New-Item -ItemType Directory -Path $UnorganizedFolder | Out-Null
			Write_Log "Created folder: $UnorganizedFolder" "green"
		}
		else
		{
			Write_Log "Folder already exists: $UnorganizedFolder" "cyan"
		}

		# -- Enumerate Desktop items; keep both files and folders.
		#    Move everything except explicit exclusions and Windows system-managed filesystem items.
		$desktopItems = @(Get-ChildItem -LiteralPath $DesktopPath -Force |
			Where-Object {
				$name = $_.Name
				# explicitly excluded names, launcher patterns, and System items (desktop.ini)
				-not $excludedSet.Contains($name) -and
				-not @($keepPatterns | Where-Object { $name -like $_ }).Count -and
				-not ($_.Attributes -band [System.IO.FileAttributes]::System)
			})

		# -- Move everything else into the Unorganized folder (never overwrite: "name (2).ext")
		foreach ($item in $desktopItems)
		{
			try
			{
				$target = Join-Path $UnorganizedFolder $item.Name
				$n = 2
				while (Test-Path -LiteralPath $target)
				{
					$base = if ($item.PSIsContainer) { $item.Name } else { [System.IO.Path]::GetFileNameWithoutExtension($item.Name) }
					$ext = if ($item.PSIsContainer) { '' } else { $item.Extension }
					$target = Join-Path $UnorganizedFolder ("{0} ({1}){2}" -f $base, $n, $ext)
					$n++
				}
				Move-Item -LiteralPath $item.FullName -Destination $target -ErrorAction Stop
				Write_Log "Moved item: $($item.Name)" "green"
			}
			catch
			{
				Write_Log "Failed to move item: $($item.Name). Error: $_" "red"
			}
		}

		Write_Log "Desktop organization complete." "green"
		Write_Log "`r`n==================== Organize_Desktop_Items Completed ====================" "blue"
	}
	catch
	{
		Write_Log "Unexpected error in Organize_Desktop_Items: $_" "red"
	}
}

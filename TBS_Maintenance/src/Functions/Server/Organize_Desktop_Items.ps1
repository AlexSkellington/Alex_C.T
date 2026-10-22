# ===================================================================================================
#                              FUNCTION: Organize_Desktop_Items
# ---------------------------------------------------------------------------------------------------
# Description:
#   Moves non-system, non-excluded items from the current user's Desktop into a single
#   "Unorganized Items" folder and makes sure the excluded folders still exist. Skips common
#   system shortcuts and script launchers, and does not restart Explorer.
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
		
		# -- Best-effort current script/launcher name (used for exclusions)
		#    $scriptName could be defined elsewhere; fall back to the current command name if available
		$scriptName = if ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Name)
		{
			$MyInvocation.MyCommand.Name
		}
		else
		{
			"TBS_Maintenance_Script.exe"
		}
		
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
		$desktopItems = Get-ChildItem -LiteralPath $DesktopPath -Force |
		Where-Object {
			# Skip items explicitly excluded by name
			-not $excludedSet.Contains($_.Name) -and
			# Skip filesystem items marked as System (for example desktop.ini)
			(-not (($_.Attributes -band [System.IO.FileAttributes]::System) -eq [System.IO.FileAttributes]::System))
		}
		
		# -- Move everything else into the Unorganized folder
		foreach ($item in $desktopItems)
		{
			try
			{
				Move-Item -LiteralPath $item.FullName -Destination $UnorganizedFolder -Force
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

# ===================================================================================================
#                              FUNCTION: Open_Selected_Node_C_Path
# ---------------------------------------------------------------------------------------------------
# Description:
#   Opens the C: drive of selected nodes in Explorer: pick lanes and/or Bizerba scales in
#   the tabbed selection form, then open \\<MACHINE>\c$ (lanes) or \\<SCALEIP>\c$ (scales,
#   trying bizuser/bizerba then bizuser/biyerba credentials).
#
# Parameters:
#   - StoreNumber  [string]: store to process (required)
#   - NodeTypes    [string[]]: node types to show - Lane, Scale
#
# Variables:
#   - $script:FunctionResults: LaneNumToMachineName, ScaleCodeToIPInfo (input)
# ===================================================================================================

function Open_Selected_Node_C_Path
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter()]
		[ValidateSet("Lane", "Scale")]
		[string[]]$NodeTypes = @("Lane", "Scale")
	)

	Write_Log "`r`n==================== Starting Open_Selected_Node_C_Path Function ====================`r`n" "blue"

	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes $NodeTypes
	if (-not $selection)
	{
		Write_Log "No selection made. Exiting." "Yellow"
		return
	}

	# Normalize lane selection to 3-digit and get machine name
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$lanes = @()
	if ($selection.Lanes -and $selection.Lanes.Count -gt 0)
	{
		foreach ($laneSel in $selection.Lanes)
		{
			$laneNum = if ($laneSel -is [PSCustomObject] -and $laneSel.LaneNumber) { $laneSel.LaneNumber }
			elseif ($laneSel -match '^\d{3}$') { $laneSel }
			elseif ($LaneNumToMachineName.ContainsKey($laneSel)) { $laneSel }
			else { $null }
			if ($laneNum -and $LaneNumToMachineName.ContainsKey($laneNum))
			{
				$machine = $LaneNumToMachineName[$laneNum]
				$sharePath = "\\$machine\c$"
				Write_Log "Opened $sharePath ..." "Green"
				Start-Process "explorer.exe" $sharePath
			}
			else
			{
				Write_Log "Machine not found for lane '$laneSel'." "Red"
			}
		}
	}

	# Normalize scale selection (scale objects or just codes)
	$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
	if ($selection.Scales -and $selection.Scales.Count -gt 0)
	{
		foreach ($scaleObj in $selection.Scales)
		{
			$scale = $null
			if ($scaleObj -is [PSCustomObject])
			{
				$scale = $scaleObj
			}
			elseif ($ScaleCodeToIPInfo.ContainsKey($scaleObj))
			{
				$scale = $ScaleCodeToIPInfo[$scaleObj]
			}
			if (-not $scale) { Write_Log "Could not resolve scale info for $scaleObj." "Red"; continue }

			$scaleHost = $null
			if ($scale.PSObject.Properties['FullIP'] -and $scale.FullIP)
			{
				$scaleHost = $scale.FullIP
			}
			elseif ($scale.PSObject.Properties['IPNetwork'] -and $scale.PSObject.Properties['IPDevice'])
			{
				$scaleHost = "$($scale.IPNetwork)$($scale.IPDevice)"
			}
			elseif ($scale.PSObject.Properties['ScaleName'])
			{
				$scaleHost = $scale.ScaleName
			}
			if (-not $scaleHost) { Write_Log "Could not determine host for $($scale.ScaleCode)." "Red"; continue }

			$sharePath = "\\$scaleHost\c$"

			# Already reachable (credentials seeded at startup)? Otherwise try bizuser/bizerba, then
			# bizuser/biyerba. Explorer opens once, after a password worked (it used to open before
			# the check, leaving an extra error window when the first password was wrong).
			$opened = Test-Path -LiteralPath $sharePath
			if ($opened) { Write_Log "Opened $sharePath." "Green" }
			foreach ($pass in @('bizerba', 'biyerba'))
			{
				if ($opened) { break }
				$null = & cmdkey.exe "/add:$scaleHost" "/user:bizuser" "/pass:$pass" 2>&1
				if (Test-Path -LiteralPath $sharePath)
				{
					Write_Log "Opened $sharePath as bizuser using password '$pass'." "Green"
					$opened = $true
				}
				else { $null = & cmdkey.exe "/delete:$scaleHost" 2>&1 }
			}
			if ($opened) { Start-Process "explorer.exe" $sharePath }
			else { Write_Log "Could not open $sharePath as bizuser with either password." "Red" }
		}
	}
	Write_Log "`r`n==================== Open_Selected_Node_C_Path Function Completed ====================" "blue"
}

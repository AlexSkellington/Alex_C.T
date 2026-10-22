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
			$opened = $false

			# Try bizuser/bizerba, then bizuser/biyerba
			cmdkey /add:$scaleHost /user:bizuser /pass:bizerba | Out-Null
			Start-Process "explorer.exe" $sharePath
			Start-Sleep -Seconds 2
			if (Test-Path $sharePath)
			{
				Write_Log "Opened $sharePath as bizuser using password 'bizerba'." "Green"
				$opened = $true
			}
			else
			{
				cmdkey /delete:$scaleHost | Out-Null
				cmdkey /add:$scaleHost /user:bizuser /pass:biyerba | Out-Null
				Start-Process "explorer.exe" $sharePath
				Start-Sleep -Seconds 2
				if (Test-Path $sharePath)
				{
					Write_Log "Opened $sharePath as bizuser using password 'biyerba'." "Green"
					$opened = $true
				}
				else
				{
					Write_Log "Could not open $sharePath as bizuser with either password." "Red"
				}
			}
			# Clean up credentials (optional)
			# cmdkey /delete:$scaleHost | Out-Null
		}
	}
	Write_Log "`r`n==================== Open_Selected_Node_C_Path Function Completed ====================" "blue"
}

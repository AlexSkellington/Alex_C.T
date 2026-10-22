# ===================================================================================================
#                              FUNCTION: Fix_Deploy_CHG
# ---------------------------------------------------------------------------------------------------
# Description:
#   Restores the deploy line in DEPLOY_CHG.sql: builds the @EXEC line from the FastDEPLOY
#   (preferred) or regular ScaleManagementApp.exe path, strips the existing deploy lines
#   (matched via the configured regex), then appends the new line. Also deletes the
#   Update_Scales_Specials_Task_Minutes scheduled task if it exists.
#
# Parameters:
#   - OfficePath  [string]: folder holding DEPLOY_CHG.sql (default: $script:OfficePath)
#
# Variables:
#   - $script:ScaleManagementApp*PathCandidates: FastDEPLOY / regular exe lookup
#   - $script:ScaleCommDeployExecRegex: regex matching the @EXEC lines to strip
# ===================================================================================================

function Fix_Deploy_CHG
{
	param (
		[string]$OfficePath = $script:OfficePath
	)
	
	Write_Log "`r`n==================== Starting Fix_Deploy_CHG Function ====================`r`n" "blue"
	
	$deployChgFile = Join-Path $OfficePath "DEPLOY_CHG.sql"
	
	if (-not (Test-Path $deployChgFile))
	{
		Write_Log "DEPLOY_CHG.sql not found at $deployChgFile" "red"
		Write_Log "`r`n==================== Fix_Deploy_CHG Function Completed ====================" "blue"
		return
	}
	
	$exeLine = ""
	$fastDeployPath = $script:ScaleManagementAppFastDeployPathCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
	$regularDeployPath = $script:ScaleManagementAppPathCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
	if (-not $fastDeployPath) { $fastDeployPath = $script:ScaleManagementAppFastDeployPath }
	if (-not $regularDeployPath) { $regularDeployPath = $script:ScaleManagementAppExePath }
	if (Test-Path -LiteralPath $fastDeployPath)
	{
		$exeLine = "/* Deploy price changes to the scales */`r`n@EXEC(RUN='$fastDeployPath');"
	}
	elseif (Test-Path -LiteralPath $regularDeployPath)
	{
		$exeLine = "/* Deploy price changes to the scales */`r`n@EXEC(RUN='$regularDeployPath');"
	}
	else
	{
		Write_Log "Neither FastDEPLOY nor regular ScaleManagementApp.exe found in configured ScaleComm roots." "red"
		Write_Log "`r`n==================== Fix_Deploy_CHG Function Completed ====================" "blue"
		return
	}
	
	# --- Restore DEPLOY_CHG.sql ---
	try
	{
		$deployExecRegex = $script:ScaleCommDeployExecRegex
		$content = Get-Content $deployChgFile -Raw
		$newContent = [System.Collections.Generic.List[string]]@(
			($content -split "`r?`n") | Where-Object {
				$_ -notmatch '^\s*/\* Deploy price changes to the scales \*/' -and
				$_ -notmatch $deployExecRegex
			}
		)
		while ($newContent.Count -gt 0 -and ($newContent[-1] -match '^\s*$'))
		{
			$null = $newContent.RemoveAt($newContent.Count - 1)
		}
		$newContent += ""
		$newContent += $exeLine
		$newContent -join "`r`n" | Set-Content -Path $deployChgFile -Encoding Default
		Write_Log "Restored line to DEPLOY_CHG.sql: $exeLine" "green"
	}
	catch
	{
		Write_Log "Failed to restore line to DEPLOY_CHG.sql: $_" "red"
		Write_Log "`r`n==================== Fix_Deploy_CHG Function Completed ====================" "blue"
		return
	}
	
	# --- Delete the minutes task if it exists ---
	$minutesTaskName = "Update_Scales_Specials_Task_Minutes"
	$taskExists = schtasks /Query /TN "$minutesTaskName" 2>&1 | Select-String -Quiet -Pattern "$minutesTaskName"
	if ($taskExists)
	{
		$deleteOut = schtasks /Delete /TN "$minutesTaskName" /F 2>&1
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Deleted scheduled task '$minutesTaskName' after DEPLOY_CHG.sql restore." "yellow"
		}
		else
		{
			Write_Log "Attempted to delete '$minutesTaskName' after restore but schtasks said: $deleteOut" "yellow"
		}
	}
	
	Write_Log "`r`n==================== Fix_Deploy_CHG Function Completed ====================" "blue"
}

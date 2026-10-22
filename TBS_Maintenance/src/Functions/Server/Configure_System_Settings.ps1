# ===================================================================================================
#                              FUNCTION: Configure_System_Settings
# ---------------------------------------------------------------------------------------------------
# Description:
#   Tunes the local PC for POS use: High Performance power plan, sleep disabled, minimum
#   processor performance at 100%, monitor off after 15 minutes on AC, several services set
#   to Automatic and started, and visual settings (thumbnails, font smoothing, ClearType).
#   Restarts Explorer so the visual settings apply.
#
# Writes:
#   - local power plan, service startup types and visual settings
#
# Notes:
#   - Desktop organizing is a separate function: Organize_Desktop_Items.
# ===================================================================================================

function Configure_System_Settings
{
	[CmdletBinding()]
	param ()

	Write_Log "`r`n==================== Starting Configure_System_Settings ====================`r`n" "blue"

	try
	{
		# ===========================================
		# 1. Power settings
		# ===========================================
		Write_Log "`r`nConfiguring power plan and performance settings..." "blue"

		# powercfg is an external program: failures show up in the exit code, never as an
		# exception (before v2.7.0 every step here logged success regardless)
		$powercfg = {
			param ([string]$What, [string[]]$PcArgs)
			$out = & powercfg.exe @PcArgs 2>&1
			if ($LASTEXITCODE -eq 0) { Write_Log $What "green"; return $true }
			Write_Log ("Failed: {0} (powercfg {1}: {2})" -f $What, ($PcArgs -join ' '), (($out | Out-String).Trim())) "red"
			return $false
		}

		# High Performance plan; on systems where it is hidden, restore it from its template first
		$highPerfGUID = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
		$null = & powercfg.exe /query $highPerfGUID 2>&1
		if ($LASTEXITCODE -ne 0) { $null = & powercfg.exe -duplicatescheme $highPerfGUID $highPerfGUID 2>&1 }
		$null = & $powercfg "Power plan set to High Performance." @('/s', $highPerfGUID)

		$null = & $powercfg "System sleep disabled (AC)." @('/change', 'standby-timeout-ac', '0')
		$null = & $powercfg "System sleep disabled (DC)." @('/change', 'standby-timeout-dc', '0')

		# Minimum processor state 100% (SUB_PROCESSOR / PROCTHROTTLEMIN) on the High Performance plan
		$subProcessor = "54533251-82be-4824-96c1-47b60b740d00"
		$procThrottleMin = "893dee8e-2bef-41e0-89c6-b55d0929964c"
		$null = & $powercfg "Minimum processor performance 100% (AC)." @('/setacvalueindex', $highPerfGUID, $subProcessor, $procThrottleMin, '100')
		$null = & $powercfg "Minimum processor performance 100% (DC)." @('/setdcvalueindex', $highPerfGUID, $subProcessor, $procThrottleMin, '100')
		$null = & $powercfg "High Performance plan re-activated." @('/setactive', $highPerfGUID)

		$null = & $powercfg "Monitor timeout (AC) set to 15 minutes." @('/change', 'monitor-timeout-ac', '15')
		Write_Log "Power and performance settings complete." "green"

		# ===========================================
		# 2. Configure Services
		# ===========================================
		Write_Log "`r`nConfiguring services to start automatically..." "blue"

		$servicesToConfigure = @("fdPHost", "FDResPub", "SSDPSRV", "upnphost")
		foreach ($service in $servicesToConfigure)
		{
			try
			{
				Set-Service -Name $service -StartupType Automatic -ErrorAction Stop
				Write_Log "Set service '$service' to Automatic." "green"

				$svc = Get-Service -Name $service -ErrorAction Stop
				if ($svc.Status -ne 'Running')
				{
					Start-Service -Name $service -ErrorAction Stop
					Write_Log "Started service '$service'." "green"
				}
				else
				{
					Write_Log "Service '$service' is already running." "cyan"
				}
			}
			catch
			{
				Write_Log "Failed to configure service '$service'. Error: $_" "red"
			}
		}

		Write_Log "Service configuration complete." "green"

		# ===========================================
		# 3. Configure Visual Settings
		# ===========================================
		Write_Log "`r`nConfiguring visual settings..." "blue"

		try
		{
			# Show thumbnails (IconsOnly = 0)
			Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
							 -Name "IconsOnly" -Value 0 -Type DWord -ErrorAction Stop
			Write_Log "Enabled 'Show thumbnails instead of icons'." "green"
		}
		catch
		{
			Write_Log "Failed to enable thumbnails. Error: $_" "red"
		}

		try
		{
			# Font smoothing + ClearType
			Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "FontSmoothing" -Value "2" -Type String -ErrorAction Stop
			Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "FontSmoothingType" -Value 2 -Type DWord -ErrorAction Stop
			Write_Log "Enabled 'Smooth edges of screen fonts' (ClearType)." "green"
		}
		catch
		{
			Write_Log "Failed to enable font smoothing. Error: $_" "red"
		}

		# Restart Explorer to apply visual tweaks (kept here since it's visual-settings-related)
		Write_Log "Restarting Explorer to apply changes..." "yellow"
		try
		{
			Stop-Process -Name explorer -Force
			Write_Log "Explorer restarted." "green"
		}
		catch
		{
			Write_Log "Explorer restart may already be complete or not required. Info: $_" "cyan"
		}

		Write_Log "All system configurations have been applied." "green"
		Write_Log "`r`n==================== Configure_System_Settings Completed ====================`r`n" "blue"
	}
	catch
	{
		Write_Log "An unexpected error occurred in Configure_System_Settings: $_" "red"
	}
}

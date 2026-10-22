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
		# ==========================================================================================
		# [CHANGE] Removed: Desktop organizing block
		#          Moved into new function: Organize_Desktop_Items
		# ==========================================================================================
		
		# ===========================================
		# 1. Configure Power Settings
		# ===========================================
		Write_Log "`r`nConfiguring power plan and performance settings..." "blue"
		
		# High Performance plan GUID
		$highPerfGUID = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
		
		try
		{
			powercfg /s $highPerfGUID
			Write_Log "Power plan set to High Performance." "green"
		}
		catch
		{
			Write_Log "Failed to set power plan to High Performance. Error: $_" "red"
		}
		
		# Disable sleep (AC/DC)
		try
		{
			powercfg /change standby-timeout-ac 0
			powercfg /change standby-timeout-dc 0
			Write_Log "System sleep disabled for AC/DC." "green"
		}
		catch
		{
			Write_Log "Failed to disable system sleep. Error: $_" "red"
		}
		
		# Minimum processor performance to 100% (AC/DC)
		try
		{
			# SUB_PROCESSOR         = 54533251-82be-4824-96c1-47b60b740d00
			# PROCTHROTTLEMIN       = 893dee8e-2bef-41e0-89c6-b55d0929964c
			powercfg /setacvalueindex $highPerfGUID "54533251-82be-4824-96c1-47b60b740d00" "893dee8e-2bef-41e0-89c6-b55d0929964c" 100
			powercfg /setdcvalueindex $highPerfGUID "54533251-82be-4824-96c1-47b60b740d00" "893dee8e-2bef-41e0-89c6-b55d0929964c" 100
			powercfg /setactive $highPerfGUID
			Write_Log "Minimum processor performance set to 100% (AC/DC)." "green"
		}
		catch
		{
			Write_Log "Failed to set processor performance. Error: $_" "red"
		}
		
		# Turn off screen after 15 minutes (AC)
		try
		{
			powercfg /change monitor-timeout-ac 15
			Write_Log "Monitor timeout (AC) set to 15 minutes." "green"
		}
		catch
		{
			Write_Log "Failed to set monitor timeout. Error: $_" "red"
		}
		
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

# ===================================================================================================
#                              FUNCTION: Invoke_Secure_Script
# ---------------------------------------------------------------------------------------------------
# Description:
#   DISABLED as of v2.6.9: this used to prompt for a hard-coded password and execute a remote
#   script via Invoke-Expression (irm <url>). The remote execution and the password were
#   removed for security; the function now logs the disabled state and returns. The
#   "Activate Windows" menu item stays so the reason is visible in the log.
# ===================================================================================================

function Invoke_Secure_Script
{
	[CmdletBinding()]
	param ()

	# --- Security hardening: remote script execution disabled ---
	Write_Log "`r`n==================== Starting Invoke_Secure_Script Function ====================`r`n" "blue"
	Write_Log "Invoke_Secure_Script is disabled: remote Invoke-Expression/irm execution was removed for security." "red"
	Write_Log "If this workflow is required, replace it with a signed local script and hash verification." "yellow"
	Write_Log "`r`n==================== Invoke_Secure_Script Function Completed ====================" "blue"
}

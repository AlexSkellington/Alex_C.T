# ===================================================================================================
#                              FUNCTION: Repair_BMS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Repairs the Bizerba BMS service on the local machine: disables the service, stops it with
#   escalating actions (Stop-Service, sc stop, net stop /y, WMI, taskkill by PID/service/image),
#   cleans the toBizerba staging and terminals cache folders, deletes and re-registers the
#   service (BMSSrv.exe -reg) and starts it. Requires elevation.
#
# Parameters:
#   - BMSSrvPath         [string]: path to BMSSrv.exe (default $script:BmssrvExePath)
#   - CleanToBizerba     [bool]: purge the toBizerba staging folder (default $true)
#   - RemoveTerminalsCache [bool]: remove the terminals cache folder (default $true)
#
# Variables:
#   - $script:BmssrvExePath: default BMSSrv.exe location
#   - $script:BmsRootPath: BMS install root (toBizerba/terminals folders)
#
# Writes:
#   - local BMS service state, toBizerba and terminals folders
# ===================================================================================================

function Repair_BMS
{
	[CmdletBinding()]
	param
	(
		# Full path to BMSSrv.exe (used to re-register the service)
		[Parameter(Mandatory = $false)]
		[string]$BMSSrvPath = $script:BmssrvExePath,
		# Purge staging folder (matches your fast .bat)
		[Parameter(Mandatory = $false)]
		[bool]$CleanToBizerba = $true,
		# Remove terminals cache folder (matches your fast .bat)
		[Parameter(Mandatory = $false)]
		[bool]$RemoveTerminalsCache = $true
	)
	
	Write_Log "`r`n==================== Starting Repair_BMS (IMMEDIATE TASKKILL) ====================`r`n" "blue"
	
	# ------------------------------------------------------------------------------------------------
	# 0) Elevation check (required for service control + protected folder deletes)
	# ------------------------------------------------------------------------------------------------
	$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).
	IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
	
	if (-not $isAdmin)
	{
		Write_Log "Insufficient permissions. Please run as Administrator." "red"
		return
	}
	
	# ------------------------------------------------------------------------------------------------
	# 1) Validate BMSSrv.exe path
	# ------------------------------------------------------------------------------------------------
	if (-not (Test-Path -LiteralPath $BMSSrvPath))
	{
		Write_Log "BMSSrv.exe not found at: $BMSSrvPath" "red"
		return
	}
	
	# ------------------------------------------------------------------------------------------------
	# 2) Constants / paths
	# ------------------------------------------------------------------------------------------------
	$serviceName = "BMS"
	$bmsRoot = $script:BmsRootPath
	$toBizerbaPath = Join-Path $bmsRoot "toBizerba"
	$terminalsPath = Join-Path $bmsRoot "terminals"
	
	# Kill only what you want (silent)
	$killImages = @(
		"BMSSrv.exe",
		"BMS.exe"
	)
	
	Write_Log "Phase 1: IMMEDIATE kill (no wait loops)..." "blue"
	
	# ------------------------------------------------------------------------------------------------
	# 3) Kill from the get-go (NO waiting) - SILENT
	# ------------------------------------------------------------------------------------------------
	try
	{
		cmd /c 'taskkill /F /T /FI "SERVICES eq BMS" >nul 2>&1'
		Write_Log 'Issued: taskkill /F /T /FI "SERVICES eq BMS" (silent)' "yellow"
	}
	catch
	{
		Write_Log "taskkill by service filter threw an exception (often harmless): $_" "yellow"
	}
	
	foreach ($img in $killImages)
	{
		try
		{
			cmd /c ("taskkill /F /T /IM {0} >nul 2>&1" -f $img)
			Write_Log "Issued: taskkill /F /T /IM $img (silent)" "yellow"
		}
		catch
		{
			Write_Log "taskkill for $img threw an exception (often harmless): $_" "yellow"
		}
	}
	
	# ------------------------------------------------------------------------------------------------
	# 4) Disable/Stop service quickly (best-effort, no waits)
	#    - Skip noisy net stop unless service is actually not stopped
	# ------------------------------------------------------------------------------------------------
	Write_Log "Phase 2: Disable/Stop service (best-effort, no waits)..." "blue"
	
	$svcBefore = $null
	try { $svcBefore = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcBefore = $null }
	
	try { cmd /c ("sc.exe config {0} start= disabled >nul 2>&1" -f $serviceName) }
	catch { }
	try { Set-Service -Name $serviceName -StartupType Disabled -ErrorAction SilentlyContinue }
	catch { }
	
	if ($svcBefore -and $svcBefore.Status -ne 'Stopped')
	{
		try { cmd /c ("sc.exe stop {0} >nul 2>&1" -f $serviceName) }
		catch { }
		try { cmd /c ("net stop {0} /y >nul 2>&1" -f $serviceName) }
		catch { } # silent
		try { Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue }
		catch { }
		
		# Re-kill once more (still no waits)
		try { cmd /c 'taskkill /F /T /FI "SERVICES eq BMS" >nul 2>&1' }
		catch { }
		foreach ($img in $killImages)
		{
			try { cmd /c ("taskkill /F /T /IM {0} >nul 2>&1" -f $img) }
			catch { }
		}
	}
	else
	{
		Write_Log "Service is already stopped or not present; skipping net stop to avoid console noise." "gray"
	}
	
	# ------------------------------------------------------------------------------------------------
	# 5) Cleanup filesystem (fast)
	# ------------------------------------------------------------------------------------------------
	Write_Log "Phase 3: Cleanup folders..." "blue"
	
	if ($CleanToBizerba)
	{
		if (Test-Path -LiteralPath $toBizerbaPath)
		{
			try
			{
				Get-ChildItem -LiteralPath $toBizerbaPath -Force -ErrorAction SilentlyContinue |
				Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
				
				Write_Log "Cleared: $toBizerbaPath" "green"
			}
			catch
			{
				Write_Log "Failed clearing $[toBizerbaPath]: $_" "yellow"
			}
		}
		else
		{
			Write_Log "Folder not found (skip): $toBizerbaPath" "gray"
		}
	}
	else
	{
		Write_Log "CleanToBizerba disabled (skip)." "gray"
	}
	
	if ($RemoveTerminalsCache)
	{
		if (Test-Path -LiteralPath $terminalsPath)
		{
			try
			{
				Remove-Item -LiteralPath $terminalsPath -Recurse -Force -ErrorAction SilentlyContinue
				Write_Log "Removed cache folder: $terminalsPath" "green"
			}
			catch
			{
				Write_Log "Failed removing $[terminalsPath]: $_" "yellow"
			}
		}
		else
		{
			Write_Log "Folder not found (skip): $terminalsPath" "gray"
		}
	}
	else
	{
		Write_Log "RemoveTerminalsCache disabled (skip)." "gray"
	}
	
	# ------------------------------------------------------------------------------------------------
	# 6) Unregister (delete service) + Register again (BMSSrv.exe -reg)
	#    - CHANGED: BMSSrv.exe output is redirected so it doesn't print to console
	# ------------------------------------------------------------------------------------------------
	Write_Log "Phase 4: Unregister (sc delete) + Register (BMSSrv.exe -reg)..." "blue"
	
	try
	{
		# Pre-kill again (helps with "marked for deletion" locks)
		try { cmd /c 'taskkill /F /T /FI "SERVICES eq BMS" >nul 2>&1' }
		catch { }
		foreach ($img in $killImages)
		{
			try { cmd /c ("taskkill /F /T /IM {0} >nul 2>&1" -f $img) }
			catch { }
		}
		
		cmd /c ("sc.exe delete {0} >nul 2>&1" -f $serviceName)
		Write_Log "Issued: sc delete $serviceName (silent)" "yellow"
	}
	catch
	{
		Write_Log "sc delete threw an exception (service may already be gone): $_" "yellow"
	}
	
	# --- Register service again, SILENT (no console spam) ---
	$tmpOut = Join-Path $env:TEMP ("BMSSrv_reg_stdout_{0}.log" -f ([Guid]::NewGuid().ToString("N")))
	$tmpErr = Join-Path $env:TEMP ("BMSSrv_reg_stderr_{0}.log" -f ([Guid]::NewGuid().ToString("N")))
	
	try
	{
		# Ensure temp files don't exist
		try { if (Test-Path -LiteralPath $tmpOut) { Remove-Item -LiteralPath $tmpOut -Force -ErrorAction SilentlyContinue } }
		catch { }
		try { if (Test-Path -LiteralPath $tmpErr) { Remove-Item -LiteralPath $tmpErr -Force -ErrorAction SilentlyContinue } }
		catch { }
		
		# CHANGED: Redirect output + hide window so BMSSrv does not write to the console
		$proc = Start-Process -FilePath $BMSSrvPath -ArgumentList "-reg" -WindowStyle Hidden -Wait -PassThru `
							  -RedirectStandardOutput $tmpOut -RedirectStandardError $tmpErr
		
		if ($proc.ExitCode -ne 0)
		{
			$stdout = ""; $stderr = ""
			try { if (Test-Path -LiteralPath $tmpOut) { $stdout = (Get-Content -LiteralPath $tmpOut -ErrorAction SilentlyContinue | Out-String).Trim() } }
			catch { }
			try { if (Test-Path -LiteralPath $tmpErr) { $stderr = (Get-Content -LiteralPath $tmpErr -ErrorAction SilentlyContinue | Out-String).Trim() } }
			catch { }
			
			Write_Log ("BMSSrv.exe -reg failed (exit code {0})." -f $proc.ExitCode) "red"
			if ($stdout) { Write_Log ("BMSSrv stdout: {0}" -f $stdout) "yellow" }
			if ($stderr) { Write_Log ("BMSSrv stderr: {0}" -f $stderr) "yellow" }
			return
		}
		
		Write_Log "BMSSrv.exe -reg completed successfully (silent)." "green"
	}
	catch
	{
		Write_Log "BMSSrv.exe -reg failed: $_" "red"
		return
	}
	finally
	{
		# Cleanup redirected output files (keep it tidy)
		try { if (Test-Path -LiteralPath $tmpOut) { Remove-Item -LiteralPath $tmpOut -Force -ErrorAction SilentlyContinue } }
		catch { }
		try { if (Test-Path -LiteralPath $tmpErr) { Remove-Item -LiteralPath $tmpErr -Force -ErrorAction SilentlyContinue } }
		catch { }
	}
	
	# ------------------------------------------------------------------------------------------------
	# 7) Re-enable + start service (avoid "service is not started" noise)
	# ------------------------------------------------------------------------------------------------
	Write_Log "Phase 5: Start service..." "blue"
	
	try { cmd /c ("sc.exe config {0} start= auto >nul 2>&1" -f $serviceName) }
	catch { }
	try { Set-Service -Name $serviceName -StartupType Automatic -ErrorAction SilentlyContinue }
	catch { }
	
	$svcAfterReg = $null
	try { $svcAfterReg = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcAfterReg = $null }
	
	if ($svcAfterReg -and $svcAfterReg.Status -ne 'Running')
	{
		try { cmd /c ("net start {0} /y >nul 2>&1" -f $serviceName) }
		catch { }
		try { Start-Service -Name $serviceName -ErrorAction SilentlyContinue }
		catch { }
	}
	else
	{
		Write_Log "Service already running or not present immediately after registration; skipping net start." "gray"
	}
	
	$svcFinal = $null
	try { $svcFinal = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcFinal = $null }
	
	if ($svcFinal -and $svcFinal.Status -eq 'Running')
	{
		Write_Log "'$serviceName' is running." "green"
	}
	else
	{
		$st = if ($svcFinal) { $svcFinal.Status }
		else { "NotFound" }
		Write_Log "WARNING: Could not confirm '$serviceName' running. Current state: $st" "yellow"
	}
	
	Write_Log "`r`n==================== Repair_BMS Completed (IMMEDIATE TASKKILL) ====================`r`n" "blue"
}

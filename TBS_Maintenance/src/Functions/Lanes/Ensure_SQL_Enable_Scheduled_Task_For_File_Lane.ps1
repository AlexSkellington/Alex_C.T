# ===================================================================================================
#                              FUNCTION: Ensure_SQL_Enable_Scheduled_Task_For_File_Lane
# ---------------------------------------------------------------------------------------------------
# Description:
#   For a lane stuck on the SQL File protocol, writes a self-deleting one-shot PowerShell
#   worker to the lane's Storeman Temp share and schedules it for 4:15 AM (next occurrence).
#   The worker enables TCP (static port), Named Pipes and Shared Memory, restarts SQL when
#   changed, then deletes its scheduled task and itself. Every failure path (no writable
#   Temp share, worker write error, schtasks /Create error) is logged red with the detail.
#
# Parameters:
#   - LaneNumber  [string]: lane to schedule (required)
#   - MachineName [string]: lane machine name (required)
#   - TcpPort     [string]: static TCP port to enable (default '1433')
#   - StartTime   [string]: task start time (default '04:15')
#
# Variables:
#   - $script:LaneSqlProtocolEnableTaskState: per-lane scheduling state
#
# Returns:
#   $true when the task was (re)scheduled successfully, $false otherwise.
#
# Writes:
#   - Enable_SQL_Protocols_Lane_<n>.ps1 on the lane Storeman Temp share + one-shot task
# ===================================================================================================

function Ensure_SQL_Enable_Scheduled_Task_For_File_Lane
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$LaneNumber,
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[string]$TcpPort = '1433',
		[string]$StartTime = '04:15'
	)

	if ([string]::IsNullOrWhiteSpace($LaneNumber) -or [string]::IsNullOrWhiteSpace($MachineName)) { return $false }
	if (-not $script:LaneSqlProtocolEnableTaskState) { $script:LaneSqlProtocolEnableTaskState = @{ } }

	$laneKey = $LaneNumber.ToString().PadLeft(3, '0')
	$taskName = "Enable_SQL_Protocols_Lane_$laneKey"
	$scriptName = "${taskName}.ps1"
	$remoteScriptSharePath = $null
	$localScriptPath = $null
	$remoteTempCandidates = @(
		[pscustomobject]@{ Share = "\\$MachineName\Storeman\Temp"; Local = "C:\Storeman\Temp" },
		[pscustomobject]@{ Share = "\\$MachineName\c$\Storeman\Temp"; Local = "C:\Storeman\Temp" },
		[pscustomobject]@{ Share = "\\$MachineName\d$\Storeman\Temp"; Local = "D:\Storeman\Temp" },
		[pscustomobject]@{ Share = "\\$MachineName\e$\Storeman\Temp"; Local = "E:\Storeman\Temp" }
	)

	# Next occurrence of StartTime (HH:mm, default 04:15) - the /SD date must match /ST
	$startSpan = [TimeSpan]::FromMinutes(255)
	$parsedSpan = [TimeSpan]::Zero
	if ([TimeSpan]::TryParse($StartTime, [ref]$parsedSpan)) { $startSpan = $parsedSpan }
	$scheduledDate = (Get-Date).Date.Add($startSpan)
	if ((Get-Date) -ge $scheduledDate) { $scheduledDate = $scheduledDate.AddDays(1) }
	$scheduledDateText = $scheduledDate.ToString('MM/dd/yyyy')

	$workerScript = @'
param(
    [string]$TaskName,
    [string]$TcpPort = '1433'
)

$ErrorActionPreference = 'Continue'

try { Add-Type -AssemblyName System.Core | Out-Null } catch { }

$instanceRoots = @(
    'SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL',
    'SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL'
)

$instanceMap = @{}
foreach ($root in $instanceRoots)
{
    try
    {
        $k = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($root)
        if ($k)
        {
            foreach ($name in $k.GetValueNames())
            {
                $id = $k.GetValue($name)
                if ($id -and -not $instanceMap.ContainsKey($name)) { $instanceMap[$name] = $id }
            }
            $k.Close()
        }
    }
    catch { }
}

$servicesToRestart = New-Object System.Collections.Generic.List[string]

foreach ($instanceName in $instanceMap.Keys)
{
    $instanceId = $instanceMap[$instanceName]
    $needsRestart = $false

    foreach ($authPath in @(
        "SOFTWARE\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer",
        "SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer"
    ))
    {
        try
        {
            $k = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($authPath, $true)
            if ($k)
            {
                if ($k.GetValue('LoginMode', 1) -ne 2)
                {
                    $k.SetValue('LoginMode', 2, [Microsoft.Win32.RegistryValueKind]::DWord)
                    $needsRestart = $true
                }
                $k.Close()
                break
            }
        }
        catch { }
    }

    foreach ($tcpPath in @(
        "SOFTWARE\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer\SuperSocketNetLib\Tcp",
        "SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer\SuperSocketNetLib\Tcp"
    ))
    {
        try
        {
            $k = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($tcpPath, $true)
            if ($k)
            {
                if ($k.GetValue('Enabled', 0) -ne 1)
                {
                    $k.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
                    $needsRestart = $true
                }
                $k.Close()
                break
            }
        }
        catch { }
    }

    foreach ($ipAllPath in @(
        "SOFTWARE\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer\SuperSocketNetLib\Tcp\IPAll",
        "SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer\SuperSocketNetLib\Tcp\IPAll"
    ))
    {
        try
        {
            $k = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($ipAllPath, $true)
            if ($k)
            {
                $curPort = [string]$k.GetValue('TcpPort', '')
                $curDyn = [string]$k.GetValue('TcpDynamicPorts', '')
                if ($curPort -ne $TcpPort -or $curDyn -ne '')
                {
                    $k.SetValue('TcpPort', $TcpPort, [Microsoft.Win32.RegistryValueKind]::String)
                    $k.SetValue('TcpDynamicPorts', '', [Microsoft.Win32.RegistryValueKind]::String)
                    $needsRestart = $true
                }
                $k.Close()
                break
            }
        }
        catch { }
    }

    foreach ($npPath in @(
        "SOFTWARE\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer\SuperSocketNetLib\Np",
        "SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer\SuperSocketNetLib\Np"
    ))
    {
        try
        {
            $k = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($npPath, $true)
            if ($k)
            {
                if ($k.GetValue('Enabled', 0) -ne 1)
                {
                    $k.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
                    $needsRestart = $true
                }
                $k.Close()
                break
            }
        }
        catch { }
    }

    foreach ($smPath in @(
        "SOFTWARE\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer\SuperSocketNetLib\Sm",
        "SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer\SuperSocketNetLib\Sm"
    ))
    {
        try
        {
            $k = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($smPath, $true)
            if ($k)
            {
                if ($k.GetValue('Enabled', 0) -ne 1)
                {
                    $k.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
                    $needsRestart = $true
                }
                $k.Close()
                break
            }
        }
        catch { }
    }

    if ($needsRestart)
    {
        $svcName = if ($instanceName -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { "MSSQL`$$instanceName" }
        if (-not $servicesToRestart.Contains($svcName)) { [void]$servicesToRestart.Add($svcName) }
    }
}

foreach ($svcName in $servicesToRestart)
{
    try
    {
        Restart-Service -Name $svcName -Force -ErrorAction Stop
    }
    catch
    {
        try { sc.exe stop $svcName | Out-Null } catch { }
        Start-Sleep -Seconds 8
        try { sc.exe start $svcName | Out-Null } catch { }
        Start-Sleep -Seconds 5
    }
}

try { schtasks.exe /Delete /TN $TaskName /F | Out-Null } catch { }
try { Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', "ping 127.0.0.1 -n 3 >nul & del /f /q `"$PSCommandPath`"" -WindowStyle Hidden } catch { }
'@

	try
	{
		foreach ($candidate in $remoteTempCandidates)
		{
			try
			{
				if (-not (Test-Path -LiteralPath $candidate.Share))
				{
					New-Item -Path $candidate.Share -ItemType Directory -Force -ErrorAction Stop | Out-Null
				}
				$probeFile = Join-Path $candidate.Share ("_tbs_probe_{0}.tmp" -f $laneKey)
				Set-Content -LiteralPath $probeFile -Value "probe" -Encoding ASCII -Force -ErrorAction Stop
				Remove-Item -LiteralPath $probeFile -Force -ErrorAction SilentlyContinue
				$remoteScriptSharePath = Join-Path $candidate.Share $scriptName
				$localScriptPath = [System.IO.Path]::Combine($candidate.Local, $scriptName)
				break
			}
			catch
			{
			}
		}
	}
	catch
	{
	}

	if ([string]::IsNullOrWhiteSpace($remoteScriptSharePath) -or [string]::IsNullOrWhiteSpace($localScriptPath))
	{
		try { Write_Log ("Could not schedule SQL enable task for lane {0}: no writable Storeman Temp share found on {1}" -f $laneKey, $MachineName) 'red' | Out-Null } catch { }
		$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{
			TaskPresent  = $false
			LastChecked  = Get-Date
			ScheduledFor = $scheduledDate
			MachineName  = $MachineName
		}
		return $false
	}

	$null = & schtasks.exe /Query /S $MachineName /TN $taskName 2>$null
	$taskExists = ($LASTEXITCODE -eq 0)
	$scriptExists = $false
	try { $scriptExists = Test-Path -LiteralPath $remoteScriptSharePath }
	catch { $scriptExists = $false }
	if ($taskExists -and $scriptExists)
	{
		$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{
			TaskPresent  = $true
			LastChecked  = Get-Date
			ScheduledFor = $scheduledDate
			MachineName  = $MachineName
		}
		return $true
	}
	elseif ($taskExists -and -not $scriptExists)
	{
		try { & schtasks.exe /Delete /S $MachineName /TN $taskName /F 2>$null | Out-Null }
		catch { }
	}

	try
	{
		Set-Content -LiteralPath $remoteScriptSharePath -Value $workerScript -Encoding UTF8 -Force -ErrorAction Stop
	}
	catch
	{
		try { Write_Log ("Could not schedule SQL enable task for lane {0}: failed to write worker script to {1}" -f $laneKey, $remoteScriptSharePath) 'red' | Out-Null } catch { }
		$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{
			TaskPresent  = $false
			LastChecked  = Get-Date
			ScheduledFor = $scheduledDate
			MachineName  = $MachineName
		}
		return $false
	}

	$taskAction = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$localScriptPath`" -TaskName `"$taskName`" -TcpPort `"$TcpPort`""
	$createOutput = & schtasks.exe /Create /S $MachineName /RU SYSTEM /RL HIGHEST /SC ONCE /TN $taskName /TR $taskAction /ST $StartTime /SD $scheduledDateText /F 2>&1
	if ($LASTEXITCODE -eq 0)
	{
		$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{
			TaskPresent  = $true
			LastChecked  = Get-Date
			ScheduledFor = $scheduledDate
			MachineName  = $MachineName
		}
		return $true
	}

	try { Remove-Item -LiteralPath $remoteScriptSharePath -Force -ErrorAction SilentlyContinue }
	catch { }
	$createOutputText = [string]($createOutput | Out-String)
	$createOutputText = $createOutputText.Trim()
	# Previously this output was computed and silently dropped - the failure was invisible.
	try { Write_Log ("Could not schedule SQL enable task for lane {0} on {1}: {2}" -f $laneKey, $MachineName, $createOutputText) 'red' | Out-Null } catch { }
	$script:LaneSqlProtocolEnableTaskState[$laneKey] = [pscustomobject]@{
		TaskPresent  = $false
		LastChecked  = Get-Date
		ScheduledFor = $scheduledDate
		MachineName  = $MachineName
	}
	return $false
}

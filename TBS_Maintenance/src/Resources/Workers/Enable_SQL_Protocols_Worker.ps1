# ---------------------------------------------------------------------------------------------------
# Lane SQL enable - the script the 4:15 AM task runs on the lane (as SYSTEM).
#   Embedded at build time into Ensure_SQL_Enable_Scheduled_Task_For_File_Lane, after the param()
#   line ($TcpPort) and Sql_Protocol_Settings.ps1 ($SetSqlProtocolSettings, $RestartSqlService).
#   Applies the SQL protocol settings, restarts SQL when something changed, makes sure inbound
#   TCP <port> is open in Windows Firewall, writes <script name>.result.txt next to itself for
#   the server to read back, then deletes its task (named like the script) and itself.
#
#   Result file (key=value lines, a blank line, then the details):
#     Time=yyyy-MM-dd HH:mm:ss  Changed=True|False  Restarted=<services>  Firewall=Added|Present|Failed
#     Error=<text, empty when nothing went wrong>
# ---------------------------------------------------------------------------------------------------
$ErrorActionPreference = 'Continue'
$TaskName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ResultPath = [System.IO.Path]::ChangeExtension($PSCommandPath, '.result.txt')
$details = New-Object System.Collections.Generic.List[string]
$changedAny = $false
$restarted = New-Object System.Collections.Generic.List[string]
$firewall = 'Failed'
$errors = New-Object System.Collections.Generic.List[string]

try
{
	$results = @(& $SetSqlProtocolSettings ([Microsoft.Win32.Registry]::LocalMachine) $TcpPort)
	if ($results.Count -eq 0) { $errors.Add('no SQL Server instance found in the registry') }
	foreach ($r in $results)
	{
		$details.Add("[$($r.Instance)]")
		foreach ($l in $r.Lines) { $details.Add('  ' + $l.Text) }
		if ($r.Changed)
		{
			$changedAny = $true
			if (& $RestartSqlService $r.Service) { $restarted.Add($r.Service) }
			else { $errors.Add("$($r.Service) did not report RUNNING within 60 seconds after the restart") }
		}
	}
}
catch { $errors.Add('settings: ' + $_.Exception.Message) }

# Inbound TCP <port> - without it the server can only reach this lane's SQL over Named Pipes / files
try
{
	$ruleName = "TBS_SQL_TCP_$TcpPort"
	$null = & netsh.exe advfirewall firewall show rule "name=$ruleName"
	if ($LASTEXITCODE -eq 0) { $firewall = 'Present' }
	else
	{
		$addOut = & netsh.exe advfirewall firewall add rule "name=$ruleName" dir=in action=allow protocol=TCP "localport=$TcpPort" profile=any
		if ($LASTEXITCODE -eq 0)
		{
			$firewall = 'Added'
			$changedAny = $true
		}
		else { $errors.Add('firewall: ' + (($addOut | Out-String).Trim())) }
	}
	$details.Add("Firewall rule $ruleName (inbound TCP $TcpPort): $firewall")
}
catch { $errors.Add('firewall: ' + $_.Exception.Message) }

try
{
	$text = @(
		'Time=' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
		'Changed=' + $changedAny
		'Restarted=' + ($restarted -join ', ')
		'Firewall=' + $firewall
		'Error=' + ($errors -join '; ')
		''
	) + $details
	[System.IO.File]::WriteAllLines($ResultPath, [string[]]$text)
}
catch { }

try { $null = & schtasks.exe /Delete /TN $TaskName /F } catch { }
# Remove this one-shot script. PowerShell has read it already, so it can delete its own file directly
# (the old hidden 'cmd /c ping -n 3 & del' chain is a classic malware idiom antivirus flags).
try { Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction Stop } catch { }

# ===================================================================================================
#                              FUNCTION: Enable_SQL_Protocols_On_Selected_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Enables TCP/IP, Named Pipes and Shared Memory protocols on the SQL Server instances of
#   the selected lanes: sets a static TCP port (default 1433), disables dynamic ports,
#   restarts the SQL service when needed, and optionally live-tests a Linked Server to pick
#   the working transport. Runs in parallel through a runspace pool (not background jobs)
#   and restarts the programs on lanes that needed a service restart.
#
# Parameters:
#   - StoreNumber         [string]: store to process (required)
#   - tcpPort             [string]: static TCP port (default "1433")
#   - CreateLinkedServers [switch]: on by default; pass :$false to skip
#
# Variables:
#   - $script:FunctionResults: LaneNumToMachineName, DBSERVER (input)
#
# Writes:
#   - SQL Server protocol registry settings + Linked Servers on the selected lanes
# ===================================================================================================

function Enable_SQL_Protocols_On_Selected_Lanes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$tcpPort = "1433",
		[Parameter(Mandatory = $false)]
		[switch]$CreateLinkedServers
	)

	# ---- Default CreateLinkedServers to $true if omitted -------------------------------------------
	if (-not $PSBoundParameters.ContainsKey('CreateLinkedServers')) { $CreateLinkedServers = $true }

	Write_Log "`r`n==================== Starting Enable_SQL_Protocols_On_Selected_Lanes ====================`r`n" "blue"

	# ---- Resolve lane -> machine -------------------------------------------------------------------
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName -or $LaneNumToMachineName.Count -eq 0)
	{
		Write_Log "No lanes available. Please retrieve nodes first." "red"
		Write_Log "`r`n==================== Enable_SQL_Protocols_On_Selected_Lanes Completed ====================" "blue"
		return
	}

	# ---- Local SQL instance for linked-server operations -------------------------------------------
	$localInstance = $script:FunctionResults['DBSERVER']
	if ([string]::IsNullOrWhiteSpace($localInstance) -or $localInstance -eq 'N/A') { $localInstance = 'localhost' }

	# ---- Lane selection UI -------------------------------------------------------------------------
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected or selection cancelled. Exiting." "yellow"
		Write_Log "`r`n==================== Enable_SQL_Protocols_On_Selected_Lanes Completed ====================" "blue"
		return
	}

	# ---- Normalize lane list (ensure 3-digit strings) ----------------------------------------------
	$lanes =
	$selection.Lanes |
	ForEach-Object {
		if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
		else { $_ }
	} |
	ForEach-Object { $_.ToString().PadLeft(3, '0') } |
	Sort-Object -Unique

	Write_Log "Selected lanes: $($lanes -join ', ')" "green"

	# ---- Prepare caches used elsewhere in your script ----------------------------------------------
	if (-not $script:LaneProtocols) { $script:LaneProtocols = @{ } }
	if (-not $script:ProtocolResults) { $script:ProtocolResults = @() }

	# ---- Create a runspace pool (MTA). Throttle to a reasonable degree of parallelism --------------
	$maxThreads = [Math]::Min($lanes.Count, [Math]::Max(2, [Environment]::ProcessorCount))
	$pool = [RunspaceFactory]::CreateRunspacePool(1, $maxThreads)
	$pool.ApartmentState = 'MTA'
	$pool.Open()

	# ---- Script that each runspace will execute (NO nested functions; just inline code) ------------
	$laneWork = @'
param($machine,$lane,$StoreNumber,$tcpPort,$createLinks,$localInstance)

# Per-lane output buffer; we send it back as a single PSCustomObject
$out = @()
$laneNeedsRestart = $false
$protocol = "File"           # default; changed after probe/DMV
$creationTransport = $null   # remember which transport we CREATE with (TCP/NP)
$createdAndTested = $false   # becomes true if linked server SELECT works

try {
    # --- Opening banner ---------------------------------------------------------------------------
    $out += [pscustomobject]@{ Text = "`r`n--- Processing Machine: $machine (Store $StoreNumber, Lane $lane) ---"; Color = "blue" }
    $out += [pscustomobject]@{ Text = "Ensuring RemoteRegistry is running on $machine..."; Color = "gray" }

    # Start RemoteRegistry on-demand; ignore console noise
    sc.exe "\\$machine" config RemoteRegistry start= demand | Out-Null
    sc.exe "\\$machine" start RemoteRegistry | Out-Null

    # --- Open remote HKLM -------------------------------------------------------------------------
    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $machine)

    # --- Enumerate SQL instances from both 64/32-bit locations -----------------------------------
    $instances = @{}
    foreach ($root in @(
        "SOFTWARE\\Microsoft\\Microsoft SQL Server\\Instance Names\\SQL",
        "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\Instance Names\\SQL"
    )) {
        $k = $reg.OpenSubKey($root)
        if ($k) {
            foreach ($n in $k.GetValueNames()) {
                $id = $k.GetValue($n)
                if ($id -and -not $instances.ContainsKey($n)) { $instances[$n] = $id }
            }
            $k.Close()
        }
    }

    if ($instances.Count -eq 0) {
        $out += [pscustomobject]@{ Text = "No SQL instances found on $machine."; Color = "yellow" }
        $reg.Close()
    }
    else {
        foreach ($instanceName in $instances.Keys) {
            $instanceID = $instances[$instanceName]
            $out += [pscustomobject]@{ Text = "Processing SQL Instance: $instanceName (ID: $instanceID)"; Color = "blue" }
            $needsRestart = $false

            # ---- Enable Mixed Mode (LoginMode=2) -------------------------------------------------
            foreach ($authPath in @(
                "SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer"
            )) {
                $k = $reg.OpenSubKey($authPath, $true)
                if ($k) {
                    if ($k.GetValue("LoginMode", 1) -ne 2) {
                        $k.SetValue("LoginMode", 2, [Microsoft.Win32.RegistryValueKind]::DWord)
                        $out += [pscustomobject]@{ Text = "Mixed Mode Authentication enabled at $authPath."; Color = "green" }
                        $needsRestart = $true
                    } else {
                        $out += [pscustomobject]@{ Text = "Mixed Mode Authentication already enabled at $authPath."; Color = "gray" }
                    }
                    $k.Close(); break
                }
            }

            # ---- Enable TCP ----------------------------------------------------------------------
            foreach ($p in @(
                "SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp"
            )) {
                $k = $reg.OpenSubKey($p, $true)
                if ($k) {
                    if ($k.GetValue('Enabled', 0) -ne 1) {
                        $k.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
                        $out += [pscustomobject]@{ Text = "TCP/IP protocol enabled at $p."; Color = "green" }
                        $needsRestart = $true
                    } else {
                        $out += [pscustomobject]@{ Text = "TCP/IP already enabled at $p."; Color = "gray" }
                    }
                    $k.Close(); break
                }
            }

            # ---- Set static port and clear dynamic -----------------------------------------------
            foreach ($p in @(
                "SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp\\IPAll",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Tcp\\IPAll"
            )) {
                $k = $reg.OpenSubKey($p, $true)
                if ($k) {
                    $curPort = $k.GetValue('TcpPort', "")
                    $curDyn  = $k.GetValue('TcpDynamicPorts', "")
                    if ($curPort -ne $tcpPort -or $curDyn -ne "") {
                        $k.SetValue('TcpPort', $tcpPort, [Microsoft.Win32.RegistryValueKind]::String)
                        $k.SetValue('TcpDynamicPorts', '', [Microsoft.Win32.RegistryValueKind]::String)
                        $out += [pscustomobject]@{ Text = "Registry port set to $tcpPort at $p."; Color = "green" }
                        $needsRestart = $true
                    } else {
                        $out += [pscustomobject]@{ Text = "TCP port already set to $tcpPort at $p."; Color = "gray" }
                    }
                    $k.Close(); break
                }
            }

            # ---- Enable Named Pipes --------------------------------------------------------------
            foreach ($p in @(
                "SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Np",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Np"
            )) {
                $k = $reg.OpenSubKey($p, $true)
                if ($k) {
                    if ($k.GetValue('Enabled', 0) -ne 1) {
                        $k.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
                        $out += [pscustomobject]@{ Text = "Named Pipes protocol enabled at $p."; Color = "green" }
                        $needsRestart = $true
                    } else {
                        $out += [pscustomobject]@{ Text = "Named Pipes already enabled at $p."; Color = "gray" }
                    }
                    $k.Close(); break
                }
            }

            # ---- Enable Shared Memory ------------------------------------------------------------
            foreach ($p in @(
                "SOFTWARE\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Sm",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server\\$instanceID\\MSSQLServer\\SuperSocketNetLib\\Sm"
            )) {
                $k = $reg.OpenSubKey($p, $true)
                if ($k) {
                    if ($k.GetValue('Enabled', 0) -ne 1) {
                        $k.SetValue('Enabled', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
                        $out += [pscustomobject]@{ Text = "Shared Memory protocol enabled at $p."; Color = "green" }
                        $needsRestart = $true
                    } else {
                        $out += [pscustomobject]@{ Text = "Shared Memory already enabled at $p."; Color = "gray" }
                    }
                    $k.Close(); break
                }
            }

            # ---- Restart SQL service if any change ----------------------------------------------
            if ($needsRestart) {
                $svcName = if ($instanceName -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { "MSSQL`$$instanceName" }
                $out += [pscustomobject]@{ Text = "Restarting SQL Service $svcName on $machine..."; Color = "gray" }
                sc.exe "\\$machine" stop $svcName | Out-Null
                Start-Sleep -Seconds 10
                sc.exe "\\$machine" start $svcName | Out-Null
                Start-Sleep -Seconds 5
                $out += [pscustomobject]@{ Text = "SQL Service $svcName restarted successfully on $machine."; Color = "green" }
                $laneNeedsRestart = $true
            } else {
                $out += [pscustomobject]@{ Text = "No protocol/auth changes for $instanceName on $machine. No restart needed."; Color = "green" }
            }
        } # foreach instance

        $reg.Close()
    }

    # --- Quick transport probe: TCP then NP --------------------------------------------------------
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $task = $tcpClient.ConnectAsync($machine, [int]$tcpPort)
        if ($task.Wait(1500) -and $tcpClient.Connected) {
            $tcpClient.Close()
            $protocol = "TCP"
        }
        else {
            $npCs = "Server=np:\\\\$machine\\pipe\\sql\\query;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=5"
            $cn = New-Object System.Data.SqlClient.SqlConnection $npCs
            try {
                $cn.Open()
                $cmd = $cn.CreateCommand(); $cmd.CommandText = "SELECT 1"; $cmd.CommandTimeout = 2
                [void]$cmd.ExecuteScalar()
                $protocol = "Named Pipes"
            } catch { } finally { $cn.Close() }
        }
    } catch { }

    if ($protocol -eq 'File') {
        $out += [pscustomobject]@{ Text = "Pre-probe did not confirm TCP or NP (likely timeouts). Will use Linked Server live test to confirm."; Color = "yellow" }
    }

    # --- Linked server creation and live test (optional) -------------------------------------------
    if ($createLinks) {
        # Provider preference; try to filter by installed list (best effort).
        $pref = @('SQLOLEDB','MSOLEDBSQL','SQLNCLI11','SQLNCLI','MSDASQL')
        try {
            $cs = "Server=$localInstance;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=8"
            $cn = New-Object System.Data.SqlClient.SqlConnection $cs
            $cn.Open()
            try {
                $cmd = $cn.CreateCommand(); $cmd.CommandTimeout = 8; $cmd.CommandText = "EXEC master.sys.sp_enum_oledb_providers;"
                $rdr = $cmd.ExecuteReader()
                $dt = New-Object System.Data.DataTable
                $dt.Load($rdr)
                $installed = @()
                foreach ($row in $dt.Rows) {
                    $col = if ($dt.Columns.Contains('name')) { 'name' } elseif ($dt.Columns.Contains('PROGID')) { 'PROGID' } else { $dt.Columns[0].ColumnName }
                    $installed += [string]$row[$col]
                }
                if ($installed.Count -gt 0) {
                    $pref2 = @()
                    foreach ($p in $pref) { if ($installed -contains $p) { $pref2 += $p } }
                    if ($pref2.Count -gt 0) { $pref = $pref2 }
                }
            } finally { $cn.Close() }
        } catch { }

        $tcpDataSource = "tcp:$machine,$tcpPort"
        $linkName = $machine
        $firstCreatedProv = $null
        $firstCreatedTsql = $null

        foreach ($prov in $pref) {
            try {
                $optTsql = "BEGIN TRY EXEC master.dbo.sp_MSset_oledb_prop N'$prov', N'AllowInProcess', 1; END TRY BEGIN CATCH END CATCH;
BEGIN TRY EXEC master.dbo.sp_MSset_oledb_prop N'$prov', N'DynamicParameters', 1; END TRY BEGIN CATCH END CATCH;"

                if ($prov -ne 'MSDASQL') {
                    if ($protocol -eq 'TCP') {
                        $creationTransport = 'TCP'
                        $createTsql = @"
IF EXISTS(SELECT 1 FROM sys.servers WHERE name=N'$linkName')
    EXEC master.dbo.sp_dropserver @server=N'$linkName', @droplogins='droplogins';
$optTsql
EXEC master.dbo.sp_addlinkedserver @server=N'$linkName', @srvproduct=N'', @provider=N'$prov', @datasrc=N'$tcpDataSource';
EXEC master.dbo.sp_serveroption N'$linkName', N'data access', N'true';
EXEC master.dbo.sp_serveroption N'$linkName', N'use remote collation', N'true';
EXEC master.dbo.sp_serveroption N'$linkName', N'rpc out', N'true';
IF NOT EXISTS (SELECT 1 FROM sys.linked_logins ll JOIN sys.servers s ON s.server_id=ll.server_id WHERE s.name=N'$linkName')
    EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=N'$linkName', @useself='TRUE';
"@
                    }
                    else {
                        $creationTransport = 'Named Pipes'
                        $createTsql = @"
IF EXISTS(SELECT 1 FROM sys.servers WHERE name=N'$linkName')
    EXEC master.dbo.sp_dropserver @server=N'$linkName', @droplogins='droplogins';
$optTsql
EXEC master.dbo.sp_addlinkedserver
    @server=N'$linkName', @srvproduct=N'', @provider=N'$prov',
    @datasrc=N'\\$machine\pipe\sql\query', @provstr=N'Network Library=dbnmpntw';
EXEC master.dbo.sp_serveroption N'$linkName', N'data access', N'true';
EXEC master.dbo.sp_serveroption N'$linkName', N'use remote collation', N'true';
EXEC master.dbo.sp_serveroption N'$linkName', N'rpc out', N'true';
IF NOT EXISTS (SELECT 1 FROM sys.linked_logins ll JOIN sys.servers s ON s.server_id=ll.server_id WHERE s.name=N'$linkName')
    EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=N'$linkName', @useself='TRUE';
"@
                    }
                }
                else {
                    if ($protocol -eq 'TCP') {
                        $creationTransport = 'TCP'
                        $provstr18 = "Driver={ODBC Driver 18 for SQL Server};Server=$machine,$tcpPort;Trusted_Connection=Yes;TrustServerCertificate=Yes;"
                        $provstr17 = "Driver={ODBC Driver 17 for SQL Server};Server=$machine,$tcpPort;Trusted_Connection=Yes;TrustServerCertificate=Yes;"
                    }
                    else {
                        $creationTransport = 'Named Pipes'
                        $npServer = "np:\\\\$machine\\pipe\\sql\\query"
                        $provstr18 = "Driver={ODBC Driver 18 for SQL Server};Server=$npServer;Trusted_Connection=Yes;TrustServerCertificate=Yes;"
                        $provstr17 = "Driver={ODBC Driver 17 for SQL Server};Server=$npServer;Trusted_Connection=Yes;TrustServerCertificate=Yes;"
                    }
                    $createTsql = @"
IF EXISTS(SELECT 1 FROM sys.servers WHERE name=N'$linkName')
    EXEC master.dbo.sp_dropserver @server=N'$linkName', @droplogins='droplogins';
$optTsql
DECLARE @ok bit = 0;
BEGIN TRY
    EXEC master.dbo.sp_addlinkedserver @server=N'$linkName', @srvproduct=N'', @provider=N'MSDASQL', @provstr=N'$provstr18';
    SET @ok = 1;
END TRY BEGIN CATCH END CATCH;
IF (@ok = 0)
BEGIN TRY
    EXEC master.dbo.sp_addlinkedserver @server=N'$linkName', @srvproduct=N'', @provider=N'MSDASQL', @provstr=N'$provstr17';
    SET @ok = 1;
END TRY BEGIN CATCH END CATCH;
IF (@ok = 1)
BEGIN
    EXEC master.dbo.sp_serveroption N'$linkName', N'data access', N'true';
    EXEC master.dbo.sp_serveroption N'$linkName', N'use remote collation', N'true';
    EXEC master.dbo.sp_serveroption N'$linkName', N'rpc out', N'true';
    IF NOT EXISTS (SELECT 1 FROM sys.linked_logins ll JOIN sys.servers s ON s.server_id=ll.server_id WHERE s.name=N'$linkName')
        EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=N'$linkName', @useself='TRUE';
END
ELSE
BEGIN
    RAISERROR('MSDASQL addlinkedserver failed for both ODBC 18 and 17', 11, 1);
END
"@
                }

                # -- Execute the CREATE script
                $cs = "Server=$localInstance;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=30"
                $cn = New-Object System.Data.SqlClient.SqlConnection $cs
                $cn.Open()
                try {
                    $cmd = $cn.CreateCommand()
                    $cmd.CommandTimeout = 30
                    $cmd.CommandText = $createTsql
                    [void]$cmd.ExecuteNonQuery()
                } finally { $cn.Close() }

                if (-not $firstCreatedProv) { $firstCreatedProv = $prov; $firstCreatedTsql = $createTsql }

                # -- Live test (no MSDTC): query remote master via linked server
                $testOk = $false
                $cs = "Server=$localInstance;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=8"
                $cn = New-Object System.Data.SqlClient.SqlConnection $cs
                $cn.Open()
                try {
                    $cmd = $cn.CreateCommand()
                    $cmd.CommandTimeout = 8
                    $cmd.CommandText = "SELECT TOP 1 name FROM [$linkName].master.sys.databases"
                    $rdr = $cmd.ExecuteReader()
                    $dt = New-Object System.Data.DataTable
                    $dt.Load($rdr)
                    $testOk = ($dt -ne $null)
                } catch {
                    $testOk = $false
                } finally { $cn.Close() }

                if ($testOk) {
                    $createdAndTested = $true
                    # Try to detect the actual transport on the remote side via DMV
                    $normalized = $false
                    try {
                        $cs = "Server=$localInstance;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=8"
                        $cn = New-Object System.Data.SqlClient.SqlConnection $cs
                        $cn.Open()
                        try {
                            $cmd = $cn.CreateCommand()
                            $cmd.CommandTimeout = 8
                            $cmd.CommandText = "EXEC ('SELECT TOP 1 net_transport FROM sys.dm_exec_connections WHERE session_id = @@SPID') AT [$linkName]"
                            $rdr = $cmd.ExecuteReader()
                            $dt = New-Object System.Data.DataTable
                            $dt.Load($rdr)
                            if ($dt -and $dt.Rows.Count -gt 0) {
                                $rt = [string]$dt.Rows[0][0]
                                if ($rt -match 'TCP')    { $protocol = 'TCP';          $normalized = $true }
                                elseif ($rt -match 'Named') { $protocol = 'Named Pipes'; $normalized = $true }
                                elseif ($rt -match 'Shared') { $protocol = 'Shared Memory'; $normalized = $true }
                                $out += [pscustomobject]@{ Text = "Remote net_transport via [$linkName]: $rt" + ($(if ($normalized) { " (persisting as '$protocol')" } else { "" })); Color = "gray" }
                            }
                        } finally { $cn.Close() }
                    } catch { }

                    if (-not $normalized -or [string]::IsNullOrWhiteSpace($protocol) -or $protocol -eq 'File') { $protocol = $creationTransport }
                    $out += [pscustomobject]@{ Text = "Linked Server [$linkName] created via provider '$prov' over $protocol - **LIVE TEST PASSED**."; Color = "green" }
                    break
                }
                else {
                    # Drop and try next provider
                    try {
                        $cs = "Server=$localInstance;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=15"
                        $cn = New-Object System.Data.SqlClient.SqlConnection $cs
                        $cn.Open()
                        try {
                            $cmd = $cn.CreateCommand()
                            $cmd.CommandTimeout = 15
                            $cmd.CommandText = "EXEC master.dbo.sp_dropserver @server=N'$linkName', @droplogins='droplogins';"
                            [void]$cmd.ExecuteNonQuery()
                        } finally { $cn.Close() }
                    } catch { }
                    $out += [pscustomobject]@{ Text = "Provider '$prov' created [$linkName] over $creationTransport but live test FAILED - trying next provider..."; Color = "yellow" }
                }
            }
            catch {
                $out += [pscustomobject]@{ Text = "Provider '$prov' failed to create Linked Server [$linkName] - trying next provider..."; Color = "gray" }
            }
        } # foreach provider

        if (-not $createdAndTested -and $firstCreatedProv) {
            # Leave a creatable provider so admin can map a SQL login later
            try {
                $cs = "Server=$localInstance;Database=master;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;Connection Timeout=30"
                $cn = New-Object System.Data.SqlClient.SqlConnection $cs
                $cn.Open()
                try {
                    $cmd = $cn.CreateCommand()
                    $cmd.CommandTimeout = 30
                    $cmd.CommandText = $firstCreatedTsql
                    [void]$cmd.ExecuteNonQuery()
                } finally { $cn.Close() }
            } catch { }
            if ($protocol -eq 'File' -or [string]::IsNullOrWhiteSpace($protocol)) { $protocol = $creationTransport }
            $out += [pscustomobject]@{ Text = "No provider passed live test. Re-created [$linkName] using '$firstCreatedProv' over $protocol so you can map a SQL login if needed."; Color = "yellow" }
        }
        elseif (-not $createdAndTested -and -not $firstCreatedProv) {
            $out += [pscustomobject]@{ Text = "All provider attempts failed creating Linked Server [$linkName]."; Color = "red" }
        }
    }
    else {
        $out += [pscustomobject]@{ Text = "CreateLinkedServers disabled; skipped Linked Server creation."; Color = "gray" }
    }
}
catch {
    $out += [pscustomobject]@{ Text = "Failed to process [$machine]: $_"; Color = "red" }
    $protocol = "File"
}

# Final guard to avoid returning 'File' on success
if ($createdAndTested -and ($protocol -eq 'File' -or [string]::IsNullOrWhiteSpace($protocol))) {
    $protocol = $creationTransport
}

# Return structured result (parent thread will log and persist)
[pscustomobject]@{
    Output    = $out
    Protocol  = $protocol
    Lane      = $lane
    Machine   = $machine
    Restarted = $laneNeedsRestart
}
'@

	# ---- Launch all lanes in parallel via runspaces -----------------------------------------------
	$tasks = @()
	foreach ($lane in $lanes)
	{
		$machine = $LaneNumToMachineName[$lane]
		if (-not $machine)
		{
			Write_Log "Machine name not found for lane $lane. Skipping." "yellow"
			continue
		}

		# Create PowerShell instance bound to the runspace pool
		$ps = [PowerShell]::Create()
		$ps.RunspacePool = $pool
		# IMPORTANT: pass a pure boolean for $createLinks
		$ps.AddScript($laneWork).AddArgument($machine).AddArgument($lane).AddArgument($StoreNumber).AddArgument($tcpPort).AddArgument([bool]$CreateLinkedServers.IsPresent).AddArgument($localInstance) | Out-Null

		$handle = $ps.BeginInvoke()
		$tasks += [pscustomobject]@{ Lane = $lane; Machine = $machine; PS = $ps; Handle = $handle; Done = $false }
	}

	# ---- Collect results as lanes complete (quasi-live) -------------------------------------------
	$restarted = @()
	while ($true)
	{
		$allDone = $true
		foreach ($t in $tasks)
		{
			if (-not $t.Done)
			{
				if ($t.Handle.IsCompleted)
				{
					# Finish and dispose
					$res = $t.PS.EndInvoke($t.Handle)
					$t.PS.Dispose()
					$t.Done = $true

					# Emit buffered logs for this lane
					foreach ($line in $res.Output) { Write_Log $line.Text $line.Color }

					# Persist in-memory protocol maps
					$script:LaneProtocols[$res.Lane] = $res.Protocol
					$script:ProtocolResults = $script:ProtocolResults | Where-Object { $_.Lane -ne $res.Lane }
					$script:ProtocolResults += [PSCustomObject]@{ Lane = $res.Lane; Protocol = $res.Protocol }

					# Persist to file (dedup by lane)
					$protocolResultsFile = $script:ProtocolResultsFile
					$laneStr = $res.Lane.ToString().PadLeft(3, '0')
					$proto = $res.Protocol
					$dir = [System.IO.Path]::GetDirectoryName($protocolResultsFile)
					if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
					if (-not (Test-Path $protocolResultsFile)) { New-Item -Path $protocolResultsFile -ItemType File -Force | Out-Null }
					$all = @()
					if (Test-Path $protocolResultsFile)
					{
						$all = Get-Content -LiteralPath $protocolResultsFile -ErrorAction SilentlyContinue | Where-Object { $_ -match '\S' }
					}
					$all = $all | Where-Object { -not ($_ -match "^\s*0*${laneStr}\s*,") }
					$all += "$laneStr,$proto"
					$sorted = $all | Sort-Object { ($_ -split ',')[0] -as [int] }
					[System.IO.File]::WriteAllLines($protocolResultsFile, $sorted, [System.Text.Encoding]::UTF8)

					if ($proto -eq 'File' -and -not [string]::IsNullOrWhiteSpace($res.Machine))
					{
						# Menu action: report the scheduling outcome in the log box.
						# (Failures log their own red line from inside the function.)
						$schedOk = $false
						try
						{
							$schedOk = [bool](Ensure_SQL_Enable_Scheduled_Task_For_File_Lane -LaneNumber $laneStr -MachineName $res.Machine)
						}
						catch { }
						if ($schedOk)
						{
							$schedState = $null
							try { $schedState = $script:LaneSqlProtocolEnableTaskState[$laneStr] } catch { }
							$whenText = if ($schedState -and $schedState.ScheduledFor) { [string]$schedState.ScheduledFor.ToString('MM/dd/yyyy HH:mm') } else { 'next 4:15 AM' }
							Write_Log ("Scheduled SQL enable for {0} (Lane {1}) for {2}" -f $res.Machine, $laneStr, $whenText) 'green'
						}
					}
					elseif ($script:LaneSqlProtocolEnableTaskState.ContainsKey($laneStr))
					{
						$script:LaneSqlProtocolEnableTaskState[$laneStr] = [pscustomobject]@{
							TaskPresent  = $false
							LastChecked  = Get-Date
							ScheduledFor = $null
							MachineName  = $res.Machine
						}
					}

					if ($res.Restarted) { $restarted += $res.Lane }

					# Final per-lane summary line
					Write_Log "Protocol detected for $($res.Machine) (Lane $($res.Lane)): $proto" "magenta"
				}
				else
				{
					$allDone = $false
				}
			}
		}
		if ($allDone) { break }
		Start-Sleep -Milliseconds 150
	}

	# ---- Cleanup pool -----------------------------------------------------------------------------
	$pool.Close()
	$pool.Dispose()

	# ---- Post-actions for lanes that needed restart -----------------------------------------------
	if ($restarted.Count -gt 0)
	{
		Send_Restart_All_Programs -StoreNumber $StoreNumber -LaneNumbers $restarted -Silent
		Write_Log "Restart All Programs sent to restarted lanes: $($restarted -join ', ')" "green"
	}

	Write_Log "`r`n==================== Enable_SQL_Protocols_On_Selected_Lanes Completed ====================" "blue"
}

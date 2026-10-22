# ===================================================================================================
#                              FUNCTION: Send_SMS_Key_to_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Syncs the server SMS key to the lane databases. "Run now" uses the in-memory lane
#   protocols and tries Windows auth first, then SQL auth (Tecnica), then an SQI file
#   fallback; the scheduled task variant reads Protocol_Results.txt at runtime and uses
#   SQL auth for lane DB updates when the protocol is not File.
#
# Parameters:
#   - StoreNumber           [string]: store to process (required)
#   - LaneNumbers           [string[]]: lanes to target (default: all)
#   - NonInteractive        [switch]: run without prompts
#   - SqlTimeoutSeconds     [int]: SQL query timeout in seconds (default 30)
#   - ConnectTimeoutSeconds [int]: SQL connect timeout in seconds (default 1)
#
# Variables:
#   - $script:LaneProtocols: per-lane protocol cache used by "Run now"
#   - $script:FunctionResults: lane map and SQL module (input)
#
# Writes:
#   - lane databases (key sync) and a scheduled task for the recurring sync
# ===================================================================================================

function Send_SMS_Key_to_Lanes
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[string[]]$LaneNumbers = $null,
		[Parameter(Mandatory = $false)]
		[switch]$NonInteractive,
		[Parameter(Mandatory = $false)]
		[int]$SqlTimeoutSeconds = 30,
		[Parameter(Mandatory = $false)]
		[int]$ConnectTimeoutSeconds = 1
	)

	Write_Log "`r`n==================== Starting Send_SMS_Key_to_Lanes ====================`r`n" "blue"

	# Normalize store number (3 digits)
	$StoreNumber = "$StoreNumber".PadLeft(3, '0')

	# Ensure FunctionResults exists
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }

	# SQL AUTH creds (used for scheduled; and as fallback for Run Now if WinAuth fails)
	$SqlAuthUser = "Tecnica"
	$SqlAuthPass = "TB$upp0rT"

	# Validate SQL module availability
	$SqlModule = $script:FunctionResults['SqlModuleName']
	if ([string]::IsNullOrWhiteSpace($SqlModule) -or $SqlModule -eq "None")
	{
		Write_Log "No SQL Server module available (SqlServer or SQLPS). Cannot sync server key." "red"
		Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
		return
	}
	try { Import-Module $SqlModule -ErrorAction Stop }
	catch
	{
		Write_Log "Failed to import SQL module '$SqlModule'. Error: $_" "red"
		Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
		return
	}

	# Ensure cached lane map exists (fallback only if missing)
	if (-not $script:FunctionResults['LaneNumToMachineName'])
	{
		Write_Log "LaneNumToMachineName missing; calling Retrieve_Nodes as fallback." "yellow"
		$null = Retrieve_Nodes -StoreNumber $StoreNumber
	}
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName)
	{
		Write_Log "LaneNumToMachineName still missing; cannot continue." "red"
		Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
		return
	}

	# OfficePath required for SQI fallback
	if (-not $OfficePath -or -not (Test-Path $OfficePath))
	{
		$guessOffice = $script:StoremanOfficeRoots | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
		if (-not $guessOffice) { $guessOffice = $script:PrimaryStoremanOfficePath }
		if (Test-Path $guessOffice) { $OfficePath = $guessOffice }
	}
	if (-not $OfficePath -or -not (Test-Path $OfficePath))
	{
		Write_Log "OfficePath missing/invalid; SQI fallback would fail. OfficePath='$OfficePath'." "red"
		Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
		return
	}

	# IMPORTANT:
	# - RUN NOW uses ONLY this variable. DEFAULT = File unless explicitly TCP/Named Pipes.
	if (-not $script:LaneProtocols) { $script:LaneProtocols = @{ } }

	# Protocol results file path (Scheduled Task will read this at runtime)
	$protoFile = $script:ProtocolResultsFile

	# MODE selection (Interactive only): Run Now vs Schedule Only
	$mode = "run"
	if (-not $NonInteractive)
	{
		try { Add-Type -AssemblyName System.Windows.Forms | Out-Null }
		catch { }

		$script:__modePick = $null
		$formMode = New-Object System.Windows.Forms.Form
		$formMode.Text = "Server Key Sync"
		$formMode.Size = New-Object System.Drawing.Size(540, 230)
		$formMode.StartPosition = "CenterScreen"
		$formMode.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
		$formMode.MaximizeBox = $false
		$formMode.MinimizeBox = $false

		$lblMode = New-Object System.Windows.Forms.Label
		$lblMode.Text = "Choose what you want to do:"
		$lblMode.AutoSize = $true
		$lblMode.Location = New-Object System.Drawing.Point(18, 18)
		$formMode.Controls.Add($lblMode)

		$btnRunNow = New-Object System.Windows.Forms.Button
		$btnRunNow.Text = "Run Now"
		$btnRunNow.Size = New-Object System.Drawing.Size(150, 45)
		$btnRunNow.Location = New-Object System.Drawing.Point(25, 65)
		$btnRunNow.Add_Click({ $script:__modePick = "run"; $formMode.DialogResult = [System.Windows.Forms.DialogResult]::OK; $formMode.Close() })
		$formMode.Controls.Add($btnRunNow)

		$btnScheduleOnly = New-Object System.Windows.Forms.Button
		$btnScheduleOnly.Text = "Schedule Task Only"
		$btnScheduleOnly.Size = New-Object System.Drawing.Size(190, 45)
		$btnScheduleOnly.Location = New-Object System.Drawing.Point(190, 65)
		$btnScheduleOnly.Add_Click({ $script:__modePick = "schedule"; $formMode.DialogResult = [System.Windows.Forms.DialogResult]::OK; $formMode.Close() })
		$formMode.Controls.Add($btnScheduleOnly)

		$btnCancel = New-Object System.Windows.Forms.Button
		$btnCancel.Text = "Cancel"
		$btnCancel.Size = New-Object System.Drawing.Size(110, 45)
		$btnCancel.Location = New-Object System.Drawing.Point(395, 65)
		$btnCancel.Add_Click({ $script:__modePick = "cancel"; $formMode.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $formMode.Close() })
		$formMode.Controls.Add($btnCancel)

		[void]$formMode.ShowDialog()
		$mode = $script:__modePick
		Remove-Variable -Name __modePick -Scope Script -ErrorAction SilentlyContinue

		if ($mode -eq "cancel" -or [string]::IsNullOrWhiteSpace($mode))
		{
			Write_Log "User cancelled Server Key Sync." "yellow"
			Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
			return
		}
	}

	# Lane selection (interactive) or normalize lanes (noninteractive)
	if (-not $NonInteractive)
	{
		Write_Log "Opening lane selection..." "yellow"
		$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
		if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
		{
			Write_Log "Lane selection cancelled or returned no lanes." "yellow"
			Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
			return
		}

		$tmp = @()
		foreach ($ln in $selection.Lanes)
		{
			$raw = $null
			if ($ln -is [pscustomobject] -and $ln.LaneNumber) { $raw = "$($ln.LaneNumber)" }
			else { $raw = "$ln" }

			$n = ($raw -replace '[^\d]', '')
			if (-not [string]::IsNullOrWhiteSpace($n)) { $tmp += $n.PadLeft(3, '0') }
		}
		$LaneNumbers = @($tmp | Where-Object { $_ } | Select-Object -Unique)
	}
	else
	{
		if (-not $LaneNumbers -or $LaneNumbers.Count -eq 0)
		{
			Write_Log "NonInteractive run requires -LaneNumbers." "red"
			Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
			return
		}
		$tmp = @()
		foreach ($ln in $LaneNumbers)
		{
			$n = ("$ln" -replace '[^\d]', '')
			if (-not [string]::IsNullOrWhiteSpace($n)) { $tmp += $n.PadLeft(3, '0') }
		}
		$LaneNumbers = @($tmp | Where-Object { $_ } | Select-Object -Unique)
	}

	Write_Log ("Selected lanes: " + ($LaneNumbers -join ", ")) "cyan"

	# Decide whether to deploy now
	$doDeployNow = $true
	if (-not $NonInteractive -and $mode -eq "schedule") { $doDeployNow = $false }

	# Results containers
	$UpdatedViaSql = New-Object System.Collections.Generic.List[string]
	$UpdatedViaFile = New-Object System.Collections.Generic.List[string]
	$FailedLanes = New-Object System.Collections.Generic.List[string]

	# =================================================================================================
	# DEPLOY NOW
	#   - Uses ONLY $script:LaneProtocols (already populated)
	#   - DEFAULT = File unless explicitly TCP/Named Pipes
	#   - Tries WinAuth first, then SQL Auth, then SQI
	# =================================================================================================
	if ($doDeployNow)
	{
		try { Add-Type -AssemblyName System.Data 2>$null }
		catch { }

		# LOCAL key read
		$localConnStr = $script:FunctionResults['ConnectionString']
		if ([string]::IsNullOrWhiteSpace($localConnStr) -or $localConnStr -eq "N/A")
		{
			Write_Log "Local DB ConnectionString missing (FunctionResults['ConnectionString'])." "red"
			Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
			return
		}

		# Enforce connect timeout on LOCAL conn string
		$localConnStrEff = $localConnStr
		if ($localConnStrEff -notmatch '(?i)\bConnect\s*Timeout\s*=')
		{
			if ($localConnStrEff.Trim().EndsWith(';')) { $localConnStrEff += "Connect Timeout=$ConnectTimeoutSeconds;" }
			else { $localConnStrEff += ";Connect Timeout=$ConnectTimeoutSeconds;" }
		}
		else
		{
			$localConnStrEff = [regex]::Replace($localConnStrEff, '(?i)\bConnect\s*Timeout\s*=\s*\d+\s*', "Connect Timeout=$ConnectTimeoutSeconds")
		}

		Write_Log "Reading SYS_TAB.F1243 from LOCAL DB..." "yellow"

		$rawValues = @()
		try
		{
			$vals = New-Object System.Collections.Generic.List[string]
			$cn = New-Object System.Data.SqlClient.SqlConnection $localConnStrEff
			try
			{
				$cn.Open()
				$cmd = $cn.CreateCommand()
				$cmd.CommandText = "SELECT F1243 FROM SYS_TAB;"
				$cmd.CommandTimeout = $SqlTimeoutSeconds
				$rdr = $cmd.ExecuteReader()
				try
				{
					while ($rdr.Read())
					{
						if (-not $rdr.IsDBNull(0)) { $vals.Add([string]$rdr.GetValue(0)) | Out-Null }
					}
				}
				finally
				{
					try { $rdr.Close() }
					catch { }
					try { $rdr.Dispose() }
					catch { }
				}
			}
			finally
			{
				try { $cn.Close() }
				catch { }
				try { $cn.Dispose() }
				catch { }
			}
			$rawValues = @($vals)
		}
		catch { $rawValues = @() }

		$serverKey = $null
		$bestLen = -1
		foreach ($v in $rawValues)
		{
			if ($v -eq $null) { continue }
			$t = ([string]$v).Trim()
			$stripped = [regex]::Replace($t, '[\s\u00A0]+', '')
			if (-not [string]::IsNullOrWhiteSpace($t) -and -not [string]::IsNullOrWhiteSpace($stripped))
			{
				if ($t.Length -gt $bestLen) { $bestLen = $t.Length; $serverKey = $t }
			}
		}

		if ([string]::IsNullOrWhiteSpace($serverKey))
		{
			Write_Log "Server key (SYS_TAB.F1243) came back empty/null from LOCAL DB. Aborting." "red"
			Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
			return
		}

		Write_Log ("Server key retrieved (length {0})." -f $serverKey.Length) "green"

		$safeServerKey = $serverKey.Replace("'", "''")
		$updateSql = "UPDATE SYS_TAB SET F1243 = '$safeServerKey'"

		# SQI fallback file (ANSI 1252, no BOM)
		$SqiFileName = "UPDATE_SMS_KEY.SQI"
		$SqiContent = "@dbEXEC($updateSql)`r`n"
		$ansiEncoding = [System.Text.Encoding]::GetEncoding(1252)
		$TempSqiPath = Join-Path ([System.IO.Path]::GetTempPath()) $SqiFileName

		try
		{
			[System.IO.File]::WriteAllText($TempSqiPath, $SqiContent, $ansiEncoding)
			try { attrib -a -r "$TempSqiPath" 2>$null | Out-Null }
			catch { }
			try { Set-ItemProperty -Path $TempSqiPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal) -ErrorAction SilentlyContinue }
			catch { }
		}
		catch
		{
			Write_Log "Failed to create temp SQI '$TempSqiPath'. Error: $_" "red"
			Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
			return
		}

		foreach ($LaneNumber in $LaneNumbers)
		{
			Write_Log "`r`n--- Lane $LaneNumber ---" "magenta"

			$laneInfo = $null
			try { $laneInfo = Get_All_Lanes_Database_Info -LaneNumber $LaneNumber }
			catch { $laneInfo = $null }
			if (-not $laneInfo)
			{
				Write_Log "Could not retrieve cached DB info for lane $LaneNumber. Skipping." "red"
				$FailedLanes.Add($LaneNumber) | Out-Null
				continue
			}

			# Normalize PSCustomObject -> hashtable (so ['TcpConnStr'] works)
			if ($laneInfo -isnot [hashtable])
			{
				$h = @{ }
				foreach ($p in $laneInfo.PSObject.Properties) { $h[$p.Name] = $p.Value }
				$laneInfo = $h
			}

			$tcpConnStr = $laneInfo['TcpConnStr']
			$namedPipesConnStr = $laneInfo['NamedPipesConnStr']

			# RUN NOW: protocol comes ONLY from $script:LaneProtocols; DEFAULT = File
			$machineName = $LaneNumToMachineName[$LaneNumber]
			$protocolPref = $null
			if ($script:LaneProtocols.ContainsKey($LaneNumber)) { $protocolPref = [string]$script:LaneProtocols[$LaneNumber] }
			elseif ($machineName -and $script:LaneProtocols.ContainsKey($machineName)) { $protocolPref = [string]$script:LaneProtocols[$machineName] }
			elseif ($machineName -and $script:LaneProtocols.ContainsKey($machineName.ToLower())) { $protocolPref = [string]$script:LaneProtocols[$machineName.ToLower()] }

			$protocolPrefNorm = "File"
			if (-not [string]::IsNullOrWhiteSpace($protocolPref))
			{
				$t = $protocolPref.Trim()
				if ($t -match '(?i)^tcp$') { $protocolPrefNorm = "TCP" }
				elseif ($t -match '(?i)pipe') { $protocolPrefNorm = "Named Pipes" }
				elseif ($t -match '(?i)^offline$') { $protocolPrefNorm = "Offline" }
				else { $protocolPrefNorm = "File" }
			}

			if ($protocolPrefNorm -eq "Offline")
			{
				Write_Log "Lane $LaneNumber ($machineName) is offline. Skipping." "yellow"
				$FailedLanes.Add($LaneNumber) | Out-Null
				continue
			}

			$sqlWorked = $false

			# SQL attempts ONLY when protocol explicitly says TCP/Named Pipes
			if ($protocolPrefNorm -ne "File")
			{
				$tryOrder = @("TCP", "Named Pipes")
				if ($protocolPrefNorm -eq "Named Pipes") { $tryOrder = @("Named Pipes", "TCP") }

				foreach ($proto in $tryOrder)
				{
					$connToUse = if ($proto -eq "Named Pipes") { $namedPipesConnStr }
					else { $tcpConnStr }
					if ([string]::IsNullOrWhiteSpace($connToUse)) { continue }

					# -----------------------
					# Attempt 1: Windows Auth
					# -----------------------
					$connWin = $connToUse
					$connWin = [regex]::Replace($connWin, '(?i)\bUser\s*ID\s*=\s*[^;]*;?', '')
					$connWin = [regex]::Replace($connWin, '(?i)\bUID\s*=\s*[^;]*;?', '')
					$connWin = [regex]::Replace($connWin, '(?i)\bPassword\s*=\s*[^;]*;?', '')
					$connWin = [regex]::Replace($connWin, '(?i)\bPWD\s*=\s*[^;]*;?', '')
					$connWin = [regex]::Replace($connWin, '(?i)\bTrusted_Connection\s*=\s*Yes\s*;?', '')
					$connWin = [regex]::Replace($connWin, '(?i)\bIntegrated\s+Security\s*=\s*False\s*;?', '')

					if ($connWin -notmatch '(?i)\bIntegrated\s+Security\s*=')
					{
						if ($connWin.Trim().EndsWith(';')) { $connWin += "Integrated Security=True;" }
						else { $connWin += ";Integrated Security=True;" }
					}
					else
					{
						$connWin = [regex]::Replace($connWin, '(?i)\bIntegrated\s+Security\s*=\s*[^;]*', 'Integrated Security=True')
					}

					if ($connWin -notmatch '(?i)\bConnect\s*Timeout\s*=')
					{
						if ($connWin.Trim().EndsWith(';')) { $connWin += "Connect Timeout=$ConnectTimeoutSeconds;" }
						else { $connWin += ";Connect Timeout=$ConnectTimeoutSeconds;" }
					}
					else
					{
						$connWin = [regex]::Replace($connWin, '(?i)\bConnect\s*Timeout\s*=\s*\d+\s*', "Connect Timeout=$ConnectTimeoutSeconds")
					}

					try
					{
						$cn2 = New-Object System.Data.SqlClient.SqlConnection $connWin
						try
						{
							$cn2.Open()
							$cmd2 = $cn2.CreateCommand()
							$cmd2.CommandText = $updateSql
							$cmd2.CommandTimeout = $SqlTimeoutSeconds
							[void]$cmd2.ExecuteNonQuery()
							$sqlWorked = $true
							Write_Log "Updated via SQL ($proto) on lane $LaneNumber (Windows Auth)." "green"
							$UpdatedViaSql.Add($LaneNumber) | Out-Null
							break
						}
						finally
						{
							try { $cn2.Close() }
							catch { }
							try { $cn2.Dispose() }
							catch { }
						}
					}
					catch
					{
						$sqlWorked = $false
					}

					if ($sqlWorked) { break }

					# -----------------------
					# Attempt 2: SQL Auth
					# -----------------------
					$connSql = $connToUse
					$connSql = [regex]::Replace($connSql, '(?i)\bIntegrated\s+Security\s*=\s*True\s*;?', '')
					$connSql = [regex]::Replace($connSql, '(?i)\bTrusted_Connection\s*=\s*Yes\s*;?', '')

					if ($connSql -notmatch '(?i)\bUser\s*ID\s*=')
					{
						if ($connSql.Trim().EndsWith(';')) { $connSql += "User ID=$SqlAuthUser;" }
						else { $connSql += ";User ID=$SqlAuthUser;" }
					}
					if ($connSql -notmatch '(?i)\bPassword\s*=')
					{
						if ($connSql.Trim().EndsWith(';')) { $connSql += "Password=$SqlAuthPass;" }
						else { $connSql += ";Password=$SqlAuthPass;" }
					}

					if ($connSql -notmatch '(?i)\bConnect\s*Timeout\s*=')
					{
						if ($connSql.Trim().EndsWith(';')) { $connSql += "Connect Timeout=$ConnectTimeoutSeconds;" }
						else { $connSql += ";Connect Timeout=$ConnectTimeoutSeconds;" }
					}
					else
					{
						$connSql = [regex]::Replace($connSql, '(?i)\bConnect\s*Timeout\s*=\s*\d+\s*', "Connect Timeout=$ConnectTimeoutSeconds")
					}

					try
					{
						$cn3 = New-Object System.Data.SqlClient.SqlConnection $connSql
						try
						{
							$cn3.Open()
							$cmd3 = $cn3.CreateCommand()
							$cmd3.CommandText = $updateSql
							$cmd3.CommandTimeout = $SqlTimeoutSeconds
							[void]$cmd3.ExecuteNonQuery()
							$sqlWorked = $true
							Write_Log "Updated via SQL ($proto) on lane $LaneNumber (SQL Auth)." "green"
							$UpdatedViaSql.Add($LaneNumber) | Out-Null
							break
						}
						finally
						{
							try { $cn3.Close() }
							catch { }
							try { $cn3.Dispose() }
							catch { }
						}
					}
					catch
					{
						$sqlWorked = $false
					}

					if ($sqlWorked) { break }
				}
			}

			# SQI fallback (protocol File OR SQL failed)
			if (-not $sqlWorked)
			{
				$LaneXFDir = Join-Path $OfficePath ("XF{0}{1}" -f $StoreNumber, $LaneNumber)
				$DestSqiPath = Join-Path $LaneXFDir $SqiFileName

				if (-not (Test-Path $LaneXFDir))
				{
					Write_Log "Lane XF folder not found: $LaneXFDir" "red"
					$FailedLanes.Add($LaneNumber) | Out-Null
					continue
				}

				try
				{
					Copy-Item -Path $TempSqiPath -Destination $DestSqiPath -Force
					try { attrib -a -r "$DestSqiPath" 2>$null | Out-Null }
					catch { }
					try { Set-ItemProperty -Path $DestSqiPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal) -ErrorAction SilentlyContinue }
					catch { }

					$UpdatedViaFile.Add($LaneNumber) | Out-Null
					Write_Log "Copied SQI to $DestSqiPath (archive cleared)" "green"
				}
				catch
				{
					Write_Log "Failed to copy SQI for lane $LaneNumber. Error: $_" "red"
					$FailedLanes.Add($LaneNumber) | Out-Null
				}
			}
		}

		if ($UpdatedViaSql.Count -gt 0) { Write_Log ("Updated via SQL: " + ($UpdatedViaSql -join ", ")) "green" }
		if ($UpdatedViaFile.Count -gt 0) { Write_Log ("Updated via SQI: " + ($UpdatedViaFile -join ", ")) "green" }
		if ($FailedLanes.Count -gt 0) { Write_Log ("Failed lanes: " + ($FailedLanes -join ", ")) "red" }
	}

	# Restart prompt (interactive Run Now only)
	if (-not $NonInteractive -and $mode -eq "run" -and $doDeployNow)
	{
		$successful = @(@($UpdatedViaSql) + @($UpdatedViaFile) | Where-Object { $_ } | Select-Object -Unique)
		if ($successful.Count -gt 0)
		{
			try { Add-Type -AssemblyName System.Windows.Forms | Out-Null }
			catch { }
			$ansRestart = [System.Windows.Forms.MessageBox]::Show(
				"Server key deployed successfully to: " + ($successful -join ", ") + "`r`n`r`nRestart ALL programs on these lanes now?",
				"Restart Programs?",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Question
			)

			if ($ansRestart -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				foreach ($ln in $successful)
				{
					$machineName = $LaneNumToMachineName[$ln]
					if ([string]::IsNullOrWhiteSpace($machineName))
					{
						Write_Log "Could not resolve machine name for lane $ln (restart skipped)." "yellow"
						continue
					}

					try
					{
						$mailslotAddress = "\\$machineName\mailslot\SMSStart_${StoreNumber}${ln}"
						$sent = [MailslotSender]::SendMailslotCommand($mailslotAddress, "@exec(RESTART_ALL=PROGRAMS).")
						if ($sent) { Write_Log "Restart sent to $machineName (lane $ln)." "green" }
						else { Write_Log "Restart FAILED to $machineName (lane $ln)." "red" }
					}
					catch
					{
						Write_Log "Restart attempt failed for lane $ln. Error: $_" "red"
					}
				}
			}
			else
			{
				Write_Log "User chose NOT to restart programs after deploy." "yellow"
			}
		}
	}

	# =================================================================================================
	# SCHEDULE (interactive only) -> ONE self-sufficient PS1
	#   - Uses SQL AUTH for lane DB updates when protocol != File
	#   - Reads Protocol_Results.txt at runtime (DEFAULT File unless explicitly TCP/Named Pipes)
	# =================================================================================================
	if (-not $NonInteractive)
	{
		$doSchedule = $false
		if ($mode -eq "schedule") { $doSchedule = $true }
		else
		{
			try { Add-Type -AssemblyName System.Windows.Forms | Out-Null }
			catch { }
			$ansSchedule = [System.Windows.Forms.MessageBox]::Show(
				"Do you want to create/update a Scheduled Task for the selected lanes?",
				"Create Scheduled Task?",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Question
			)
			if ($ansSchedule -eq [System.Windows.Forms.DialogResult]::Yes) { $doSchedule = $true }
		}

		if ($doSchedule)
		{
			try { Add-Type -AssemblyName System.Windows.Forms | Out-Null }
			catch { }

			$formFreq = New-Object System.Windows.Forms.Form
			$formFreq.Text = "Server Key Sync Frequency"
			$formFreq.Width = 520
			$formFreq.Height = 240
			$formFreq.StartPosition = "CenterScreen"

			$lbl = New-Object System.Windows.Forms.Label
			$lbl.AutoSize = $true
			$lbl.Location = New-Object System.Drawing.Point(15, 20)
			$lbl.Text = "Run every how many minutes? (1 - 1439)"
			$formFreq.Controls.Add($lbl)

			$tb = New-Object System.Windows.Forms.TextBox
			$tb.Location = New-Object System.Drawing.Point(20, 50)
			$tb.Width = 100
			$tb.Text = "60"
			$formFreq.Controls.Add($tb)

			$lbl2 = New-Object System.Windows.Forms.Label
			$lbl2.AutoSize = $true
			$lbl2.Location = New-Object System.Drawing.Point(15, 95)
			$lbl2.Text = "Selected lanes: " + ($LaneNumbers -join ", ")
			$formFreq.Controls.Add($lbl2)

			$okBtn = New-Object System.Windows.Forms.Button
			$okBtn.Text = "OK"
			$okBtn.Location = New-Object System.Drawing.Point(300, 160)
			$okBtn.Add_Click({ $formFreq.DialogResult = [System.Windows.Forms.DialogResult]::OK })
			$formFreq.Controls.Add($okBtn)

			$cancelBtn = New-Object System.Windows.Forms.Button
			$cancelBtn.Text = "Cancel"
			$cancelBtn.Location = New-Object System.Drawing.Point(380, 160)
			$cancelBtn.Add_Click({ $formFreq.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
			$formFreq.Controls.Add($cancelBtn)

			$formFreq.AcceptButton = $okBtn
			$formFreq.CancelButton = $cancelBtn

			if ($formFreq.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK)
			{
				Write_Log "Scheduled Task setup cancelled by user." "yellow"
				Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
				return
			}

			[int]$freqMinutes = 0
			if (-not [int]::TryParse($tb.Text, [ref]$freqMinutes) -or $freqMinutes -lt 1 -or $freqMinutes -gt 1439)
			{
				Write_Log "Invalid frequency entered. Using 60 minutes." "yellow"
				$freqMinutes = 60
			}

			# Required cached data for self-sufficient script
			$localConnStr = $script:FunctionResults['ConnectionString']
			if ([string]::IsNullOrWhiteSpace($localConnStr) -or $localConnStr -eq "N/A")
			{
				Write_Log "Cannot schedule because local ConnectionString is missing in FunctionResults['ConnectionString']." "red"
				Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
				return
			}

			$laneConnMap = @{ }
			foreach ($ln in $LaneNumbers)
			{
				$li = $null
				try { $li = Get_All_Lanes_Database_Info -LaneNumber $ln }
				catch { $li = $null }
				if ($li)
				{
					if ($li -isnot [hashtable])
					{
						$h2 = @{ }
						foreach ($p in $li.PSObject.Properties) { $h2[$p.Name] = $p.Value }
						$li = $h2
					}

					$laneConnMap[$ln] = @{
						TcpConnStr	      = $li['TcpConnStr']
						NamedPipesConnStr = $li['NamedPipesConnStr']
					}
				}
			}

			$taskScriptPath = Join-Path $script:ScriptsFolder ("Send_SMS_Key_Task_Store{0}.ps1" -f $StoreNumber)
			$latestLogPath = Join-Path $script:ScriptsLog    ("Send_SMS_Key_LATEST_Store{0}.txt" -f $StoreNumber)

			$enc = [System.Text.Encoding]::UTF8
			$localConnB64 = [Convert]::ToBase64String($enc.GetBytes($localConnStr))
			$laneConnB64 = [Convert]::ToBase64String($enc.GetBytes(($laneConnMap | ConvertTo-Json -Depth 6)))

			$laneListLiteral = ($LaneNumbers | ForEach-Object { "'$_'" }) -join ", "

			$sb = New-Object System.Text.StringBuilder
			[void]$sb.AppendLine("# AUTO-GENERATED Send_SMS_Key task (self-sufficient, one file)")
			[void]$sb.AppendLine("`$ErrorActionPreference = 'Stop'")
			[void]$sb.AppendLine("try { Add-Type -AssemblyName System.Data 2>`$null } catch { }")
			[void]$sb.AppendLine("`$SqlUser = 'Tecnica'")
			[void]$sb.AppendLine("`$SqlPass = 'TB`$upp0rT'")
			[void]$sb.AppendLine("`$StoreNumber = '$StoreNumber'")
			[void]$sb.AppendLine("`$OfficePath  = '$OfficePath'")
			[void]$sb.AppendLine("`$SqlTimeoutSeconds = $SqlTimeoutSeconds")
			[void]$sb.AppendLine("`$ConnectTimeoutSeconds = $ConnectTimeoutSeconds")
			[void]$sb.AppendLine("`$LatestLogPath = '$latestLogPath'")
			[void]$sb.AppendLine("`$ProtocolFile = '$protoFile'")
			[void]$sb.AppendLine("`$enc = [System.Text.Encoding]::UTF8")
			[void]$sb.AppendLine("`$LocalConnStr = `$enc.GetString([Convert]::FromBase64String('$localConnB64'))")
			[void]$sb.AppendLine("`$LaneConnMap  = ConvertFrom-Json (`$enc.GetString([Convert]::FromBase64String('$laneConnB64')))")
			[void]$sb.AppendLine("`$LaneNumbers = @($laneListLiteral)")
			[void]$sb.AppendLine("`$UpdatedViaSql  = New-Object System.Collections.Generic.List[string]")
			[void]$sb.AppendLine("`$UpdatedViaFile = New-Object System.Collections.Generic.List[string]")
			[void]$sb.AppendLine("`$FailedLanes    = New-Object System.Collections.Generic.List[string]")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("# Load protocol map from txt (DEFAULT = File unless explicitly TCP/Named Pipes)")
			[void]$sb.AppendLine("`$ProtoMap = @{}")
			[void]$sb.AppendLine("if (Test-Path `$ProtocolFile) {")
			[void]$sb.AppendLine("  try {")
			[void]$sb.AppendLine("    foreach (`$line in (Get-Content -LiteralPath `$ProtocolFile -ErrorAction SilentlyContinue)) {")
			[void]$sb.AppendLine("      if (`$line -match '^\s*([^,]+),\s*([^,]+)\s*$') {")
			[void]$sb.AppendLine("        `$laneRaw = `$matches[1].Trim()")
			[void]$sb.AppendLine("        `$protoRaw = `$matches[2].Trim()")
			[void]$sb.AppendLine("        `$num = (`$laneRaw -replace '[^\d]','').PadLeft(3,'0')")
			[void]$sb.AppendLine("        if (-not [string]::IsNullOrWhiteSpace(`$num)) { `$ProtoMap[`$num] = `$protoRaw }")
			[void]$sb.AppendLine("      }")
			[void]$sb.AppendLine("    }")
			[void]$sb.AppendLine("  } catch { }")
			[void]$sb.AppendLine("}")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("# Enforce connect timeout on LOCAL conn string")
			[void]$sb.AppendLine("`$localEff = `$LocalConnStr")
			[void]$sb.AppendLine("if (`$localEff -notmatch '(?i)\bConnect\s*Timeout\s*=') { if (`$localEff.Trim().EndsWith(';')) { `$localEff += ""Connect Timeout=`$ConnectTimeoutSeconds;"" } else { `$localEff += "";Connect Timeout=`$ConnectTimeoutSeconds;"" } }")
			[void]$sb.AppendLine("else { `$localEff = [regex]::Replace(`$localEff, '(?i)\bConnect\s*Timeout\s*=\s*\d+\s*', ""Connect Timeout=`$ConnectTimeoutSeconds"") }")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("# Read key")
			[void]$sb.AppendLine("`$rawValues = @()")
			[void]$sb.AppendLine("try {")
			[void]$sb.AppendLine("  `$vals = New-Object System.Collections.Generic.List[string]")
			[void]$sb.AppendLine("  `$cn = New-Object System.Data.SqlClient.SqlConnection `$localEff")
			[void]$sb.AppendLine("  try { `$cn.Open(); `$cmd = `$cn.CreateCommand(); `$cmd.CommandText = ""SELECT F1243 FROM SYS_TAB;""; `$cmd.CommandTimeout = `$SqlTimeoutSeconds; `$rdr = `$cmd.ExecuteReader();")
			[void]$sb.AppendLine("    try { while (`$rdr.Read()) { if (-not `$rdr.IsDBNull(0)) { `$vals.Add([string]`$rdr.GetValue(0)) | Out-Null } } } finally { try { `$rdr.Close() } catch { } try { `$rdr.Dispose() } catch { } }")
			[void]$sb.AppendLine("  } finally { try { `$cn.Close() } catch { } try { `$cn.Dispose() } catch { } }")
			[void]$sb.AppendLine("  `$rawValues = @(`$vals)")
			[void]$sb.AppendLine("} catch { `$rawValues = @() }")
			[void]$sb.AppendLine("`$serverKey = `$null; `$bestLen = -1")
			[void]$sb.AppendLine("foreach (`$v in `$rawValues) { if (`$v -eq `$null) { continue }; `$t = ([string]`$v).Trim(); `$stripped = [regex]::Replace(`$t,'[\s\u00A0]+',''); if (-not [string]::IsNullOrWhiteSpace(`$t) -and -not [string]::IsNullOrWhiteSpace(`$stripped)) { if (`$t.Length -gt `$bestLen) { `$bestLen = `$t.Length; `$serverKey = `$t } } }")
			[void]$sb.AppendLine("if ([string]::IsNullOrWhiteSpace(`$serverKey)) { throw ""Server key empty/null from LOCAL DB."" }")
			[void]$sb.AppendLine("`$safeServerKey = `$serverKey.Replace(""'"",""''"")")
			[void]$sb.AppendLine("`$updateSql = ""UPDATE SYS_TAB SET F1243 = '"" + `$safeServerKey + ""'""")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("# SQI (ANSI/1252)")
			[void]$sb.AppendLine("`$SqiFileName = ""UPDATE_SERVERKEY.SQI""")
			[void]$sb.AppendLine("`$SqiContent  = ""@dbEXEC("" + `$updateSql + "")`r`n""")
			[void]$sb.AppendLine("`$ansi = [System.Text.Encoding]::GetEncoding(1252)")
			[void]$sb.AppendLine("`$TempSqiPath = Join-Path ([System.IO.Path]::GetTempPath()) `$SqiFileName")
			[void]$sb.AppendLine("[System.IO.File]::WriteAllText(`$TempSqiPath, `$SqiContent, `$ansi)")
			[void]$sb.AppendLine("try { attrib -a -r `"`$TempSqiPath`" 2>`$null | Out-Null } catch { }")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("foreach (`$ln in `$LaneNumbers) {")
			[void]$sb.AppendLine("  `$pref = 'File'")
			[void]$sb.AppendLine("  if (`$ProtoMap.ContainsKey(`$ln)) { `$pref = [string]`$ProtoMap[`$ln] }")
			[void]$sb.AppendLine("  `$prefNorm = 'File'")
			[void]$sb.AppendLine("  if (-not [string]::IsNullOrWhiteSpace(`$pref)) { if (`$pref -match '(?i)^tcp$') { `$prefNorm='TCP' } elseif (`$pref -match '(?i)pipe') { `$prefNorm='Named Pipes' } }")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("  `$laneProp = `$LaneConnMap.PSObject.Properties[`$ln]")
			[void]$sb.AppendLine("  `$laneObj  = if (`$laneProp) { `$laneProp.Value } else { `$null }")
			[void]$sb.AppendLine("  `$tcpConn  = if (`$laneObj) { [string]`$laneObj.TcpConnStr } else { `$null }")
			[void]$sb.AppendLine("  `$npConn   = if (`$laneObj) { [string]`$laneObj.NamedPipesConnStr } else { `$null }")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("  `$sqlWorked = `$false")
			[void]$sb.AppendLine("  if (`$prefNorm -ne 'File') {")
			[void]$sb.AppendLine("    `$tryOrder = @('TCP','Named Pipes'); if (`$prefNorm -eq 'Named Pipes') { `$tryOrder=@('Named Pipes','TCP') }")
			[void]$sb.AppendLine("    foreach (`$proto in `$tryOrder) {")
			[void]$sb.AppendLine("      `$connToUse = if (`$proto -eq 'Named Pipes') { `$npConn } else { `$tcpConn }")
			[void]$sb.AppendLine("      if ([string]::IsNullOrWhiteSpace(`$connToUse)) { continue }")
			[void]$sb.AppendLine("      `$connEff = `$connToUse")
			[void]$sb.AppendLine("      `$connEff = [regex]::Replace(`$connEff, '(?i)\bIntegrated\s+Security\s*=\s*True\s*;?', '')")
			[void]$sb.AppendLine("      `$connEff = [regex]::Replace(`$connEff, '(?i)\bTrusted_Connection\s*=\s*Yes\s*;?', '')")
			[void]$sb.AppendLine("      if (`$connEff -notmatch '(?i)\bUser\s*ID\s*=') { if (`$connEff.Trim().EndsWith(';')) { `$connEff += ""User ID=`$SqlUser;"" } else { `$connEff += "";User ID=`$SqlUser;"" } }")
			[void]$sb.AppendLine("      if (`$connEff -notmatch '(?i)\bPassword\s*=') { if (`$connEff.Trim().EndsWith(';')) { `$connEff += ""Password=`$SqlPass;"" } else { `$connEff += "";Password=`$SqlPass;"" } }")
			[void]$sb.AppendLine("      if (`$connEff -notmatch '(?i)\bConnect\s*Timeout\s*=') { if (`$connEff.Trim().EndsWith(';')) { `$connEff += ""Connect Timeout=`$ConnectTimeoutSeconds;"" } else { `$connEff += "";Connect Timeout=`$ConnectTimeoutSeconds;"" } }")
			[void]$sb.AppendLine("      else { `$connEff = [regex]::Replace(`$connEff, '(?i)\bConnect\s*Timeout\s*=\s*\d+\s*', ""Connect Timeout=`$ConnectTimeoutSeconds"") }")
			[void]$sb.AppendLine("      try { `$cn2 = New-Object System.Data.SqlClient.SqlConnection `$connEff; try { `$cn2.Open(); `$cmd2 = `$cn2.CreateCommand(); `$cmd2.CommandText = `$updateSql; `$cmd2.CommandTimeout = `$SqlTimeoutSeconds; [void]`$cmd2.ExecuteNonQuery(); `$sqlWorked = `$true } finally { try { `$cn2.Close() } catch { } try { `$cn2.Dispose() } catch { } } } catch { `$sqlWorked = `$false }")
			[void]$sb.AppendLine("      if (`$sqlWorked) { `$UpdatedViaSql.Add(`$ln) | Out-Null; break }")
			[void]$sb.AppendLine("    }")
			[void]$sb.AppendLine("  }")
			[void]$sb.AppendLine("  if (-not `$sqlWorked) {")
			[void]$sb.AppendLine("    `$laneXF = Join-Path `$OfficePath (""XF{0}{1}"" -f `$StoreNumber, `$ln)")
			[void]$sb.AppendLine("    `$dest = Join-Path `$laneXF `$SqiFileName")
			[void]$sb.AppendLine("    if (-not (Test-Path `$laneXF)) { `$FailedLanes.Add(`$ln) | Out-Null }")
			[void]$sb.AppendLine("    else { try { Copy-Item -Path `$TempSqiPath -Destination `$dest -Force; try { attrib -a -r `"`$dest`" 2>`$null | Out-Null } catch { }; `$UpdatedViaFile.Add(`$ln) | Out-Null } catch { `$FailedLanes.Add(`$ln) | Out-Null } }")
			[void]$sb.AppendLine("  }")
			[void]$sb.AppendLine("}")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("`$okAll = @(@(`$UpdatedViaSql) + @(`$UpdatedViaFile) | Where-Object { `$_ } | Select-Object -Unique)")
			[void]$sb.AppendLine("`$status = ""FAILED""; if (`$okAll.Count -gt 0) { `$status = ""OK""; if (`$FailedLanes.Count -gt 0) { `$status = ""PARTIAL"" } }")
			[void]$sb.AppendLine("`$okAllText = ""(none)""; if (`$okAll.Count -gt 0) { `$okAllText = (`$okAll -join "", "") }")
			[void]$sb.AppendLine("`$sqlText = ""(none)""; if (`$UpdatedViaSql.Count -gt 0) { `$sqlText = (@(`$UpdatedViaSql) -join "", "") }")
			[void]$sb.AppendLine("`$sqiText = ""(none)""; if (`$UpdatedViaFile.Count -gt 0) { `$sqiText = (@(`$UpdatedViaFile) -join "", "") }")
			[void]$sb.AppendLine("`$failedText = ""(none)""; if (`$FailedLanes.Count -gt 0) { `$failedText = (@(`$FailedLanes) -join "", "") }")
			[void]$sb.AppendLine("`$lines = New-Object System.Collections.Generic.List[string]")
			[void]$sb.AppendLine("`$lines.Add((""Timestamp : {0}"" -f (Get-Date -Format ""yyyy-MM-dd HH:mm:ss""))) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""Computer  : {0}"" -f `$env:COMPUTERNAME)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""Store     : {0}"" -f `$StoreNumber)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""ProtocolFile : {0}"" -f `$ProtocolFile)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""Status    : {0}"" -f `$status)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""Selected  : {0}"" -f (`$LaneNumbers -join "", ""))) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""OK (ALL)  : {0}"" -f `$okAllText)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""OK (SQL)  : {0}"" -f `$sqlText)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""OK (SQI)  : {0}"" -f `$sqiText)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""Failed    : {0}"" -f `$failedText)) | Out-Null")
			[void]$sb.AppendLine("`$logDir = [System.IO.Path]::GetDirectoryName(`$LatestLogPath)")
			[void]$sb.AppendLine("if (-not (Test-Path `$logDir)) { New-Item -Path `$logDir -ItemType Directory -Force | Out-Null }")
			[void]$sb.AppendLine("`$lines | Set-Content -Path `$LatestLogPath -Encoding UTF8")

			try
			{
				Set-Content -Path $taskScriptPath -Value ($sb.ToString() -replace "`n", "`r`n") -Encoding UTF8
				Write_Log "Task script written/updated (single file): $taskScriptPath" "green"
				Write_Log "Latest results log (overwrites each run): $latestLogPath" "cyan"
			}
			catch
			{
				Write_Log "Failed to write task script '$taskScriptPath'. Error: $_" "red"
				Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
				return
			}

			try
			{
				$taskName = "Send_SMS_Key - Store $StoreNumber"
				$psExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
				$tr = "`"$psExe`" -NoProfile -ExecutionPolicy Bypass -File `"$taskScriptPath`""

				$null = schtasks /Delete /TN "$taskName" /F 2>$null
				$out = schtasks /Create /TN "$taskName" /TR $tr /SC MINUTE /MO $freqMinutes /RL HIGHEST /F /RU "SYSTEM" 2>&1

				if ($LASTEXITCODE -eq 0)
				{
					Write_Log "Scheduled Task created/updated: $taskName (every $freqMinutes minutes, SYSTEM)." "green"
				}
				else
				{
					Write_Log "Failed to create scheduled task '$taskName'. schtasks said: $out" "red"
					Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
					return
				}
			}
			catch
			{
				Write_Log "Failed to create/update Scheduled Task. Error: $_" "red"
				Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
				return
			}
		}
	}

	Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
}

# ===================================================================================================
#                              FUNCTION: Close_Open_Transactions
# ---------------------------------------------------------------------------------------------------
# Description:
#   Closes open POS transactions on the lanes referenced by recent XE error files: scans the
#   store XE folder for S*.*.* files (max 30 days old), extracts the lane number from each
#   file name, marks the lane's open SAL_HDR transactions as CLOSE via SQL (file fallback),
#   and can send a restart-all-programs command over the mailslot after deployment.
#
# Parameters:
#   - StoreNumber  [string]: store to process (required)
#
# Variables:
#   - $script:FunctionResults: SqlModuleName and the lane maps (input)
#   - $OfficePath: Storeman Office root (XE folder)
#
# Writes:
#   - lane databases (SAL_HDR F1067 = CLOSE) and Scripts_by_Alex_C.T\Closed_Transactions_LOG.txt
# ===================================================================================================

function Close_Open_Transactions
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)
	
	# ==================== Start and validate XE path ====================
	Write_Log "`r`n==================== Starting Close_Open_Transactions ====================`r`n" "blue"
	
	$XEFolderPath = "$OfficePath\XE${StoreNumber}901"
	if (-not (Test-Path $XEFolderPath))
	{
		Write_Log "XE folder not found: $XEFolderPath" "red"
		return
	}
	
	# ==================== Validate SQL module availability ====================
	$SqlModule = $script:FunctionResults['SqlModuleName']
	if (-not $SqlModule -or $SqlModule -eq "None")
	{
		Write_Log "No SQL Server module available (SqlServer or SQLPS). Cannot close transactions." "red"
		return
	}
	
	# SQL string for closing all open transactions (manual fallback)
	$CloseTransactionManual = "@dbEXEC(UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1067 <> 'CLOSE')"
	
	# Paths for logs (keep a persistent transaction log)
	$LogFolderPath = "$BasePath\Scripts_by_Alex_C.T"
	$LogFilePath = Join-Path -Path $LogFolderPath -ChildPath "Closed_Transactions_LOG.txt"
	if (-not (Test-Path $LogFolderPath))
	{
		try { New-Item -Path $LogFolderPath -ItemType Directory -Force | Out-Null }
		catch { Write_Log "Failed to create log directory '$LogFolderPath'. Error: $_" "red"; return }
	}
	
	$MatchedTransactions = $false
	
	try
	{
		$currentTime = Get-Date
		# Only process XE files less than 30 days old and starting with S (SMSStart)
		$files = Get-ChildItem -Path $XEFolderPath -Filter "S*.???" | Where-Object {
			($currentTime - $_.LastWriteTime).TotalDays -le 30
		}
		
		# ==================== PROCESS EACH XE ERROR FILE FOUND ====================
		if ($files -and $files.Count -gt 0)
		{
			foreach ($file in $files)
			{
				try
				{
					# Attempt to extract the lane number from the filename
					if ($file.Name -match '^S.*\.(\d{3})$') { $LaneNumber = $Matches[1] }
					else { continue }
					
					# Read contents for parsing
					$content = Get-Content -Path $file.FullName
					$fromLine = $content | Where-Object { $_ -like 'From:*' }
					$subjectLine = $content | Where-Object { $_ -like 'Subject:*' }
					$msgLine = $content | Where-Object { $_ -like 'MSG:*' }
					$lastRecordedStatusLine = $content | Where-Object { $_ -like 'Last recorded status:*' }
					
					# Parse out store/lane numbers from the "From:" line
					if ($fromLine -match 'From:\s*(\d{3})(\d{3})')
					{
						$fileStoreNumber = $Matches[1]
						$fileLaneNumber = $Matches[2]
						if ($fileStoreNumber -eq $StoreNumber -and $fileLaneNumber -eq $LaneNumber)
						{
							# Get subject/MSG and check if it matches open transaction/health check
							if ($subjectLine -match 'Subject:\s*(.*)')
							{
								$subject = $Matches[1].Trim()
								if ($subject -eq 'Health' -and $msgLine -match 'MSG:\s*(.*)' -and $Matches[1].Trim() -eq 'This application is not running.')
								{
									# Get transaction number from status line
									if ($lastRecordedStatusLine -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)')
									{
										$transactionNumber = $Matches[1]
										$laneKey = $LaneNumber.PadLeft(3, '0')
										$protocolType = if ($script:LaneProtocols.ContainsKey($laneKey)) { $script:LaneProtocols[$laneKey] }
										else { "File" }
										if ($protocolType -eq 'Offline')
										{
											Write_Log "Lane $LaneNumber is offline. Skipping transaction close." "yellow"
											continue
										}
										
										# Get full DB connection info for this lane
										$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $LaneNumber
										$namedPipesConnStr = $laneInfo['NamedPipesConnStr']
										$tcpConnStr = $laneInfo['TcpConnStr']
										$tcpServer = $laneInfo['TcpServer']
										$dbName = $laneInfo['DBName']
										$closeSQL = "UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1032 = $transactionNumber"
										$protocolWorked = $false
										
										# Try protocol (Named Pipes/TCP), otherwise fallback to file
										Import-Module $SqlModule -ErrorAction Stop
										$supportsConnStr = (Get-Command Invoke-Sqlcmd).Parameters.ContainsKey('ConnectionString')
										
										# ----------- Try Named Pipes protocol first -----------
										if ($protocolType -eq "Named Pipes")
										{
											try
											{
												if ($supportsConnStr)
												{
													Invoke-Sqlcmd -ConnectionString $namedPipesConnStr -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
												}
												else
												{
													Invoke-Sqlcmd -ServerInstance $tcpServer -Database $dbName -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
												}
												$protocolWorked = $true
												Write_Log "Closed transaction $transactionNumber via SQL protocol (Named Pipes) on lane $LaneNumber." "green"
											}
											catch { Write_Log "Named Pipes failed for lane $LaneNumber." "yellow" }
										}
										# ----------- Try TCP protocol -----------
										elseif ($protocolType -eq "TCP")
										{
											try
											{
												if ($supportsConnStr)
												{
													Invoke-Sqlcmd -ConnectionString $tcpConnStr -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
												}
												else
												{
													Invoke-Sqlcmd -ServerInstance $tcpServer -Database $dbName -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
												}
												$protocolWorked = $true
												Write_Log "Closed transaction $transactionNumber via SQL protocol (TCP) on lane $LaneNumber." "green"
											}
											catch { Write_Log "TCP failed for lane $LaneNumber." "yellow" }
										}
										
										# ----------- Fallback: File-based SQI if protocol failed -----------
										if (-not $protocolWorked)
										{
											$CloseTransactionAuto = "@dbEXEC(UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1032 = $transactionNumber)"
											$LaneDirectory = "$OfficePath\XF${StoreNumber}${LaneNumber}"
											if (Test-Path $LaneDirectory)
											{
												$CloseTransactionFilePath = Join-Path -Path $LaneDirectory -ChildPath "Close_Transaction.sqi"
												Set-Content -Path $CloseTransactionFilePath -Value $CloseTransactionAuto -Encoding ASCII
												Set-ItemProperty -Path $CloseTransactionFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
												Write_Log "Wrote Close_Transaction.sqi file for lane $LaneNumber (file fallback)." "yellow"
											}
											else
											{
												Write_Log "Lane directory $LaneDirectory not found (file fallback)." "yellow"
											}
										}
										
										# Log the close action
										$logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Closed transaction $transactionNumber on lane $LaneNumber"
										Add-Content -Path $LogFilePath -Value $logMessage
										
										# Remove processed XE file
										Remove-Item -Path $file.FullName -Force
										
										Write_Log "Processed file $($file.Name) for lane $LaneNumber and closed transaction $transactionNumber" "green"
										$MatchedTransactions = $true
										
										# ----------- Optionally restart all programs on the lane -----------
										Start-Sleep -Seconds 3
										$nodes = Retrieve_Nodes -StoreNumber $StoreNumber
										if ($nodes)
										{
											$machineName = $nodes.LaneNumToMachineName[$LaneNumber]
											if ($machineName)
											{
												$mailslotAddress = "\\$machineName\mailslot\SMSStart_${StoreNumber}${LaneNumber}"
												$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
												$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
												if ($result) { Write_Log "Restart command sent to Machine $machineName (Store $StoreNumber, Lane $LaneNumber) after deployment." "green" }
												else { Write_Log "Failed to send restart command to Machine $machineName (Store $StoreNumber, Lane $LaneNumber)." "red" }
											}
										}
									}
									else
									{
										Write_Log "Could not extract transaction number from Last recorded status in file $($file.Name)" "red"
									}
								}
							}
						}
					}
				}
				catch
				{
					Write_Log "Error processing file $($file.Name): $_" "red"
				}
			}
		}
		
		# ==================== If no matching files: Manual prompt for lane ====================
		if (-not $MatchedTransactions)
		{
			Write_Log "No files or no matching transactions found. Prompting for lane number." "yellow"
			$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
			if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
			{
				Write_Log "Lane selection cancelled or returned no selection." "yellow"
				Write_Log "`r`n==================== Close_Open_Transactions Function Completed ====================" "blue"
				return
			}
			
			foreach ($LaneNumber in $selection.Lanes)
			{
				$LaneNumber = if ($LaneNumber -is [pscustomobject] -and $LaneNumber.LaneNumber) { $LaneNumber.LaneNumber }
				else { $LaneNumber }
				$LaneNumber = $LaneNumber.PadLeft(3, '0')
				$laneKey = $LaneNumber
				$protocolType = if ($script:LaneProtocols.ContainsKey($laneKey)) { $script:LaneProtocols[$laneKey] }
				else { "File" }
				if ($protocolType -eq 'Offline')
				{
					Write_Log "Lane $LaneNumber is offline. Skipping." "yellow"
					continue
				}
				$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $LaneNumber
				$namedPipesConnStr = $laneInfo['NamedPipesConnStr']
				$tcpConnStr = $laneInfo['TcpConnStr']
				$tcpServer = $laneInfo['TcpServer']
				$dbName = $laneInfo['DBName']
				$closeSQL = "UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1067 <> 'CLOSE'"
				$protocolWorked = $false
				
				Import-Module $SqlModule -ErrorAction Stop
				$supportsConnStr = (Get-Command Invoke-Sqlcmd).Parameters.ContainsKey('ConnectionString')
				
				if ($protocolType -eq "Named Pipes")
				{
					try
					{
						if ($supportsConnStr)
						{
							Invoke-Sqlcmd -ConnectionString $namedPipesConnStr -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
						}
						else
						{
							Invoke-Sqlcmd -ServerInstance $tcpServer -Database $dbName -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
						}
						$protocolWorked = $true
						Write_Log "Closed all open transactions via SQL protocol (Named Pipes) on lane $LaneNumber." "green"
					}
					catch { Write_Log "Named Pipes failed for lane $LaneNumber." "yellow" }
				}
				elseif ($protocolType -eq "TCP")
				{
					try
					{
						if ($supportsConnStr)
						{
							Invoke-Sqlcmd -ConnectionString $tcpConnStr -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
						}
						else
						{
							Invoke-Sqlcmd -ServerInstance $tcpServer -Database $dbName -Query $closeSQL -QueryTimeout 30 -ErrorAction Stop
						}
						$protocolWorked = $true
						Write_Log "Closed all open transactions via SQL protocol (TCP) on lane $LaneNumber." "green"
					}
					catch { Write_Log "TCP failed for lane $LaneNumber." "yellow" }
				}
				
				if (-not $protocolWorked)
				{
					$LaneDirectory = "$OfficePath\XF${StoreNumber}${LaneNumber}"
					if (Test-Path $LaneDirectory)
					{
						$CloseTransactionFilePath = Join-Path -Path $LaneDirectory -ChildPath "Close_Transaction.sqi"
						Set-Content -Path $CloseTransactionFilePath -Value $CloseTransactionManual -Encoding ASCII
						Set-ItemProperty -Path $CloseTransactionFilePath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
						Write_Log "Deployed Close_Transaction.sqi to lane $LaneNumber (fallback)." "yellow"
					}
					else
					{
						Write_Log "Lane directory $LaneDirectory not found" "yellow"
					}
				}
				$logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - User deployed Close_Transaction to lane $LaneNumber"
				Add-Content -Path $LogFilePath -Value $logMessage
				# Clean up XE folder, except FATALs
				Get-ChildItem -Path $XEFolderPath -File | Where-Object { $_.Name -notlike "*FATAL*" } | Remove-Item -Force
				$nodes = Retrieve_Nodes -StoreNumber $StoreNumber
				if ($nodes)
				{
					$machineName = $nodes.LaneNumToMachineName[$LaneNumber]
					if ($machineName)
					{
						$mailslotAddress = "\\$machineName\mailslot\SMSStart_${StoreNumber}${LaneNumber}"
						$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
						$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
						if ($result) { Write_Log "Restart All Programs sent to Machine $machineName (Store $StoreNumber, Lane $LaneNumber) after user deployment." "green" }
						else { Write_Log "Failed to send restart command to Machine $machineName (Store $StoreNumber, Lane $LaneNumber)." "red" }
					}
				}
				Write_Log "Prompt deployment process completed." "yellow"
			}
		}
	}
	catch
	{
		Write_Log "An error occurred during monitoring: $_" "red"
	}
	Write_Log "No further matching files were found after processing." "yellow"
	Write_Log "`r`n==================== Close_Open_Transactions Function Completed ====================" "blue"
}

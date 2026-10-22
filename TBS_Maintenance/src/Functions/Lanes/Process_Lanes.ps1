# ===================================================================================================
#                              FUNCTION: Process_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Runs the per-lane SQL maintenance (LANESQL script) for the lanes of a store. Resolves
#   each lane's DB server/name from the cached lane info, honors the cached protocol
#   preference (TCP / Named Pipes / File) and records the processed lanes for the summary.
#
# Parameters:
#   - StoreNumber     [string]: store to process
#   - ProcessAllLanes [switch]: process every lane of the store
#
# Variables:
#   - $script:FunctionResults: LaneNumToMachineName, MachineNameToLaneNum, SqlModuleName (input)
#   - $script:ProcessedLanes: lanes processed so far (output)
#   - $OfficePath: Storeman Office root (XF lane folders)
# ===================================================================================================

function Process_Lanes
{
	param (
		[string]$StoreNumber,
		[switch]$ProcessAllLanes
	)

	Write_Log "`r`n==================== Starting Process_Lanes Function ====================`r`n" "blue"

	if (-not (Test-Path $OfficePath))
	{
		Write_Log "XF Base Path not found: $OfficePath" "yellow"
		Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
		return
	}

	# Make sure ProcessedLanes exists
	if (-not $script:ProcessedLanes) { $script:ProcessedLanes = @() }

	# ----------------------------------------
	# Import detected SQL module for Invoke-Sqlcmd usage
	# ----------------------------------------
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	if ($SqlModuleName -and $SqlModuleName -ne "None")
	{
		try
		{
			Import-Module $SqlModuleName -ErrorAction Stop
			$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
		}
		catch
		{
			Write_Log "Failed to import SQL module '$SqlModuleName' or find Invoke-Sqlcmd: $_" "red"
			return
		}
	}
	else
	{
		Write_Log "No valid SQL module available for SQL operations!" "red"
		return
	}

	# ----------------------------------------
	# Check for available Lane Machines
	# ----------------------------------------
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName -or $LaneNumToMachineName.Count -eq 0)
	{
		Write_Log "No lanes available. Please retrieve nodes first." "red"
		Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
		return
	}

	$MachineNameToLaneNum = $script:FunctionResults['MachineNameToLaneNum']
	if (-not $MachineNameToLaneNum) { $MachineNameToLaneNum = @{ } }

	# ----------------------------------------
	# Get user's lane selection
	# ----------------------------------------
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if ($selection -eq $null)
	{
		Write_Log "Lane processing canceled by user." "yellow"
		Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
		return
	}

	# Support both string and object selections for lanes
	if ($selection.Lanes -and $selection.Lanes.Count -gt 0)
	{
		if ($selection.Lanes[0] -is [PSCustomObject] -and $selection.Lanes[0].PSObject.Properties.Name -contains 'LaneNumber')
		{
			$Lanes = $selection.Lanes | ForEach-Object { $_.LaneNumber }
		}
		else
		{
			$Lanes = $selection.Lanes
		}
	}
	else
	{
		Write_Log "No lanes selected." "yellow"
		Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
		return
	}

	# ----------------------------------------
	# Parse Lane SQL script sections for processing
	# ----------------------------------------
	$sectionPattern = '(?s)/\*\s*(?<SectionName>[^*/]+?)\s*\*/\s*(?<SQLCommands>(?:.(?!/\*)|.)*?)(?=(/\*|$))'
	$matches = [regex]::Matches($script:LaneSQLScript, $sectionPattern)
	if ($matches.Count -eq 0)
	{
		Write_Log "No sections found in Lane SQL script." "red"
		Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
		return
	}

	$fixedSections = @(
		"Set a long timeout so the entire script runs",
		"Clear the long database timeout"
	)

	$allSectionNames = $matches |
	ForEach-Object { $_.Groups['SectionName'].Value.Trim() } |
	Where-Object { $fixedSections -notcontains $_ }

	$SectionsToSend = Show_Section_Selection_Form -SectionNames $allSectionNames
	if (-not $SectionsToSend -or $SectionsToSend.Count -eq 0)
	{
		Write_Log "No sections selected for lanes." "yellow"
		Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
		return
	}

	# ----------------------------------------
	# Pre-build the SQI script for fallback (same for all lanes)
	# ----------------------------------------
	$topBlock = "/* Set a long timeout so the entire script runs */`r`n@WIZRPL(DBASE_TIMEOUT=E);`r`n" +
	"--------------------------------------------------------------------------------`r`n"
	$bottomBlock = "--------------------------------------------------------------------------------`r`n" +
	"/* Clear the long database timeout */`r`n@WIZCLR(DBASE_TIMEOUT);"

	$middleBlock = ($matches | Where-Object {
			$SectionsToSend -contains $_.Groups['SectionName'].Value.Trim()
		}) | ForEach-Object {
		"/* $($_.Groups['SectionName'].Value.Trim()) */`r`n$($_.Groups['SQLCommands'].Value.Trim())"
	} | Out-String

	$finalScript = $topBlock + $middleBlock + $bottomBlock

	# ----------------------------------------
	# MULTIPLE lanes: Always file fallback
	# ----------------------------------------
	if ($Lanes.Count -gt 1)
	{
		Write_Log "Multiple lanes selected, using file-based fallback for all lanes." "yellow"

		foreach ($LaneNumber in $Lanes)
		{
			# Normalize lane number (accept "001" or "POS001" etc)
			$laneNumCandidate = $LaneNumber
			if ($MachineNameToLaneNum.ContainsKey($LaneNumber)) { $laneNumCandidate = $MachineNameToLaneNum[$LaneNumber] }
			elseif ($LaneNumber -match '(\d{1,3})$') { $laneNumCandidate = $Matches[1] }

			if ($laneNumCandidate -match '^\d+$') { $laneNum = $laneNumCandidate.PadLeft(3, '0') }
			else { $laneNum = $LaneNumber }

			$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $laneNum
			if (-not $laneInfo)
			{
				Write_Log "Could not get DB info for lane $laneNum. Skipping." "yellow"
				continue
			}

			# Convert PSCustomObject -> hashtable (important when loaded from JSON cache)
			if ($laneInfo -and ($laneInfo -isnot [System.Collections.IDictionary]))
			{
				$tmp = @{ }
				foreach ($p in $laneInfo.PSObject.Properties) { $tmp[$p.Name] = $p.Value }
				$laneInfo = $tmp
			}

			$machineName = $laneInfo['MachineName']
			if (-not $machineName) { $machineName = $LaneNumToMachineName[$laneNum] }

			$LaneLocalPath = Join-Path $OfficePath ("XF" + $StoreNumber + $laneNum)

			if (Test-Path $LaneLocalPath)
			{
				Write_Log "Writing Lane_Database_Maintenance.sqi to Lane $laneNum ($machineName)..." "blue"
				try
				{
					Set-Content -Path (Join-Path $LaneLocalPath "Lane_Database_Maintenance.sqi") -Value $finalScript -Encoding Ascii
					Set-ItemProperty -Path (Join-Path $LaneLocalPath "Lane_Database_Maintenance.sqi") -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
					Write_Log "Created and wrote to file at Lane #$laneNum ($machineName) successfully. (file copy)" "green"

					if (-not ($script:ProcessedLanes -contains $laneNum)) { $script:ProcessedLanes += $laneNum }
				}
				catch
				{
					Write_Log "Failed to write to [$machineName]: $_" "red"
				}
			}
			else
			{
				Write_Log "Lane #$laneNum not found at path: $LaneLocalPath" "yellow"
			}
		}
	}
	else
	{
		# SINGLE lane: Try protocol, fallback to file
		foreach ($LaneNumber in $Lanes)
		{
			# Normalize lane number (accept "001" or "POS001" etc)
			$laneNumCandidate = $LaneNumber
			if ($MachineNameToLaneNum.ContainsKey($LaneNumber)) { $laneNumCandidate = $MachineNameToLaneNum[$LaneNumber] }
			elseif ($LaneNumber -match '(\d{1,3})$') { $laneNumCandidate = $Matches[1] }

			if ($laneNumCandidate -match '^\d+$') { $laneNum = $laneNumCandidate.PadLeft(3, '0') }
			else { $laneNum = $LaneNumber }

			$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $laneNum
			if (-not $laneInfo)
			{
				Write_Log "Could not get DB info for lane $laneNum. Skipping." "yellow"
				continue
			}

			# Convert PSCustomObject -> hashtable (important when loaded from JSON cache)
			if ($laneInfo -and ($laneInfo -isnot [System.Collections.IDictionary]))
			{
				$tmp = @{ }
				foreach ($p in $laneInfo.PSObject.Properties) { $tmp[$p.Name] = $p.Value }
				$laneInfo = $tmp
			}

			$machineName = $laneInfo['MachineName']
			$namedPipesConnStr = $laneInfo['NamedPipesConnStr']
			$tcpConnStr = $laneInfo['TcpConnStr']

			$LaneLocalPath = Join-Path $OfficePath ("XF" + $StoreNumber + $laneNum)

			# Get protocol for this lane
			$laneKey = $laneNum.PadLeft(3, '0')
			$protocolType = $null
			if ($script:LaneProtocols) { $protocolType = $script:LaneProtocols[$laneKey] }

			$workingConnStr = $null
			if ($protocolType -eq "Named Pipes") { $workingConnStr = $namedPipesConnStr }
			elseif ($protocolType -eq "TCP") { $workingConnStr = $tcpConnStr }

			# Validate connection string (DO NOT replace it with "File")
			$workingConnStrIsValid =
			(-not [string]::IsNullOrWhiteSpace($workingConnStr)) -and
			($workingConnStr -match '^\s*Server\s*=') -and
			($workingConnStr -match 'Database\s*=')

			$displayProtocol = if ([string]::IsNullOrWhiteSpace($protocolType)) { "File" }
			else { $protocolType }
			$displayConnStr = if ($workingConnStrIsValid) { $workingConnStr }
			else { "<file>" }

			Write_Log "Lane $laneNum uses protocol: $displayProtocol" "gray"
			Write_Log "Lane $laneNum connection string: $displayConnStr" "gray"

			if ($protocolType -eq 'Offline')
			{
				Write_Log "Lane $laneNum ($machineName) is offline. Skipping." "yellow"
				continue
			}

			# If protocol not ready / invalid conn string, fallback to file
			if (-not $protocolType -or $protocolType -eq "File" -or -not $workingConnStrIsValid)
			{
				Write_Log "Protocol not ready/unavailable for $machineName. Skipping protocol and using file copy." "yellow"

				if (Test-Path $LaneLocalPath)
				{
					Write_Log "`r`nProcessing $machineName using file fallback..." "blue"
					Write_Log "Lane path found: $LaneLocalPath" "blue"
					Write_Log "Writing Lane_Database_Maintenance.sqi to Lane..." "blue"
					try
					{
						Set-Content -Path (Join-Path $LaneLocalPath "Lane_Database_Maintenance.sqi") -Value $finalScript -Encoding Ascii
						Set-ItemProperty -Path (Join-Path $LaneLocalPath "Lane_Database_Maintenance.sqi") -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
						Write_Log "Created and wrote to file at Lane #$laneNum ($machineName) successfully. (file copy)" "green"

						if (-not ($script:ProcessedLanes -contains $laneNum)) { $script:ProcessedLanes += $laneNum }
					}
					catch
					{
						Write_Log "Failed to write to [$machineName]: $_" "red"
					}
				}
				else
				{
					Write_Log "Lane #$laneNum not found at path: $LaneLocalPath" "yellow"
				}
				continue
			}

			# ----------------------------------------
			# Protocol execution: Try SQL via protocol
			# ----------------------------------------
			$protocolWorked = $false

			# These MUST be valid for ServerInstance fallback
			$server = $laneInfo['TcpServer']
			$database = $laneInfo['DBName']
			if ([string]::IsNullOrWhiteSpace($database)) { $database = "LANESQL" } # critical fix
			if ([string]::IsNullOrWhiteSpace($server) -and $workingConnStr -match 'Server\s*=\s*([^;]+)') { $server = $Matches[1].Trim() }

			$currentLogin = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

			try
			{
				$matchesFiltered = [regex]::Matches($script:LaneSQLFiltered, $sectionPattern)
				$sections = ($matchesFiltered | Where-Object { $SectionsToSend -contains $_.Groups['SectionName'].Value.Trim() })

				foreach ($match in $sections)
				{
					$sectionName = $match.Groups['SectionName'].Value.Trim()
					$sqlCommands = $match.Groups['SQLCommands'].Value.Trim()

					Write_Log "Executing section: '$sectionName' on $machineName" "blue"
					Write_Log "--------------------------------------------------------------------------------"
					Write_Log "$sqlCommands" "orange"
					Write_Log "--------------------------------------------------------------------------------"

					$querySucceeded = $false
					$retriedForMapping = $false

					for ($try = 1; $try -le 2; $try++)
					{
						try
						{
							& $InvokeSqlCmd -ConnectionString $workingConnStr -Query $sqlCommands -QueryTimeout 0 -ErrorAction Stop
							Write_Log "Section '$sectionName' executed successfully on $machineName using ($protocolType)." "green"
							$protocolWorked = $true
							$querySucceeded = $true
							break
						}
						catch
						{
							$errorMsg = $_.Exception.Message

							if ($errorMsg -match 'Login failed for user' -and -not $retriedForMapping)
							{
								Write_Log "Login failed for $currentLogin on $machineName. Attempting to map user and retry..." "yellow"
								$checkUserQuery = "SELECT COUNT(*) AS UserExists FROM sys.database_principals WHERE name = '$currentLogin'"

								try
								{
									$userExists = (& $InvokeSqlCmd -ServerInstance $server -Database $database -Query $checkUserQuery -ErrorAction Stop).UserExists
								}
								catch { $userExists = 0 }

								if ($userExists -eq 0)
								{
									$createUserSql = @"
USE [$database];
CREATE USER [$currentLogin] FOR LOGIN [$currentLogin];
ALTER ROLE db_owner ADD MEMBER [$currentLogin];
"@
									try
									{
										& $InvokeSqlCmd -ServerInstance $server -Database $database -Query $createUserSql -ErrorAction Stop
										Write_Log "Mapped and granted db_owner to $currentLogin in [$database]." "blue"
									}
									catch
									{
										Write_Log "Failed to map $currentLogin in [$database]: $_" "yellow"
									}
								}
								else
								{
									try { & $InvokeSqlCmd -ServerInstance $server -Database $database -Query "ALTER ROLE db_owner ADD MEMBER [$currentLogin];" -ErrorAction SilentlyContinue }
									catch { }
									Write_Log "$currentLogin already mapped in [$database]." "gray"
								}

								$retriedForMapping = $true
								continue
							}
							elseif ($errorMsg -match 'Login failed for user')
							{
								Write_Log "Login failed for $currentLogin on $machineName. ServerInstance fallback will NOT be attempted (would repeat login failure)..." "yellow"
								$protocolWorked = $false
								break
							}
							else
							{
								Write_Log "ConnectionString method failed for section '$sectionName' on [$machineName]: $_. Trying ServerInstance fallback..." "yellow"
								try
								{
									& $InvokeSqlCmd -ServerInstance $server -Database $database -Query $sqlCommands -QueryTimeout 0 -ErrorAction Stop
									Write_Log "Section '$sectionName' executed successfully on $machineName using ($protocolType, ServerInstance fallback)." "green"
									$protocolWorked = $true
									$querySucceeded = $true
								}
								catch
								{
									Write_Log "ServerInstance method also failed for section '$sectionName' on [$machineName]: $_" "red"
									$protocolWorked = $false
								}
								break
							}
						}
					}

					if (-not $querySucceeded) { break }
				}
			}
			catch
			{
				Write_Log "Failed to execute a section on $machineName via protocol: $_. Falling back to file." "yellow"
				$protocolWorked = $false
			}

			if ($protocolWorked)
			{
				if (-not ($script:ProcessedLanes -contains $laneNum)) { $script:ProcessedLanes += $laneNum }
				continue
			}

			# Fallback: Classic file-based method
			if (Test-Path $LaneLocalPath)
			{
				Write_Log "`r`nProcessing $machineName using file copy..." "blue"
				Write_Log "Lane path found: $LaneLocalPath" "blue"
				Write_Log "Writing Lane_Database_Maintenance.sqi to Lane..." "blue"
				try
				{
					Set-Content -Path (Join-Path $LaneLocalPath "Lane_Database_Maintenance.sqi") -Value $finalScript -Encoding Ascii
					Set-ItemProperty -Path (Join-Path $LaneLocalPath "Lane_Database_Maintenance.sqi") -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
					Write_Log "Created and wrote to file at Lane #$laneNum ($machineName) successfully. (file copy)" "green"

					if (-not ($script:ProcessedLanes -contains $laneNum)) { $script:ProcessedLanes += $laneNum }
				}
				catch
				{
					Write_Log "Failed to write to [$machineName]: $_" "red"
				}
			}
			else
			{
				Write_Log "Lane #$laneNum not found at path: $LaneLocalPath" "yellow"
			}
		}
	}

	Write_Log "`r`nTotal Lanes processed: $($script:ProcessedLanes.Count)" "green"
	if ($script:ProcessedLanes.Count -gt 0)
	{
		Write_Log "Processed Lanes: $($script:ProcessedLanes -join ', ')" "green"
	}
	Write_Log "`r`n==================== Process_Lanes Function Completed ====================" "blue"
}

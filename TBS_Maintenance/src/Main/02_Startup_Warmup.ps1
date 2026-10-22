$form.Add_Shown({
		if ($script:StartupWarmupStarted) { return }
		$script:StartupWarmupStarted = $true
		$script:StartupWarmupIndex = 0
		$script:StartupWarmupSteps = @()
		$script:StartupScaleCredStarted = $false

		# Startup milestone (file log): the moment the window first appears
		try { Write_Log ("Startup: window shown {0:N1}s after script start" -f ((Get-Date) - $script:StartupT0).TotalSeconds) 'green' -FileOnly } catch { }

		try
		{
			if ($script:StartupScaleCredTimer)
			{
				$script:StartupScaleCredTimer.Stop()
				$script:StartupScaleCredTimer.Dispose()
				$script:StartupScaleCredTimer = $null
			}
		}
		catch { }

		$script:StartupScaleCredTimer = New-Object System.Windows.Forms.Timer
		$script:StartupScaleCredTimer.Interval = 250
		$script:StartupScaleCredTimer.Add_Tick({
				try
				{
					if ($script:StartupScaleCredTimer)
					{
						$script:StartupScaleCredTimer.Stop()
						$script:StartupScaleCredTimer.Dispose()
						$script:StartupScaleCredTimer = $null
					}
				}
				catch { }

				if ($script:StartupScaleCredStarted) { return }

				try
				{
					$windowsScales = $script:FunctionResults['WindowsScales']
					if ($windowsScales -and $windowsScales.Count -gt 0)
					{
						Add_Scale_Credentials -ScaleCodeToIPInfo $windowsScales
					}
					$script:StartupScaleCredStarted = $true
				}
				catch { }
			})
		try { $script:StartupScaleCredTimer.Start() }
		catch { }

		# v2.6.9: these steps run as BACKGROUND JOBS (Mode = 'Job'). The tick
		# handler below starts each one and polls it, so no step can freeze the
		# UI thread. Every job refreshes the function's FILE cache; OnComplete
		# then re-runs the same call WITHOUT -RefreshCache on the UI thread,
		# which reads the warm cache, hydrates FunctionResults and updates the
		# GUI labels (a few ms of local work). 'Installing SMS functions' writes
		# no shared cache and returns nothing, so it just runs in the job.
		# EXCEPTION: 'Generating SQL scripts' is INLINE (Action), not a job -
		# it is pure CPU (5 here-strings + 2 regex passes, no I/O, ~2 ms) so it
		# is safe on the UI thread, and running it inline sets the $script: SQL
		# variables directly (no job-scope handoff / ReturnVarNames wait).
		# NOTE: CallArgs on 'Job' steps is a SCRIPTBLOCK evaluated at start time,
		# so every job receives the current values of the variables it depends on.
		$script:StartupWarmupSteps += [pscustomobject]@{
			Label        = 'Refreshing store info...'
			Mode         = 'Job'
			FunctionName = 'Get_Store_And_Database_Info'
			CallArgs     = { @{ WinIniPath = $WinIniPath; SmsStartIniPath = $SmsStartIniPath; StartupIniPath = $StartupIniPath; SystemIniPath = $SystemIniPath; RefreshCache = $true } }
			SeedVars     = @{ SetupFiles = $script:SetupFiles; StoremanOfficeDbsFallbacks = $script:StoremanOfficeDbsFallbacks }
			OnComplete   = {
				try
				{
					Get_Store_And_Database_Info -WinIniPath $WinIniPath -SmsStartIniPath $SmsStartIniPath -StartupIniPath $StartupIniPath -SystemIniPath $SystemIniPath
					$script:StoreNumber = $script:FunctionResults['StoreNumber']
					$script:StoreName = $script:FunctionResults['StoreName']
					$script:SqlModuleName = $script:FunctionResults['SqlModuleName']
				}
				catch { }
			}
		}
		$script:StartupWarmupSteps += [pscustomobject]@{
			Label        = 'Refreshing node discovery...'
			Mode         = 'Job'
			FunctionName = 'Retrieve_Nodes'
			CallArgs     = { @{ StoreNumber = $script:StoreNumber; RefreshCache = $true } }
			SeedVars     = @{ SetupFiles = $script:SetupFiles; BasePath = $script:BasePath }
			DependentBodies = @{ Normalize_Sql_ConnectionString = ${function:Normalize_Sql_ConnectionString}.ToString() }
			OnComplete   = {
				try
				{
					Retrieve_Nodes -StoreNumber $script:StoreNumber | Out-Null
					$script:Nodes = $script:FunctionResults['Nodes']
					$script:LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
					Start_Lane_Protocol_Jobs -LaneNumToMachineName $script:LaneNumToMachineName -SqlModuleName $script:SqlModuleName
				}
				catch { }
			}
		}
		$script:StartupWarmupSteps += @(
			[pscustomobject]@{
				Label        = 'Loading lane DB cache...'
				Mode         = 'Job'
				FunctionName = 'Get_All_Lanes_Database_Info'
				CallArgs     = @{ RefreshCache = $true }
				SeedVars     = @{ SetupFiles = $script:SetupFiles }
				OnComplete   = { try { Get_All_Lanes_Database_Info | Out-Null } catch { } }
			},
			[pscustomobject]@{
				Label        = 'Loading table aliases...'
				Mode         = 'Job'
				FunctionName = 'Get_Table_Aliases'
				CallArgs     = @{ RefreshCache = $true }
				SeedVars     = @{ SetupFiles = $script:SetupFiles }
				OnComplete   = { try { $script:AliasToTable = Get_Table_Aliases } catch { } }
			},
			[pscustomobject]@{
				Label  = 'Generating SQL scripts...'
				Action = { try { Generate_SQL_Scripts -StoreNumber $script:StoreNumber | Out-Null } catch { } }
			},
			[pscustomobject]@{
				Label        = 'Installing SMS functions...'
				Mode         = 'Job'
				FunctionName = 'Install_FUNCTIONS_Into_SMS'
				CallArgs     = { @{ StoreNumber = $script:StoreNumber; OfficePath = $OfficePath; Silent = $true } }
				SeedVars     = @{ BasePath = $script:BasePath }
			},
			[pscustomobject]@{
				Label  = 'Starting test item job...'
				Action = {
					try
					{
						Get-Job | Where-Object { $_.Name -eq 'InsertTestItemJob' } |
						Where-Object { $_.State -in @('Completed', 'Failed', 'Stopped') } |
						Remove-Job -Force -ErrorAction SilentlyContinue

						# (-Name and -State cannot be combined on Get-Job: that threw before v2.7.0,
						#  so the test item job never started - filter the state instead)
						$runningInsertJob = Get-Job | Where-Object { $_.Name -eq 'InsertTestItemJob' } | Where-Object { $_.State -eq 'Running' } | Select-Object -First 1
						if (-not $runningInsertJob)
						{
							$fnBody = ${function:Insert_Test_Item}.ToString()
							$cs = $script:FunctionResults['ConnectionString']
							$dbs = $script:FunctionResults['DBSERVER']
							$dbn = $script:FunctionResults['DBNAME']
							$sqlModule = $script:FunctionResults['SqlModuleName']

							Start-Job -Name 'InsertTestItemJob' -ArgumentList $fnBody, $cs, $dbs, $dbn, $sqlModule -ScriptBlock {
								param (
									[string]$InsertTestItemBody,
									[string]$ConnectionString,
									[string]$DbServer,
									[string]$DbName,
									[string]$SqlModule
								)

								$script:FunctionResults = @{
									ConnectionString = $ConnectionString
									DBSERVER		 = $DbServer
									DBNAME		     = $DbName
									SqlModuleName    = $SqlModule
								}

								Set-Item -Path Function:\Insert_Test_Item -Value ([scriptblock]::Create($InsertTestItemBody))
								Insert_Test_Item -ConnectionString $ConnectionString -Silent | Out-Null
							} | Out-Null
						}
					}
					catch { }
				}
			},
			[pscustomobject]@{
				Label  = 'Adding scale credentials...'
				Action = {
					try
					{
						if (-not $script:StartupScaleCredStarted)
						{
							Add_Scale_Credentials -ScaleCodeToIPInfo $script:FunctionResults['WindowsScales']
							$script:StartupScaleCredStarted = $true
						}
					}
					catch { }
				}
			},
			[pscustomobject]@{
				Label  = 'Cleaning XE folder...'
				Action = { try { $script:ClearXEJob = Clear_XE_Folder }
					catch { } }
			},
			[pscustomobject]@{
				Label  = 'Preparing VNC exports...'
				Action = {
					if ($script:AutoExportVncOnLaunch)
					{
						try
						{
							Export_VNC_Files_For_All_Nodes `
														   -LaneNumToMachineName $script:FunctionResults['LaneNumToMachineName'] `
														   -ScaleCodeToIPInfo $script:FunctionResults['ScaleCodeToIPInfo'] `
														   -BackofficeNumToMachineName $script:FunctionResults['BackofficeNumToMachineName'] `
														   -AllVNCPasswords $script:FunctionResults['AllVNCPasswords'] `
														   -RunInBackground `
														   -SilentBackground
						}
						catch { }
					}
				}
			},
			[pscustomobject]@{
				Label  = 'Preparing hardware exports...'
				Action = {
					if ($script:AutoExportRemoteMachineInfoOnLaunch)
					{
						try
						{
							Get_Remote_Machine_Info `
													-RunInBackground `
													-SilentBackground `
													-UseAllAvailableNodes | Out-Null
						}
						catch { }
					}
				}
			},
			[pscustomobject]@{
				Label  = 'Checking quarterly maintenance...'
				Action = {
					if ($script:AutoQuarterlySystemHealthMaintenanceOnLaunch)
					{
						try
						{
							Start_Quarterly_System_Health_Maintenance `
																	  -IntervalMonths $script:SystemHealthMaintenanceIntervalMonths `
																	  -RetentionCount $script:SystemHealthMaintenanceRetentionCount `
																	  -SummaryRetentionLines $script:SystemHealthMaintenanceSummaryRetentionLines | Out-Null
						}
						catch { }
					}
				}
			}
		)
		$script:StartupWarmupTimer = New-Object System.Windows.Forms.Timer
		$script:StartupWarmupTimer.Interval = 450
		$script:StartupWarmupTimer.Add_Tick({
				if (-not $script:StartupWarmupSteps -or $script:StartupWarmupIndex -ge $script:StartupWarmupSteps.Count)
				{
					$script:StartupWarmupStatus = $null

					try
					{
						if ($script:StartupWarmupTimer)
						{
							$script:StartupWarmupTimer.Stop()
							$script:StartupWarmupTimer.Dispose()
							$script:StartupWarmupTimer = $null
						}
					}
					catch { }

					$script:StartupWarmupSteps = @()
					$script:StartupWarmupIndex = 0
					return
				}

				$step = $script:StartupWarmupSteps[$script:StartupWarmupIndex]

				try { $script:StartupWarmupStatus = if ($step -and $step.Label) { [string]$step.Label }
					else { $null } }
				catch { $script:StartupWarmupStatus = $null }

				# v2.6.9: 'Job' steps run in a background job - this tick only
				# starts or polls them, so the UI thread never waits on them.
				if ($step -and $step.Mode -eq 'Job')
				{
				if (-not $step.Job)
				{
						# CallArgs is a scriptblock that is evaluated at START time,
						# so the job always receives the current values (e.g. a
						# StoreNumber that step 1 just refreshed).
						$callArgs = @{}
						try
						{
							if ($step.CallArgs -is [scriptblock]) { $callArgs = & $step.CallArgs }
							elseif ($step.CallArgs) { $callArgs = $step.CallArgs }
						}
						catch { }
						$newJob = Start_Background_Function_Job -FunctionName $step.FunctionName -Label $step.Label -CallArgs $callArgs -SeedVars $step.SeedVars -ReturnVarNames $step.ReturnVarNames -DependentBodies $step.DependentBodies -OnComplete $step.OnComplete
						if ($newJob)
						{
							$step | Add-Member -NotePropertyName 'Job' -NotePropertyValue $newJob -Force
						}
						else
						{
							Write_Log ("Startup: warmup step '{0}' could not be started as a job" -f $step.Label) 'red'
							$script:StartupWarmupIndex++
						}
					}
					elseif ($script:BgfJobs -and $script:BgfJobs.ContainsKey($step.Job.Id))
					{
						# Job still registered: still running. The shared reaper
						# will merge the result and fire OnComplete for us.
					}
					else
					{
						# Job completed (reaper handled merge + OnComplete) - advance.
						$script:StartupWarmupIndex++
					}
					return
				}

				# Sync steps run inline - each one takes a few hundred ms at most.
				try
				{
					if ($step -and $step.Action)
					{
						$__warmupSw = [System.Diagnostics.Stopwatch]::StartNew()
						& $step.Action
						$__warmupSw.Stop()
						Write_Log ("Startup: warmup step '{0}' finished in {1:N1}s" -f $step.Label, $__warmupSw.Elapsed.TotalSeconds) 'blue' -FileOnly
					}
				}
				catch { }
				$script:StartupWarmupIndex++
			})
		$script:StartupWarmupTimer.Start()
	})

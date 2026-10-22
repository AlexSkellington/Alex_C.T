# ===================================================================================================
#                              FUNCTION: Start_Background_Function_Job
# ---------------------------------------------------------------------------------------------------
# Description:
#   The one background-job builder for this script (v2.6.9). Runs a script function in a
#   background job without blocking the UI thread. The job receives the function code,
#   Write_Log, any dependent functions, a copy of the FunctionResults bag and seed variables,
#   and its results are merged back into the live script scope when it finishes.
#
# Parameters:
#   - FunctionName     [string]: function to run in the job (required)
#   - Label            [string]: shown in the status bar / logged while the job runs
#   - CallArgs         [hashtable]: named arguments splatted onto the function call
#   - SeedVars         [hashtable]: script variables the function reads (name -> value)
#   - ReturnVarNames   [string[]]: script variables the function writes back (names only)
#   - DependentBodies  [hashtable]: extra functions the target calls (name -> body text)
#   - OnComplete       [scriptblock]: run on the UI thread when the job completes
#   - NotifyOnComplete [switch]: box line when the job ends without a result
#
# Variables:
#   - $script:BgfJobs: registry of running jobs for the shared reaper timer
#
# Returns:
#   The Start-Job job object, or $null if the function was not found.
# ===================================================================================================

function Start_Background_Function_Job
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$FunctionName,
		# shown in the status bar / logged while the job runs
		[string]$Label = '',
		# named arguments splatted onto the function call, e.g. @{ RefreshCache = $true }
		[hashtable]$CallArgs = @{},
		# extra script variables the function reads (name -> value), e.g. @{ SetupFiles = $script:SetupFiles }
		[hashtable]$SeedVars = @{},
		# script variables the function WRITES that must be handed back to the
		# main session (names only), e.g. @('LaneSQLScript', 'ServerSQLScript')
		[string[]]$ReturnVarNames = @(),
		# extra functions the target calls internally (name -> body text).
		# Write_Log is always passed, so the job can log like the main session.
		[hashtable]$DependentBodies = @{},
		# scriptblock run on the UI thread when the job completes; it receives
		# the result object (possibly $null). Used by warmup steps that must
		# hydrate state from the warm cache after the job's network work.
		[scriptblock]$OnComplete = $null,
		# mark the job as one whose abnormal finish ("ended without a result")
		# is worth a line in the log BOX. Normal success lines always go to the
		# file log only - the status bar covers progress, the box sees failures.
		[switch]$NotifyOnComplete
	)
	
	if (-not $CallArgs) { $CallArgs = @{} }
	if (-not $SeedVars) { $SeedVars = @{} }
	if (-not $ReturnVarNames) { $ReturnVarNames = @() }
	if (-not $DependentBodies) { $DependentBodies = @{} }
	if (-not $script:BgfJobs) { $script:BgfJobs = @{} }
	
	# NOTE: use Get-Command, NOT ${function:$FunctionName} - the variable-drive
	# syntax with a dynamic name does not resolve from inside a function body
	# under Windows PowerShell 5.1 (only literal names work there).
	$fnCmd = $null
	try { $fnCmd = Get-Command -Name $FunctionName -CommandType Function -ErrorAction SilentlyContinue } catch { }
	if (-not $fnCmd)
	{
		# NOTE: Out-Null keeps the return value clean even if a Write_Log
		# variant ever writes to the pipeline.
		try { Write_Log ("Start_Background_Function_Job: function '{0}' not found, cannot start job" -f $FunctionName) 'red' | Out-Null } catch { }
		return $null
	}
	$fnBody = $fnCmd.ScriptBlock.ToString()
	
	# Copy the results bag - the job must never mutate the live one
	$seedResults = @{}
	if ($script:FunctionResults)
	{
		foreach ($k in @($script:FunctionResults.Keys)) { $seedResults[$k] = $script:FunctionResults[$k] }
	}
	
	# Standard context every job needs so its Write_Log can find the file-log
	# folder and honour test mode. Callers do not have to remember this.
	$autoSeed = @{
		ScriptsLog        = $script:ScriptsLog
		IsTestEnvironment = $script:IsTestEnvironment
	}
	
	$job = Start-Job -Name ('bgf_' + $FunctionName) -ScriptBlock {
		param (
			[string]$FnName,
			[string]$FnBody,
			[hashtable]$Deps,
			[hashtable]$SeedResults,
			[hashtable]$SeedVars,
			[hashtable]$CallArgs,
			[string[]]$ReturnVarNames,
			[string]$WriteLogBody,
			[hashtable]$AutoSeed
		)
		
		Add-Type -AssemblyName System.Windows.Forms 2>$null
		Add-Type -AssemblyName System.Drawing 2>$null
		
		Set-Item -Path function:Write_Log -Value ([scriptblock]::Create($WriteLogBody))
		foreach ($depName in @($Deps.Keys))
		{
			Set-Item -Path ('function:' + $depName) -Value ([scriptblock]::Create($Deps[$depName]))
		}
		Set-Item -Path ('function:' + $FnName) -Value ([scriptblock]::Create($FnBody))
		
		$script:FunctionResults = @{ }
		foreach ($k in @($SeedResults.Keys)) { $script:FunctionResults[$k] = $SeedResults[$k] }
		foreach ($k in @($SeedVars.Keys))
		{
			Set-Variable -Scope Script -Name $k -Value $SeedVars[$k]
		}
		foreach ($k in @($AutoSeed.Keys))
		{
			if (-not (Get-Variable -Scope Script -Name $k -ValueOnly -ErrorAction SilentlyContinue))
			{
				Set-Variable -Scope Script -Name $k -Value $AutoSeed[$k]
			}
		}
		
		try
		{
			& (Get-Item -Path ('function:' + $FnName)) @CallArgs
			$extraVars = @{}
			foreach ($vn in $ReturnVarNames)
			{
				$vv = Get-Variable -Scope Script -Name $vn -ValueOnly -ErrorAction SilentlyContinue
				if ($null -ne $vv) { $extraVars[$vn] = $vv }
			}
			[pscustomobject]@{
				Success         = $true
				FunctionResults = $script:FunctionResults
				ExtraVars       = $extraVars
				Error           = $null
			}
		}
		catch
		{
			[pscustomobject]@{
				Success         = $false
				FunctionResults = $script:FunctionResults
				ExtraVars       = @{}
				Error           = $_.Exception.Message
			}
		}
	} -ArgumentList @(
		$FunctionName,
		$fnBody,
		$DependentBodies,
		$seedResults,
		$SeedVars,
		$CallArgs,
		$ReturnVarNames,
		${function:Write_Log}.ToString(),
		$autoSeed
	)
	
	if (-not $job)
	{
		Write_Log ("Start_Background_Function_Job: failed to start job for '{0}'" -f $FunctionName) 'red'
		return $null
	}
	
	$script:BgfJobs[$job.Id] = @{
		Job              = $job
		Label            = $Label
		Started          = Get-Date
		OnComplete       = $OnComplete
		NotifyOnComplete = [bool]$NotifyOnComplete
	}
	
	# v2.6.9: ONE shared reaper polls every background FUNCTION job on the
	# UI thread (1s tick), merges finished results and fires the job's
	# OnComplete callback (warmup hydration) or a one-line completion notice.
	# In hosts without a message pump (e.g. the selftest's child process) it
	# simply never ticks, and callers may poll Get_Background_Function_Job_Result
	# directly instead.
	# NOTE: the other timers in this script (ScaleCredReaper, VncExportReaper,
	# RemoteMachineInfoReaper, StartupScaleCredTimer, svcTimer) each manage an
	# INLINE job pool with local-variable closures and specialized collect
	# logic; they deliberately do not go through this engine. Migrating them
	# would mean restructuring those functions into the worker pattern -
	# considered and deferred (no performance gain, real regression risk).
	try
	{
		if (-not $script:BgfJobReaper)
		{
			$script:BgfJobReaper = New-Object System.Windows.Forms.Timer
			$script:BgfJobReaper.Interval = 1000
			$script:BgfJobReaper.Add_Tick({
					try
					{
						if (-not $script:BgfJobs -or $script:BgfJobs.Count -eq 0) { return }
						foreach ($key in @($script:BgfJobs.Keys))
						{
							$entry = $script:BgfJobs[$key]
							if (-not $entry -or -not $entry.Job) { continue }
							if ($entry.Job.State -in @('Running', 'NotStarted')) { continue }
							
							$result = $null
							try { $result = Get_Background_Function_Job_Result -Job $entry.Job } catch { }
							
							$secs = 0
							try { $secs = ((Get-Date) - $entry.Started).TotalSeconds } catch { }
							if ($result)
							{
								if ($result.Success)
								{
									# v2.6.9: success notices go to the FILE LOG ONLY (-FileOnly).
									# The status bar already shows what the app is doing in the
									# background, so the box stays quiet on success and only
									# sees failures (user preference).
									Write_Log ("Background job '{0}' finished in {1:N1}s" -f $entry.Label, $secs) 'blue' -FileOnly
								}
								else
								{
									Write_Log ("Background job '{0}' failed: {1}" -f $entry.Label, $result.Error) 'red'
								}
							}
							elseif ($entry.NotifyOnComplete)
							{
								Write_Log ("Background job '{0}' ended without a result" -f $entry.Label) 'yellow'
							}
							
							if ($entry.OnComplete)
							{
								try { & $entry.OnComplete $result } catch { }
							}
						}
					}
					catch { }
			})
			$script:BgfJobReaper.Start()
		}
	}
	catch { }
	return $job
}

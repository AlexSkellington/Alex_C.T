# ===================================================================================================
#                              FUNCTION: Get_Background_Function_Job_Result
# ---------------------------------------------------------------------------------------------------
# Description:
#   Polls a job created by Start_Background_Function_Job. Returns $null while the job is still
#   running. When it has finished, receives the result (before removing the job), merges the
#   job's FunctionResults and ExtraVars back into the live script scope and returns the
#   result object. Safe to call repeatedly.
#
# Parameters:
#   - Job  [Job]: the job object returned by Start_Background_Function_Job
#
# Variables:
#   - $script:BgfJobs: the job's registry entry is removed on completion
#
# Returns:
#   The job result object (Success/FunctionResults/ExtraVars/Error), or $null.
# ===================================================================================================

function Get_Background_Function_Job_Result
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.Management.Automation.Job]$Job
	)
	
	if (-not $Job) { return $null }
	if ($Job.State -in @('Running', 'NotStarted')) { return $null }
	
	$result = $null
	try { $result = @(Receive-Job -Job $Job -ErrorAction SilentlyContinue) | Select-Object -Last 1 } catch { }
	try { Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue | Out-Null } catch { }
	
	if ($result)
	{
		if ($result.FunctionResults)
		{
			if (-not $script:FunctionResults) { $script:FunctionResults = @{} }
			foreach ($k in @($result.FunctionResults.Keys))
			{
				$script:FunctionResults[$k] = $result.FunctionResults[$k]
			}
		}
		if ($result.ExtraVars)
		{
			foreach ($k in @($result.ExtraVars.Keys))
			{
				Set-Variable -Scope Script -Name $k -Value $result.ExtraVars[$k] -Force
			}
		}
	}
	
	if ($script:BgfJobs -and $script:BgfJobs.ContainsKey($Job.Id))
	{
		$script:BgfJobs.Remove($Job.Id)
	}
	return $result
}

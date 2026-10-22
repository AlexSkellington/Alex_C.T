# ===================================================================================================
#                              FUNCTION: Delete_Files
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deletes files/folders under a directory (recursively), optionally only those matching
#   name patterns and never those matching exclusion patterns. Runs synchronously (with log
#   lines) or as a background job so the caller is not held up.
#
# Parameters:
#   - Path           [string]: directory to delete from (required)
#   - SpecifiedFiles [string[]]: name patterns to delete (wildcards); all items when omitted
#   - Exclusions     [string[]]: name patterns to keep (wildcards)
#   - AsJob          [switch]: run the deletion as a background job
#
# Returns:
#   Number of items deleted (synchronous), or the Start-Job job (AsJob).
#
# Notes:
#   - Patterns match item NAMES at any depth (Get-ChildItem -Include/-Exclude): an excluded
#     folder name does not protect the files inside it.
# ===================================================================================================

function Delete_Files
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory, Position = 0, HelpMessage = "The directory path where files and folders will be deleted.")]
		[ValidateNotNullOrEmpty()]
		[string]$Path,
		[Parameter(HelpMessage = "Specific file/folder patterns to delete. Wildcards supported.")]
		[string[]]$SpecifiedFiles,
		[Parameter(HelpMessage = "Patterns to exclude from deletion. Wildcards supported.")]
		[string[]]$Exclusions,
		[Parameter(HelpMessage = "Run the deletion as a background job.")]
		[switch]$AsJob
	)

	# The deletion itself - shared by the job and the synchronous path.
	# Returns -1 when the path does not exist, else the number of items deleted.
	$deleteBlock = {
		param ($Path, $SpecifiedFiles, $Exclusions)

		$rp = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
		if (-not $rp) { return -1 }

		$gciParams = @{ Path = (Join-Path $rp.Path '*'); Recurse = $true; Force = $true; ErrorAction = 'SilentlyContinue' }
		if ($SpecifiedFiles -and $SpecifiedFiles.Count -gt 0) { $gciParams['Include'] = $SpecifiedFiles }
		if ($Exclusions -and $Exclusions.Count -gt 0) { $gciParams['Exclude'] = $Exclusions }

		$deletedCount = 0
		foreach ($it in @(Get-ChildItem @gciParams))
		{
			try
			{
				Remove-Item -LiteralPath $it.FullName -Force -Recurse -ErrorAction Stop
				$deletedCount++
			}
			catch { }
		}
		return $deletedCount
	}

	if ($AsJob)
	{
		$jobName = "DeleteFilesJob_{0}" -f (($Path -replace '[^A-Za-z0-9_-]', '_').Trim('_'))
		return Start-Job -Name $jobName -ScriptBlock $deleteBlock -ArgumentList $Path, $SpecifiedFiles, $Exclusions
	}

	try
	{
		$deletedCount = & $deleteBlock $Path $SpecifiedFiles $Exclusions
	}
	catch
	{
		Write_Log "An error occurred during deletion. $_" "red"
		return 0
	}

	if ($deletedCount -lt 0)
	{
		Write_Log "The specified path '$Path' does not exist." "red"
		return 0
	}
	if ($deletedCount -eq 0) { Write_Log "No items matched in '$Path'." "yellow" }
	Write_Log "Total items deleted: $deletedCount" "blue"
	return $deletedCount
}

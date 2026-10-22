# ===================================================================================================
#                              FUNCTION: Delete_Files
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deletes files/folders in a directory, supporting wildcards and exclusions. Can run
#   synchronously or as a background job so the main script is not interrupted.
#
# Parameters:
#   - Path           [string]: directory to delete from (required)
#   - SpecifiedFiles [string[]]: name/patterns to delete (wildcards supported)
#   - Exclusions     [string[]]: name/patterns to keep (wildcards supported)
#   - AsJob          [switch]: run the deletion as a background job
#
# Returns:
#   Number of items deleted.
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

	if ($AsJob)
	{
		$scriptBlock = {
			param ($Path,
				$SpecifiedFiles,
				$Exclusions)

			$deletedCount = 0
			$rp = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
			if (-not $rp) { return 0 }
			$targetPath = $rp.Path
			$gciBase = Join-Path $targetPath '*'

			try
			{
				# Build Get-ChildItem params safely
				$gciParams = @{ Path = $gciBase; Recurse = $true; Force = $true; ErrorAction = 'SilentlyContinue' }
				if ($SpecifiedFiles -and $SpecifiedFiles.Count -gt 0) { $gciParams['Include'] = $SpecifiedFiles }
				if ($Exclusions -and $Exclusions.Count -gt 0) { $gciParams['Exclude'] = $Exclusions }

				$items = Get-ChildItem @gciParams
				if ($items)
				{
					foreach ($it in $items)
					{
						try
						{
							Remove-Item -LiteralPath $it.FullName -Force -Recurse -ErrorAction Stop
							$deletedCount++
						}
						catch { }
					}
				}
				return $deletedCount
			}
			catch
			{
				return $deletedCount
			}
		}

		return Start-Job -Name ("DeleteFilesJob_{0}" -f ((($Path -replace '[^A-Za-z0-9_-]', '_').Trim('_')))) -ScriptBlock $scriptBlock -ArgumentList $Path, $SpecifiedFiles, $Exclusions
	}

	# -------- Synchronous path --------
	$deletedCount = 0
	$rp = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
	if (-not $rp)
	{
		Write_Log "The specified path '$Path' does not exist." "red"
		return 0
	}
	$targetPath = $rp.Path
	$gciBase = Join-Path $targetPath '*'

	try
	{
		$gciParams = @{ Path = $gciBase; Recurse = $true; Force = $true; ErrorAction = 'SilentlyContinue' }
		if ($SpecifiedFiles -and $SpecifiedFiles.Count -gt 0) { $gciParams['Include'] = $SpecifiedFiles }
		if ($Exclusions -and $Exclusions.Count -gt 0) { $gciParams['Exclude'] = $Exclusions }

		$items = Get-ChildItem @gciParams
		if ($items)
		{
			foreach ($it in $items)
			{
				try
				{
					Remove-Item -LiteralPath $it.FullName -Force -Recurse -ErrorAction Stop
					$deletedCount++
				}
				catch { }
			}
		}
		else
		{
			Write_Log "No items matched in '$targetPath'." "yellow"
		}

		Write_Log "Total items deleted: $deletedCount" "blue"
		return $deletedCount
	}
	catch
	{
		Write_Log "An error occurred during deletion. $_" "red"
		return $deletedCount
	}
}

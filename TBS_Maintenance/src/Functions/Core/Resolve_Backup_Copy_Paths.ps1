# ===================================================================================================
#                              FUNCTION: Resolve_Backup_Copy_Paths
# ---------------------------------------------------------------------------------------------------
# Description:
#   Checks the extra copy folders of a backup scheduler with Resolve_Backup_Folder (mapped drive ->
#   \\server\share, created, write-tested) and drops duplicates and the main backup folder itself.
#
# Parameters:
#   - Paths   [string[]]: the extra folders as entered
#   - Primary [string]: the main backup folder (left out of the copies)
#
# Returns:
#   [pscustomobject] Ok (bool), Paths (string[] resolved), Error (text listing the bad folders)
# ===================================================================================================

function Resolve_Backup_Copy_Paths
{
	param (
		[string[]]$Paths,
		[string]$Primary
	)

	$ok = New-Object System.Collections.Generic.List[string]
	$errors = New-Object System.Collections.Generic.List[string]
	foreach ($p in @($Paths))
	{
		if ([string]::IsNullOrWhiteSpace($p)) { continue }
		$chk = Resolve_Backup_Folder -Path $p
		if (-not $chk.Ok) { $errors.Add($chk.Error); continue }
		if ($Primary -and [string]::Equals($chk.Path.TrimEnd('\'), $Primary.TrimEnd('\'), [System.StringComparison]::OrdinalIgnoreCase)) { continue }
		if (@($ok | Where-Object { [string]::Equals($_, $chk.Path, [System.StringComparison]::OrdinalIgnoreCase) }).Count -gt 0) { continue }
		$ok.Add($chk.Path)
	}
	return [pscustomobject]@{ Ok = ($errors.Count -eq 0); Paths = [string[]]$ok.ToArray(); Error = ($errors -join "`r`n") }
}

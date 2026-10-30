# ===================================================================================================
#                             FUNCTION: Remove_MiniGhost_Temp_Copies
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deletes the copies of MiniGhost left in %TEMP% (MiniGhost.ps1 and the
#   MiniGhost_<version>_Elevated.ps1 written by the elevation bootstrap for irm | iex launches).
#
# Parameters:
#   - TempPath [string]: folder to clean (default: the user's %TEMP%)
#
# Returns:
#   Number of files deleted.
#
# Notes:
#   - Top level of the temp folder only. The old cleanup searched all of %TEMP% recursively for
#     MiniGhost.ps1 on every exit (slow on a full temp folder) and never matched the _Elevated copy.
# ===================================================================================================

function Remove_MiniGhost_Temp_Copies
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$TempPath = [System.IO.Path]::GetTempPath()
	)
	
	$deleted = 0
	if ([string]::IsNullOrWhiteSpace($TempPath) -or -not (Test-Path -LiteralPath $TempPath)) { return $deleted }
	
	foreach ($file in @(Get-ChildItem -LiteralPath $TempPath -File -Filter 'MiniGhost*.ps1' -ErrorAction SilentlyContinue))
	{
		if ($file.Name -notmatch '^(?i)MiniGhost(_[\d.]+_Elevated)?\.ps1$') { continue }
		try
		{
			Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
			$deleted++
		}
		catch { }
	}
	return $deleted
}

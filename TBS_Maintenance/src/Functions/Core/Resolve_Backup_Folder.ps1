# ===================================================================================================
#                              FUNCTION: Resolve_Backup_Folder
# ---------------------------------------------------------------------------------------------------
# Description:
#   Checks a backup folder typed or picked in a backup scheduler: expands %VARIABLES%, turns a
#   mapped network drive letter into its \\server\share path (scheduled tasks and the SQL Server
#   service do not see a user's mapped drives), creates the folder when missing and confirms it
#   can be written to.
#
# Parameters:
#   - Path [string]: the folder as entered
#
# Returns:
#   [pscustomobject] Ok (bool), Path (resolved folder), IsNetwork (bool), Error (text when not Ok)
# ===================================================================================================

function Resolve_Backup_Folder
{
	param (
		[string]$Path
	)

	$p = [Environment]::ExpandEnvironmentVariables(([string]$Path).Trim().Trim('"')).TrimEnd('\')
	if ([string]::IsNullOrWhiteSpace($p)) { return [pscustomobject]@{ Ok = $false; Path = $p; IsNetwork = $false; Error = 'No folder entered.' } }
	if ($p -match '^[A-Za-z]:$') { $p += '\' }

	# Mapped drive (Z:\...) -> \\server\share\...
	if ($p -match '^(?<drive>[A-Za-z]):')
	{
		try
		{
			$disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter ("DeviceID='{0}:'" -f $Matches['drive'].ToUpper()) -ErrorAction Stop
			if ($disk -and $disk.DriveType -eq 4 -and $disk.ProviderName) { $p = $disk.ProviderName.TrimEnd('\') + $p.Substring(2) }
		}
		catch { }
	}
	$isNetwork = $p.StartsWith('\\')
	if ($p -notmatch '^([A-Za-z]:\\|\\\\[^\\]+\\[^\\]+)') { return [pscustomobject]@{ Ok = $false; Path = $p; IsNetwork = $isNetwork; Error = "Use a full path like D:\Backups or \\server\share\Backups (got '$p')." } }

	try
	{
		if (-not (Test-Path -LiteralPath $p)) { $null = New-Item -ItemType Directory -Path $p -Force -ErrorAction Stop }
		$probe = Join-Path $p ('.tbs_write_test_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
		[System.IO.File]::WriteAllText($probe, 'ok')
		Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
	}
	catch
	{
		return [pscustomobject]@{ Ok = $false; Path = $p; IsNetwork = $isNetwork; Error = "Cannot create or write to '$p': $($_.Exception.Message)" }
	}
	return [pscustomobject]@{ Ok = $true; Path = $p; IsNetwork = $isNetwork; Error = $null }
}

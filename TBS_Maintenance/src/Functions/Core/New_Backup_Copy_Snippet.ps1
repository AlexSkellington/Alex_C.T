# ===================================================================================================
#                              FUNCTION: New_Backup_Copy_Snippet
# ---------------------------------------------------------------------------------------------------
# Description:
#   Returns the PowerShell lines a generated backup script runs after a successful backup to copy
#   the new file to each extra folder (into <folder>\<SubFolder>), keep only the newest Keep
#   backups there and log each copy. A failed copy is logged and never fails the backup.
#
# Parameters:
#   - Paths            [string[]]: extra folders (resolved); none = empty text
#   - SubFolder        [string]: machine subfolder inside each folder
#   - Filter           [string]: file pattern for the rotation (e.g. STORESQL_*.bak)
#   - Keep             [int]: backups to keep in each extra folder
#   - SourceExpression [string]: the generated script's expression for the new backup file
#   - LogExpression    [string]: the generated script's expression for its log file
#   - Indent           [string]: indentation of the generated lines
#
# Notes:
#   - Every path is written as a single-quoted literal (quotes doubled): a '$' or quote in a path
#     cannot break or change the generated script.
# ===================================================================================================

function New_Backup_Copy_Snippet
{
	param (
		[string[]]$Paths,
		[string]$SubFolder,
		[string]$Filter,
		[int]$Keep,
		[string]$SourceExpression,
		[string]$LogExpression,
		[string]$Indent = '    '
	)

	$list = @($Paths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
	if ($list.Count -eq 0) { return '' }
	$q = { param ([string]$s) "'" + $s.Replace("'", "''") + "'" }
	$lit = '@(' + (($list | ForEach-Object { & $q $_ }) -join ', ') + ')'
	$i = $Indent
	$lines = @(
		"# Extra copies (`"Also copy to`") - a failed copy is logged, never fails the backup"
		"foreach (`$copyRoot in $lit) {"
		"    try {"
		"        `$copyDir = Join-Path `$copyRoot $(& $q $SubFolder)"
		"        if (-not (Test-Path -LiteralPath `$copyDir)) { New-Item -Path `$copyDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }"
		"        Copy-Item -LiteralPath $SourceExpression -Destination `$copyDir -Force -ErrorAction Stop"
		"        @(Get-ChildItem -Path `$copyDir -Filter $(& $q $Filter) -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -Skip $Keep) |"
		"            ForEach-Object { Remove-Item -LiteralPath `$_.FullName -Force -ErrorAction SilentlyContinue }"
		"        `"[`$(Get-Date)] Copy saved to `$copyDir`" | Out-File -FilePath $LogExpression -Append -Encoding utf8"
		"    }"
		"    catch {"
		"        `"[`$(Get-Date)] Copy to `$copyRoot FAILED: `$(`$_.Exception.Message)`" | Out-File -FilePath $LogExpression -Append -Encoding utf8"
		"    }"
		"}"
	)
	return (($lines | ForEach-Object { $i + $_ }) -join "`r`n")
}

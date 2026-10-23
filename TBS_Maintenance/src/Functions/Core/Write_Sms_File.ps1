# ===================================================================================================
#                              FUNCTION: Write_Sms_File
# ---------------------------------------------------------------------------------------------------
# Description:
#   Writes a file SMS will pick up (.sqi / .sql / .sqm drops in Office or a lane's XF folder) the
#   way SMS expects it: Windows-1252 (ANSI), CRLF line endings, a final CRLF, no BOM, and plain
#   attributes - no archive bit (SMS treats an archived file as already processed). A read-only
#   or hidden file of the same name is unlocked first (the write used to fail on those).
#
# Parameters:
#   - Path    [string]: file to write (required)
#   - Content [string]: file text, any line endings (required, may be empty)
#
# Returns:
#   Nothing. Throws when the file cannot be written - callers log their own message.
# ===================================================================================================

function Write_Sms_File
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$Path,
		[Parameter(Mandatory = $true)]
		[AllowEmptyString()]
		[string]$Content
	)

	$text = $Content -replace "`r?`n", "`r`n"
	if ($text -and -not $text.EndsWith("`r`n")) { $text += "`r`n" }
	if ([System.IO.File]::Exists($Path)) { [System.IO.File]::SetAttributes($Path, [System.IO.FileAttributes]::Normal) }
	[System.IO.File]::WriteAllText($Path, $text, [System.Text.Encoding]::GetEncoding(1252))
	[System.IO.File]::SetAttributes($Path, [System.IO.FileAttributes]::Normal)
}

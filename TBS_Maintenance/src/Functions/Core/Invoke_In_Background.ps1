# ===================================================================================================
#                                FUNCTION: Invoke_In_Background
# ---------------------------------------------------------------------------------------------------
# Description:
#   Runs a self-contained script block in its own runspace and waits for it with Wait_UI, so long
#   work (big SQL statements, slow network calls) does not freeze the window. Returns the block's
#   output; an error in the block is thrown here like a normal error.
#
# Parameters:
#   - ScriptBlock  [scriptblock]: the work - it must not use this script's functions or variables
#                                (a runspace starts empty); pass what it needs as arguments
#   - ArgumentList [object[]]: arguments for the block's param()
#
# Notes:
#   - The block runs with ErrorActionPreference = Stop, so any failed statement ends it.
# ===================================================================================================

function Invoke_In_Background
{
	param (
		[Parameter(Mandatory = $true)]
		[scriptblock]$ScriptBlock,
		[object[]]$ArgumentList = @()
	)

	$ps = [PowerShell]::Create()
	try
	{
		$null = $ps.AddScript('$ErrorActionPreference = "Stop"')
		$null = $ps.AddStatement().AddScript($ScriptBlock.ToString())
		foreach ($a in $ArgumentList) { $null = $ps.AddArgument($a) }
		$handle = $ps.BeginInvoke()
		while (-not $handle.IsCompleted) { Wait_UI -Milliseconds 100 }
		$out = $ps.EndInvoke($handle)
		if ($ps.HadErrors -and $ps.Streams.Error.Count -gt 0) { throw $ps.Streams.Error[0].Exception }
		return $out
	}
	finally { $ps.Dispose() }
}

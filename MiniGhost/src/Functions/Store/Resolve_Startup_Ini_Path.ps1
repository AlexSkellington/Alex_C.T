# ===================================================================================================
#                           FUNCTION: Resolve_Startup_Ini_Path
# ---------------------------------------------------------------------------------------------------
# Description:
# Resolves the active Startup.ini path using an explicit value first, then existing script/global
# variables, and finally BasePath when available.
# ===================================================================================================

function Resolve_Startup_Ini_Path
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath
	)
	
	$resolvedPath = $StartupIniPath
	
	try
	{
		if ([string]::IsNullOrWhiteSpace($resolvedPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name StartupIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value))
				{
					$resolvedPath = [string]$v.Value
					break
				}
			}
		}
		
		if ([string]::IsNullOrWhiteSpace($resolvedPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name BasePath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value))
				{
					$candidatePath = Join-Path ([string]$v.Value) "Startup.ini"
					if (Test-Path $candidatePath)
					{
						$resolvedPath = $candidatePath
						break
					}
				}
			}
		}
		
		if (-not [string]::IsNullOrWhiteSpace($resolvedPath) -and (Test-Path $resolvedPath))
		{
			return $resolvedPath
		}
	}
	catch { }
	
	return $null
}

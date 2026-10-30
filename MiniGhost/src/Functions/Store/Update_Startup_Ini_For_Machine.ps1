# ===================================================================================================
#                          FUNCTION: Update_Startup_Ini_For_Machine
# ---------------------------------------------------------------------------------------------------
# Description:
# Applies machine-specific Startup.ini changes that should happen both after a rename and during a
# same-name sync. This keeps the DBSERVER host update in one place and preserves encoding/newlines.
# ===================================================================================================

function Update_Startup_Ini_For_Machine
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[psobject]$MachineContext,
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath
	)
	
	$resolvedStartupIniPath = Resolve_Startup_Ini_Path -StartupIniPath $StartupIniPath
	if ([string]::IsNullOrWhiteSpace($resolvedStartupIniPath))
	{
		return [PSCustomObject]@{
			Success       = $false
			Path          = $null
			DbServerValue = $null
			TerValue      = $null
			Details       = "Startup.ini path could not be resolved."
			ErrorMessage  = "startup.ini not found."
		}
	}
	
	try
	{
		$bytes = [System.IO.File]::ReadAllBytes($resolvedStartupIniPath)
		
		$encStartup = $null
		if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
		{ $encStartup = New-Object System.Text.UTF8Encoding($true) }
		elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
		{ $encStartup = [System.Text.Encoding]::Unicode }
		elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
		{ $encStartup = [System.Text.Encoding]::BigEndianUnicode }
		else
		{ $encStartup = [System.Text.Encoding]::Default }
		
		$textStartup = $encStartup.GetString($bytes)
		$nlStartup = "`r`n"
		if ($textStartup -notmatch "`r`n" -and $textStartup -match "`n") { $nlStartup = "`n" }
		
		$startupLines = $textStartup -split "\r?\n", -1
		
		$existingDbLine = $null
		foreach ($line in $startupLines)
		{
			if ($line -match '^\s*(?i:DBSERVER)\s*=')
			{
				$existingDbLine = $line
				break
			}
		}
		
		$existingDbServer = $null
		if ($existingDbLine -match '^\s*(?i:DBSERVER)\s*=\s*(.+?)\s*$')
		{
			$existingDbServer = $matches[1].Trim()
		}
		
		$instance = $null
		if ($existingDbServer -and ($existingDbServer -match '\\'))
		{
			$parts = $existingDbServer.Split('\')
			if ($parts.Count -ge 2 -and $parts[1]) { $instance = $parts[1] }
		}
		
		$dbServerValue = if ($instance)
		{
			"DBSERVER=$($MachineContext.MachineName)\$instance"
		}
		else
		{
			"DBSERVER=$($MachineContext.MachineName)"
		}
		
		$terValue = $null
		if (-not $MachineContext.IsServerMoniker -and $MachineContext.MachineNumber)
		{
			$terValue = "TER=$($MachineContext.MachineNumber)"
		}
		
		$updatedLines = @()
		$dbServerUpdated = $false
		$terUpdated = $false
		
		foreach ($line in $startupLines)
		{
			$updatedLine = $line
			
			if (-not $MachineContext.IsServerMoniker -and $terValue -and ($updatedLine -match '^\s*(?i:TER)\s*='))
			{
				$updatedLine = $terValue
				$terUpdated = $true
			}
			
			if ($updatedLine -match '^\s*(?i:DBSERVER)\s*=')
			{
				$updatedLine = $dbServerValue
				$dbServerUpdated = $true
			}
			
			$updatedLines += $updatedLine
		}
		
		if (-not $dbServerUpdated)
		{
			$updatedLines += $dbServerValue
			$dbServerUpdated = $true
		}
		
		[System.IO.File]::WriteAllText($resolvedStartupIniPath, ($updatedLines -join $nlStartup), $encStartup)
		
		$detailParts = @()
		$detailParts += "DBSERVER set to '$dbServerValue'."
		if ($MachineContext.IsServerMoniker)
		{
			$detailParts += "TER left unchanged due to SERVER moniker."
		}
		elseif ($terUpdated)
		{
			$detailParts += "TER set to '$terValue'."
		}
		else
		{
			$detailParts += "TER line was not present, so only DBSERVER was enforced here."
		}
		
		return [PSCustomObject]@{
			Success       = $true
			Path          = $resolvedStartupIniPath
			DbServerValue = $dbServerValue
			TerValue      = $terValue
			Details       = ($detailParts -join ' ')
			ErrorMessage  = $null
		}
	}
	catch
	{
		return [PSCustomObject]@{
			Success       = $false
			Path          = $resolvedStartupIniPath
			DbServerValue = $null
			TerValue      = $null
			Details       = "Startup.ini machine sync failed."
			ErrorMessage  = $_.Exception.Message
		}
	}
}

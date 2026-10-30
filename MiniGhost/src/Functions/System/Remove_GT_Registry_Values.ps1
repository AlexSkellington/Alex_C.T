# ===================================================================================================
#                                       FUNCTION: Remove_GT_Registry_Values
# ---------------------------------------------------------------------------------------------------
# Description:
#   Removes all registry values starting with 'GT' from specified registry paths.
# ===================================================================================================

function Remove_GT_Registry_Values
{
	# Define registry paths for 32-bit and 64-bit
	$regPath32 = "HKLM:\SOFTWARE\Store Management\Counters"
	$regPath64 = "HKLM:\SOFTWARE\Wow6432Node\Store Management\Counters"
	
	$is64bit = [System.Environment]::Is64BitOperatingSystem
	
	$totalDeletedCount = 0
	$success = $true
	$status = "Successful"
	$message = ""
	
	# Build list of candidate paths (prefer Wow6432Node on 64-bit, but include both if present)
	$pathsToCheck = @()
	
	if ($is64bit)
	{
		if (Test-Path $regPath64) { $pathsToCheck += $regPath64 }
		if (Test-Path $regPath32) { $pathsToCheck += $regPath32 }
	}
	else
	{
		if (Test-Path $regPath32) { $pathsToCheck += $regPath32 }
	}
	
	if (-not $pathsToCheck -or $pathsToCheck.Count -eq 0)
	{
		return @{
			Success	     = $false
			Status	     = "Failed"
			DeletedCount = 0
			Message	     = "No valid registry paths found for the current environment."
		}
	}
	
	foreach ($path in $pathsToCheck)
	{
		try
		{
			# Get value names from the key by inspecting item properties
			# (Get-ItemProperty returns a PSCustomObject with properties = value names + PS metadata)
			$itemProps = Get-ItemProperty -Path $path -ErrorAction Stop
			
			# Collect only real registry value names starting with GT (exclude PS metadata props)
			$gtNames = @()
			foreach ($p in $itemProps.PSObject.Properties)
			{
				if ($p.Name -eq 'PSPath') { continue }
				if ($p.Name -eq 'PSParentPath') { continue }
				if ($p.Name -eq 'PSChildName') { continue }
				if ($p.Name -eq 'PSDrive') { continue }
				if ($p.Name -eq 'PSProvider') { continue }
				
				if ($p.Name -like 'GT*')
				{
					$gtNames += $p.Name
				}
			}
			
			# Delete each GT* value
			foreach ($name in $gtNames)
			{
				try
				{
					Remove-ItemProperty -Path $path -Name $name -ErrorAction Stop
					$totalDeletedCount++
				}
				catch
				{
					# Keep going but mark overall as partial failure
					$success = $false
					$status = "Failed"
					if ($message) { $message += " | " }
					$message += "Failed to delete '$name' in '$path': $_"
				}
			}
		}
		catch
		{
			$success = $false
			$status = "Failed"
			if ($message) { $message += " | " }
			$message += "Error accessing registry path '$path': $_"
		}
	}
	
	return @{
		Success	     = $success
		Status	     = $status
		DeletedCount = $totalDeletedCount
		Message	     = $message
	}
}

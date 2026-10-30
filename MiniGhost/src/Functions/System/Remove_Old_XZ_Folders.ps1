# ===================================================================================================
# FUNCTION: Remove_Old_XZ_Folders
# ---------------------------------------------------------------------------------------------------
# Description:
# Removes old X* and Z* folders under \storeman\office that do NOT match the NEW store and allowed
# terminal numbers.
#
# Strict folder name pattern acted upon:
#   ^(X|Z)[A-Za-z](store 3-4 digits)(terminal 3 digits)$
# Examples: XF0231006, XW0242901, ZF1234006, ZW1234901
#
# Keeps ONLY folders that match:
#   - Store = new store (3-digit OR 4-digit representation)
#   - Terminal IN: current lane terminal (from MachineName), 900, 901
#
# Safeguard: refuses to run if extracted terminal is 901 unless -AllowBackoffice is passed.
# Office folder: -OfficePath, else the detected storeman Office folder ($OfficePath), else
# \\localhost\storeman\office, C:\storeman\office, D:\storeman\office.
# Automatic: no prompts.
# ===================================================================================================

function Remove_Old_XZ_Folders
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus,
		[Parameter(Mandatory = $false)]
		[switch]$AllowBackoffice,
		[Parameter(Mandatory = $false)]
		[string]$OfficePath
	)
	
	# -----------------------------
	# Init / status bucket
	# -----------------------------
	$deletedFolders = @()
	$failedToDeleteFolders = @()
	$anyFoldersSeen = $false
	$anyCandidatesMatched = $false
	
	if ($OperationStatus)
	{
		if (-not $OperationStatus.ContainsKey("OldXFoldersDeletion"))
		{
			$OperationStatus["OldXFoldersDeletion"] = [pscustomobject]@{
				Status  = ""
				Message = ""
				Details = ""
			}
		}
	}
	
	# -----------------------------
	# Validate store number
	# -----------------------------
	$storeNumberTrim = $StoreNumber
	if ($null -eq $storeNumberTrim) { $storeNumberTrim = "" }
	$storeNumberTrim = $storeNumberTrim.Trim()
	
	if ($storeNumberTrim -notmatch '^(?!0+$)\d{3,4}$')
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Invalid StoreNumber '$StoreNumber' (must be 3 or 4 digits, not all zeros)."
			$OperationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
		}
		Write-Host "Failed: Invalid StoreNumber '$StoreNumber'." -ForegroundColor Red
		return
	}
	
	$storeInt = [int]$storeNumberTrim
	$store3 = $storeInt.ToString("D3")
	$store4 = $storeInt.ToString("D4")
	
	# -----------------------------
	# Pick base path
	# -----------------------------
	# The detected storeman folder comes first: it can be on another drive or named differently
	# (e.g. E:\Storeman), where the fixed paths below would miss it or pick a stale copy.
	$possibleBasePaths = @()
	if (-not [string]::IsNullOrWhiteSpace($OfficePath)) { $possibleBasePaths += $OfficePath }
	else
	{
		# (the parameter hides the script-level $OfficePath, so read that one explicitly)
		foreach ($sc in @('Script', 'Global'))
		{
			$v = Get-Variable -Name OfficePath -Scope $sc -ErrorAction SilentlyContinue
			if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $possibleBasePaths += [string]$v.Value; break }
		}
		$possibleBasePaths += @("\\localhost\storeman\office", "C:\storeman\office", "D:\storeman\office")
	}
	$basePath = $null
	foreach ($p in $possibleBasePaths)
	{
		if (Test-Path $p) { $basePath = $p; break }
	}
	
	if (-not $basePath)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "None of the base paths exist: $($possibleBasePaths -join ', ')"
			$OperationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
		}
		Write-Host "Failed: None of the base paths exist: $($possibleBasePaths -join ', ')" -ForegroundColor Red
		return
	}
	
	# -----------------------------
	# Extract machine terminal (last 1-3 digits, padded to 3)
	# -----------------------------
	$mn = $MachineName
	if ($null -eq $mn) { $mn = "" }
	$mn = $mn.Trim()
	
	$mn = $mn -replace '^[\\\/]+', '' # strip leading \\ or /
	if ($mn -match '[\\\/]') { $mn = ($mn -split '[\\\/]')[0] } # keep host portion
	if ($mn -match '\.') { $mn = ($mn -split '\.')[0] } # drop domain
	$mn = $mn.Trim().ToUpper()
	
	if ($mn -notmatch '(\d{1,3})$')
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "MachineName '$MachineName' does not end with 1-3 digits; cannot extract terminal number."
			$OperationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
		}
		Write-Host "Failed: MachineName '$MachineName' cannot provide terminal digits." -ForegroundColor Red
		return
	}
	
	$machineNumber = ([int]$Matches[1]).ToString("D3")
	
	# Safeguard for backoffice unless explicitly allowed
	if ($machineNumber -eq "901" -and -not $AllowBackoffice)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Refusing to run because extracted terminal number is 901 (backoffice). Use -AllowBackoffice to override."
			$OperationStatus["OldXFoldersDeletion"].Details = "This could delete many lane folders. Override only if intentional."
		}
		Write-Host "Failed: Refusing to run on terminal 901 (backoffice safeguard). Use -AllowBackoffice to override." -ForegroundColor Red
		return
	}
	
	# Keep set: current lane + 900 + 901
	$keepStores = @($store3, $store4)
	$keepTerminals = @($machineNumber, "900", "901")
	
	# -----------------------------
	# Enumerate X* and Z* folders
	# -----------------------------
	$folders = @()
	
	$xFolders = Get-ChildItem -Path $basePath -Directory -Filter "X*" -ErrorAction SilentlyContinue
	if ($xFolders) { $folders += $xFolders }
	
	$zFolders = Get-ChildItem -Path $basePath -Directory -Filter "Z*" -ErrorAction SilentlyContinue
	if ($zFolders) { $folders += $zFolders }
	
	if (-not $folders -or $folders.Count -eq 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "No Folders Found"
			$OperationStatus["OldXFoldersDeletion"].Message = "No X* or Z* folders were found under '$basePath'."
			$OperationStatus["OldXFoldersDeletion"].Details = "Nothing to delete."
		}
		Write-Host "Info: No X* or Z* folders found under '$basePath'." -ForegroundColor Cyan
		return
	}
	
	$anyFoldersSeen = $true
	
	foreach ($folder in $folders)
	{
		$folderName = $folder.Name
		
		# STRICT pattern only: (X|Z)(letter)(store 3-4 digits)(terminal 3 digits)
		if ($folderName -match '^(?<prefix>[XZ][A-Za-z])(?<store>\d{3,4})(?<terminal>\d{3})$')
		{
			$folderStore = $matches['store']
			$folderTerminal = $matches['terminal']
			
			$storeOk = ($keepStores -contains $folderStore)
			$terminalOk = ($keepTerminals -contains $folderTerminal)
			
			# Delete if it does NOT match new store AND allowed terminal(s)
			if (-not ($storeOk -and $terminalOk))
			{
				$anyCandidatesMatched = $true
				
				$maxRetries = 3
				$retryCount = 0
				$deleted = $false
				
				while ($retryCount -lt $maxRetries -and -not $deleted)
				{
					try
					{
						Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction Stop
						$deletedFolders += $folderName
						$deleted = $true
					}
					catch
					{
						$retryCount++
						Start-Sleep -Seconds 2
						
						if ($retryCount -ge $maxRetries)
						{
							$failedToDeleteFolders += $folderName
							Write-Host "Failed to delete folder: $folderName. Error: $($_.Exception.Message)" -ForegroundColor Red
						}
					}
				}
			}
		}
	}
	
	# -----------------------------
	# Results
	# -----------------------------
	$resultMessage = ""
	if ($deletedFolders.Count -gt 0)
	{
		$resultMessage += "Deleted folders:`n$($deletedFolders -join "`n")`n"
	}
	if ($failedToDeleteFolders.Count -gt 0)
	{
		$resultMessage += "Failed to delete folders:`n$($failedToDeleteFolders -join "`n")`n"
	}
	
	if (-not $anyCandidatesMatched)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "No Matching Folders"
			$OperationStatus["OldXFoldersDeletion"].Message = "No X*/Z* folders needed deletion. Kept stores: $($keepStores -join ', ') terminals: $($keepTerminals -join ', ')."
			$OperationStatus["OldXFoldersDeletion"].Details = "Nothing to delete."
		}
		Write-Host "Info: Nothing to delete. Kept stores: $($keepStores -join ', ') terminals: $($keepTerminals -join ', ')." -ForegroundColor Cyan
		return
	}
	
	if ($deletedFolders.Count -gt 0 -and $failedToDeleteFolders.Count -eq 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Successful"
			$OperationStatus["OldXFoldersDeletion"].Message = "Old X*/Z* folders deleted successfully."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Success: Old X*/Z* folders deleted successfully." -ForegroundColor Green
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Green }
	}
	elseif ($deletedFolders.Count -gt 0 -and $failedToDeleteFolders.Count -gt 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Partial Failure"
			$OperationStatus["OldXFoldersDeletion"].Message = "Some old X*/Z* folders could not be deleted."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Warning: Some old X*/Z* folders could not be deleted." -ForegroundColor Yellow
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Yellow }
	}
	else
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Failed to delete any old X*/Z* folders."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Error: Failed to delete any old X*/Z* folders." -ForegroundColor Red
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Red }
	}
}

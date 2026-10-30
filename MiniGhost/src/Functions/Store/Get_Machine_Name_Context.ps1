# ===================================================================================================
#                           FUNCTION: Get_Machine_Name_Context
# ---------------------------------------------------------------------------------------------------
# Description:
# Normalizes a requested machine name and derives the terminal behavior used by INI + SQL sync paths.
# ===================================================================================================

function Get_Machine_Name_Context
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[string]$CurrentMachineName = $env:COMPUTERNAME,
		[Parameter(Mandatory = $false)]
		[ValidatePattern('^\d{3}$')]
		[string]$ServerTerminal = '901'
	)
	
	$normalizedHost = [string]$MachineName
	if ($null -eq $normalizedHost) { $normalizedHost = "" }
	$normalizedHost = $normalizedHost.Trim()
	$normalizedHost = $normalizedHost -replace '^[\\\/]+', ''
	if ($normalizedHost -match '[\\\/]') { $normalizedHost = ($normalizedHost -split '[\\\/]')[0] }
	if ($normalizedHost -match '\.') { $normalizedHost = ($normalizedHost -split '\.')[0] }
	$normalizedHost = $normalizedHost.Trim().ToUpper()
	
	$currentHost = [string]$CurrentMachineName
	if ($null -eq $currentHost) { $currentHost = "" }
	$currentHost = $currentHost.Trim()
	$currentHost = $currentHost -replace '^[\\\/]+', ''
	if ($currentHost -match '[\\\/]') { $currentHost = ($currentHost -split '[\\\/]')[0] }
	if ($currentHost -match '\.') { $currentHost = ($currentHost -split '\.')[0] }
	$currentHost = $currentHost.Trim().ToUpper()
	
	$nodeRole = "Unknown"
	try
	{
		if (-not [string]::IsNullOrWhiteSpace([string]$script:NodeRole))
		{
			$nodeRole = ([string]$script:NodeRole).Trim()
		}
	}
	catch { }
	
	$isServerNodeRole = ($nodeRole -match '^(?i)(StoreServer|HostServer)$')
	$isLaneNodeRole = ($nodeRole -match '^(?i)Lane$')
	$machineNameFormatText = if ($isServerNodeRole)
	{
		"Valid examples:`n  SERVER001`n  SERVER-001`n  0231SERVER001`n`nRule: [letters] + [optional store digits before or after the prefix] + [optional hyphen] + [3-digit terminal]."
	}
	elseif ($isLaneNodeRole)
	{
		"Valid examples:`n  POS001`n  POS-001`n  LN026001`n  0384LANE001`n  001`n`nRule: [letters] + [optional store digits before or after the prefix] + [optional hyphen] + [3-digit terminal]."
	}
	else
	{
		"Lane examples:`n  POS001`n  POS-001`n  LN026001`n  0384LANE001`n  001`n`nServer examples:`n  SERVER001`n  SERVER-001`n  0231SERVER001`n`nRule: [letters] + [optional store digits before or after the prefix] + [optional hyphen] + [3-digit terminal]."
	}
	$namePrefix = $null
	$nameSeparator = ""
	$storePrefixRaw = $null
	$storePlacement = "None"
	$terminalRaw = $null
	$inputWasTerminalOnly = $false
	
	if ([string]::IsNullOrWhiteSpace($normalizedHost))
	{
		return [PSCustomObject]@{
			Success            = $false
			ErrorMessage       = "Machine name cannot be empty."
			MachineName        = $null
			CurrentMachineName = $currentHost
			IsSameMachineName  = $false
			IsServerMoniker    = $false
			MachineNumber      = $null
			NodeRole           = $nodeRole
			NamePrefix         = $null
			NameSeparator      = ""
			StorePrefix        = $null
			StorePlacement     = "None"
			TerminalNumber     = $null
			InputWasTerminalOnly = $false
		}
	}
	
	if ($normalizedHost -match '^\d{3}$')
	{
		if ($isServerNodeRole)
		{
			return [PSCustomObject]@{
				Success              = $false
				ErrorMessage         = "Terminal-only input is only supported for lane-style machine names."
				MachineName          = $normalizedHost
				CurrentMachineName   = $currentHost
				IsSameMachineName    = $false
				IsServerMoniker      = $true
				MachineNumber        = $null
				NodeRole             = $nodeRole
				NamePrefix           = $null
				NameSeparator        = ""
				StorePrefix          = $null
				StorePlacement       = "None"
				TerminalNumber       = $normalizedHost
				InputWasTerminalOnly = $true
			}
		}
		
		$currentStore = $null
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('StoreNumber') -and $script:FunctionResults['StoreNumber'])
		{
			$currentStore = ($script:FunctionResults['StoreNumber'].ToString()).Trim()
		}
		else
		{
			$currentStore = Get_Store_Number_From_INI
			if ($currentStore) { $currentStore = ($currentStore.ToString()).Trim() }
		}
		
		if ($currentHost -match '^(?<StorePrefix>\d{3,4})(?<NamePrefix>[A-Z]+)(?<NameSeparator>-?)(?<Terminal>\d{3})$')
		{
			$storePrefixRaw = $matches['StorePrefix']
			$namePrefix = $matches['NamePrefix']
			$nameSeparator = $matches['NameSeparator']
			$storePlacement = "Leading"
		}
		elseif ($currentHost -match '^(?<NamePrefix>[A-Z]+)(?<StorePrefix>\d{3,4})(?<NameSeparator>-?)(?<Terminal>\d{3})$')
		{
			$namePrefix = $matches['NamePrefix']
			$storePrefixRaw = $matches['StorePrefix']
			$nameSeparator = $matches['NameSeparator']
			$storePlacement = "Embedded"
		}
		elseif ($currentHost -match '^(?<NamePrefix>[A-Z]+)(?<NameSeparator>-?)(?<Terminal>\d{3})$')
		{
			$namePrefix = $matches['NamePrefix']
			$nameSeparator = $matches['NameSeparator']
			$storePlacement = "None"
		}
		else
		{
			return [PSCustomObject]@{
				Success              = $false
				ErrorMessage         = "Terminal-only input requires the current machine name to already use a lane naming style with a reusable prefix."
				MachineName          = $normalizedHost
				CurrentMachineName   = $currentHost
				IsSameMachineName    = $false
				IsServerMoniker      = $false
				MachineNumber        = $null
				NodeRole             = $nodeRole
				NamePrefix           = $null
				NameSeparator        = ""
				StorePrefix          = $null
				StorePlacement       = "None"
				TerminalNumber       = $normalizedHost
				InputWasTerminalOnly = $true
			}
		}
		
		if (($storePlacement -ne "None") -and [string]::IsNullOrWhiteSpace($storePrefixRaw) -and ($currentStore -match '^\d{3,4}$'))
		{
			$storePrefixRaw = $currentStore
		}
		
		if (($storePlacement -ne "None") -and [string]::IsNullOrWhiteSpace($storePrefixRaw))
		{
			return [PSCustomObject]@{
				Success              = $false
				ErrorMessage         = "Terminal-only input could not determine which store number to preserve from the current naming style."
				MachineName          = $normalizedHost
				CurrentMachineName   = $currentHost
				IsSameMachineName    = $false
				IsServerMoniker      = $false
				MachineNumber        = $null
				NodeRole             = $nodeRole
				NamePrefix           = $namePrefix
				NameSeparator        = $nameSeparator
				StorePrefix          = $null
				StorePlacement       = $storePlacement
				TerminalNumber       = $normalizedHost
				InputWasTerminalOnly = $true
			}
		}
		
		$terminalRaw = $normalizedHost
		$inputWasTerminalOnly = $true
		
		switch ($storePlacement)
		{
			"Leading"  { $normalizedHost = "$storePrefixRaw$namePrefix$nameSeparator$terminalRaw" }
			"Embedded" { $normalizedHost = "$namePrefix$storePrefixRaw$nameSeparator$terminalRaw" }
			default    { $normalizedHost = "$namePrefix$nameSeparator$terminalRaw" }
		}
	}
	elseif ($normalizedHost -match '^(?<StorePrefix>\d{3,4})(?<NamePrefix>[A-Z]+)(?<NameSeparator>-?)(?<Terminal>\d{3})$')
	{
		$storePrefixRaw = $matches['StorePrefix']
		$namePrefix = $matches['NamePrefix']
		$nameSeparator = $matches['NameSeparator']
		$terminalRaw = $matches['Terminal']
		$storePlacement = "Leading"
	}
	elseif ($normalizedHost -match '^(?<NamePrefix>[A-Z]+)(?<StorePrefix>\d{3,4})(?<NameSeparator>-?)(?<Terminal>\d{3})$')
	{
		$namePrefix = $matches['NamePrefix']
		$storePrefixRaw = $matches['StorePrefix']
		$nameSeparator = $matches['NameSeparator']
		$terminalRaw = $matches['Terminal']
		$storePlacement = "Embedded"
	}
	elseif ($normalizedHost -match '^(?<NamePrefix>[A-Z]+)(?<NameSeparator>-?)(?<Terminal>\d{3})$')
	{
		$namePrefix = $matches['NamePrefix']
		$nameSeparator = $matches['NameSeparator']
		$terminalRaw = $matches['Terminal']
		$storePlacement = "None"
	}
	else
	{
		return [PSCustomObject]@{
			Success              = $false
			ErrorMessage         = "Invalid format.`n`n$machineNameFormatText"
			MachineName          = $normalizedHost
			CurrentMachineName   = $currentHost
			IsSameMachineName    = $false
			IsServerMoniker      = $isServerNodeRole
			MachineNumber        = $null
			NodeRole             = $nodeRole
			NamePrefix           = $null
			NameSeparator        = ""
			StorePrefix          = $null
			StorePlacement       = "None"
			TerminalNumber       = $null
			InputWasTerminalOnly = $false
		}
	}
	
	if ($normalizedHost.Length -gt 15)
	{
		return [PSCustomObject]@{
			Success              = $false
			ErrorMessage         = "Machine name '$normalizedHost' exceeds the Windows 15-character computer-name limit."
			MachineName          = $normalizedHost
			CurrentMachineName   = $currentHost
			IsSameMachineName    = $false
			IsServerMoniker      = $isServerNodeRole
			MachineNumber        = $null
			NodeRole             = $nodeRole
			NamePrefix           = $namePrefix
			NameSeparator        = $nameSeparator
			StorePrefix          = $storePrefixRaw
			StorePlacement       = $storePlacement
			TerminalNumber       = $terminalRaw
			InputWasTerminalOnly = $inputWasTerminalOnly
		}
	}
	
	$isServerMoniker = $isServerNodeRole
	if ($normalizedHost -match '(?i)SERVER')
	{
		$isServerMoniker = $true
	}
	
	$machineNumber = $null
	if ($isServerMoniker)
	{
		$machineNumber = $ServerTerminal
	}
	else
	{
		$machineNumber = ([int]$terminalRaw).ToString("D3")
		if ($machineNumber -eq '000')
		{
			return [PSCustomObject]@{
				Success              = $false
				ErrorMessage         = "Invalid terminal number extracted from '$normalizedHost' (000 is not allowed)."
				MachineName          = $normalizedHost
				CurrentMachineName   = $currentHost
				IsSameMachineName    = $false
				IsServerMoniker      = $false
				MachineNumber        = $null
				NodeRole             = $nodeRole
				NamePrefix           = $namePrefix
				NameSeparator        = $nameSeparator
				StorePrefix          = $storePrefixRaw
				StorePlacement       = $storePlacement
				TerminalNumber       = $terminalRaw
				InputWasTerminalOnly = $inputWasTerminalOnly
			}
		}
	}
	
	return [PSCustomObject]@{
		Success              = $true
		ErrorMessage         = $null
		MachineName          = $normalizedHost
		CurrentMachineName   = $currentHost
		IsSameMachineName    = ($normalizedHost -eq $currentHost)
		IsServerMoniker      = $isServerMoniker
		MachineNumber        = $machineNumber
		NodeRole             = $nodeRole
		NamePrefix           = $namePrefix
		NameSeparator        = $nameSeparator
		StorePrefix          = $storePrefixRaw
		StorePlacement       = $storePlacement
		TerminalNumber       = $terminalRaw
		InputWasTerminalOnly = $inputWasTerminalOnly
	}
}

# ===================================================================================================
#                              FUNCTION: Ping_All_Nodes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Pings all nodes of one type (Lane / Scale / Backoffice) for a store and reports
#   per-node reachability in the log box. Uses Test-Hosts_Parallel so a store full of
#   dead nodes takes about two seconds worst case instead of minutes.
#
# Parameters:
#   - NodeType     [string]: "Lane", "Scale" or "Backoffice" (required)
#   - StoreNumber  [string]: shown in the summary line (optional)
#
# Variables:
#   - $script:FunctionResults: LaneNumToMachineName / ScaleCodeToIPInfo / BackofficeNumToMachineName
# ===================================================================================================

function Ping_All_Nodes
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateSet("Lane", "Scale", "Backoffice")]
		[string]$NodeType,
		[string]$StoreNumber
	)
	
	Write_Log "`r`n==================== Starting Ping_All_Nodes ($NodeType) Function ====================`r`n" "blue"
	
	$nodesToPing = @()
	$nodeLabel = ""
	$nodeSummary = ""
	
	switch ($NodeType)
	{
		"Lane" {
			$nodeLabel = "Lane"
			$nodeSummary = "Lanes"
			if (-not $script:FunctionResults.ContainsKey('LaneNumToMachineName'))
			{
				Write_Log "Lane mappings not available. Please run Retrieve_Nodes first." "red"
				return
			}
			$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
			foreach ($laneNum in $LaneNumToMachineName.Keys | Where-Object { $_ -match '^\d{3}$' })
			{
				$machineName = $LaneNumToMachineName[$laneNum]
				if ($machineName -and $machineName -notin @("Unknown", "Not Found"))
				{
					$nodesToPing += [PSCustomObject]@{
						Key    = $laneNum
						Target = $machineName
						Label  = "" # Do not use $null or $null string
					}
				}
			}
			$nodesToPing = $nodesToPing | Sort-Object { [int]$_.Key }
		}
		
		"Scale" {
			$nodeLabel = "Scale"
			$nodeSummary = "Scales"
			if (-not $script:FunctionResults.ContainsKey('ScaleCodeToIPInfo') -or $script:FunctionResults['ScaleCodeToIPInfo'].Count -eq 0)
			{
				Write_Log "No scales found to ping." "yellow"
				return
			}
			$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
			foreach ($code in $ScaleCodeToIPInfo.Keys | Where-Object { $_ -match '^\d{1,3}$' })
			{
				$scaleObj = $ScaleCodeToIPInfo[$code]
				$ip = $null
				if ($scaleObj.PSObject.Properties['FullIP'])
				{
					$ip = $scaleObj.FullIP
				}
				elseif ($scaleObj.PSObject.Properties['IPNetwork'] -and $scaleObj.PSObject.Properties['IPDevice'])
				{
					$ip = "$($scaleObj.IPNetwork)$($scaleObj.IPDevice)"
				}
				if ($ip -and $ip -notin @("Unknown", "Not Found", ""))
				{
					$nodesToPing += [PSCustomObject]@{
						Key    = $code
						Target = $ip
						Label  = $scaleObj.ScaleName
					}
				}
			}
			$nodesToPing = $nodesToPing | Sort-Object { [int]$_.Key }
			# Remove duplicate IPs
			$uniqueTargets = @{ }
			$finalNodesToPing = @()
			foreach ($node in $nodesToPing)
			{
				if (-not $uniqueTargets.ContainsKey($node.Target))
				{
					$uniqueTargets[$node.Target] = $true
					$finalNodesToPing += $node
				}
			}
			$nodesToPing = $finalNodesToPing
		}
		
		"Backoffice" {
			$nodeLabel = "Backoffice"
			$nodeSummary = "Backoffices"
			if (-not $script:FunctionResults.ContainsKey('BackofficeNumToMachineName'))
			{
				Write_Log "Backoffice information is not available. Please run Retrieve_Nodes first." "red"
				return
			}
			$BackofficeNumToMachineName = $script:FunctionResults['BackofficeNumToMachineName']
			foreach ($boNum in $BackofficeNumToMachineName.Keys | Where-Object { $_ -match '^\d{3}$' })
			{
				$machineName = $BackofficeNumToMachineName[$boNum]
				if ($machineName -and $machineName -notin @("Unknown", "Not Found"))
				{
					$nodesToPing += [PSCustomObject]@{
						Key    = $boNum
						Target = $machineName
						Label  = "" # Do not use $null or $null string
					}
				}
			}
			$nodesToPing = $nodesToPing | Sort-Object { [int]$_.Key }
		}
	}
	
	# Final deduplication by Target (just to be sure)
	$uniqueTargets = @{ }
	$finalList = @()
	foreach ($node in $nodesToPing)
	{
		if (-not $uniqueTargets.ContainsKey($node.Target))
		{
			$uniqueTargets[$node.Target] = $node
			$finalList += $node
		}
	}
	$nodesToPing = $finalList
	
	if ($nodesToPing.Count -eq 0)
	{
		Write_Log "No valid $nodeSummary found to ping." "yellow"
		return
	}
	
	Write_Log "All $nodeSummary will be pinged." "green"
	
	# v2.6.9: probe all valid targets in parallel (capped concurrency, 400ms
	# timeout) and then log the results in node order. This replaces the old
	# serial Test-Connection loop, which could block for minutes when several
	# hosts were down.
	$probeTargets = @()
	foreach ($node in $nodesToPing)
	{
		if (-not [string]::IsNullOrWhiteSpace($node.Target) -and $node.Target -notin @("Unknown", "Not Found"))
		{
			$probeTargets += $node.Target
		}
	}
	$probe = @{}
	if ($probeTargets.Count -gt 0)
	{
		$probe = Test-Hosts_Parallel -Targets $probeTargets
	}
	
	$successCount = 0
	$failureCount = 0
	$skippedCount = 0
	
	foreach ($node in $nodesToPing)
	{
		$primary = $node.Key
		$target = $node.Target
		
		# Compose labelInfo only for scales if present and non-empty
		$labelInfo = ""
		if ($NodeType -eq "Scale" -and $node.PSObject.Properties.Name -contains "Label" -and $node.Label -and $node.Label -ne "")
		{
			$labelInfo = " [($($node.Label))]"
		}
		
		if ([string]::IsNullOrWhiteSpace($target) -or $target -in @("Unknown", "Not Found"))
		{
			# For lanes/backoffices, $labelInfo is always ""
			Write_Log "$nodeLabel #[$primary]: Target '$target'. Status: Skipped." "yellow"
			$skippedCount++
			continue
		}
		
		$state = 'ERROR'
		if ($probe.ContainsKey($target)) { $state = $probe[$target] }
		if ($state -eq 'OK')
		{
			Write_Log ("$nodeLabel #[$primary]${labelInfo}: Target '$target' is reachable. Status: Success.") "green"
			$successCount++
		}
		elseif ($state -eq 'FAIL')
		{
			Write_Log ("$nodeLabel #[$primary]${labelInfo}: Target '$target' is not reachable. Status: Failed.") "red"
			$failureCount++
		}
		else
		{
			Write_Log ("$nodeLabel #[$primary]${labelInfo}: Failed to ping target '$target'. Error: name resolution or probe failed.") "red"
			$failureCount++
		}
	}
	

	$summaryMsg = "Ping Summary for $nodeSummary"
	if ($StoreNumber) { $summaryMsg += " (Store Number: $StoreNumber)" }
	$summaryMsg += " - Success: $successCount, Failed: $failureCount."
	Write_Log $summaryMsg "blue"
	Write_Log "`r`n==================== Ping_All_Nodes ($NodeType) Function Completed ====================" "blue"
}

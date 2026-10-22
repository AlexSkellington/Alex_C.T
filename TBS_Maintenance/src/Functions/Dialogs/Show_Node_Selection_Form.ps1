# ===================================================================================================
#                              FUNCTION: Show_Node_Selection_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   The shared node picker: tabbed WinForms dialog for Lanes, Scales, Backoffices, or any
#   combination. Returns a hashtable with keys matching the selected node types (Lanes,
#   Scales, Backoffices), or $null when the user cancels or selects nothing.
#
# Parameters:
#   - StoreNumber        [string] (required): store to look up
#   - NodeTypes          [string[]]: "Lane" / "Scale" / "Backoffice" (default: all three)
#   - OnlyBizerbaScales  [switch]: limit the Scales tab to Bizerba scales
#   - Title              [string]: dialog title (default "Select Nodes to Process")
#   - ExcludedNodes      [string[]]: machine names and/or 3-digit lane numbers to hide
#   - SingleLaneOnly     [switch]: cap lane selection at 1
#   - LaneSelectionLimit [int]: cap lane selection at N (0 = unlimited)
# ===================================================================================================

function Show_Node_Selection_Form
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter()]
		[ValidateSet("Lane", "Scale", "Backoffice")]
		[string[]]$NodeTypes = @("Lane", "Scale", "Backoffice"),
		[switch]$OnlyBizerbaScales,
		[string]$Title = "Select Nodes to Process",
		[string[]]$ExcludedNodes,
		[switch]$SingleLaneOnly,
		# NEW: cap lane selection at 1 (optional)
		[int]$LaneSelectionLimit = 0 # NEW: cap lane selection at N (0 = unlimited)
	)

	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	# Normalize exclusions (names lowercased; lanes as 3-digit)
	$__ex_names = @()
	$__ex_lanes = @()
	if ($ExcludedNodes)
	{
		foreach ($e in $ExcludedNodes)
		{
			if ($null -ne $e -and "$e".Trim())
			{
				$s = "$e".Trim()
				if ($s -match '(\d{1,3})$')
				{
					try { $__ex_lanes += ('{0:D3}' -f ([int]$matches[1])) }
					catch { }
				}
				$__ex_names += $s.ToLower()
			}
		}
		$__ex_names = $__ex_names | Select-Object -Unique
		$__ex_lanes = $__ex_lanes | Select-Object -Unique
	}

	$form = New-Object System.Windows.Forms.Form
	$form.Text = $Title
	$form.Size = New-Object System.Drawing.Size(430, 450)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false

	$tabs = New-Object System.Windows.Forms.TabControl
	$tabs.Location = New-Object System.Drawing.Point(10, 10)
	$tabs.Size = New-Object System.Drawing.Size(390, 320)
	$form.Controls.Add($tabs)

	$tabControls = @{ }

	# ----- Lanes Tab -----
	if ("Lane" -in $NodeTypes)
	{
		$tabLanes = New-Object System.Windows.Forms.TabPage
		$tabLanes.Text = "Lanes"
		$clbLanes = New-Object System.Windows.Forms.CheckedListBox
		$clbLanes.Location = New-Object System.Drawing.Point(10, 10)
		$clbLanes.Size = New-Object System.Drawing.Size(350, 270)
		$clbLanes.CheckOnClick = $true
		$tabLanes.Controls.Add($clbLanes)
		$tabs.TabPages.Add($tabLanes)
		$tabControls["Lanes"] = $clbLanes

		# Load lane node data from FunctionResults
		$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
		$LaneMachineLabels = $script:FunctionResults['LaneMachineLabels']
		$LaneMachinePath = $script:FunctionResults['LaneMachinePath']
		$LaneMachineToServerPath = $script:FunctionResults['LaneMachineToServerPath']

		$allLanes = @()
		if ($script:FunctionResults.ContainsKey('LaneMachineNames') -and $script:FunctionResults['LaneMachineNames'].Count -gt 0)
		{
			$allLanes = $script:FunctionResults['LaneMachineNames']
		}
		elseif (Test-Path -Path $OfficePath)
		{
			$laneFolders = Get-ChildItem -Path $OfficePath -Directory -Filter "XF${StoreNumber}0*"
			if ($laneFolders) { $allLanes = $laneFolders | ForEach-Object { $_.Name.Substring($_.Name.Length - 3, 3) } }
		}

		$sortedLanes = $allLanes | Sort-Object
		foreach ($laneMachine in $sortedLanes)
		{
			# Find numeric lane (if possible)
			$laneNum = $null
			foreach ($key in $LaneNumToMachineName.Keys)
			{
				if ($LaneNumToMachineName[$key] -eq $laneMachine -and $key -match '^\d{3}$') { $laneNum = $key; break }
			}

			# Skip excluded by machine name or lane number
			$__skip = $false
			if ($laneMachine -and ($__ex_names -contains $laneMachine.ToLower())) { $__skip = $true }
			elseif ($laneNum -and ($__ex_lanes -contains $laneNum)) { $__skip = $true }
			if ($__skip) { continue }

			# Add item
			$obj = [PSCustomObject]@{
				LaneNumber  = $laneNum
				MachineName = $laneMachine
				Label	    = $LaneMachineLabels[$laneMachine]
				Path	    = $LaneMachinePath[$laneMachine]
				ServerPath  = $LaneMachineToServerPath[$laneMachine]
				DisplayName = if ($laneNum) { "$laneMachine ($laneNum)" } else { $laneMachine }
			}
			$obj | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.DisplayName } -Force
			$clbLanes.Items.Add($obj) | Out-Null
		}

		# --- Enforce optional selection cap for lanes (single or N) ---
		$__laneCap = if ($SingleLaneOnly) { 1 }
		elseif ($LaneSelectionLimit -gt 0) { [int]$LaneSelectionLimit }
		else { 0 }
		if ($__laneCap -gt 0)
		{
			$clbLanes.Add_ItemCheck({
					# If trying to check and it would exceed the cap, cancel the check.
					if ($_.NewValue -ne [System.Windows.Forms.CheckState]::Checked) { return }
					$idx = $_.Index
					$checkedCount = 0
					for ($i = 0; $i -lt $clbLanes.Items.Count; $i++)
					{
						if ($i -ne $idx -and $clbLanes.GetItemChecked($i)) { $checkedCount++ }
					}
					if (($checkedCount + 1) -gt $__laneCap)
					{
						$_.NewValue = [System.Windows.Forms.CheckState]::Unchecked
						if (-not $script:__laneLimitWarned)
						{
							[System.Windows.Forms.MessageBox]::Show("You can select at most $__laneCap lane(s).", "Selection limit",
								[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
							$script:__laneLimitWarned = $true
						}
					}
				})
		}
	}

	# ----- Scales Tab -----
	if ("Scale" -in $NodeTypes)
	{
		$tabScales = New-Object System.Windows.Forms.TabPage
		$tabScales.Text = "Scales"
		$clbScales = New-Object System.Windows.Forms.CheckedListBox
		$clbScales.Location = New-Object System.Drawing.Point(10, 10)
		$clbScales.Size = New-Object System.Drawing.Size(350, 270)
		$clbScales.CheckOnClick = $true
		$tabScales.Controls.Add($clbScales)
		$tabs.TabPages.Add($tabScales)
		$tabControls["Scales"] = $clbScales

		$allScales = @()
		if ($script:FunctionResults.ContainsKey('ScaleCodeToIPInfo'))
		{
			$allScales = $script:FunctionResults['ScaleCodeToIPInfo'].Values
			if ($OnlyBizerbaScales) { $allScales = $allScales | Where-Object { $_.ScaleBrand -match 'bizerba' } }
		}
		$sortedScales = $allScales | Sort-Object { [int]($_.ScaleCode) }

		$uniqueScaleIPs = @{ }
		$dedupedScales = @()
		foreach ($scale in $sortedScales)
		{
			$ipKey = if ($scale.FullIP) { $scale.FullIP }
			elseif ($scale.IPNetwork -and $scale.IPDevice) { "$($scale.IPNetwork)$($scale.IPDevice)" }
			else { $null }
			if ($ipKey -and -not $uniqueScaleIPs.ContainsKey($ipKey))
			{
				$uniqueScaleIPs[$ipKey] = $true
				$dedupedScales += $scale
			}
		}

		foreach ($scale in $dedupedScales)
		{
			# Skip excluded by name/code
			$__skip = $false
			if ($scale.ScaleName -and ($__ex_names -contains ($scale.ScaleName.ToLower()))) { $__skip = $true }
			elseif ($scale.ScaleCode -and ($__ex_lanes -contains ('{0:D3}' -f ([int]$scale.ScaleCode)))) { $__skip = $true }
			if ($__skip) { continue }

			$ip = if ($scale.IPNetwork -and $scale.IPDevice) { "$($scale.IPNetwork)$($scale.IPDevice)" }
			else { "" }
			$displayName = "$($scale.ScaleName) [$ip]"
			$scaleObj = [PSCustomObject]@{
				ScaleCode   = $scale.ScaleCode
				ScaleName   = $scale.ScaleName
				IPAddress   = $ip
				Vendor	    = $scale.Vendor
				FullIP	    = $scale.FullIP
				IPNetwork   = $scale.IPNetwork
				IPDevice    = $scale.IPDevice
				ScaleBrand  = $scale.ScaleBrand
				DisplayName = $displayName
			}
			$scaleObj | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.DisplayName } -Force
			$clbScales.Items.Add($scaleObj) | Out-Null
		}
	}

	# ----- Backoffices Tab -----
	if ("Backoffice" -in $NodeTypes)
	{
		$tabBO = New-Object System.Windows.Forms.TabPage
		$tabBO.Text = "Backoffices"
		$clbBO = New-Object System.Windows.Forms.CheckedListBox
		$clbBO.Location = New-Object System.Drawing.Point(10, 10)
		$clbBO.Size = New-Object System.Drawing.Size(350, 270)
		$clbBO.CheckOnClick = $true
		$tabBO.Controls.Add($clbBO)
		$tabs.TabPages.Add($tabBO)
		$tabControls["Backoffices"] = $clbBO

		$boDict = if ($script:FunctionResults.ContainsKey('BackofficeNumToMachineName')) { $script:FunctionResults['BackofficeNumToMachineName'] }
		else { @{ } }
		$boLabels = $script:FunctionResults['BackofficeNumToLabel']
		$boPaths = $script:FunctionResults['BackofficeNumToPath']

		$seenBonumbers = @{ }
		foreach ($boNumKey in $boDict.Keys | Sort-Object)
		{
			if ($boNumKey -match '(\d{3})')
			{
				$bonum = $matches[1]

				# Skip excluded by machine name or BO number
				$__skip = $false
				$machineName = $boDict[$boNumKey]
				if ($machineName -and ($__ex_names -contains $machineName.ToLower())) { $__skip = $true }
				elseif ($__ex_lanes -contains $bonum) { $__skip = $true }
				if ($__skip) { continue }

				if (-not $seenBonumbers.ContainsKey($bonum))
				{
					$seenBonumbers[$bonum] = $true
					$label = $boLabels[$boNumKey]
					$path = $boPaths[$boNumKey]

					$obj = [PSCustomObject]@{
						BONumber    = $bonum
						MachineName = $machineName
						Label	    = $label
						Path	    = $path
						DisplayName = if ($machineName) { "$machineName ($bonum)" } else { "Unknown ($bonum)" }
					}
					$obj | Add-Member ScriptMethod ToString { $this.DisplayName } -Force
					$clbBO.Items.Add($obj) | Out-Null
				}
			}
		}
	}

	$btnSelectAll = New-Object System.Windows.Forms.Button
	$btnSelectAll.Location = New-Object System.Drawing.Point(20, 340)
	$btnSelectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnSelectAll.Text = "Select All"
	$btnSelectAll.BackColor = [System.Drawing.SystemColors]::Control

	$btnDeselectAll = New-Object System.Windows.Forms.Button
	$btnDeselectAll.Location = New-Object System.Drawing.Point(220, 340)
	$btnDeselectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnDeselectAll.Text = "Deselect All"
	$btnDeselectAll.BackColor = [System.Drawing.SystemColors]::Control

	$setBtnColor = {
		param ($btn,
			$state)
		switch ($state)
		{
			1 { $btn.BackColor = [System.Drawing.Color]::Yellow }
			2 { $btn.BackColor = [System.Drawing.Color]::LightGreen }
			Default { $btn.BackColor = [System.Drawing.SystemColors]::Control }
		}
	}
	&$setBtnColor $btnSelectAll 0
	&$setBtnColor $btnDeselectAll 0

	# --- SINGLE TAB SELECT ALL LOGIC ---
	if ($tabs.TabPages.Count -eq 1)
	{
		$clb = $tabControls[$tabs.TabPages[0].Text]
		$isAllChecked = {
			for ($i = 0; $i -lt $clb.Items.Count; $i++) { if (-not $clb.GetItemChecked($i)) { return $false } }
			return $clb.Items.Count -gt 0
		}
		$isAnyChecked = {
			for ($i = 0; $i -lt $clb.Items.Count; $i++) { if ($clb.GetItemChecked($i)) { return $true } }
			return $false
		}
		$btnSelectAll.Add_Click({
				$allChecked = & $isAllChecked
				if (-not $allChecked)
				{
					for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) }
					& $setBtnColor $btnSelectAll 2
				}
			})
		$btnDeselectAll.Add_Click({
				for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) }
				& $setBtnColor $btnSelectAll 0
				& $setBtnColor $btnDeselectAll 0
			})
		foreach ($event in @("Add_ItemCheck", "Add_MouseUp", "Add_KeyUp"))
		{
			$clb.$event.Invoke({
					Start-Sleep -Milliseconds 30
					$allChecked = & $isAllChecked
					$anyChecked = & $isAnyChecked
					if ($allChecked) { & $setBtnColor $btnSelectAll 2 }
					elseif ($anyChecked) { & $setBtnColor $btnSelectAll 1 }
					else { & $setBtnColor $btnSelectAll 0 }
				})
		}
	}
	else
	{
		# ---- NORMAL MULTI-TAB LOGIC ----
		$tabSelectState = @{ }
		$lastSelectTabIndex = $null
		$selectAllYellowTabIndex = $null

		$btnSelectAll.Add_Click({
				$tabName = $tabs.SelectedTab.Text
				$clb = $tabControls[$tabName]
				$tabIndex = $tabs.SelectedIndex
				$currentTabAllChecked = $true
				for ($i = 0; $i -lt $clb.Items.Count; $i++) { if (-not $clb.GetItemChecked($i)) { $currentTabAllChecked = $false; break } }
				$allTabsChecked = $true
				foreach ($clbTest in $tabControls.Values)
				{
					for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
				}
				if ($allTabsChecked)
				{
					foreach ($k in $tabControls.Keys)
					{
						$tabIndex2 = -1
						for ($t = 0; $t -lt $tabs.TabPages.Count; $t++) { if ($tabs.TabPages[$t].Text -eq $k) { $tabIndex2 = $t; break } }
						if ($tabIndex2 -eq -1) { continue }
						$tabSelectState[$tabIndex2] = 0
					}
					$tabSelectState[$tabIndex] = 0
					$selectAllYellowTabIndex = $null
					&$setBtnColor $btnSelectAll 0
					return
				}
				if ($currentTabAllChecked -and -not $allTabsChecked)
				{
					foreach ($k in $tabControls.Keys)
					{
						$list = $tabControls[$k]
						for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $true) }
						$tabIndex2 = -1
						for ($t = 0; $t -lt $tabs.TabPages.Count; $t++) { if ($tabs.TabPages[$t].Text -eq $k) { $tabIndex2 = $t; break } }
						if ($tabIndex2 -eq -1) { continue }
						$tabSelectState[$tabIndex2] = 2
					}
					&$setBtnColor $btnSelectAll 2
					$selectAllYellowTabIndex = $null
					return
				}
				for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) }
				$tabSelectState[$tabIndex] = 1
				$lastSelectTabIndex = $tabIndex
				$selectAllYellowTabIndex = $tabIndex
				$allTabsChecked = $true
				foreach ($clbTest in $tabControls.Values)
				{
					for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
				}
				if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
				else { & $setBtnColor $btnSelectAll 1 }
			})
		$btnDeselectAll.Add_Click({
				$tabName = $tabs.SelectedTab.Text
				$clb = $tabControls[$tabName]
				$tabIndex = $tabs.SelectedIndex
				$noneChecked = $true
				for ($i = 0; $i -lt $clb.Items.Count; $i++) { if ($clb.GetItemChecked($i)) { $noneChecked = $false; break } }
				if ($noneChecked)
				{
					$originalTab = $tabs.SelectedTab
					foreach ($k in $tabControls.Keys)
					{
						$tabIndex2 = -1
						for ($t = 0; $t -lt $tabs.TabPages.Count; $t++) { if ($tabs.TabPages[$t].Text -eq $k) { $tabIndex2 = $t; break } }
						if ($tabIndex2 -eq -1) { continue }
						$tabs.SelectedTab = $tabs.TabPages[$tabIndex2]
						$list = $tabControls[$k]
						for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $false) }
						$tabSelectState[$tabIndex2] = 0
						$list.Refresh()
					}
					$tabs.SelectedTab = $originalTab
				}
				else
				{
					for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) }
					$tabSelectState[$tabIndex] = 0
				}
				&$setBtnColor $btnDeselectAll 0
				&$setBtnColor $btnSelectAll 0
				$selectAllYellowTabIndex = $null
			})
		foreach ($clb in $tabControls.Values)
		{
			$clb.Add_ItemCheck({
					Start-Sleep -Milliseconds 50
					$allTabsChecked = $true
					foreach ($clbTest in $tabControls.Values)
					{
						for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
					}
					if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
					else
					{
						$tabIndex = $tabs.SelectedIndex
						$clbLocal = $tabControls[$tabs.SelectedTab.Text]
						$allChecked = $true
						for ($i = 0; $i -lt $clbLocal.Items.Count; $i++) { if (-not $clbLocal.GetItemChecked($i)) { $allChecked = $false; break } }
						if ($allChecked -and $clbLocal.Items.Count -gt 0) { & $setBtnColor $btnSelectAll 1; $selectAllYellowTabIndex = $tabIndex }
						else { & $setBtnColor $btnSelectAll 0; $selectAllYellowTabIndex = $null }
					}
				})
			$clb.Add_MouseUp({
					$allTabsChecked = $true
					foreach ($clbTest in $tabControls.Values)
					{
						for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
					}
					if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
					else
					{
						$tabIndex = $tabs.SelectedIndex
						$clbLocal = $tabControls[$tabs.SelectedTab.Text]
						$allChecked = $true
						for ($i = 0; $i -lt $clbLocal.Items.Count; $i++) { if (-not $clbLocal.GetItemChecked($i)) { $allChecked = $false; break } }
						if ($allChecked -and $clbLocal.Items.Count -gt 0) { & $setBtnColor $btnSelectAll 1; $selectAllYellowTabIndex = $tabIndex }
						else { & $setBtnColor $btnSelectAll 0; $selectAllYellowTabIndex = $null }
					}
				})
			$clb.Add_KeyUp({
					$allTabsChecked = $true
					foreach ($clbTest in $tabControls.Values)
					{
						for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
					}
					if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
					else
					{
						$tabIndex = $tabs.SelectedIndex
						$clbLocal = $tabControls[$tabs.SelectedTab.Text]
						$allChecked = $true
						for ($i = 0; $i -lt $clbLocal.Items.Count; $i++) { if (-not $clbLocal.GetItemChecked($i)) { $allChecked = $false; break } }
						if ($allChecked -and $clbLocal.Items.Count -gt 0) { & $setBtnColor $btnSelectAll 1; $selectAllYellowTabIndex = $tabIndex }
						else { & $setBtnColor $btnSelectAll 0; $selectAllYellowTabIndex = $null }
					}
				})
		}
		$tabs.add_SelectedIndexChanged({
				$tabIndex = $tabs.SelectedIndex
				$clb = $tabControls[$tabs.SelectedTab.Text]
				$allTabsChecked = $true
				foreach ($clbTest in $tabControls.Values)
				{
					for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
				}
				if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
				else
				{
					$allChecked = $true
					for ($i = 0; $i -lt $clb.Items.Count; $i++) { if (-not $clb.GetItemChecked($i)) { $allChecked = $false; break } }
					if ($allChecked -and $clb.Items.Count -gt 0) { & $setBtnColor $btnSelectAll 1; $selectAllYellowTabIndex = $tabIndex }
					else { & $setBtnColor $btnSelectAll 0; $selectAllYellowTabIndex = $null }
				}
				&$setBtnColor $btnDeselectAll 0
			})
	}

	$form.Controls.Add($btnSelectAll)
	$form.Controls.Add($btnDeselectAll)

	# OK Button
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = New-Object System.Drawing.Point(60, 380)
	$btnOK.Size = New-Object System.Drawing.Size(140, 32)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)

	# Cancel Button
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(220, 380)
	$btnCancel.Size = New-Object System.Drawing.Size(140, 32)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)
	if ($Apply_Modern_Popup_Style) { & $Apply_Modern_Popup_Style $form }

	# ----- Show form & collect selections -----
	$dialogResult = $form.ShowDialog()
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

	$result = @{ }
	if ("Lane" -in $NodeTypes)
	{
		$clb = $tabControls["Lanes"]
		$selectedLaneNums = @()
		for ($i = 0; $i -lt $clb.Items.Count; $i++)
		{
			if ($clb.GetItemChecked($i))
			{
				$item = $clb.Items[$i]
				if ($item.PSObject.Properties['LaneNumber'] -and $item.LaneNumber)
				{
					$selectedLaneNums += $item.LaneNumber
				}
			}
		}
		$result.Lanes = $selectedLaneNums
	}
	if ("Scale" -in $NodeTypes)
	{
		$clb = $tabControls["Scales"]
		$selectedScaleCodes = @()
		for ($i = 0; $i -lt $clb.Items.Count; $i++)
		{
			if ($clb.GetItemChecked($i))
			{
				$item = $clb.Items[$i]
				if ($item.PSObject.Properties['ScaleCode'] -and $item.ScaleCode)
				{
					$selectedScaleCodes += "$($item.ScaleCode)"
				}
			}
		}
		$result.Scales = $selectedScaleCodes
	}
	if ("Backoffice" -in $NodeTypes)
	{
		$clb = $tabControls["Backoffices"]
		$selectedBONums = @()
		for ($i = 0; $i -lt $clb.Items.Count; $i++)
		{
			if ($clb.GetItemChecked($i))
			{
				$item = $clb.Items[$i]
				if ($item.PSObject.Properties['BONumber'] -and $item.BONumber)
				{
					$selectedBONums += $item.BONumber
				}
			}
		}
		$result.Backoffices = $selectedBONums
	}

	# Basic validation
	if ($result.Keys.Count -eq 0 -or ($result.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum -eq 0)
	{
		[System.Windows.Forms.MessageBox]::Show("No nodes selected.", "Information",
			[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
		Remove-Variable -Name __laneLimitWarned -Scope Script -ErrorAction SilentlyContinue
		return $null
	}

	Remove-Variable -Name __laneLimitWarned -Scope Script -ErrorAction SilentlyContinue
	return $result
}

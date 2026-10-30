# ===================================================================================================
#                              FUNCTION: Remove_Archive_Bit
# ---------------------------------------------------------------------------------------------------
# Description:
#   Clears the file archive bit so the Storeman load process re-copies changed files. Run Now
#   does a one-time recursive clear using the lane map (or Ter_Load.sql parsing, local paths
#   only). Service mode installs an NSSM service with FileSystemWatchers that do a full
#   cleanup at startup and then clear the bits on each touched item as it changes; a
#   Ter_Load.sql change or a watcher overflow rebuilds the watchers with a full cleanup.
#   Includes an Uninstall Service button.
#
# Variables:
#   - $OfficePath: local Office root (input)
#   - $script:ScriptsFolder: folder for the generated service PS1
#   - $script:FunctionResults['StoreNumber']: store number for the logs
#
# Notes:
#   - Uses an already-downloaded nssm.exe; it never downloads NSSM itself.
# ===================================================================================================

function Remove_Archive_Bit
{
	[CmdletBinding()]
	param ()

	Write_Log "`r`n==================== Starting Remove_Archive_Bit Function ====================`r`n" "blue"

	# ---------------------------
	# NO-PROMPT "always sensing" poll interval (seconds)
	# ---------------------------
	$AutoPollSeconds = 3

	# ------------------------------------------------------------------------------------------------
	# Validate required context
	# ------------------------------------------------------------------------------------------------
	if (-not $OfficePath -or [string]::IsNullOrWhiteSpace($OfficePath))
	{
		Write_Log "ERROR: OfficePath is not set. Cannot continue." "red"
		Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
		return
	}

	# --- Main context variables
	$storeNumber = [string]$script:FunctionResults['StoreNumber']
	$scriptFolder = $script:ScriptsFolder

	# ------------------------------------------------------------------------------------------------
	# Resolve OfficePath to LOCAL filesystem path when OfficePath is a local UNC share
	#   Example: \\localhost\storeman\office  ->  C:\Storeman\Office (if share "storeman" points to C:\Storeman)
	# ------------------------------------------------------------------------------------------------
	$OfficePathLocalResolved = $OfficePath
	try
	{
		$op = [string]$OfficePath
		if ($op -match '^(?i)\\\\([^\\]+)\\([^\\]+)(\\.*)?$')
		{
			$uncHost = [string]$matches[1]
			$share = [string]$matches[2]
			$rest = [string]$matches[3]

			$isLocalHost = $false
			if ($uncHost -match '^(?i)localhost$') { $isLocalHost = $true }
			elseif ($uncHost -eq $env:COMPUTERNAME) { $isLocalHost = $true }
			elseif ($uncHost -eq "$($env:COMPUTERNAME).$($env:USERDNSDOMAIN)") { $isLocalHost = $true }

			if ($isLocalHost)
			{
				$rest2 = ""
				if ($rest) { $rest2 = $rest.TrimStart('\') }

				$sharePath = $null
				try
				{
					$sm = Get-SmbShare -Name $share -ErrorAction Stop
					if ($sm -and $sm.Path) { $sharePath = [string]$sm.Path }
				}
				catch { }

				if (-not $sharePath)
				{
					try
					{
						$w = Get-CimInstance -ClassName Win32_Share -Filter ("Name='{0}'" -f $share.Replace("'", "''")) -ErrorAction Stop
						if ($w -and $w.Path) { $sharePath = [string]$w.Path }
					}
					catch { }
				}

				if (-not [string]::IsNullOrWhiteSpace($sharePath))
				{
					if (-not [string]::IsNullOrWhiteSpace($rest2))
					{
						$OfficePathLocalResolved = Join-Path $sharePath $rest2
					}
					else
					{
						$OfficePathLocalResolved = $sharePath
					}
				}
			}
		}
	}
	catch { }

	# IMPORTANT: build Ter_Load paths AFTER local resolution (service must not depend on UNC)
	$terFileRaw = Join-Path $OfficePath "Load\Ter_Load.sql"
	$terFileLocal = Join-Path $OfficePathLocalResolved "Load\Ter_Load.sql"
	$terFile = $terFileLocal
	if (-not (Test-Path -LiteralPath $terFileLocal) -and (Test-Path -LiteralPath $terFileRaw))
	{
		# Fallback only if local doesn't exist
		$terFile = $terFileRaw
	}

	# Optional: log what we're using (helps confirm fix)
	Write_Log "OfficePath (raw): $OfficePath" "yellow"
	Write_Log "OfficePath (local resolved): $OfficePathLocalResolved" "yellow"
	Write_Log "Ter_Load (local preferred): $terFileLocal" "yellow"
	Write_Log "Ter_Load (selected): $terFile" "yellow"

	# ------------------------------------------------------------------------------------------------
	# Service install paths
	# ------------------------------------------------------------------------------------------------
	$installRoot = Join-Path $scriptFolder "Remove_Archive_Bit_Service"
	$svcScripts = Join-Path $installRoot "Scripts"
	$svcLogsDir = Join-Path $installRoot "Logs"
	$svcPS1 = Join-Path $svcScripts "Remove_Archive_Bit_Service.ps1"
	$svcConfigPath = Join-Path $installRoot "Remove_Archive_Bit_Config.json"
	$serviceName = "Remove_Archive_Bit"

	foreach ($d in @($installRoot, $svcScripts, $svcLogsDir))
	{
		if (-not (Test-Path -LiteralPath $d)) { New-Item -Path $d -ItemType Directory -Force | Out-Null }
	}

	# ------------------------------------------------------------------------------------------------
	# NSSM: Use already-downloaded NSSM only (NO downloads here)
	#   We will accept:
	#     - $script:FunctionResults['NssmExe'] (if your startup Download_NSSM set it)
	#     - $script:NssmExe (if set globally)
	#     - Default global path: C:\Tecnica_Systems\Alex_C.T\Tools\NSSM\nssm.exe
	# ------------------------------------------------------------------------------------------------
	$nssmExe = $null
	$nssmCandidates = @()

	try
	{
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('NssmExe') -and $script:FunctionResults['NssmExe'])
		{
			$nssmCandidates += [string]$script:FunctionResults['NssmExe']
		}
	}
	catch { }

	try
	{
		if ($script:NssmExe) { $nssmCandidates += [string]$script:NssmExe }
	}
	catch { }

	# Preferred default location from your Download_NSSM function
	$nssmCandidates += $script:NssmDefaultPath

	foreach ($p in $nssmCandidates)
	{
		if ([string]::IsNullOrWhiteSpace($p)) { continue }
		if (Test-Path -LiteralPath $p)
		{
			$nssmExe = $p
			break
		}
	}

	# ------------------------------------------------------------------------------------------------
	# Normalize lane map (we only use lane numbers; NEVER build \\POS### paths)
	# ------------------------------------------------------------------------------------------------
	$LaneMapNormalized = $null
	if ($script:FunctionResults.ContainsKey('LaneNumToMachineName') -and $script:FunctionResults['LaneNumToMachineName'])
	{
		$tmp = [ordered]@{ }

		foreach ($kv in $script:FunctionResults['LaneNumToMachineName'].GetEnumerator())
		{
			$k = [string]$kv.Key
			$v = [string]$kv.Value

			$l = $null
			$m = $null

			if ($k -match '^\d{1,3}$')
			{
				$l = ("{0:000}" -f ([int]$k))
				$m = $v
			}
			elseif ($k -match '(?i)(\d{3})$')
			{
				$l = $matches[1]
				$m = if (-not [string]::IsNullOrWhiteSpace($v)) { $v }
				else { $k }
			}

			if (-not $l) { continue }
			if ([string]::IsNullOrWhiteSpace($m)) { continue }
			if ($m -like '@SMSSERVER') { continue }

			if (-not $tmp.Contains($l)) { $tmp[$l] = $m }
		}

		if ($tmp.Count -gt 0) { $LaneMapNormalized = $tmp }
	}

	# ------------------------------------------------------------------------------------------------
	# UI: Run / Install / Uninstall
	# ------------------------------------------------------------------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Remove Archive Bit"
	$form.Size = New-Object System.Drawing.Size(480, 240)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.Tag = $null

	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Choose an action:`r`n- Run Now (one-time)`r`n- Install Service (always-on, Ter_Load driven)`r`n- Uninstall Service"
	$label.Location = New-Object System.Drawing.Point(20, 15)
	$label.Size = New-Object System.Drawing.Size(440, 60)
	$label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$form.Controls.Add($label)

	$btnRunNow = New-Object System.Windows.Forms.Button
	$btnRunNow.Text = "Run Now"
	$btnRunNow.Location = New-Object System.Drawing.Point(10, 95)
	$btnRunNow.Size = New-Object System.Drawing.Size(140, 45)
	$form.Controls.Add($btnRunNow)

	$btnService = New-Object System.Windows.Forms.Button
	$btnService.Text = "Install Service"
	$btnService.Location = New-Object System.Drawing.Point(165, 95)
	$btnService.Size = New-Object System.Drawing.Size(140, 45)
	$form.Controls.Add($btnService)

	$btnUninstall = New-Object System.Windows.Forms.Button
	$btnUninstall.Text = "Uninstall Service"
	$btnUninstall.Location = New-Object System.Drawing.Point(315, 95)
	$btnUninstall.Size = New-Object System.Drawing.Size(140, 45)
	$form.Controls.Add($btnUninstall)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(190, 160)
	$btnCancel.Size = New-Object System.Drawing.Size(90, 32)
	$form.Controls.Add($btnCancel)

	$btnRunNow.Add_Click({ $form.Tag = "run"; $form.DialogResult = 'OK'; $form.Close() })
	$btnService.Add_Click({ $form.Tag = "service"; $form.DialogResult = 'OK'; $form.Close() })
	$btnUninstall.Add_Click({ $form.Tag = "uninstall"; $form.DialogResult = 'OK'; $form.Close() })
	$btnCancel.Add_Click({ $form.Tag = "cancel"; $form.DialogResult = 'Cancel'; $form.Close() })

	$form.AcceptButton = $btnRunNow
	$form.CancelButton = $btnCancel

	[void]$form.ShowDialog()
	$selectedAction = [string]$form.Tag
	$form.Dispose()

	if ($selectedAction -eq "cancel" -or -not $selectedAction)
	{
		Write_Log "User cancelled Remove_Archive_Bit." "yellow"
		Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# UNINSTALL SERVICE
	# ------------------------------------------------------------------------------------------------
	if ($selectedAction -eq "uninstall")
	{
		$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
			[Security.Principal.WindowsBuiltInRole]::Administrator
		)
		if (-not $isAdmin)
		{
			[void][System.Windows.Forms.MessageBox]::Show(
				"Uninstalling the service requires Administrator rights.`r`n`r`nPlease re-run as Administrator and try again.",
				"Administrator Required",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Warning
			)
			Write_Log "ERROR: Uninstall requested without Administrator rights." "red"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		$existingSvc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
		if (-not $existingSvc)
		{
			Write_Log "Service '$serviceName' is not installed." "yellow"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		# Try native uninstall first (no NSSM needed)
		try { & sc.exe stop $serviceName *> $null }
		catch { }
		try { & sc.exe delete $serviceName *> $null }
		catch { }

		Wait_UI -Milliseconds 700
		if (-not (Get-Service -Name $serviceName -ErrorAction SilentlyContinue))
		{
			Write_Log "Service '$serviceName' removed via sc.exe (no NSSM required)." "green"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		# If sc.exe didn't fully remove it, fall back to NSSM - but ONLY if NSSM exists (NO downloads)
		if (-not $nssmExe -or -not (Test-Path -LiteralPath $nssmExe))
		{
			Write_Log "ERROR: sc.exe did not fully remove the service, and NSSM is missing." "red"
			Write_Log "Expected NSSM at: $script:NssmDefaultPath (or provided by Download_NSSM at startup)" "yellow"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		try { & $nssmExe stop $serviceName *> $null }
		catch { }
		try { & $nssmExe remove $serviceName confirm *> $null }
		catch { }

		Wait_UI -Milliseconds 700
		if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue)
		{
			Write_Log "ERROR: NSSM attempted removal but service still exists. Check Services.msc for '$serviceName'." "red"
			Write_Log "NSSM used: $nssmExe" "yellow"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		Write_Log "Service '$serviceName' uninstalled successfully." "green"
		Write_Log "NSSM used: $nssmExe" "green"
		Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# INSTALL SERVICE (LOCAL Office + LOCAL Ter_Load, REALTIME WATCHERS)
	# ------------------------------------------------------------------------------------------------
	if ($selectedAction -eq "service")
	{
		$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
			[Security.Principal.WindowsBuiltInRole]::Administrator
		)
		if (-not $isAdmin)
		{
			[void][System.Windows.Forms.MessageBox]::Show(
				"Installing the service requires Administrator rights.`r`n`r`nPlease re-run as Administrator and try again.",
				"Administrator Required",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Warning
			)
			Write_Log "ERROR: Service install requested without Administrator rights." "red"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		if (-not $storeNumber)
		{
			Write_Log "ERROR: Store number not present in context." "red"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}
		if (-not (Test-Path -LiteralPath $terFile))
		{
			Write_Log "ERROR: Ter_Load.sql could not be located at $terFile" "red"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		# NSSM must exist for install (NO downloads here)
		if (-not $nssmExe -or -not (Test-Path -LiteralPath $nssmExe))
		{
			Write_Log "ERROR: NSSM is missing. This function will not download it." "red"
			Write_Log "Expected NSSM at: $script:NssmDefaultPath (or provided by Download_NSSM at startup)" "yellow"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		# Service config
		$configObj = [ordered]@{
			StoreNumber = "$storeNumber"
			OfficePathRaw = "$OfficePath"
			OfficePathLocal = "$OfficePathLocalResolved"
			TerFileLocal = "$terFile"
			LogsDir	    = "$svcLogsDir"
			AutoPollSeconds = [int]$AutoPollSeconds
			LaneMap	    = $LaneMapNormalized
		}
		($configObj | ConvertTo-Json -Depth 7) | Set-Content -LiteralPath $svcConfigPath -Encoding UTF8 -Force
		Write_Log "Service config saved: $svcConfigPath" "green"

		# Service script (LOCAL paths only)
		$svcContent = @"
#requires -Version 5.1
`$ErrorActionPreference = 'Continue'

# --- Load config ---
`$configPath = '$($svcConfigPath.Replace("'", "''"))'
if (-not (Test-Path -LiteralPath `$configPath)) { exit 2 }

try { `$cfg = Get-Content -LiteralPath `$configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
catch { exit 3 }

`$StoreNumber     = [string]`$cfg.StoreNumber
`$OfficePath      = [string]`$cfg.OfficePathLocal
`$OfficePathRaw   = [string]`$cfg.OfficePathRaw
`$TerFile         = [string]`$cfg.TerFileLocal
`$LogsDir         = [string]`$cfg.LogsDir
`$AutoPollSeconds = [int]`$cfg.AutoPollSeconds
if (`$AutoPollSeconds -lt 1) { `$AutoPollSeconds = 3 }

# --- Resolve OfficePath locally again (defensive) ---
try {
	`$op = [string]`$OfficePath
	if ([string]::IsNullOrWhiteSpace(`$op)) { `$op = [string]`$OfficePathRaw }
	if (`$op -match '^(?i)\\\\([^\\]+)\\([^\\]+)(\\.*)?$') {
		`$uncHost = [string]`$matches[1]
		`$share = [string]`$matches[2]
		`$rest = [string]`$matches[3]

		`$isLocalHost = `$false
		if (`$uncHost -match '^(?i)localhost$') { `$isLocalHost = `$true }
		elseif (`$uncHost -eq `$env:COMPUTERNAME) { `$isLocalHost = `$true }
		elseif (`$uncHost -eq ("{0}.{1}" -f `$env:COMPUTERNAME, `$env:USERDNSDOMAIN)) { `$isLocalHost = `$true }

		if (`$isLocalHost) {
			`$rest2 = ""
			if (`$rest) { `$rest2 = `$rest.TrimStart('\') }

			`$sharePath = `$null
			try {
				`$w = Get-CimInstance -ClassName Win32_Share -Filter ("Name='{0}'" -f `$share.Replace("'", "''")) -ErrorAction Stop
				if (`$w -and `$w.Path) { `$sharePath = [string]`$w.Path }
			} catch { }

			if (-not [string]::IsNullOrWhiteSpace(`$sharePath)) {
				if (-not [string]::IsNullOrWhiteSpace(`$rest2)) {
					`$OfficePath = Join-Path `$sharePath `$rest2
				} else {
					`$OfficePath = `$sharePath
				}
			}
		}
	}
} catch { }

if (-not (Test-Path -LiteralPath `$LogsDir)) { New-Item -Path `$LogsDir -ItemType Directory -Force | Out-Null }
`$LogFile = Join-Path `$LogsDir 'Remove_Archive_Bit_Service.log'

try { Add-Content -LiteralPath `$LogFile -Value ("[{0}] Service starting..." -f (Get-Date)) -Encoding UTF8 } catch { }
try { Add-Content -LiteralPath `$LogFile -Value ("[{0}] OfficePath local: {1}" -f (Get-Date), `$OfficePath) -Encoding UTF8 } catch { }
try { Add-Content -LiteralPath `$LogFile -Value ("[{0}] Ter_Load local: {1}" -f (Get-Date), `$TerFile) -Encoding UTF8 } catch { }

`$watchers = New-Object System.Collections.ArrayList
`$terWatcher = `$null
`$needRebuild = `$true

while (`$true)
{
	# Touched item -> clear just its bits. Only a Ter_Load.sql change or a watcher overflow
	# (_e) rebuilds everything (every file event used to trigger a full attrib rescan).
	try {
		`$clearBits = [System.IO.FileAttributes]::Archive -bor [System.IO.FileAttributes]::ReadOnly
		foreach (`$evt in @(Get-Event -ErrorAction SilentlyContinue | Where-Object { [string]`$_.SourceIdentifier -like 'RAB_*' }))
		{
			`$sid = [string]`$evt.SourceIdentifier
			if (`$sid -like 'RAB_TER_WATCH*' -or `$sid -like '*_e') { `$needRebuild = `$true }
			else {
				try {
					`$p = [string]`$evt.SourceEventArgs.FullPath
					if (`$p -and (Test-Path -LiteralPath `$p)) {
						`$it = Get-Item -LiteralPath `$p -Force -ErrorAction Stop
						if (`$it.Attributes -band `$clearBits) {
							`$newAttr = [int](`$it.Attributes) -band (-bnot [int]`$clearBits)
							if (`$newAttr -eq 0) { `$newAttr = [int][System.IO.FileAttributes]::Normal }
							`$it.Attributes = [System.IO.FileAttributes]`$newAttr
						}
					}
				} catch { }
			}
			try { Remove-Event -EventIdentifier `$evt.EventIdentifier -ErrorAction SilentlyContinue } catch { }
		}
	} catch { }

	if (`$needRebuild)
	{
		`$needRebuild = `$false

		# Dispose old watchers + unregister all event ids we create
		try {
			foreach (`$w in @(`$watchers)) {
				try { `$w.EnableRaisingEvents = `$false } catch { }
				try {
					`$id = 'RAB_FSW_' + `$w.GetHashCode().ToString()
					Unregister-Event -SourceIdentifier `$id -ErrorAction SilentlyContinue
					Unregister-Event -SourceIdentifier (`$id + '_c') -ErrorAction SilentlyContinue
					Unregister-Event -SourceIdentifier (`$id + '_r') -ErrorAction SilentlyContinue
					Unregister-Event -SourceIdentifier (`$id + '_d') -ErrorAction SilentlyContinue
					Unregister-Event -SourceIdentifier (`$id + '_e') -ErrorAction SilentlyContinue
				} catch { }
				try { `$w.Dispose() } catch { }
			}
			`$watchers.Clear() | Out-Null
		} catch { }

		# Dispose Ter watcher
		try {
			if (`$terWatcher) {
				try { `$terWatcher.EnableRaisingEvents = `$false } catch { }
				try {
					Unregister-Event -SourceIdentifier 'RAB_TER_WATCH' -ErrorAction SilentlyContinue
					Unregister-Event -SourceIdentifier 'RAB_TER_WATCH_C' -ErrorAction SilentlyContinue
					Unregister-Event -SourceIdentifier 'RAB_TER_WATCH_R' -ErrorAction SilentlyContinue
				} catch { }
				try { `$terWatcher.Dispose() } catch { }
			}
		} catch { }

		# --- Build LOCAL lane roots from Ter_Load.sql (ROBUST: regex extract XF<store><lane>) ---
		`$laneRoots = New-Object System.Collections.Generic.List[string]
		`$lines = @()
		try { `$lines = Get-Content -LiteralPath `$TerFile -ErrorAction SilentlyContinue } catch { }

		`$picked = @()
		try { `$picked = `$lines | Where-Object { `$_ -match 'Terminal\s*0' } } catch { }

		`$text = ""
		if (`$picked -and `$picked.Count -gt 0) { `$text = (`$picked -join "`n") }
		else { `$text = (`$lines -join "`n") }

		if (-not [string]::IsNullOrWhiteSpace(`$text))
		{
			try {
				`$snEsc = [regex]::Escape(`$StoreNumber)
				`$rx = '(?i)XF' + `$snEsc + '\d{3}'

				foreach (`$m in [regex]::Matches(`$text, `$rx))
				{
					`$folder = ([string]`$m.Value).ToUpper()
					if (`$folder.EndsWith('900') -or `$folder.EndsWith('901')) { continue }

					`$local = Join-Path `$OfficePath `$folder
					if (-not `$laneRoots.Contains(`$local)) { [void]`$laneRoots.Add(`$local) }
				}
			} catch { }
		}

		# Fallback: enumerate local XF folders for this store
		if (`$laneRoots.Count -eq 0)
		{
			try {
				foreach (`$d in Get-ChildItem -LiteralPath `$OfficePath -Directory -ErrorAction SilentlyContinue)
				{
					if (`$d.Name -match ('(?i)^XF' + [regex]::Escape(`$StoreNumber) + '\d{3}$'))
					{
						if (`$d.Name -match ('(?i)^XF' + [regex]::Escape(`$StoreNumber) + '(900|901)$')) { continue }
						if (-not `$laneRoots.Contains(`$d.FullName)) { [void]`$laneRoots.Add(`$d.FullName) }
					}
				}
			} catch { }
		}

		# --- Server roots LOCAL ---
		`$serverRoots = New-Object System.Collections.Generic.List[string]
		`$sr900 = Join-Path `$OfficePath ("XF{0}900" -f `$StoreNumber)
		`$sr901 = Join-Path `$OfficePath ("XF{0}901" -f `$StoreNumber)
		[void]`$serverRoots.Add(`$sr900)
		[void]`$serverRoots.Add(`$sr901)

		`$allRoots = @()
		try { `$allRoots += `$laneRoots.ToArray() } catch { }
		try { `$allRoots += `$serverRoots.ToArray() } catch { }

		# --- Full cleanup ---
		try { Add-Content -LiteralPath `$LogFile -Value ("[{0}] Full cleanup starting..." -f (Get-Date)) -Encoding UTF8 } catch { }

		foreach (`$root in `$allRoots)
		{
			if ([string]::IsNullOrWhiteSpace(`$root)) { continue }
			`$r = `$root.Trim().Trim('"')
			if (-not (Test-Path -LiteralPath `$r)) { continue }
			try {
				`$cmd = 'attrib -a -r "{0}\*" /S /D' -f `$r
				& `$env:ComSpec /c `$cmd 2>`$null | Out-Null
			} catch { }
		}

		try {
			Add-Content -LiteralPath `$LogFile -Value ("[{0}] Full cleanup done. LaneRoots={1}, ServerRoots={2}. Building watchers..." -f (Get-Date), `$laneRoots.Count, `$serverRoots.Count) -Encoding UTF8
		} catch { }

		# --- Build filesystem watchers ---
		foreach (`$root in `$allRoots)
		{
			if ([string]::IsNullOrWhiteSpace(`$root)) { continue }
			`$r = `$root.Trim().Trim('"')
			if (-not (Test-Path -LiteralPath `$r)) { continue }

			try {
				`$fsw = New-Object System.IO.FileSystemWatcher
				`$fsw.Path = `$r
				`$fsw.Filter = '*'
				`$fsw.IncludeSubdirectories = `$true
				`$fsw.NotifyFilter = [System.IO.NotifyFilters]'FileName, DirectoryName, LastWrite, Attributes'
				`$fsw.InternalBufferSize = 65536

				`$id = 'RAB_FSW_' + `$fsw.GetHashCode().ToString()
				Register-ObjectEvent -InputObject `$fsw -EventName Created -SourceIdentifier `$id | Out-Null
				Register-ObjectEvent -InputObject `$fsw -EventName Changed -SourceIdentifier (`$id + '_c') | Out-Null
				Register-ObjectEvent -InputObject `$fsw -EventName Renamed -SourceIdentifier (`$id + '_r') | Out-Null
				Register-ObjectEvent -InputObject `$fsw -EventName Deleted -SourceIdentifier (`$id + '_d') | Out-Null
				Register-ObjectEvent -InputObject `$fsw -EventName Error -SourceIdentifier (`$id + '_e') | Out-Null

				`$fsw.EnableRaisingEvents = `$true
				[void]`$watchers.Add(`$fsw)
			} catch { }
		}

		# --- Watch Ter_Load.sql itself ---
		try {
			`$terDir = Split-Path -Parent `$TerFile
			if (Test-Path -LiteralPath `$terDir)
			{
				`$terWatcher = New-Object System.IO.FileSystemWatcher
				`$terWatcher.Path = `$terDir
				`$terWatcher.Filter = (Split-Path -Leaf `$TerFile)
				`$terWatcher.IncludeSubdirectories = `$false
				`$terWatcher.NotifyFilter = [System.IO.NotifyFilters]'FileName, LastWrite, Size'

				Register-ObjectEvent -InputObject `$terWatcher -EventName Changed -SourceIdentifier 'RAB_TER_WATCH' | Out-Null
				Register-ObjectEvent -InputObject `$terWatcher -EventName Created -SourceIdentifier 'RAB_TER_WATCH_C' | Out-Null
				Register-ObjectEvent -InputObject `$terWatcher -EventName Renamed -SourceIdentifier 'RAB_TER_WATCH_R' | Out-Null

				`$terWatcher.EnableRaisingEvents = `$true
			}
		} catch { }

		try { Add-Content -LiteralPath `$LogFile -Value ("[{0}] Watchers active: {1}" -f (Get-Date), `$watchers.Count) -Encoding UTF8 } catch { }
	}

	# Wake on the next event (or after the poll interval)
	`$null = Wait-Event -Timeout `$AutoPollSeconds
}
"@

		$svcContent | Set-Content -LiteralPath $svcPS1 -Encoding UTF8 -Force
		Write_Log "Service script saved: $svcPS1" "green"

		$serviceStdOut = Join-Path $svcLogsDir "Remove_Archive_Bit_Service_stdout.log"
		$serviceStdErr = Join-Path $svcLogsDir "Remove_Archive_Bit_Service_stderr.log"

		$existingSvc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
		if ($existingSvc)
		{
			try { & $nssmExe stop $serviceName *> $null }
			catch { }
			try { & $nssmExe remove $serviceName confirm *> $null }
			catch { }
		}

		& $nssmExe install $serviceName "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
				   -NoProfile -ExecutionPolicy Bypass -File "$svcPS1" *> $null

		& $nssmExe set $serviceName AppDirectory "$svcScripts" *> $null
		& $nssmExe set $serviceName AppStdout "$serviceStdOut" *> $null
		& $nssmExe set $serviceName AppStderr "$serviceStdErr" *> $null
		& $nssmExe set $serviceName Start SERVICE_AUTO_START *> $null

		# LOCAL filesystem permissions; avoids UNC auth failures
		& $nssmExe set $serviceName ObjectName "LocalSystem" "" *> $null

		& $nssmExe set $serviceName DisplayName "Remove Archive Bit (LOCAL XF Watcher)" *> $null
		& $nssmExe set $serviceName Description "Clears archive/read-only bits in LOCAL XF folders; watches Ter_Load.sql and filesystem changes." *> $null

		& $nssmExe start $serviceName *> $null

		Write_Log "Service '$serviceName' installed and started as: LocalSystem" "green"
		Write_Log "Logs: $svcLogsDir" "green"
		Write_Log "NSSM used: $nssmExe" "green"

		Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# RUN NOW (one-time)
	# ------------------------------------------------------------------------------------------------
	if (-not $storeNumber)
	{
		Write_Log "ERROR: Store number not present in context." "red"
		Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
		return
	}
	if (-not (Test-Path -LiteralPath $terFile))
	{
		Write_Log "ERROR: Ter_Load.sql could not be located at $terFile" "red"
		Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
		return
	}

	# --- Run Now: ALWAYS LOCAL paths (never \\POS###)
	$didLaneWork = $false

	if ($LaneMapNormalized)
	{
		foreach ($laneNum in ($LaneMapNormalized.Keys | Sort-Object { [int]$_ }))
		{
			$localLanePath = Join-Path $OfficePathLocalResolved ("XF{0}{1}" -f $storeNumber, $laneNum)
			$lp = $localLanePath.Trim().Trim('"')

			if (Test-Path -LiteralPath $lp)
			{
				Write_Log "Refreshing attributes (recursive) in $lp..." "green"
				$cmd = 'attrib -a -r "{0}\*" /S /D' -f $lp
				try { & $env:ComSpec /c $cmd 2>$null | Out-Null }
				catch { Write_Log "ERROR: Failed attrib on: $lp" "red" }
				$didLaneWork = $true
			}
			else
			{
				Write_Log "Local lane path not found (skip): $lp" "yellow"
			}
		}
	}
	else
	{
		Write_Log "Lane map not found/usable. Falling back to Ter_Load.sql parsing (LOCAL rebuild)..." "yellow"
	}

	# Fallback: robust Ter_Load scan for XF<store><lane>
	if (-not $didLaneWork)
	{
		$lanePaths = New-Object System.Collections.Generic.List[string]
		$lines = @()
		try { $lines = Get-Content -LiteralPath $terFile -ErrorAction SilentlyContinue }
		catch { }

		$picked = @()
		try { $picked = $lines | Where-Object { $_ -match 'Terminal\s*0' } }
		catch { }

		$text = ""
		if ($picked -and $picked.Count -gt 0) { $text = ($picked -join "`n") }
		else { $text = ($lines -join "`n") }

		if (-not [string]::IsNullOrWhiteSpace($text))
		{
			try
			{
				$snEsc = [regex]::Escape($storeNumber)
				$rx = '(?i)XF' + $snEsc + '\d{3}'

				foreach ($m in [regex]::Matches($text, $rx))
				{
					$folder = ([string]$m.Value).ToUpper()
					if ($folder.EndsWith('900') -or $folder.EndsWith('901')) { continue }

					$local = Join-Path $OfficePathLocalResolved $folder
					if (-not $lanePaths.Contains($local)) { [void]$lanePaths.Add($local) }
				}
			}
			catch { }
		}

		# fallback: enumerate folders if still empty
		if ($lanePaths.Count -eq 0)
		{
			try
			{
				foreach ($d in Get-ChildItem -LiteralPath $OfficePathLocalResolved -Directory -ErrorAction SilentlyContinue)
				{
					if ($d.Name -match ('(?i)^XF' + [regex]::Escape($storeNumber) + '\d{3}$'))
					{
						if ($d.Name -match ('(?i)^XF' + [regex]::Escape($storeNumber) + '(900|901)$')) { continue }
						if (-not $lanePaths.Contains($d.FullName)) { [void]$lanePaths.Add($d.FullName) }
					}
				}
			}
			catch { }
		}

		if ($lanePaths.Count -gt 0)
		{
			foreach ($p in $lanePaths)
			{
				$pp = $p.Trim().Trim('"')
				if (Test-Path -LiteralPath $pp)
				{
					Write_Log "Refreshing attributes (recursive) in $pp..." "green"
					$cmd = 'attrib -a -r "{0}\*" /S /D' -f $pp
					try { & $env:ComSpec /c $cmd 2>$null | Out-Null }
					catch { Write_Log "ERROR: Failed attrib on: $pp" "red" }
				}
				else
				{
					Write_Log "Path not found (skip): $pp" "yellow"
				}
			}
		}
		else
		{
			Write_Log "No LOCAL lane paths could be derived from Ter_Load.sql and no XF lane folders found locally." "yellow"
		}
	}

	# Server roots LOCAL
	foreach ($suffix in 900, 901)
	{
		$serverPath = Join-Path $OfficePathLocalResolved ("XF{0}{1}" -f $storeNumber, $suffix)
		$sp = $serverPath.Trim().Trim('"')
		if (Test-Path -LiteralPath $sp)
		{
			Write_Log "Refreshing attributes (recursive) in $sp" "green"
			$cmd = 'attrib -a -r "{0}\*" /S /D' -f $sp
			try { & $env:ComSpec /c $cmd 2>$null | Out-Null }
			catch { Write_Log "ERROR: Failed attrib on: $sp" "red" }
		}
		else
		{
			Write_Log "Server path not found: $sp" "yellow"
		}
	}

	Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
}

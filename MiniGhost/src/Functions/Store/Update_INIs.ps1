# ===================================================================================================
# 									FUNCTION: Update_INIs
# ---------------------------------------------------------------------------------------------------
# Description:
# Updates ALL required INI files with the new store number.
# Supports both 3-digit and 4-digit store numbers.
# NOTE: No backups are created.
#
# Always-best-effort behavior:
# - Does NOT prompt
# - If MachineName missing -> uses $env:COMPUTERNAME
# - If terminal digits cannot be extracted -> still updates other INIs; skips TER insertion where unsafe
# - If old store cannot be detected -> still forces STORE= and REDIR lines; skips token-replace mapping
# - Startup.ini / Server.ini / SMSStart.ini / INFO_*901_WIN.ini preserve ORIGINAL encoding + ORIGINAL newline style
# - Updates Server.ini when present (same folder as Startup.ini first, then common roots)
#
# IMPORTANT SAFETY RULE:
# - If MachineName contains "SERVER" (anywhere, case-insensitive), we MUST NOT change terminal numbers.
#   That means:
#     * Do NOT derive terminal from name
#     * Do NOT rewrite TER= lines in Startup.ini / Server.ini / SMSStart.ini
#     * DO still update STORE= and REDIRMAIL/REDIRMSG to <STORE>901
#
# SmsHttps.INI behavior:
# - SERVER moniker:
#     * KEEP multiple processor entries
#     * ONLY update the store prefix in:
#         - processor key (e.g., 0231901 -> 0242901)
#         - REDIRMAIL/REDIRMSG values (e.g., 0231901 -> 0242901)
#     * Clear LicenseGUID (always)
# - NON-SERVER moniker:
#     * (Safe default here) we DO NOT touch processors (to avoid accidental changes)
#     * Clear LicenseGUID (always)
#
# NEW:
# - SMSStart.ini may contain PARAMETERS=/ini=Startup902 (or Startup902.ini).
#   We resolve that referenced INI and update it too (Store/Redir/token-replace).
#   We DO NOT force TER changes in the referenced Startup*.ini (we generally leave TER as-is).
# ===================================================================================================

function Update_INIs
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidatePattern('^(?!0+$)\d{3,4}$')]
		[string]$newStoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[string]$OldStoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineName,
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath,
		[Parameter(Mandatory = $false)]
		[string]$GlobalSmsStartIniPath,
		[Parameter(Mandatory = $false)]
		[string]$WinIniPath,
		[Parameter(Mandatory = $false)]
		[string]$SmsHttpsIniPath,
		[Parameter(Mandatory = $false)]
		[switch]$CreateSmsHttpsIfMissing
	)
	
	$success = $true
	
	# ------------------------------------------------------------------------------------------------
	# Prefer already-built variables for paths (no hardcoded probing)
	# ------------------------------------------------------------------------------------------------
	try
	{
		if ([string]::IsNullOrWhiteSpace($StartupIniPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name StartupIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $StartupIniPath = [string]$v.Value; break }
			}
		}
		if ([string]::IsNullOrWhiteSpace($GlobalSmsStartIniPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name GlobalSmsStartIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $GlobalSmsStartIniPath = [string]$v.Value; break }
			}
		}
		if ([string]::IsNullOrWhiteSpace($WinIniPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name WinIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $WinIniPath = [string]$v.Value; break }
			}
		}
		if ([string]::IsNullOrWhiteSpace($SmsHttpsIniPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name SmsHttpsIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $SmsHttpsIniPath = [string]$v.Value; break }
			}
		}
	}
	catch { }
	
	# Optional: reuse already-built core paths if present (does not rebuild new ones)
	$BasePath = $null
	$DbsPath = $null
	$SystemIniVar = $null
	try
	{
		foreach ($sc in @('Script', 'Global'))
		{
			if (-not $BasePath)
			{
				$v = Get-Variable -Name BasePath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and $v.Value) { $BasePath = [string]$v.Value }
			}
			if (-not $DbsPath)
			{
				$v = Get-Variable -Name DbsPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and $v.Value) { $DbsPath = [string]$v.Value }
			}
			if (-not $SystemIniVar)
			{
				$v = Get-Variable -Name SystemIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and $v.Value) { $SystemIniVar = [string]$v.Value }
			}
		}
	}
	catch { }
	
	# If caller didn't supply MachineName, use current machine name (no prompt, no rename)
	if ([string]::IsNullOrWhiteSpace($MachineName))
	{
		$MachineName = $env:COMPUTERNAME
	}
	
	# ------------------------------------------------------------------------------------------------
	# Normalize NEW store (preserve user width for STORE=/REDIR lines)
	# ------------------------------------------------------------------------------------------------
	$newStoreTrim = ($newStoreNumber + "").Trim()
	$newStoreInt = [int]$newStoreTrim
	$newStore3 = $newStoreInt.ToString("D3")
	$newStore4 = $newStoreInt.ToString("D4")
	
	# ------------------------------------------------------------------------------------------------
	# Normalize machine name (strip UNC/path + domain) + ALWAYS uppercase
	# ------------------------------------------------------------------------------------------------
	$mn = ($MachineName + "").Trim()
	$mn = $mn -replace '^[\\\/]+', ''
	if ($mn -match '[\\\/]') { $mn = ($mn -split '[\\\/]')[0] }
	if ($mn -match '\.') { $mn = ($mn -split '\.')[0] }
	$mn = $mn.Trim().ToUpper()
	
	# ------------------------------------------------------------------------------------------------
	# SERVER MONIKER RULE: if name contains SERVER anywhere -> NEVER change terminal numbers
	# ------------------------------------------------------------------------------------------------
	$isServerMoniker = $false
	if ($mn -match '(?i)SERVER') { $isServerMoniker = $true }
	elseif ($script:NodeRole -and ($script:NodeRole -match '^(?i)(StoreServer|HostServer)$')) { $isServerMoniker = $true }
	
	# ------------------------------------------------------------------------------------------------
	# Derive terminal ONLY for NON-server moniker machines
	# ------------------------------------------------------------------------------------------------
	$newTerminal = $null
	if (-not $isServerMoniker -and ($mn -match '(\d{1,3})$'))
	{
		$newTerminal = ([int]$Matches[1]).ToString("D3")
		if ($newTerminal -eq "000") { $newTerminal = $null }
	}
	
	# ===============================================================================================
	# 1) Resolve + Update Startup.ini (required)  (no hardcoded probing; must already be set)
	# ===============================================================================================
	if ([string]::IsNullOrWhiteSpace($StartupIniPath) -and $BasePath)
	{
		try
		{
			$p = Join-Path $BasePath "Startup.ini"
			if (Test-Path $p) { $StartupIniPath = $p }
		}
		catch { }
	}
	
	if ([string]::IsNullOrWhiteSpace($StartupIniPath) -or -not (Test-Path $StartupIniPath))
	{
		Write-Host "startup.ini not found. Ensure `$StartupIniPath is already set (or pass -StartupIniPath)." -ForegroundColor Red
		return $false
	}
	
	$oldStoreDetected = $null
	$startupDir = $null
	try { $startupDir = Split-Path -Path $StartupIniPath -Parent }
	catch { $startupDir = $null }
	
	try
	{
		$bytes = [System.IO.File]::ReadAllBytes($StartupIniPath)
		
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
		
		# Detect OLD store number (prefer STORE=)
		foreach ($l in $startupLines)
		{
			if ($l -match '^\s*STORE\s*=\s*(\d{3,4})\s*$') { $oldStoreDetected = $matches[1]; break }
		}
		if (-not $oldStoreDetected)
		{
			foreach ($l in $startupLines)
			{
				if ($l -match '^\s*SERVERNAME\s*=\s*(\d{3,4})[A-Za-z]') { $oldStoreDetected = $matches[1]; break }
			}
		}
		if (-not $oldStoreDetected)
		{
			foreach ($l in $startupLines)
			{
				if ($l -match '\\\\(\d{3,4})[A-Za-z]') { $oldStoreDetected = $matches[1]; break }
			}
		}
		
		# If user provided OldStoreNumber, override detection (if valid)
		if (-not [string]::IsNullOrWhiteSpace($OldStoreNumber))
		{
			$os = $OldStoreNumber.Trim()
			if ($os -match '^(?!0+$)\d{3,4}$') { $oldStoreDetected = $os }
		}
		
		$doTokenReplace = $true
		if (-not $oldStoreDetected) { $doTokenReplace = $false }
		
		$oldTokens = @()
		$newTokens = @()
		if ($doTokenReplace)
		{
			$oldStoreInt = [int]$oldStoreDetected
			$oldTokens = @($oldStoreInt.ToString("D4"), $oldStoreInt.ToString("D3"))
			$newTokens = @($newStore4, $newStore3)
		}
		
		for ($i = 0; $i -lt $startupLines.Count; $i++)
		{
			$line = $startupLines[$i]
			
			$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
			$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
			
			# TER only when NON-server and we have a derived terminal
			if (-not $isServerMoniker -and $newTerminal)
			{
				$line = $line -replace '^\s*TER\s*=\s*\d{1,4}\s*$', ("TER=" + $newTerminal)
			}
			
			# Always uppercase SERVERNAME value if present
			if ($line -match '^(?<ws>\s*)(?i:SERVERNAME)\s*=\s*(?<val>.*)\s*$')
			{
				$ws = $matches['ws']
				$val = ($matches['val'] + "").Trim()
				if ($val.Length -gt 0) { $line = $ws + "SERVERNAME=" + $val.ToUpper() }
			}
			
			# Token replace mapping (SAFE: do not rewrite terminal suffixes like SERVER001)
			if ($doTokenReplace)
			{
				for ($t = 0; $t -lt $oldTokens.Count; $t++)
				{
					$oldEsc = [regex]::Escape($oldTokens[$t])
					$newTok = $newTokens[$t]
					
					$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
					$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
					$line = $line -replace ("(?<![\dA-Za-z])" + $oldEsc + "(?!\d)"), $newTok
				}
			}
			
			$startupLines[$i] = $line
		}
		
		[System.IO.File]::WriteAllText($StartupIniPath, ($startupLines -join $nlStartup), $encStartup)
		Write-Host "Updated startup.ini" -ForegroundColor Green
		
		if (-not $oldStoreDetected) { $oldStoreDetected = $newStoreTrim }
	}
	catch
	{
		$success = $false
		Write-Host "Failed updating startup.ini: $($_.Exception.Message)" -ForegroundColor Red
		if (-not $oldStoreDetected) { $oldStoreDetected = $newStoreTrim }
	}
	
	# ===============================================================================================
	# 1B) Update Server.ini (optional) (relative to Startup.ini / BasePath only)
	# ===============================================================================================
	$serverIniPath = $null
	try
	{
		if ($startupDir)
		{
			$p = Join-Path $startupDir "Server.ini"
			if (Test-Path $p) { $serverIniPath = $p }
		}
		if (-not $serverIniPath -and $BasePath)
		{
			$p = Join-Path $BasePath "Server.ini"
			if (Test-Path $p) { $serverIniPath = $p }
		}
	}
	catch { $serverIniPath = $null }
	
	if (-not [string]::IsNullOrWhiteSpace($serverIniPath) -and (Test-Path $serverIniPath))
	{
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($serverIniPath)
			
			$encServer = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{ $encServer = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{ $encServer = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{ $encServer = [System.Text.Encoding]::BigEndianUnicode }
			else
			{ $encServer = [System.Text.Encoding]::Default }
			
			$textServer = $encServer.GetString($bytes)
			$nlServer = "`r`n"
			if ($textServer -notmatch "`r`n" -and $textServer -match "`n") { $nlServer = "`n" }
			
			$serverLines = $textServer -split "\r?\n", -1
			
			$doTokenReplaceS = $false
			$oldTokensS = @()
			$newTokensS = @()
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$doTokenReplaceS = $true
				$osi = [int]$oldStoreDetected
				$oldTokensS = @($osi.ToString("D4"), $osi.ToString("D3"))
				$newTokensS = @($newStore4, $newStore3)
			}
			
			for ($i = 0; $i -lt $serverLines.Count; $i++)
			{
				$line = $serverLines[$i]
				
				$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
				$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
				
				if (-not $isServerMoniker -and $newTerminal)
				{
					$line = $line -replace '^\s*TER\s*=\s*\d{1,4}\s*$', ("TER=" + $newTerminal)
				}
				
				if ($line -match '^(?<ws>\s*)(?i:SERVERNAME)\s*=\s*(?<val>.*)\s*$')
				{
					$ws = $matches['ws']
					$val = ($matches['val'] + "").Trim()
					if ($val.Length -gt 0) { $line = $ws + "SERVERNAME=" + $val.ToUpper() }
				}
				
				if ($doTokenReplaceS)
				{
					for ($t = 0; $t -lt $oldTokensS.Count; $t++)
					{
						$oldEsc = [regex]::Escape($oldTokensS[$t])
						$newTok = $newTokensS[$t]
						
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
						$line = $line -replace ("(?<![\dA-Za-z])" + $oldEsc + "(?!\d)"), $newTok
					}
				}
				
				$serverLines[$i] = $line
			}
			
			[System.IO.File]::WriteAllText($serverIniPath, ($serverLines -join $nlServer), $encServer)
			Write-Host "Updated Server.ini" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating Server.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	
	# ===============================================================================================
	# 1C) Update Office\System.ini (optional) using existing $SystemIniPath if available
	# ===============================================================================================
	$systemIniPath = $null
	if (-not [string]::IsNullOrWhiteSpace($SystemIniVar) -and (Test-Path $SystemIniVar))
	{
		$systemIniPath = $SystemIniVar
	}
	elseif ($startupDir)
	{
		$p = Join-Path $startupDir "Office\System.ini"
		if (Test-Path $p) { $systemIniPath = $p }
	}
	
	if (-not [string]::IsNullOrWhiteSpace($systemIniPath) -and (Test-Path $systemIniPath))
	{
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($systemIniPath)
			
			$encSys = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{ $encSys = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{ $encSys = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{ $encSys = [System.Text.Encoding]::BigEndianUnicode }
			else
			{ $encSys = [System.Text.Encoding]::Default }
			
			$textSys = $encSys.GetString($bytes)
			
			$nlSys = "`r`n"
			if ($textSys -notmatch "`r`n" -and $textSys -match "`n") { $nlSys = "`n" }
			
			$linesSys = $textSys -split "\r?\n", -1
			
			$newStoreIsOne = $false
			try { if ([int]$newStoreTrim -eq 1) { $newStoreIsOne = $true } }
			catch { $newStoreIsOne = $false }
			
			$oldTokensSys = @()
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$osi = [int]$oldStoreDetected
				$oldTokensSys = @($osi.ToString("D4"), $osi.ToString("D3"))
			}
			
			$inSmsSection = $false
			$changedSys = $false
			
			for ($i = 0; $i -lt $linesSys.Count; $i++)
			{
				$line = $linesSys[$i]
				
				if ($line -match '^\s*\[(?<sec>.+?)\]\s*$')
				{
					$sec = $matches['sec'].Trim()
					$inSmsSection = ($sec -ieq 'SMS')
					continue
				}
				
				if (-not $inSmsSection) { continue }
				
				if ($line -match '^(?<ws>\s*)(?i:Name)\s*=\s*(?<val>.*)\s*$')
				{
					$ws = $matches['ws']
					$val = $matches['val']
					$origVal = $val
					
					$val2 = ($val + "").Trim()
					
					if ($newStoreIsOne)
					{
						$val2 = [regex]::Replace($val2, '\s+(?:0001|001|01|1)\s*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
						foreach ($tok in @('0001', '001', '01', '1'))
						{
							$val2 = [regex]::Replace($val2, '\s+' + [regex]::Escape($tok) + '\s*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
						}
						$val2 = $val2.Trim()
					}
					else
					{
						$replaced = $false
						foreach ($tok in $oldTokensSys)
						{
							if (-not [string]::IsNullOrWhiteSpace($tok) -and ($val2 -match ("(?<!\d)" + [regex]::Escape($tok) + "(?!\d)")))
							{
								$val2 = [regex]::Replace($val2, "(?<!\d)" + [regex]::Escape($tok) + "(?!\d)", $newStoreTrim)
								$replaced = $true
							}
						}
						
						if (-not $replaced -and ($val2 -match '\s+\d{3,4}\s*$'))
						{
							$val2 = [regex]::Replace($val2, '\s+\d{3,4}\s*$', (" " + $newStoreTrim))
							$replaced = $true
						}
						
						if (-not $replaced)
						{
							$val2 = ($val2.TrimEnd() + " " + $newStoreTrim).Trim()
						}
					}
					
					if ($val2 -ne $origVal)
					{
						$linesSys[$i] = ($ws + "Name=" + $val2)
						$changedSys = $true
					}
					
					break
				}
			}
			
			if ($changedSys)
			{
				[System.IO.File]::WriteAllText($systemIniPath, ($linesSys -join $nlSys), $encSys)
				Write-Host "Updated Office\System.ini (SMS Name)" -ForegroundColor Green
			}
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating Office\System.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	
	# ===============================================================================================
	# 2) Update Global SMSStart.ini (optional)
	#     + NEW: detect PARAMETERS=/ini=Startup### and update referenced INI too
	#     + NEW: update [TIMESYNC] SERVER=... (store prefix) + ENABLE=1 (insert if missing)
	# ===============================================================================================
	if (-not [string]::IsNullOrWhiteSpace($GlobalSmsStartIniPath) -and (Test-Path $GlobalSmsStartIniPath))
	{
		$referencedIniNames = @()
		
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($GlobalSmsStartIniPath)
			
			$encSmsStart = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{ $encSmsStart = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{ $encSmsStart = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{ $encSmsStart = [System.Text.Encoding]::BigEndianUnicode }
			else
			{ $encSmsStart = [System.Text.Encoding]::Default }
			
			$textSmsStart = $encSmsStart.GetString($bytes)
			$nlSmsStart = "`r`n"
			if ($textSmsStart -notmatch "`r`n" -and $textSmsStart -match "`n") { $nlSmsStart = "`n" }
			
			$globalLines = $textSmsStart -split "\r?\n", -1
			
			$doTokenReplace = $false
			$oldTokens = @()
			$newTokens = @()
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$doTokenReplace = $true
				$osi = [int]$oldStoreDetected
				$oldTokens = @($osi.ToString("D4"), $osi.ToString("D3"))
				$newTokens = @($newStore4, $newStore3)
			}
			
			$out = @()
			$inSmsStartSection = $false
			$terFound = $false
			$terLine = $null
			if (-not $isServerMoniker -and $newTerminal) { $terLine = "TER=$newTerminal" }
			
			$inTimeSyncSection = $false
			$tsEnableFound = $false
			
			foreach ($raw in $globalLines)
			{
				$line = $raw
				
				# Detect PARAMETERS=/ini=Startup902 (or Startup902.ini)
				if ($line -match '^\s*(?i:PARAMETERS)\s*=\s*(?<p>.*)\s*$')
				{
					$p = $matches['p']
					if (-not [string]::IsNullOrWhiteSpace($p))
					{
						if ($p -match '(?i)(?:^|\s)/ini\s*=\s*(?<ini>[^\s"]+)')
						{
							$iniName = $matches['ini'].Trim().Trim('"').Trim("'")
							if (-not [string]::IsNullOrWhiteSpace($iniName))
							{
								$referencedIniNames += $iniName
							}
						}
					}
				}
				
				if ($line -match '^\s*\[(.+?)\]\s*$')
				{
					# Leaving [SMSSTART] -> ensure TER inserted if needed
					if ($inSmsStartSection -and -not $terFound -and $terLine)
					{
						$out += $terLine
						$terFound = $true
					}
					
					# Leaving [TIMESYNC] -> ensure ENABLE=1 if missing
					if ($inTimeSyncSection -and -not $tsEnableFound)
					{
						$out += "ENABLE=1"
						$tsEnableFound = $true
					}
					
					$sectionName = $matches[1].Trim()
					$inSmsStartSection = ($sectionName -ieq 'SMSSTART')
					if ($inSmsStartSection) { $terFound = $false }
					
					$inTimeSyncSection = ($sectionName -ieq 'TIMESYNC')
					if ($inTimeSyncSection) { $tsEnableFound = $false }
					
					$out += $line
					continue
				}
				
				if ($inSmsStartSection)
				{
					$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
					$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
					
					# TER only when allowed (non-server)
					if ($terLine -and ($line -match '^\s*(?i:TER)\s*='))
					{
						$line = $terLine
						$terFound = $true
					}
					
					if ($doTokenReplace)
					{
						for ($t = 0; $t -lt $oldTokens.Count; $t++)
						{
							$oldEsc = [regex]::Escape($oldTokens[$t])
							$newTok = $newTokens[$t]
							$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
							$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
							$line = $line -replace ("(?<![\dA-Za-z])" + $oldEsc + "(?!\d)"), $newTok
						}
					}
				}
				elseif ($inTimeSyncSection)
				{
					# Force ENABLE=1 (and mark found)
					if ($line -match '^\s*(?i:ENABLE)\s*=')
					{
						$line = "ENABLE=1"
						$tsEnableFound = $true
					}
					
					# Update SERVER= store prefix (and force uppercase)
					if ($line -match '^\s*(?i:SERVER)\s*=\s*(?<sv>.*)\s*$')
					{
						$sv = ($matches['sv'] + "").Trim()
						
						if ($sv -match '^(?<st>\d{3,4})(?<rest>.*)$')
						{
							$st = $matches['st']
							$rest = $matches['rest']
							$newSt = ([int]$newStoreTrim).ToString(("D{0}" -f $st.Length))
							$sv = $newSt + $rest
						}
						else
						{
							if ($doTokenReplace)
							{
								for ($t = 0; $t -lt $oldTokens.Count; $t++)
								{
									$oldEsc = [regex]::Escape($oldTokens[$t])
									$newTok = $newTokens[$t]
									$sv = $sv -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
									$sv = $sv -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
									$sv = $sv -replace ("(?<![\dA-Za-z])" + $oldEsc + "(?!\d)"), $newTok
								}
							}
						}
						
						$line = "SERVER=" + $sv.ToUpper()
					}
				}
				
				$out += $line
			}
			
			# If file ended while still in sections
			if ($inSmsStartSection -and -not $terFound -and $terLine)
			{
				$out += $terLine
			}
			if ($inTimeSyncSection -and -not $tsEnableFound)
			{
				$out += "ENABLE=1"
			}
			
			[System.IO.File]::WriteAllText($GlobalSmsStartIniPath, ($out -join $nlSmsStart), $encSmsStart)
			Write-Host "Updated SMSStart.ini (SMSSTART + TIMESYNC)" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating SMSStart.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
		
		# ---- Update referenced Startup*.ini files (PARAMETERS=/ini=...) ----
		if ($referencedIniNames -and $referencedIniNames.Count -gt 0)
		{
			$smsStartDir = $null
			try { $smsStartDir = Split-Path -Path $GlobalSmsStartIniPath -Parent }
			catch { $smsStartDir = $null }
			
			foreach ($iniRef in ($referencedIniNames | Select-Object -Unique))
			{
				try
				{
					$iniTarget = ($iniRef + "").Trim()
					if ([string]::IsNullOrWhiteSpace($iniTarget)) { continue }
					if ($iniTarget -notmatch '\.ini$') { $iniTarget = ($iniTarget + ".ini") }
					
					$candidatePaths = @()
					
					if ($iniTarget -match '^[A-Za-z]:\\' -or $iniTarget -match '^\\\\')
					{
						$candidatePaths += $iniTarget
					}
					else
					{
						if ($smsStartDir) { $candidatePaths += (Join-Path $smsStartDir $iniTarget) }
						if ($BasePath) { $candidatePaths += (Join-Path $BasePath    $iniTarget) }
						if ($startupDir) { $candidatePaths += (Join-Path $startupDir  $iniTarget) }
					}
					
					$resolved = $null
					foreach ($cp in $candidatePaths)
					{
						if (-not [string]::IsNullOrWhiteSpace($cp) -and (Test-Path $cp))
						{
							$resolved = $cp
							break
						}
					}
					
					if (-not $resolved) { continue }
					
					$b = [System.IO.File]::ReadAllBytes($resolved)
					
					$encRef = $null
					if ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
					{ $encRef = New-Object System.Text.UTF8Encoding($true) }
					elseif ($b.Length -ge 2 -and $b[0] -eq 0xFF -and $b[1] -eq 0xFE)
					{ $encRef = [System.Text.Encoding]::Unicode }
					elseif ($b.Length -ge 2 -and $b[0] -eq 0xFE -and $b[1] -eq 0xFF)
					{ $encRef = [System.Text.Encoding]::BigEndianUnicode }
					else
					{ $encRef = [System.Text.Encoding]::Default }
					
					$txt = $encRef.GetString($b)
					$nlRef = "`r`n"
					if ($txt -notmatch "`r`n" -and $txt -match "`n") { $nlRef = "`n" }
					
					$lines = $txt -split "\r?\n", -1
					
					$doTR = $false
					$ot = @()
					$nt = @()
					if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
					{
						$doTR = $true
						$osi = [int]$oldStoreDetected
						$ot = @($osi.ToString("D4"), $osi.ToString("D3"))
						$nt = @($newStore4, $newStore3)
					}
					
					for ($i = 0; $i -lt $lines.Count; $i++)
					{
						$line = $lines[$i]
						
						$line = [regex]::Replace($line, '^\s*(?i:STORE)\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim))
						$line = [regex]::Replace($line, '^\s*(?i:Store)\s*=\s*\d{1,4}\s*$', ("Store=" + $newStoreTrim))
						
						$line = [regex]::Replace(
							$line,
							'^\s*(?i:REDIRMAIL|REDIRMSG|RedirMail|RedirMsg)\s*=\s*\d{1,4}(901)\s*$',
							{
								param ($m)
								$key = $m.Value.Split('=')[0].Trim()
								return ($key + "=" + $newStoreTrim + "901")
							}
						)
						
						if ($line -match '^(?<ws>\s*)(?i:SERVERNAME)\s*=\s*(?<val>.*)\s*$')
						{
							$ws = $matches['ws']
							$val = ($matches['val'] + "").Trim()
							if ($val.Length -gt 0) { $line = $ws + "SERVERNAME=" + $val.ToUpper() }
						}
						
						# Do NOT force TER changes in referenced Startup*.ini
						
						if ($doTR)
						{
							for ($t = 0; $t -lt $ot.Count; $t++)
							{
								$oldEsc = [regex]::Escape($ot[$t])
								$newTok = $nt[$t]
								$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
								$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
								$line = $line -replace ("(?<![\dA-Za-z])" + $oldEsc + "(?!\d)"), $newTok
							}
						}
						
						$lines[$i] = $line
					}
					
					[System.IO.File]::WriteAllText($resolved, ($lines -join $nlRef), $encRef)
					Write-Host ("Updated referenced INI: " + $resolved) -ForegroundColor Green
				}
				catch
				{
					$success = $false
					Write-Host ("Failed updating referenced INI '" + $iniRef + "': " + $_.Exception.Message) -ForegroundColor Red
				}
			}
		}
	}
	
	# ===============================================================================================
	# 3) Update INFO_*901_WIN.ini (optional) (use variable; small fallback via $DbsPath only)
	# ===============================================================================================
	if ([string]::IsNullOrWhiteSpace($WinIniPath) -and $DbsPath -and (Test-Path $DbsPath))
	{
		try
		{
			$f = Get-ChildItem -Path $DbsPath -File -Filter "INFO_*901_WIN.ini" -ErrorAction SilentlyContinue | Select-Object -First 1
			if ($f) { $WinIniPath = $f.FullName }
		}
		catch { }
	}
	
	if (-not [string]::IsNullOrWhiteSpace($WinIniPath) -and (Test-Path $WinIniPath))
	{
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($WinIniPath)
			
			$encWin = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{ $encWin = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{ $encWin = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{ $encWin = [System.Text.Encoding]::BigEndianUnicode }
			else
			{ $encWin = [System.Text.Encoding]::Default }
			
			$textWin = $encWin.GetString($bytes)
			$nlWin = "`r`n"
			if ($textWin -notmatch "`r`n" -and $textWin -match "`n") { $nlWin = "`n" }
			
			$winLines = $textWin -split "\r?\n", -1
			
			$doTokenReplaceW = $false
			$oldTokensW = @()
			$newTokensW = @()
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$doTokenReplaceW = $true
				$osi = [int]$oldStoreDetected
				$oldTokensW = @($osi.ToString("D4"), $osi.ToString("D3"))
				$newTokensW = @($newStore4, $newStore3)
			}
			
			for ($j = 0; $j -lt $winLines.Count; $j++)
			{
				$line = $winLines[$j]
				$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
				$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
				
				if ($doTokenReplaceW)
				{
					for ($t = 0; $t -lt $oldTokensW.Count; $t++)
					{
						$oldEsc = [regex]::Escape($oldTokensW[$t])
						$newTok = $newTokensW[$t]
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
						$line = $line -replace ("(?<![\dA-Za-z])" + $oldEsc + "(?!\d)"), $newTok
					}
				}
				
				$winLines[$j] = $line
			}
			
			[System.IO.File]::WriteAllText($WinIniPath, ($winLines -join $nlWin), $encWin)
			Write-Host "Updated INFO_*901_WIN.ini" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating INFO_*901_WIN.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	
	# ===============================================================================================
	# 4) Update SmsHttps.INI (FORCED)
	#   - SERVER moniker:
	#       * KEEP multiple processors
	#       * ONLY update store prefix in processor keys + REDIRMAIL/REDIRMSG values
	#       * Always clear LicenseGUID
	#   - NON-SERVER:
	#       * Enforce EXACTLY ONE processor entry keyed <STORE><TER> (store width inferred from file when possible)
	#       * Purge other processor lines
	#       * Force REDIRMAIL/REDIRMSG=<STORE>901
	#       * Always clear LicenseGUID
	#   - ALWAYS writes the file (even if identical), so it "triggers" every run.
	# ===============================================================================================
	
	# --- Resolve SmsHttpsIniPath if missing (fast, no broad probing) ---
	if ([string]::IsNullOrWhiteSpace($SmsHttpsIniPath))
	{
		try
		{
			$startupDirLocal = $null
			if (-not [string]::IsNullOrWhiteSpace($StartupIniPath))
			{
				try { $startupDirLocal = Split-Path -Path $StartupIniPath -Parent }
				catch { $startupDirLocal = $null }
			}
		}
		catch { }
	}
	if (-not [string]::IsNullOrWhiteSpace($SmsHttpsIniPath) -and (Test-Path $SmsHttpsIniPath))
	{
		try
		{
			# If terminal couldn't be derived from machine name, try TER= in Startup.ini (NON-server only)
			if (-not $isServerMoniker -and -not $newTerminal -and -not [string]::IsNullOrWhiteSpace($StartupIniPath) -and (Test-Path $StartupIniPath))
			{
				try
				{
					$terLine = (Get-Content -Path $StartupIniPath -ErrorAction Stop | Where-Object { $_ -match '^\s*TER\s*=\s*\d{1,4}\s*$' } | Select-Object -First 1)
					if ($terLine -and ($terLine -match '^\s*TER\s*=\s*(\d{1,4})\s*$'))
					{
						$tmpTer = $matches[1]
						if ($tmpTer -match '^\d{1,3}$')
						{
							$newTerminal = ([int]$tmpTer).ToString("D3")
							if ($newTerminal -eq "000") { $newTerminal = $null }
						}
					}
				}
				catch { }
			}
			
			$bytes = [System.IO.File]::ReadAllBytes($SmsHttpsIniPath)
			
			$enc = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $enc = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $enc = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { $enc = [System.Text.Encoding]::BigEndianUnicode }
			else { $enc = [System.Text.Encoding]::Default }
			
			$text = $enc.GetString($bytes)
			
			$nl = "`r`n"
			if ($text -notmatch "`r`n" -and $text -match "`n") { $nl = "`n" }
			
			$lines = $text -split "\r?\n", -1
			
			# --- Ensure/clear [GENERAL] LicenseGUID ---
			$genStart = -1
			for ($i = 0; $i -lt $lines.Length; $i++)
			{
				if ($lines[$i] -match '^\s*\[\s*GENERAL\s*\]\s*$') { $genStart = $i; break }
			}
			if ($genStart -lt 0)
			{
				$lines = @("[GENERAL]", "LicenseGUID=", "") + $lines
			}
			else
			{
				$genEnd = $lines.Length
				for ($i = $genStart + 1; $i -lt $lines.Length; $i++)
				{
					if ($lines[$i] -match '^\s*\[.*\]\s*$') { $genEnd = $i; break }
				}
				
				$lgFound = $false
				for ($i = $genStart + 1; $i -lt $genEnd; $i++)
				{
					if ($lines[$i] -match '^(?<ws>\s*)(?i:LicenseGUID)\s*=\s*(?<val>.*)\s*$')
					{
						$lines[$i] = ($matches['ws'] + "LicenseGUID=")
						$lgFound = $true
						break
					}
				}
				if (-not $lgFound)
				{
					$pre = @()
					$post = @()
					for ($k = 0; $k -lt $genEnd; $k++) { $pre += $lines[$k] }
					for ($k = $genEnd; $k -lt $lines.Length; $k++) { $post += $lines[$k] }
					$lines = @($pre + "LicenseGUID=" + $post)
				}
			}
			
			# --- Find [PROCESSORS] section (create if missing) ---
			$secStart = -1
			for ($i = 0; $i -lt $lines.Length; $i++)
			{
				if ($lines[$i] -match '^\s*\[\s*PROCESSORS?\s*\]\s*$') { $secStart = $i; break }
			}
			
			if ($secStart -lt 0)
			{
				# Always create it (forced behavior)
				$lines = @($lines + "" + "[PROCESSORS]")
				$secStart = ($lines.Length - 1)
			}
			
			$secEnd = $lines.Length
			for ($i = $secStart + 1; $i -lt $lines.Length; $i++)
			{
				if ($lines[$i] -match '^\s*\[.*\]\s*$') { $secEnd = $i; break }
			}
			
			# --- Infer store width from existing processor keys when possible (6=3+3, 7=4+3) ---
			$storeWidth = $newStoreTrim.Length
			try
			{
				for ($i = $secStart + 1; $i -lt $secEnd; $i++)
				{
					if ($lines[$i] -match '^\s*(\d{6,7})\s*=')
					{
						$keyLen = $matches[1].Length
						if ($keyLen -eq 7) { $storeWidth = 4; break }
						if ($keyLen -eq 6) { $storeWidth = 3; break }
					}
				}
			}
			catch { }
			
			$storeNorm = ([int]$newStoreTrim).ToString(("D{0}" -f $storeWidth))
			$redirKey901 = $storeNorm + "901"
			
			# =========================
			# SERVER: keep many processors; update store prefix in key + REDIR store prefix
			# =========================
			if ($isServerMoniker)
			{
				for ($i = $secStart + 1; $i -lt $secEnd; $i++)
				{
					$line = $lines[$i]
					if ($line -match '^(?<ws>\s*)(?<store>\d{3,4})(?<term>\d{3})\s*=\s*(?<rhs>.*)$')
					{
						$ws = $matches['ws']
						$st = $matches['store']
						$term = $matches['term']
						$rhs = $matches['rhs']
						
						$newStorePart = ([int]$newStoreTrim).ToString(("D{0}" -f $st.Length))
						$newKey = $newStorePart + $term
						
						$rhs2 = $rhs
						$rhs2 = [regex]::Replace($rhs2, '(?i)\b(REDIRMAIL)\s*=\s*(\d{3,4})(\d{3})\b', {
								param ($m)
								$k = $m.Groups[1].Value
								$st2 = $m.Groups[2].Value
								$tr2 = $m.Groups[3].Value
								$nst2 = ([int]$newStoreTrim).ToString(("D{0}" -f $st2.Length))
								return ($k + "=" + $nst2 + $tr2)
							})
						$rhs2 = [regex]::Replace($rhs2, '(?i)\b(REDIRMSG)\s*=\s*(\d{3,4})(\d{3})\b', {
								param ($m)
								$k = $m.Groups[1].Value
								$st2 = $m.Groups[2].Value
								$tr2 = $m.Groups[3].Value
								$nst2 = ([int]$newStoreTrim).ToString(("D{0}" -f $st2.Length))
								return ($k + "=" + $nst2 + $tr2)
							})
						
						$lines[$i] = ($ws + $newKey + "=" + $rhs2)
					}
				}
			}
			# =========================
			# NON-SERVER: enforce ONE processor keyed <STORE><TER> and force REDIRs to <STORE>901
			# =========================
			else
			{
				if ($newTerminal)
				{
					$newKey = $storeNorm + $newTerminal
					
					# Collect existing processor entries (numeric)
					$entries = @()
					for ($i = $secStart + 1; $i -lt $secEnd; $i++)
					{
						if ($lines[$i] -match '^(?<ws>\s*)(?<key>\d{6,7})\s*=\s*(?<rhs>.*)$')
						{
							$entries += [pscustomobject]@{
								Index = $i
								Ws    = $matches['ws']
								Key   = $matches['key']
								Rhs   = $matches['rhs']
							}
						}
					}
					
					# Pick a source rhs (same logic as you had, but without requiring old info to exist)
					$source = $null
					foreach ($e in $entries)
					{
						if ($e.Key.Length -ge 3 -and $e.Key.Substring($e.Key.Length - 3) -eq $newTerminal) { $source = $e; break }
					}
					if (-not $source -and $entries.Count -gt 0) { $source = $entries[0] }
					
					$rhsUpdated = ""
					$wsToUse = ""
					if ($source)
					{
						$rhsUpdated = $source.Rhs
						$wsToUse = $source.Ws
					}
					
					# Replace any old processor keys inside RHS -> newKey
					foreach ($e in $entries)
					{
						$rhsUpdated = $rhsUpdated -replace ("(?<!\d)" + [regex]::Escape($e.Key) + "(?!\d)"), $newKey
					}
					
					# Strip existing REDIRMAIL/REDIRMSG then enforce <STORE>901
					$rhsUpdated = [regex]::Replace($rhsUpdated, '\bREDIRMAIL\s*=\s*\d{6,7}', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
					$rhsUpdated = [regex]::Replace($rhsUpdated, '\bREDIRMSG\s*=\s*\d{6,7}', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
					$rhsUpdated = $rhsUpdated -replace ',{2,}', ','
					$rhsUpdated = $rhsUpdated.Trim().Trim(',').Trim()
					
					if ([string]::IsNullOrWhiteSpace($rhsUpdated))
					{
						$rhsUpdated = "REDIRMAIL=$redirKey901,REDIRMSG=$redirKey901,TARGETSEND=,TARGETRECV=,DEADLOCKPRIORITY="
					}
					else
					{
						$rhsUpdated = "REDIRMAIL=$redirKey901,REDIRMSG=$redirKey901," + $rhsUpdated
					}
					
					$newProcessorLine = $wsToUse + $newKey + "=" + $rhsUpdated
					
					# Purge all numeric processor lines; insert ONE just after section header
					$out = @()
					for ($i = 0; $i -lt $lines.Length; $i++)
					{
						$out += $lines[$i]
						
						if ($i -eq $secStart)
						{
							$out += $newProcessorLine
							continue
						}
						
						if ($i -gt $secStart -and $i -lt $secEnd)
						{
							if ($lines[$i] -match '^\s*\d{6,7}\s*=') { $out = $out[0 .. ($out.Count - 2)]; continue } # remove what we just added
						}
					}
					
					$lines = $out
				}
			}
			
			# ---- FORCED write (always) ----
			[System.IO.File]::WriteAllText($SmsHttpsIniPath, ($lines -join $nl), $enc)
			Write-Host ("Updated SmsHttps.INI (forced): " + $SmsHttpsIniPath) -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating SmsHttps.INI: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	return $success
}

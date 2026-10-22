# ===================================================================================================
#                                 SECTION: Storeman Folder Detection
# ---------------------------------------------------------------------------------------------------
# Description:
#   Resolves $BasePath, the storeman folder every store task works against, in three levels:
#     1) the local SMB share "storeman", IF Office\Dbs\INFO_*901_WIN.INI has a KeyString and a
#        KeyValidDate no older than 14 days                                  -> Production
#     2) the first configured test root that exists (C:\Storeman, D:\Storeman) -> Test fallback
#     3) a scan of every fixed drive's top level for a *storeman* folder with a Startup.ini,
#        preferring one that is published as a share                        -> Drive-scan fallback
#   A share whose key is older than 14 days is the dedicated KEY EXPIRED state: store data is
#   usually still there, only a re-key is needed (the GUI says so instead of "test mode").
#
# Variables:
#   - $BasePath                       the resolved storeman folder
#   - $script:IsTestEnvironment       $true for every non-production result
#   - $script:StoremanSource          'Production' | 'Test fallback' | 'Drive-scan fallback'
#   - $script:StoremanSourceReason    plain-language reason (banner tooltip, log, status bar)
#   - $script:StoremanKeyStale        KEY EXPIRED state
#   - $script:StoremanKeyDate         the stale KeyValidDate (KEY EXPIRED only)
#
# Notes:
#   - tests\TBS_Maintenance_SelfTest.ps1 runs this block verbatim, from the '$BasePath = $null'
#     line to the 'Selected (storeman) folder' line - keep both marker lines.
# ===================================================================================================
$BasePath = $null
$testPath = $script:StoremanRoots | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
if (-not $testPath) { $testPath = $script:PrimaryStoremanRoot }

# ---------------------------------------------------------------------------
# Test-environment awareness: remember WHICH storeman source we ended up with,
# so the GUI (banner badge, status bar, log) can clearly show that we are NOT
# running against the production share.
# ---------------------------------------------------------------------------
$script:IsTestEnvironment = $false
$script:StoremanSource = 'Production'
$script:StoremanSourceReason = "SMB share 'storeman' with valid key"
# Dedicated KEY EXPIRED state: the share + store data exist, but KeyValidDate is
# older than 14 days. Distinct from plain test mode because store data (folder +
# database) is usually STILL available there - the GUI says "re-key required"
# instead of "tasks unavailable".
$script:StoremanKeyStale = $false
$script:StoremanKeyDate = $null

# 1) Try local SMB share "storeman"
$share = $null
try
{
	$share = Get-SmbShare -ErrorAction Stop | Where-Object { $_.Name -ieq 'storeman' } | Select-Object -First 1
}
catch { $share = $null }

if ($share -and -not [string]::IsNullOrWhiteSpace($share.Path))
{
	# Find INFO file in share path
	$infoFile = Get-ChildItem -Path (Join-Path $share.Path 'Office\Dbs') -Filter 'INFO_*901_WIN.INI' -ErrorAction SilentlyContinue | Select-Object -First 1
	if ($infoFile)
	{
		# Read KeyString + KeyValidDate (inline)
		$keyString = ""
		$keyValid = $null

		try
		{
			$txt = Get-Content -Path $infoFile.FullName -Raw -ErrorAction Stop

			$mKey = [regex]::Match($txt, '(?im)^\s*KeyString\s*=\s*(.+?)\s*$')
			if ($mKey.Success) { $keyString = ($mKey.Groups[1].Value).Trim() }

			$mDt = [regex]::Match($txt, '(?im)^\s*KeyValidDate\s*=\s*(\d{2}/\d{2}/\d{4})\s*$')
			if ($mDt.Success)
			{
				try { $keyValid = [datetime]::ParseExact($mDt.Groups[1].Value, 'MM/dd/yyyy', $null) }
				catch { $keyValid = $null }
			}
		}
		catch { }

		# Key requirements for SHARE:
		#   - KeyString present
		#   - KeyValidDate exists and is NOT older than 14 days
		$hasKeyString = -not [string]::IsNullOrWhiteSpace($keyString)
		$minDate = (Get-Date).Date.AddDays(-14)
		$keyWithinLimit = ($keyValid -is [datetime]) -and ($keyValid.Date -ge $minDate)

		if ($hasKeyString -and $keyWithinLimit)
		{
			$BasePath = $share.Path
		}
		else
		{
			$script:IsTestEnvironment = $true
			$script:StoremanSource = 'Test fallback'
			if ($hasKeyString -and ($keyValid -is [datetime]))
			{
				# Key exists but is older than 14 days -> dedicated KEY EXPIRED state
				$script:StoremanKeyStale = $true
				$script:StoremanKeyDate = $keyValid
				$script:StoremanSourceReason = "SMB share 'storeman' exists with store data present, but KeyValidDate ($($keyValid.ToString('MM/dd/yyyy'))) is older than 14 days - re-key required"
				Write-Warning "Share 'storeman' found but KeyValidDate ($($keyValid.ToString('MM/dd/yyyy'))) is older than 14 days - re-key required. Falling back to ${testPath}."
			}
			elseif (-not $hasKeyString)
			{
				$script:StoremanSourceReason = "SMB share 'storeman' exists but its KeyString is missing"
				Write-Warning "Share 'storeman' found but KeyString is missing. Falling back to ${testPath} for testing."
			}
			else
			{
				$script:StoremanSourceReason = "SMB share 'storeman' exists but KeyValidDate is missing or unreadable"
				Write-Warning "Share 'storeman' found but KeyValidDate is missing or unreadable. Falling back to ${testPath} for testing."
			}
		}
	}
	else
	{
		$script:IsTestEnvironment = $true
		$script:StoremanSource = 'Test fallback'
		$script:StoremanSourceReason = "SMB share 'storeman' exists but Office\Dbs\INFO_*901_WIN.INI is missing"
		Write-Warning "Share 'storeman' found but INFO file missing. Falling back to ${testPath} for testing."
	}
}
else
{
	$script:IsTestEnvironment = $true
	$script:StoremanSource = 'Test fallback'
	$script:StoremanSourceReason = "no 'storeman' SMB share on this machine"
	Write-Warning "Storeman share 'storeman' not found. Falling back to configured test roots / drive scan."
}

# 2) Testing fallback (configured roots) -> 3) drive-scan fallback (ported from MiniGhost)
if (-not $BasePath)
{
	if ($testPath -and (Test-Path -LiteralPath $testPath))
	{
		$BasePath = $testPath

		# Defensive: make sure the test-environment flag is set no matter which
		# fallback branch we took
		if (-not $script:IsTestEnvironment)
		{
			$script:IsTestEnvironment = $true
			$script:StoremanSource = 'Test fallback'
			$script:StoremanSourceReason = "no 'storeman' SMB share on this machine"
		}

		# Optional: warn if the testing path still doesn't have INFO
		$testInfo = Get-ChildItem -Path (Join-Path $BasePath 'Office\Dbs') -Filter 'INFO_*901_WIN.INI' -ErrorAction SilentlyContinue | Select-Object -First 1
		if (-not $testInfo)
		{
			Write-Warning "Testing fallback selected (${testPath}) but INFO_*901_WIN.INI was not found. Running with limited functionality."
		}
	}
	else
	{
		# 3) FINAL FALLBACK (MiniGhost logic): scan every fixed drive for any
		# *storeman* folder that contains a Startup.ini, preferring the one
		# actually published as an SMB share.
		$script:IsTestEnvironment = $true
		$script:StoremanSource = 'Drive-scan fallback'
		Write-Warning "No 'storeman' SMB share and no existing C:\Storeman / D:\Storeman test root. Scanning all fixed drives for a *storeman* folder..."

		$storemanDirs = @()
		$fixedDrives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 -and $_.Root -match '^[A-Z]:\\$' }
		foreach ($drive in $fixedDrives)
		{
			# Case-insensitive match for any *storeman* variation in directory name
			$dirs = Get-ChildItem -Path "$($drive.Root)" -Directory -ErrorAction SilentlyContinue |
			Where-Object { $_.Name -imatch 'storeman' } |
			ForEach-Object {
				$candidatePath = Join-Path $_.FullName '\'
				$files = Get-ChildItem -Path $candidatePath -Filter 'Startup.ini' -ErrorAction SilentlyContinue
				if ($files) { $_ }
			}
			if ($dirs) { $storemanDirs += $dirs }
		}

		$scanned = $null
		if ($storemanDirs.Count -gt 0)
		{
			# Prefer a path that is actually shared as a Windows share
			$shares = Get-SmbShare -ErrorAction SilentlyContinue
			foreach ($dir in $storemanDirs)
			{
				if ($shares.Path -contains $dir.FullName)
				{
					$scanned = $dir
					break
				}
			}
			if (-not $scanned) { $scanned = $storemanDirs[0] }
		}

		if ($scanned)
		{
			$BasePath = $scanned.FullName
			$script:StoremanSourceReason = "no 'storeman' SMB share and no test root - found '$BasePath' by drive scan (Startup.ini present)"
			Write-Warning "Storeman folder found by drive scan: '$BasePath'. Running in test/fallback mode."
		}
		else
		{
			# Last resort: primary test root, explicitly flagged as missing
			$BasePath = $testPath
			$script:StoremanSourceReason += "; test root '$BasePath' does not exist and no *storeman* folder found by drive scan"
			Write-Warning "Could not locate any storeman folder containing 'storeman\Startup.ini'. Running with limited functionality."
		}
	}
}

Write-Host "Selected (storeman) folder: '$BasePath'" -ForegroundColor Magenta

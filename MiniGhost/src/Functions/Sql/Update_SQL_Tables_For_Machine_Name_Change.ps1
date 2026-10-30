# ===================================================================================================
#                       FUNCTION: Update_SQL_Tables_For_Machine_Name_Change
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates TER_TAB, STO_TAB, LNK_TAB, RUN_TAB, and STD_TAB after machine name change / store change.
#
# Auto DB context detection (based on Startup.ini DBNAME already gathered via Get_Database_Connection_String):
#   - LANESQL             => LocalTerminalDB (lane/POS/SCO cleanup in same store is allowed)
#   - STORESQL or HOSTSQL => ServerStoreDB  (do NOT lane-cleanup other terminals)
#
# STD_TAB enhancement:
#   - Also updates F1531 (CompanyName) store suffix when present (e.g., "Company Name 231" or "Company #0231")
#
# FIXES:
#   - Non-server runs preserve/copy TER_TAB 901 into the NEW store before deleting old-store rows.
#   - STO_TAB now always ensures the 901 row exists.
#   - LNK_TAB now preserves/copies 901 into the NEW store before cleanup, so the server side is not left empty.
# ===================================================================================================

function Update_SQL_Tables_For_Machine_Name_Change
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidatePattern('^(?!0+$)\d{3,4}$')]
		[string]$storeNumber,
		[Parameter(Mandatory = $true)]
		[string]$machineName,
		[Parameter(Mandatory = $true)]
		[string]$machineNumber,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineName,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineNumber,
		[Parameter(Mandatory = $false)]
		[ValidateSet('Auto', 'LocalTerminalDB', 'ServerStoreDB')]
		[string]$DatabaseContext = 'Auto'
	)
	
	# ---------------------------
	# Constants / normalization
	# ---------------------------
	$hostStoreNumber = "999"
	$serverTerminal = "901"
	$protectedMinTerminal = 940
	
	$storeNumber = (($storeNumber | ForEach-Object { if ($null -eq $_) { "" }
				else { $_.ToString() } }).Trim())
	$machineName = (($machineName | ForEach-Object { if ($null -eq $_) { "" }
				else { $_.ToString() } }).Trim())
	$machineNumber = (($machineNumber | ForEach-Object { if ($null -eq $_) { "" }
				else { $_.ToString() } }).Trim())
	
	# Detect SERVER moniker (example: 0242SERVER001)
	$isServerMoniker = $false
	if (-not [string]::IsNullOrWhiteSpace($machineName) -and ($machineName -match '(?i)SERVER'))
	{
		$isServerMoniker = $true
	}
	elseif ($script:NodeRole -and ($script:NodeRole -match '^(?i)(StoreServer|HostServer)$'))
	{
		$isServerMoniker = $true
	}
	
	$null = Ensure_Database_Connection_Context
	
	# ---------------------------
	# Pull DBNAME (already gathered by Get_Database_Connection_String) for better Auto context
	# ---------------------------
	$dbNameDetected = $null
	try
	{
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('DBNAME') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['DBNAME']))
		{
			$dbNameDetected = ([string]$script:FunctionResults['DBNAME']).Trim()
		}
	}
	catch { }
	
	# Fallback: parse from ConnectionString if present
	if ([string]::IsNullOrWhiteSpace($dbNameDetected))
	{
		try
		{
			if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('ConnectionString'))
			{
				$cs = [string]$script:FunctionResults['ConnectionString']
				if ($cs -match '(?i)\bDatabase\s*=\s*([^;]+)') { $dbNameDetected = $Matches[1].Trim() }
				elseif ($cs -match '(?i)\bInitial\s+Catalog\s*=\s*([^;]+)') { $dbNameDetected = $Matches[1].Trim() }
			}
		}
		catch { }
	}
	
	# ---------------------------
	# Normalize NEW/OLD server host string (for matching TER_TAB.F1125 and replacing STO_TAB.F1018)
	# ---------------------------
	$newServerNameNorm = $machineName
	if ($null -eq $newServerNameNorm) { $newServerNameNorm = "" }
	$newServerNameNorm = $newServerNameNorm.Trim()
	$newServerNameNorm = $newServerNameNorm -replace '^[\\\/]+', ''
	if ($newServerNameNorm -match '[\\\/]') { $newServerNameNorm = ($newServerNameNorm -split '[\\\/]')[0] }
	if ($newServerNameNorm -match '\.') { $newServerNameNorm = ($newServerNameNorm -split '\.')[0] }
	$newServerNameNorm = $newServerNameNorm.Trim().ToUpper()
	
	$oldServerNameNorm = $OldMachineName
	if ([string]::IsNullOrWhiteSpace($oldServerNameNorm)) { $oldServerNameNorm = $env:COMPUTERNAME }
	if ($null -eq $oldServerNameNorm) { $oldServerNameNorm = "" }
	$oldServerNameNorm = $oldServerNameNorm.Trim()
	$oldServerNameNorm = $oldServerNameNorm -replace '^[\\\/]+', ''
	if ($oldServerNameNorm -match '[\\\/]') { $oldServerNameNorm = ($oldServerNameNorm -split '[\\\/]')[0] }
	if ($oldServerNameNorm -match '\.') { $oldServerNameNorm = ($oldServerNameNorm -split '\.')[0] }
	$oldServerNameNorm = $oldServerNameNorm.Trim().ToUpper()
	
	# ---------------------------
	# Normalize NEW terminal to 3 digits
	# ---------------------------
	$newTerminal = $null
	if ($machineNumber -match '^\d{1,3}$')
	{
		$newTerminal = ([int]$machineNumber).ToString("D3")
		if ($newTerminal -eq "000") { $newTerminal = $null }
	}
	else
	{
		$tmp = $machineNumber
		if ($null -eq $tmp) { $tmp = "" }
		$tmp = $tmp.Trim()
		$tmp = $tmp -replace '^[\\\/]+', ''
		if ($tmp -match '[\\\/]') { $tmp = ($tmp -split '[\\\/]')[0] }
		if ($tmp -match '\.') { $tmp = ($tmp -split '\.')[0] }
		$tmp = $tmp.Trim().ToUpper()
		
		if ($tmp -match '(\d{1,3})$')
		{
			$newTerminal = ([int]$Matches[1]).ToString("D3")
			if ($newTerminal -eq "000") { $newTerminal = $null }
		}
	}
	
	# Determine OLD terminal for safe RUN_TAB swap (used only when NOT server)
	if ([string]::IsNullOrWhiteSpace($OldMachineName)) { $OldMachineName = $env:COMPUTERNAME }
	
	$oldTerminal = $null
	
	if (-not [string]::IsNullOrWhiteSpace($OldMachineNumber))
	{
		if ($OldMachineNumber -match '^\d{1,3}$')
		{
			$oldTerminal = ([int]$OldMachineNumber).ToString("D3")
			if ($oldTerminal -eq "000") { $oldTerminal = $null }
		}
		else
		{
			$tmpOld = $OldMachineNumber
			if ($null -eq $tmpOld) { $tmpOld = "" }
			$tmpOld = $tmpOld.Trim()
			$tmpOld = $tmpOld -replace '^[\\\/]+', ''
			if ($tmpOld -match '[\\\/]') { $tmpOld = ($tmpOld -split '[\\\/]')[0] }
			if ($tmpOld -match '\.') { $tmpOld = ($tmpOld -split '\.')[0] }
			$tmpOld = $tmpOld.Trim().ToUpper()
			
			if ($tmpOld -match '(\d{1,3})$')
			{
				$oldTerminal = ([int]$Matches[1]).ToString("D3")
				if ($oldTerminal -eq "000") { $oldTerminal = $null }
			}
		}
	}
	
	if (-not $oldTerminal)
	{
		$tmpOld2 = $OldMachineName
		if ($null -eq $tmpOld2) { $tmpOld2 = "" }
		$tmpOld2 = $tmpOld2.Trim()
		$tmpOld2 = $tmpOld2 -replace '^[\\\/]+', ''
		if ($tmpOld2 -match '[\\\/]') { $tmpOld2 = ($tmpOld2 -split '[\\\/]')[0] }
		if ($tmpOld2 -match '\.') { $tmpOld2 = ($tmpOld2 -split '\.')[0] }
		$tmpOld2 = $tmpOld2.Trim().ToUpper()
		
		if ($tmpOld2 -match '(\d{1,3})$')
		{
			$oldTerminal = ([int]$Matches[1]).ToString("D3")
			if ($oldTerminal -eq "000") { $oldTerminal = $null }
		}
	}
	
	# ---------------------------
	# Decide DB context
	# ---------------------------
	$isLocalTerminalDB = $false
	
	if ($DatabaseContext -eq 'LocalTerminalDB')
	{
		$isLocalTerminalDB = $true
	}
	elseif ($DatabaseContext -eq 'ServerStoreDB')
	{
		$isLocalTerminalDB = $false
	}
	else
	{
		# Auto: prefer DBNAME
		if (-not [string]::IsNullOrWhiteSpace($dbNameDetected) -and ($dbNameDetected -match '^(?i)LANESQL$'))
		{
			$isLocalTerminalDB = $true
		}
		elseif (-not [string]::IsNullOrWhiteSpace($dbNameDetected) -and ($dbNameDetected -match '^(?i)(STORESQL|HOSTSQL)$'))
		{
			$isLocalTerminalDB = $false
		}
		else
		{
			# Fallback heuristic: if targeting this box, assume local terminal DB
			$execHostNorm = $env:COMPUTERNAME
			if ($null -eq $execHostNorm) { $execHostNorm = "" }
			$execHostNorm = ($execHostNorm + "").Trim()
			$execHostNorm = $execHostNorm -replace '^[\\\/]+', ''
			if ($execHostNorm -match '[\\\/]') { $execHostNorm = ($execHostNorm -split '[\\\/]')[0] }
			if ($execHostNorm -match '\.') { $execHostNorm = ($execHostNorm -split '\.')[0] }
			$execHostNorm = $execHostNorm.Trim().ToUpper()
			
			$targetHostNorm = $machineName
			if ($null -eq $targetHostNorm) { $targetHostNorm = "" }
			$targetHostNorm = ($targetHostNorm + "").Trim()
			$targetHostNorm = $targetHostNorm -replace '^[\\\/]+', ''
			if ($targetHostNorm -match '[\\\/]') { $targetHostNorm = ($targetHostNorm -split '[\\\/]')[0] }
			if ($targetHostNorm -match '\.') { $targetHostNorm = ($targetHostNorm -split '\.')[0] }
			$targetHostNorm = $targetHostNorm.Trim().ToUpper()
			
			$isLocalTerminalDB = ($targetHostNorm -eq $execHostNorm)
		}
	}
	
	# SAFETY: Non-server mode must have a terminal, otherwise we can't safely upsert/cleanup.
	if (-not $isServerMoniker -and -not $newTerminal)
	{
		return @{
			Success		    = $false
			FailedCommands  = @("Refusing to run: could not derive a valid NEW terminal from -machineNumber/-machineName for non-server mode.")
			IsServerMoniker = $false
			RunTabUpdated   = $false
			OldTerminalUsed = $oldTerminal
			NewTerminalUsed = $newTerminal
			DatabaseContext = $DatabaseContext
			LocalTerminalDB = $isLocalTerminalDB
			DBNAME		    = $dbNameDetected
		}
	}
	
	# ---------------------------
	# Table names
	# ---------------------------
	$terTableName = "TER_TAB"
	$stoTableName = "STO_TAB"
	$lnkTableName = "LNK_TAB"
	$runTableName = "RUN_TAB"
	$stdTableName = "STD_TAB"
	
	# ===============================================================================================
	# TER_TAB
	# ===============================================================================================
	
	$createViewCommandTer = @"
IF OBJECT_ID('Ter_Load', 'V') IS NOT NULL DROP VIEW Ter_Load;
EXEC(N'
CREATE VIEW Ter_Load AS
SELECT F1056, F1057, F1058, F1125, F1169
FROM $terTableName;
');
"@
	
	$moveProtectedTer_ToNewStore = @"
DECLARE @NewStore  varchar(10) = '$storeNumber';
DECLARE @HostStore varchar(10) = '$hostStoreNumber';
DECLARE @ProtectedMin int = $protectedMinTerminal;

INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169)
SELECT @NewStore, t.F1057, t.F1058, t.F1125, t.F1169
FROM $terTableName t
WHERE t.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) >= @ProtectedMin
  AND NOT EXISTS (
      SELECT 1
      FROM $terTableName x
      WHERE x.F1056 = @NewStore
        AND x.F1057 = t.F1057
  );

DELETE t
FROM $terTableName t
WHERE t.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) >= @ProtectedMin;
"@
	
	$moveServerNodes902_920_ToNewStore_AndUpdate = @"
DECLARE @NewStore  varchar(10) = '$storeNumber';
DECLARE @HostStore varchar(10) = '$hostStoreNumber';
DECLARE @NewServer varchar(200) = '$newServerNameNorm';
DECLARE @OldServer varchar(200) = '$oldServerNameNorm';

;WITH Candidates AS
(
    SELECT t.*
    FROM $terTableName t
    WHERE t.F1056 NOT IN (@NewStore, @HostStore)
      AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920
      AND t.F1125 IS NOT NULL
      AND
      (
          UPPER(t.F1125) LIKE '%\\' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%\\' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%@' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%@' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%' + @NewServer + '%'
      )
)
INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169)
SELECT @NewStore, c.F1057, c.F1058, c.F1125, c.F1169
FROM Candidates c
WHERE NOT EXISTS (
    SELECT 1
    FROM $terTableName x
    WHERE x.F1056 = @NewStore
      AND x.F1057 = c.F1057
);

UPDATE t
SET
    t.F1125 = CASE
                WHEN t.F1125 IS NULL THEN t.F1125
                WHEN LEFT(LTRIM(RTRIM(t.F1125)),1)='@' THEN '@' + @NewServer
                ELSE REPLACE(UPPER(t.F1125), @OldServer, @NewServer)
              END,
    t.F1169 = CASE
                WHEN t.F1169 IS NULL THEN t.F1169
                ELSE '\\' + @NewServer + '\storeman\office\XF' + @NewStore + '901\'
              END
FROM $terTableName t
WHERE t.F1056 = @NewStore
  AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920;

DELETE t
FROM $terTableName t
WHERE t.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920
  AND t.F1125 IS NOT NULL
  AND
  (
      UPPER(t.F1125) LIKE '%\\' + @OldServer + '%'
      OR UPPER(t.F1125) LIKE '%\\' + @NewServer + '%'
      OR UPPER(t.F1125) LIKE '%@' + @OldServer + '%'
      OR UPPER(t.F1125) LIKE '%@' + @NewServer + '%'
      OR UPPER(t.F1125) LIKE '%' + @OldServer + '%'
      OR UPPER(t.F1125) LIKE '%' + @NewServer + '%'
  );
"@
	
	$deleteOtherStoresTer = @"
DELETE FROM $terTableName
WHERE F1056 NOT IN ('$storeNumber', '$hostStoreNumber');
"@
	
	$deleteNonServerRowsForNewStore = @"
DECLARE @NewStore varchar(10) = '$storeNumber';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';
DECLARE @ProtectedMin int = $protectedMinTerminal;
DECLARE @NewServer varchar(200) = '$newServerNameNorm';
DECLARE @OldServer varchar(200) = '$oldServerNameNorm';

DELETE t
FROM $terTableName t
WHERE t.F1056 = @NewStore
  AND t.F1057 <> @ServerTerm
  AND
  (
      (
          TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) IS NULL
          OR TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) < @ProtectedMin
      )
      AND
      (
          NOT (
              TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920
              AND t.F1125 IS NOT NULL
              AND
              (
                  UPPER(t.F1125) LIKE '%\\' + @NewServer + '%'
                  OR UPPER(t.F1125) LIKE '%\\' + @OldServer + '%'
                  OR UPPER(t.F1125) LIKE '%@' + @NewServer + '%'
                  OR UPPER(t.F1125) LIKE '%@' + @OldServer + '%'
                  OR UPPER(t.F1125) LIKE '%' + @NewServer + '%'
                  OR UPPER(t.F1125) LIKE '%' + @OldServer + '%'
              )
          )
      )
  );
"@
	
	$upsertLaneTer = @"
IF EXISTS (SELECT 1 FROM $terTableName WHERE F1056='$storeNumber' AND F1057='$newTerminal')
BEGIN
    UPDATE $terTableName
    SET F1058='Terminal $newTerminal',
        F1125='\\$machineName\storeman\office\XF$storeNumber$newTerminal\',
        F1169='\\$machineName\storeman\office\XF$storeNumber$serverTerminal\'
    WHERE F1056='$storeNumber' AND F1057='$newTerminal';
END
ELSE
BEGIN
    INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169) VALUES
    ('$storeNumber', '$newTerminal',
     'Terminal $newTerminal',
     '\\$machineName\storeman\office\XF$storeNumber$newTerminal\',
     '\\$machineName\storeman\office\XF$storeNumber$serverTerminal\');
END
"@
	
	$preserve901ForNonServerStoreChange = @"
DECLARE @NewStore   varchar(10) = '$storeNumber';
DECLARE @HostStore  varchar(10) = '$hostStoreNumber';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';

IF NOT EXISTS (
    SELECT 1
    FROM $terTableName
    WHERE F1056 = @NewStore
      AND F1057 = @ServerTerm
)
BEGIN
    INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169)
    SELECT TOP 1
           @NewStore,
           @ServerTerm,
           CASE
               WHEN NULLIF(LTRIM(RTRIM(t.F1058)), '') IS NULL THEN 'Server'
               ELSE t.F1058
           END,
           t.F1125,
           t.F1169
    FROM $terTableName t
    WHERE t.F1057 = @ServerTerm
      AND t.F1056 <> @HostStore
    ORDER BY
        CASE WHEN t.F1125 IS NOT NULL AND LTRIM(RTRIM(t.F1125)) <> '' THEN 0 ELSE 1 END,
        t.F1056;
END

UPDATE $terTableName
SET F1058 = CASE
                WHEN NULLIF(LTRIM(RTRIM(F1058)), '') IS NULL THEN 'Server'
                ELSE F1058
            END
WHERE F1056 = @NewStore
  AND F1057 = @ServerTerm;
"@
	
	$cleanupTer_ForLocalTerminalDB = @"
DECLARE @NewStore varchar(10) = '$storeNumber';
DECLARE @KeepTerm varchar(10) = '$newTerminal';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';
DECLARE @ProtectedMin int = $protectedMinTerminal;

DELETE t
FROM $terTableName t
WHERE t.F1056 = @NewStore
  AND t.F1057 NOT IN (@KeepTerm, @ServerTerm)
  AND (TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) IS NULL OR TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) < @ProtectedMin);
"@
	
	$upsertServer901Ter_ForNewStoreOnly = @"
IF EXISTS (SELECT 1 FROM $terTableName WHERE F1056='$storeNumber' AND F1057='$serverTerminal')
BEGIN
    UPDATE $terTableName
    SET F1058='Server',
        F1125='@$machineName',
        F1169=NULL
    WHERE F1056='$storeNumber' AND F1057='$serverTerminal';
END
ELSE
BEGIN
    INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169) VALUES
    ('$storeNumber', '$serverTerminal',
     'Server',
     '@$machineName',
     NULL);
END
"@
	
	$dropViewCommandTer = @"
IF OBJECT_ID('Ter_Load', 'V') IS NOT NULL DROP VIEW Ter_Load;
"@
	
	# ===============================================================================================
	# STO_TAB
	# ===============================================================================================
	
	$createViewCommandSto = @"
IF OBJECT_ID('Sto_Load', 'V') IS NOT NULL DROP VIEW Sto_Load;
EXEC(N'
CREATE VIEW Sto_Load AS
SELECT F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691
FROM $stoTableName;
');
"@
	
	$upsertStoTerminal = @"
MERGE INTO $stoTableName AS target
USING (VALUES
    ('$newTerminal', 'Terminal $newTerminal', 1, 1, 1, NULL, NULL, NULL, NULL)
) AS source (F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691)
ON target.F1000 = source.F1000
WHEN MATCHED THEN
    UPDATE SET
        F1018 = source.F1018,
        F1180 = source.F1180,
        F1181 = source.F1181,
        F1182 = source.F1182
WHEN NOT MATCHED THEN
    INSERT (F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691)
    VALUES (source.F1000, source.F1018, source.F1180, source.F1181, source.F1182, source.F1937, source.F1965, source.F1966, source.F2691);
"@
	
	$upsertStoServer901 = @"
MERGE INTO $stoTableName AS target
USING (VALUES
    ('$serverTerminal', 'Server', 1, 1, 1, NULL, NULL, NULL, NULL)
) AS source (F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691)
ON target.F1000 = source.F1000
WHEN MATCHED THEN
    UPDATE SET
        F1018 = CASE
                    WHEN NULLIF(LTRIM(RTRIM(target.F1018)), '') IS NULL THEN source.F1018
                    ELSE target.F1018
                END,
        F1180 = ISNULL(target.F1180, source.F1180),
        F1181 = ISNULL(target.F1181, source.F1181),
        F1182 = ISNULL(target.F1182, source.F1182)
WHEN NOT MATCHED THEN
    INSERT (F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691)
    VALUES (source.F1000, source.F1018, source.F1180, source.F1181, source.F1182, source.F1937, source.F1965, source.F1966, source.F2691);
"@
	
	$cleanupSto_ForLocalTerminalDB = @"
DECLARE @KeepTerm varchar(10) = '$newTerminal';

DELETE FROM $stoTableName
WHERE F1000 <> @KeepTerm
  AND F1000 NOT IN ('DSM','PAL','RAL','XAL','901','999');
"@
	
	$cleanupSto_ForServer = @"
;WITH StoNums AS
(
    SELECT
        F1000,
        F1018,
        TRY_CONVERT(int, LTRIM(RTRIM(F1000))) AS NumVal
    FROM $stoTableName
    WHERE ISNUMERIC(F1000) = 1
)
DELETE s
FROM $stoTableName s
JOIN StoNums n ON n.F1000 = s.F1000
WHERE
    n.F1000 NOT IN ('901','999')
    AND
    (
        (n.NumVal BETWEEN 1 AND 99)
        OR
        (
            n.NumVal BETWEEN 902 AND 920
            AND (n.F1018 IS NULL OR UPPER(n.F1018) NOT LIKE '%SERVER%')
        )
    );

DECLARE @NewServer varchar(200) = '$newServerNameNorm';
DECLARE @OldServer varchar(200) = '$oldServerNameNorm';

UPDATE $stoTableName
SET F1018 = REPLACE(UPPER(F1018), @OldServer, @NewServer)
WHERE ISNUMERIC(F1000) = 1
  AND TRY_CONVERT(int, LTRIM(RTRIM(F1000))) BETWEEN 902 AND 920
  AND F1018 IS NOT NULL
  AND UPPER(F1018) LIKE '%' + @OldServer + '%';
"@
	
	$dropViewCommandSto = @"
IF OBJECT_ID('Sto_Load', 'V') IS NOT NULL DROP VIEW Sto_Load;
"@
	
	# ===============================================================================================
	# LNK_TAB
	# ===============================================================================================
	
	$createViewCommandLnk = @"
IF OBJECT_ID('Lnk_Load', 'V') IS NOT NULL DROP VIEW Lnk_Load;
EXEC(N'
CREATE VIEW Lnk_Load AS
SELECT F1000, F1056, F1057
FROM $lnkTableName;
');
"@
	
	$preserve901LnkForNewStore = @"
DECLARE @NewStore   varchar(10) = '$storeNumber';
DECLARE @HostStore  varchar(10) = '$hostStoreNumber';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';

INSERT INTO $lnkTableName (F1000, F1056, F1057)
SELECT DISTINCT
       l.F1000,
       @NewStore,
       @ServerTerm
FROM $lnkTableName l
WHERE l.F1057 = @ServerTerm
  AND l.F1056 <> @HostStore
  AND l.F1000 IS NOT NULL
  AND NOT EXISTS (
      SELECT 1
      FROM $lnkTableName x
      WHERE x.F1000 = l.F1000
        AND x.F1056 = @NewStore
        AND x.F1057 = @ServerTerm
  );

IF NOT EXISTS (
    SELECT 1
    FROM $lnkTableName
    WHERE F1056 = @NewStore
      AND F1057 = @ServerTerm
)
BEGIN
    INSERT INTO $lnkTableName (F1000, F1056, F1057)
    VALUES (@ServerTerm, @NewStore, @ServerTerm);
END
"@
	
	$rebuildLnkForTerminal = @"
DECLARE @NewStore  varchar(10) = '$storeNumber';
DECLARE @HostStore varchar(10) = '$hostStoreNumber';
DECLARE @Term      varchar(10) = '$newTerminal';
DECLARE @ProtectedMin int = $protectedMinTerminal;

DECLARE @NewServer varchar(200) = '$newServerNameNorm';
DECLARE @OldServer varchar(200) = '$oldServerNameNorm';

DECLARE @SourceStore varchar(10);

SELECT TOP 1 @SourceStore = F1056
FROM $lnkTableName
WHERE F1057 = @Term
  AND F1056 <> @HostStore
  AND F1056 <> @NewStore
ORDER BY F1056;

IF @SourceStore IS NULL
BEGIN
    SELECT TOP 1 @SourceStore = F1056
    FROM $lnkTableName
    WHERE F1057 = @Term
      AND F1056 = @NewStore;
END

DECLARE @Codes TABLE (F1000 varchar(50) NOT NULL PRIMARY KEY);

IF @SourceStore IS NOT NULL
BEGIN
    INSERT INTO @Codes(F1000)
    SELECT DISTINCT F1000
    FROM $lnkTableName
    WHERE F1056 = @SourceStore
      AND F1057 = @Term
      AND F1000 IS NOT NULL;
END

IF EXISTS (SELECT 1 FROM $terTableName WHERE F1056=@NewStore AND TRY_CONVERT(int, LTRIM(RTRIM(F1057))) BETWEEN 902 AND 920)
BEGIN
    DECLARE @ServerTerms TABLE (Term varchar(10) NOT NULL PRIMARY KEY);

    INSERT INTO @ServerTerms(Term)
    SELECT DISTINCT t.F1057
    FROM $terTableName t
    WHERE t.F1056 = @NewStore
      AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920
      AND t.F1125 IS NOT NULL
      AND
      (
          UPPER(t.F1125) LIKE '%\\' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%\\' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%@' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%@' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%' + @OldServer + '%'
      );

    INSERT INTO $lnkTableName (F1000, F1056, F1057)
    SELECT l.F1000, @NewStore, l.F1057
    FROM $lnkTableName l
    JOIN @ServerTerms st ON st.Term = l.F1057
    WHERE l.F1056 NOT IN (@NewStore, @HostStore)
      AND NOT EXISTS (
          SELECT 1
          FROM $lnkTableName x
          WHERE x.F1000 = l.F1000
            AND x.F1056 = @NewStore
            AND x.F1057 = l.F1057
      );

    DELETE l
    FROM $lnkTableName l
    JOIN @ServerTerms st ON st.Term = l.F1057
    WHERE l.F1056 NOT IN (@NewStore, @HostStore);
END

INSERT INTO $lnkTableName (F1000, F1056, F1057)
SELECT l.F1000, @NewStore, l.F1057
FROM $lnkTableName l
WHERE l.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(l.F1057))) >= @ProtectedMin
  AND NOT EXISTS (
      SELECT 1
      FROM $lnkTableName x
      WHERE x.F1000 = l.F1000
        AND x.F1056 = @NewStore
        AND x.F1057 = l.F1057
  );

DELETE l
FROM $lnkTableName l
WHERE l.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(l.F1057))) >= @ProtectedMin;

DELETE FROM $lnkTableName
WHERE F1056 NOT IN (@NewStore, @HostStore);

IF NOT EXISTS (SELECT 1 FROM @Codes)
BEGIN
    INSERT INTO @Codes(F1000) VALUES (@Term);
    IF NOT EXISTS (SELECT 1 FROM @Codes WHERE F1000='DSM') INSERT INTO @Codes(F1000) VALUES ('DSM');
    IF NOT EXISTS (SELECT 1 FROM @Codes WHERE F1000='PAL') INSERT INTO @Codes(F1000) VALUES ('PAL');
    IF NOT EXISTS (SELECT 1 FROM @Codes WHERE F1000='RAL') INSERT INTO @Codes(F1000) VALUES ('RAL');
    IF NOT EXISTS (SELECT 1 FROM @Codes WHERE F1000='XAL') INSERT INTO @Codes(F1000) VALUES ('XAL');
END

INSERT INTO $lnkTableName (F1000, F1056, F1057)
SELECT c.F1000, @NewStore, @Term
FROM @Codes c
WHERE NOT EXISTS (
    SELECT 1
    FROM $lnkTableName x
    WHERE x.F1000 = c.F1000
      AND x.F1056 = @NewStore
      AND x.F1057 = @Term
);
"@
	
	$cleanupLnk_ForLocalTerminalDB = @"
DECLARE @NewStore varchar(10) = '$storeNumber';
DECLARE @KeepTerm varchar(10) = '$newTerminal';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';
DECLARE @ProtectedMin int = $protectedMinTerminal;

DELETE l
FROM $lnkTableName l
WHERE l.F1056 = @NewStore
  AND l.F1057 NOT IN (@KeepTerm, @ServerTerm)
  AND (TRY_CONVERT(int, LTRIM(RTRIM(l.F1057))) IS NULL OR TRY_CONVERT(int, LTRIM(RTRIM(l.F1057))) < @ProtectedMin);
"@
	
	$dropViewCommandLnk = @"
IF OBJECT_ID('Lnk_Load', 'V') IS NOT NULL DROP VIEW Lnk_Load;
"@
	
	# ===============================================================================================
	# RUN_TAB (ONLY when NOT server)
	# ===============================================================================================
	
	$createViewCommandRun = @"
IF OBJECT_ID('Run_Load', 'V') IS NOT NULL DROP VIEW Run_Load;
EXEC(N'
CREATE VIEW Run_Load AS
SELECT F1102, F1000, F1104
FROM $runTableName;
');
"@
	
	$updateRunTab_SafeSwap_OldToNew = @"
DECLARE @OldTerm varchar(10) = '$oldTerminal';
DECLARE @NewTerm varchar(10) = '$newTerminal';
DECLARE @Server  varchar(10) = '$serverTerminal';

IF @OldTerm IS NOT NULL AND @OldTerm <> '' AND @OldTerm <> @NewTerm
BEGIN
    UPDATE $runTableName
    SET F1000 = @NewTerm
    WHERE F1000 = @OldTerm;

    UPDATE $runTableName
    SET F1104 = @NewTerm
    WHERE F1104 = @OldTerm
      AND F1104 <> @Server;
END
"@
	
	$dropViewCommandRun = @"
IF OBJECT_ID('Run_Load', 'V') IS NOT NULL DROP VIEW Run_Load;
"@
	
	# ===============================================================================================
	# STD_TAB (store key + CompanyName suffix fix in F1531)
	# ===============================================================================================
	
	$createViewCommandStd = @"
IF OBJECT_ID('Std_Load', 'V') IS NOT NULL DROP VIEW Std_Load;
EXEC(N'
CREATE VIEW Std_Load AS
SELECT F1056, F1531
FROM $stdTableName;
');
"@
	
	$fixStdTabStoreKey_AndCompanySuffix = @"
DECLARE @NewStore  varchar(10) = '$storeNumber';
DECLARE @HostStore varchar(10) = '$hostStoreNumber';

-- =========================================
-- 1) Force store key: keep ONLY (@NewStore + @HostStore)
--    - If @NewStore row exists: delete all other non-host rows
--    - Else: convert ONE non-host row to @NewStore, then delete the rest
-- =========================================
IF EXISTS (SELECT 1 FROM $stdTableName WHERE F1056 = @NewStore)
BEGIN
    DELETE FROM $stdTableName
    WHERE F1056 <> @NewStore
      AND F1056 <> @HostStore;
END
ELSE
BEGIN
    DECLARE @PickOld varchar(10) = NULL;

    SELECT TOP 1 @PickOld = F1056
    FROM $stdTableName
    WHERE F1056 <> @HostStore
    ORDER BY F1056;

    IF @PickOld IS NOT NULL
    BEGIN
        UPDATE $stdTableName
        SET F1056 = @NewStore
        WHERE F1056 = @PickOld;
    END

    DELETE FROM $stdTableName
    WHERE F1056 <> @NewStore
      AND F1056 <> @HostStore;
END

-- =========================================
-- 2) Update F1531 ONLY IF a store number is already present at the end
--    (space/#/- + 3 or 4 digits). If no suffix exists -> do nothing.
--    Robust against NBSP/TAB/CR/LF at the end.
-- =========================================
;WITH X AS
(
    SELECT
        s.F1056,
        s.F1531,
        Clean = RTRIM(
                    REPLACE(REPLACE(REPLACE(REPLACE(
                        LTRIM(RTRIM(s.F1531)),
                        CHAR(160), ' '),
                        CHAR(9),  ' '),
                        CHAR(13), ' '),
                        CHAR(10), ' ')
                )
    FROM STD_TAB s
    WHERE s.F1056 = @NewStore
      AND s.F1531 IS NOT NULL
)
UPDATE X
SET F1531 =
    CASE
        WHEN Clean LIKE '%[ #\-][0-9][0-9][0-9][0-9]'
             AND TRY_CONVERT(int, RIGHT(Clean, 4)) IS NOT NULL
             AND RIGHT(Clean, 4) NOT IN ('0901','0999')
        THEN LEFT(Clean, LEN(Clean) - 4) + @NewStore

        WHEN Clean LIKE '%[ #\-][0-9][0-9][0-9]'
             AND TRY_CONVERT(int, RIGHT(Clean, 3)) IS NOT NULL
             AND RIGHT(Clean, 3) NOT IN ('901','999')
        THEN LEFT(Clean, LEN(Clean) - 3) + @NewStore

        WHEN Clean LIKE '%[ #\-][0-9][0-9]'
             AND TRY_CONVERT(int, RIGHT(Clean, 2)) IS NOT NULL
        THEN LEFT(Clean, LEN(Clean) - 2) + @NewStore

        WHEN Clean LIKE '%[ #\-][0-9]'
             AND TRY_CONVERT(int, RIGHT(Clean, 1)) IS NOT NULL
        THEN LEFT(Clean, LEN(Clean) - 1) + @NewStore

        ELSE F1531
    END
WHERE
    Clean LIKE '%[ #\-][0-9]'
    OR Clean LIKE '%[ #\-][0-9][0-9]'
    OR Clean LIKE '%[ #\-][0-9][0-9][0-9]'
    OR Clean LIKE '%[ #\-][0-9][0-9][0-9][0-9]';
"@
	
	$dropViewCommandStd = @"
IF OBJECT_ID('Std_Load', 'V') IS NOT NULL DROP VIEW Std_Load;
"@
	
	# ===============================================================================================
	# Execute SQL commands (ORDER MATTERS)
	# ===============================================================================================
	
	$sqlCommands = @(
		$createViewCommandTer,
		$moveProtectedTer_ToNewStore
	)
	
	if ($isServerMoniker)
	{
		$sqlCommands += $moveServerNodes902_920_ToNewStore_AndUpdate
		$sqlCommands += $deleteOtherStoresTer
		$sqlCommands += $deleteNonServerRowsForNewStore
		$sqlCommands += $upsertServer901Ter_ForNewStoreOnly
	}
	else
	{
		$sqlCommands += $preserve901ForNonServerStoreChange
		$sqlCommands += $deleteOtherStoresTer
		$sqlCommands += $upsertLaneTer
		
		if ($isLocalTerminalDB)
		{
			$sqlCommands += $cleanupTer_ForLocalTerminalDB
		}
	}
	
	$sqlCommands += @(
		$dropViewCommandTer,
		
		# STO_TAB
		$createViewCommandSto,
		$upsertStoServer901
	)
	
	if ($isServerMoniker)
	{
		$sqlCommands += $cleanupSto_ForServer
	}
	else
	{
		$sqlCommands += $upsertStoTerminal
		if ($isLocalTerminalDB)
		{
			$sqlCommands += $cleanupSto_ForLocalTerminalDB
		}
	}
	
	$sqlCommands += @(
		$dropViewCommandSto,
		
		# LNK_TAB
		$createViewCommandLnk,
		$preserve901LnkForNewStore,
		$rebuildLnkForTerminal
	)
	
	if (-not $isServerMoniker -and $isLocalTerminalDB)
	{
		$sqlCommands += $cleanupLnk_ForLocalTerminalDB
	}
	
	$sqlCommands += @(
		$dropViewCommandLnk
	)
	
	if (-not $isServerMoniker)
	{
		$sqlCommands += @(
			$createViewCommandRun,
			$updateRunTab_SafeSwap_OldToNew,
			$dropViewCommandRun
		)
	}
	
	$sqlCommands += @(
		$createViewCommandStd,
		$fixStdTabStoreKey_AndCompanySuffix,
		$dropViewCommandStd
	)
	
	$allSqlSuccessful = $true
	$failedSqlCommands = @()
	
	foreach ($command in $sqlCommands)
	{
		if (-not (Execute_SQL_Commands -commandText $command))
		{
			$allSqlSuccessful = $false
			$failureMessage = if (-not [string]::IsNullOrWhiteSpace($script:LastSqlExecutionError)) { $script:LastSqlExecutionError } else { "Unknown SQL execution failure." }
			$failedSqlCommands += ("Error: " + $failureMessage + "`r`nSQL:`r`n" + $command)
		}
	}
	
	return @{
		Success		    = $allSqlSuccessful
		FailedCommands  = $failedSqlCommands
		IsServerMoniker = $isServerMoniker
		RunTabUpdated   = (-not $isServerMoniker)
		OldTerminalUsed = $oldTerminal
		NewTerminalUsed = $newTerminal
		DatabaseContext = $DatabaseContext
		LocalTerminalDB = $isLocalTerminalDB
		DBNAME		    = $dbNameDetected
		ConfiguredDBSERVER = if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('ConfiguredDBSERVER')) { [string]$script:FunctionResults['ConfiguredDBSERVER'] } else { $null }
		RuntimeDBSERVER = if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('RuntimeDBSERVER')) { [string]$script:FunctionResults['RuntimeDBSERVER'] } else { $null }
	}
}

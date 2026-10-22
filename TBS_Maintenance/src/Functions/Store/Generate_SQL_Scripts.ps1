# ===================================================================================================
#                              FUNCTION: Generate_SQL_Scripts
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generates the SQL setup/maintenance scripts for lanes and the store server (memory
#   configuration, SQL user, schedules) from the current script variables. The five scripts are
#   held in $script: variables for the in-process consumers; nothing is written to disk.
#
# Parameters:
#   - StoreNumber  [string]: store number (normalized to 3 digits)
#
# Variables:
#   - $script:FunctionResults: reads ConnectionString/DBSERVER/DBNAME
#   - Writes $script:LaneSQLScript, $script:LaneSQLFiltered, $script:LaneSQLScript_Mailslot,
#     $script:ServerSQLScript and $script:ScheduleServerScript for the maintenance schedulers
# ===================================================================================================

function Generate_SQL_Scripts
{
	param (
		[string]$StoreNumber
	)
	
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
	
	# Normalize store number (existing behavior preserved: 3-digit format)
	if (-not [string]::IsNullOrWhiteSpace($StoreNumber) -and $StoreNumber -match '^\d+$')
	{
		$StoreNumber = ([int]$StoreNumber).ToString('000')
	}
	elseif (-not [string]::IsNullOrWhiteSpace($StoreNumber))
	{
		$StoreNumber = $StoreNumber.PadLeft(3, '0')
	}
	
	if (-not $script:FunctionResults.ContainsKey('ConnectionString'))
	{
		Write_Log "Failed to retrieve the connection string." "red"
		return
	}
	

	# ------------------------------------------------------------------------------------------------
	# Build scripts from CURRENT variables every time so the $script: content variables
	# always reflect the live connection/store context for the maintenance consumers.
	# ------------------------------------------------------------------------------------------------
	$ConnectionString = $script:FunctionResults['ConnectionString']
	$defaultStoreDbName = "STORESQL"
	$defaultLaneDbName = "LANESQL"
	$dbServer = $script:FunctionResults['DBSERVER']
	
	if ($script:FunctionResults.ContainsKey('DBNAME') -and -not [string]::IsNullOrWhiteSpace($script:FunctionResults['DBNAME']))
	{
		$storeDbName = $script:FunctionResults['DBNAME']
	}
	else
	{
		Write_Log "No 'DBNAME' in FunctionResults. Defaulting to '$defaultStoreDbName'." "yellow"
		$storeDbName = $defaultStoreDbName
	}
	
	$laneDbName = $defaultLaneDbName
	
	# ------------------ Lane script ------------------
	$LaneSQLScript = @"
/* Set a long timeout so the entire script runs */
@WIZRPL(DBASE_TIMEOUT=E);

/* Set memory configuration */
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', 1024;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;

/* Create "$BackupSqlUser" user in the database */
IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = '$BackupSqlUser')
    CREATE LOGIN [$BackupSqlUser] WITH PASSWORD = '$BackupSqlPass', CHECK_POLICY = OFF;
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = '$BackupSqlUser')
    CREATE USER [$BackupSqlUser] FOR LOGIN [$BackupSqlUser];
IF IS_ROLEMEMBER('db_owner', '$BackupSqlUser') = 0
    EXEC sp_addrolemember 'db_owner', '$BackupSqlUser';

/* Truncate unnecessary tables */
IF OBJECT_ID('COST_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('COST_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE COST_REV;
IF OBJECT_ID('POS_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('POS_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE POS_REV;
IF OBJECT_ID('OBJ_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('OBJ_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE OBJ_REV;
IF OBJECT_ID('PRICE_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE PRICE_REV;
IF OBJECT_ID('REV_HDR', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('REV_HDR', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE REV_HDR;
IF OBJECT_ID('SAL_REG_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_REG_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_REG_SAV;
IF OBJECT_ID('SAL_HDR_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_HDR_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_HDR_SAV;
IF OBJECT_ID('SAL_TTL_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_TTL_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_TTL_SAV;
IF OBJECT_ID('SAL_DET_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_DET_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_DET_SAV;

/* Truncate PRICE_EVENT table for records older than 7 days */
IF OBJECT_ID('PRICE_EVENT','U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_EVENT','OBJECT','DELETE') = 1 DELETE FROM PRICE_EVENT WHERE F254 < DATEADD(DAY,-7,GETDATE());

/* Drop specific tables older than 30 days */
DECLARE @cmd varchar(4000) 
DECLARE cmds CURSOR FOR 
SELECT 'drop table [' + name + ']' 
FROM sys.tables 
WHERE (name LIKE 'TMP_%' OR name LIKE 'MSVHOST%' OR name LIKE 'MMPHOST%' OR name LIKE 'M$StoreNumber%' OR name LIKE 'R$StoreNumber%') 
OPEN cmds 
WHILE 1 = 1 
BEGIN 
FETCH cmds INTO @cmd 
IF @@fetch_status != 0 BREAK 
EXEC(@cmd) 
END 
CLOSE cmds; 
DEALLOCATE cmds;

/* Cleaning HEADER_SAV */
IF OBJECT_ID('HEADER_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('HEADER_SAV', 'OBJECT', 'DELETE') = 1 
    DELETE FROM HEADER_SAV 
    WHERE (F903 = 'SVHOST' OR F903 = 'MPHOST' OR F903 = CONCAT('M', '$StoreNumber', '901')) 
    AND (DATEDIFF(DAY, F907, GETDATE()) > 30 OR DATEDIFF(DAY, F909, GETDATE()) > 30);

/* Delete bad SMS items */
@dbEXEC(DELETE FROM OBJ_TAB WHERE F01='0020000000000') 
@dbEXEC(DELETE FROM OBJ_TAB WHERE F01 LIKE '% %') 
@dbEXEC(DELETE FROM OBJ_TAB WHERE LEN(F01)<>13) 
@dbEXEC(DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000') 
@dbEXEC(DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=0 AND SUBSTRING(F01,9,5) = '00000') 
@dbEXEC(DELETE FROM POS_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)) 
@dbEXEC(DELETE FROM PRICE_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)) 
@dbEXEC(DELETE FROM COST_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)) 
@dbEXEC(DELETE FROM SCL_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)) 
@dbEXEC(DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) <> '002') 
@dbEXEC(DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000') 
@dbEXEC(UPDATE SCL_TAB SET SCL_TAB.F267 = SCL_TXT.F267 FROM SCL_TAB SCL JOIN SCL_TXT_TAB SCL_TXT ON (SCL.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000'))) 
@dbEXEC(UPDATE SCL_TAB SET SCL_TAB.F268 = SCL_NUT.F268 FROM SCL_TAB SCL JOIN SCL_NUT_TAB SCL_NUT ON (SCL.F01=CONCAT('002',FORMAT(SCL_NUT.F268,'00000'),'00000'))) 
@dbEXEC(DELETE FROM SCL_TXT_TAB WHERE F267 NOT IN (SELECT F267 FROM SCL_TAB)) 
@dbEXEC(DELETE FROM SCL_NUT_TAB WHERE F268 NOT IN (SELECT F268 FROM SCL_TAB)) 
@dbEXEC(UPDATE SCL_TAB SET F267 = NULL WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F267,'00000'),'00000') FROM SCL_TXT_TAB)) 
@dbEXEC(UPDATE SCL_TAB SET F268 = NULL WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F268,'00000'),'00000') FROM SCL_NUT_TAB)) 
@dbEXEC(UPDATE SCL_TXT_TAB SET SCL_TXT_TAB.F04 = POS.F04 FROM SCL_TXT_TAB SCL_TXT JOIN POS_TAB POS ON (POS.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')) WHERE ISNUMERIC(SCL_TXT.F04)=0) 
@dbEXEC(UPDATE SCL_TAB SET F256 = REPLACE(REPLACE(REPLACE(F256, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TAB SET F1952 = REPLACE(REPLACE(REPLACE(F1952, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TAB SET F2581 = REPLACE(REPLACE(REPLACE(F2581, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TAB SET F2582 = REPLACE(REPLACE(REPLACE(F2582, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')) 
@dbEXEC(UPDATE SCL_TXT_TAB SET F297 = REPLACE(REPLACE(REPLACE(F297, CHAR(13),' '), CHAR(10),' '), CHAR(9),' '))

/* Delete batches older than 7 days from the lane */
DELETE FROM Header_bat WHERE F909 < DATEADD(day, -7, GETDATE());
DELETE FROM Header_dct WHERE F909 < DATEADD(day, -7, GETDATE());
DELETE FROM Header_old WHERE F909 < DATEADD(day, -7, GETDATE());
DELETE FROM Header_sav WHERE F909 < DATEADD(day, -7, GETDATE());

/* Rebuild indexes and update database statistics */
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD'
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN'

/* Shrink the main database file */
DBCC SHRINKFILE ($laneDbName)

/* Shrink the database log file */
DBCC SHRINKFILE (${laneDbName}_Log)

/* Restrict the indefinite log file growth */
ALTER DATABASE $laneDbName SET RECOVERY SIMPLE

/* Clear the long database timeout */
@WIZCLR(DBASE_TIMEOUT);
"@
	
	$script:LaneSQLScript = $LaneSQLScript
	
	# ------------------ Filtered lane script ------------------
	$ServerMemoryConfig = @"
DECLARE @Memory25PercentMB BIGINT;
SELECT @Memory25PercentMB = (total_physical_memory_kb / 1024) * 25 / 100
FROM sys.dm_os_sys_memory;
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', @Memory25PercentMB;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;
"@
	
	$sectionPattern = '(?s)/\*\s*(?<SectionName>[^*/]+?)\s*\*/\s*(?<SQLCommands>(?:.(?!/\*)|.)*?)(?=(/\*|$))'
	$sectionsToSkip = @(
		'Set a long timeout so the entire script runs',
		'Clear the long database timeout'
	)
	
	$LaneSQLFiltered = ""
	$matches = [regex]::Matches($LaneSQLScript, $sectionPattern)
	
	foreach ($match in $matches)
	{
		$sectionName = $match.Groups['SectionName'].Value.Trim()
		$sqlCommands = $match.Groups['SQLCommands'].Value.Trim()
		
		if ($sectionsToSkip -contains $sectionName)
		{
			continue
		}
		elseif ($sectionName -eq 'Set memory configuration')
		{
			$LaneSQLFiltered += "/* $sectionName */`r`n$ServerMemoryConfig`r`n`r`n"
		}
		else
		{
			$sqlCommands = $sqlCommands -replace '@dbEXEC\((?:\"(.*)\"|(.*))\)', '$1$2'
			$LaneSQLFiltered += "/* $sectionName */`r`n$sqlCommands`r`n`r`n"
		}
	}
	
	$script:LaneSQLFiltered = $LaneSQLFiltered
	
	# ------------------ Mailslot-friendly lane script ------------------
	$lines = $script:LaneSQLScript -split "`r?`n"
	$macroPattern = '^\s*(@|/|\*)'
	
	$fixedLines = foreach ($line in $lines)
	{
		if ($line -match $macroPattern -or [string]::IsNullOrWhiteSpace($line))
		{
			$line
		}
		else
		{
			"@EXEC($line)"
		}
	}
	
	$script:LaneSQLScript_Mailslot = ($fixedLines -join "`r`n")
	
	# ------------------ Server script ------------------
	$ServerSQLScript = @"
/* Set memory configuration */
DECLARE @Memory25PercentMB BIGINT;
SELECT @Memory25PercentMB = (total_physical_memory_kb / 1024) * 25 / 100
FROM sys.dm_os_sys_memory;
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', @Memory25PercentMB;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;

/* Create "$BackupSqlUser" user in the database */
IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = '$BackupSqlUser')
    CREATE LOGIN [$BackupSqlUser] WITH PASSWORD = '$BackupSqlPass', CHECK_POLICY = OFF;
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = '$BackupSqlUser')
    CREATE USER [$BackupSqlUser] FOR LOGIN [$BackupSqlUser];
IF IS_ROLEMEMBER('db_owner', '$BackupSqlUser') = 0
    EXEC sp_addrolemember 'db_owner', '$BackupSqlUser';

/* Create-Refresh ScaleCommApp Triggers */
IF OBJECT_ID('dbo.TBS_ITM_SMAppUPDATED', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.TBS_ITM_SMAppUPDATED (
        Id INT IDENTITY(1,1) PRIMARY KEY,
        CodeF01 VARCHAR(13) NULL,
        Sent BIT NOT NULL DEFAULT 0,
        SentAt DATETIME NOT NULL DEFAULT GETDATE()
    );
END

IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.TBS_ITM_SMAppUPDATED') AND name = 'IDX_TBS_ITM_SMAppUPDATED_CodeF01')
    DROP INDEX IDX_TBS_ITM_SMAppUPDATED_CodeF01 ON dbo.TBS_ITM_SMAppUPDATED;
IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.TBS_ITM_SMAppUPDATED') AND name = 'IDX_TBS_ITM_SMAppUPDATED_Sent')
    DROP INDEX IDX_TBS_ITM_SMAppUPDATED_Sent ON dbo.TBS_ITM_SMAppUPDATED;
IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.TBS_ITM_SMAppUPDATED') AND name = 'IDX_TBS_ITM_SMAppUPDATED_SentAt')
    DROP INDEX IDX_TBS_ITM_SMAppUPDATED_SentAt ON dbo.TBS_ITM_SMAppUPDATED;

IF OBJECT_ID('dbo.SMApp_UpdateOBJ', 'TR') IS NOT NULL
    DROP TRIGGER dbo.SMApp_UpdateOBJ
GO
IF OBJECT_ID('dbo.SMApp_UpdatePOS', 'TR') IS NOT NULL
    DROP TRIGGER dbo.SMApp_UpdatePOS
GO
IF OBJECT_ID('dbo.SMApp_UpdatePrice', 'TR') IS NOT NULL
    DROP TRIGGER dbo.SMApp_UpdatePrice
GO
IF OBJECT_ID('dbo.SMApp_UpdateSCL_TXT', 'TR') IS NOT NULL
    DROP TRIGGER dbo.SMApp_UpdateSCL_TXT
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[SMApp_UpdateOBJ]
   ON  [dbo].[OBJ_TAB]
   AFTER INSERT,UPDATE
AS 
BEGIN
       SET NOCOUNT ON;

	INSERT INTO dbo.TBS_ITM_SMAppUPDATED (CodeF01,Sent,SentAt)
	SELECT F01,0, GETDATE() FROM inserted WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=1 AND SUBSTRING(F01,9,5) = '00000'
END
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[SMApp_UpdatePOS]
   ON  [dbo].[POS_TAB]
   AFTER INSERT,UPDATE
AS 
BEGIN
		SET NOCOUNT ON;
		INSERT INTO dbo.TBS_ITM_SMAppUPDATED (CodeF01,Sent,SentAt)
		SELECT F01,0, GETDATE() FROM inserted WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=1 AND SUBSTRING(F01,9,5) = '00000' 
END
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[SMApp_UpdatePrice]
   ON  [dbo].[PRICE_TAB]
   AFTER INSERT,UPDATE
AS 
BEGIN
		SET NOCOUNT ON;
		INSERT INTO dbo.TBS_ITM_SMAppUPDATED (CodeF01,Sent,SentAt)
		SELECT F01,0, GETDATE() FROM inserted WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=1 AND SUBSTRING(F01,9,5) = '00000' 
END
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[SMApp_UpdateSCL_TXT]
ON [dbo].[SCL_TXT_TAB]
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;

    INSERT INTO dbo.TBS_ITM_SMAppUPDATED (CodeF01, Sent, SentAt)
    SELECT
        '002' + RIGHT('00000' + CONVERT(varchar(5), inserted.F267), 5) + '00000',
        0,
        GETDATE()
    FROM inserted
    JOIN dbo.OBJ_TAB
        ON '002' + RIGHT('00000' + CONVERT(varchar(5), inserted.F267), 5) + '00000' = OBJ_TAB.F01
    WHERE SUBSTRING(OBJ_TAB.F01,1,3) = '002'
      AND ISNUMERIC(SUBSTRING(OBJ_TAB.F01,4,5)) = 1
      AND SUBSTRING(OBJ_TAB.F01,9,5) = '00000';
END
GO

CREATE INDEX IDX_TBS_ITM_SMAppUPDATED_CodeF01 ON dbo.TBS_ITM_SMAppUPDATED (CodeF01);
CREATE INDEX IDX_TBS_ITM_SMAppUPDATED_Sent ON dbo.TBS_ITM_SMAppUPDATED (Sent);
CREATE INDEX IDX_TBS_ITM_SMAppUPDATED_SentAt ON dbo.TBS_ITM_SMAppUPDATED (SentAt);
GO

/* Truncate unnecessary tables */
IF OBJECT_ID('COST_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('COST_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE COST_REV;
IF OBJECT_ID('POS_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('POS_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE POS_REV;
IF OBJECT_ID('OBJ_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('OBJ_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE OBJ_REV;
IF OBJECT_ID('PRICE_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE PRICE_REV;
IF OBJECT_ID('REV_HDR', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('REV_HDR', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE REV_HDR;
IF OBJECT_ID('SAL_REG_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_REG_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_REG_SAV;
IF OBJECT_ID('SAL_HDR_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_HDR_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_HDR_SAV;
IF OBJECT_ID('SAL_TTL_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_TTL_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_TTL_SAV;
IF OBJECT_ID('SAL_DET_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_DET_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_DET_SAV;

/* Delete bad SMS items */
DELETE FROM OBJ_TAB WHERE F01='0020000000000';
DELETE FROM OBJ_TAB WHERE F01 LIKE '% %';
DELETE FROM OBJ_TAB WHERE LEN(F01)<>13;
DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000';
DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=0 AND SUBSTRING(F01,9,5) = '00000';
DELETE FROM POS_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB);
DELETE FROM PRICE_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB);
DELETE FROM COST_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB);
DELETE FROM SCL_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB);
DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) <> '002';
DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000';
UPDATE SCL_TAB SET SCL_TAB.F267 = SCL_TXT.F267 ,SCL_TAB.F1001=1 FROM SCL_TAB SCL JOIN SCL_TXT_TAB SCL_TXT ON (SCL.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000'));
UPDATE SCL_TAB SET SCL_TAB.F268 = SCL_NUT.F268 ,SCL_TAB.F1001=1 FROM SCL_TAB SCL JOIN SCL_NUT_TAB SCL_NUT ON (SCL.F01=CONCAT('002',FORMAT(SCL_NUT.F268,'00000'),'00000'));
DELETE FROM SCL_TXT_TAB WHERE F267 NOT IN (SELECT F267 FROM SCL_TAB);
DELETE FROM SCL_NUT_TAB WHERE F268 NOT IN (SELECT F268 FROM SCL_TAB);
UPDATE SCL_TAB SET F267 = NULL, F1001 = 1 WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F267,'00000'),'00000') FROM SCL_TXT_TAB);
UPDATE SCL_TAB SET F268 = NULL, F1001 = 1 WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F268,'00000'),'00000') FROM SCL_NUT_TAB);
UPDATE SCL_TXT_TAB SET SCL_TXT_TAB.F04 = POS.F04, SCL_TXT_TAB.F1001=1 FROM SCL_TXT_TAB SCL_TXT JOIN POS_TAB POS ON (POS.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')) WHERE ISNUMERIC(SCL_TXT.F04)=0;
UPDATE SCL_TAB SET F256 = REPLACE(REPLACE(REPLACE(F256, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ');
UPDATE SCL_TAB SET F1952 = REPLACE(REPLACE(REPLACE(F1952, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ');
UPDATE SCL_TAB SET F2581 = REPLACE(REPLACE(REPLACE(F2581, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ');
UPDATE SCL_TAB SET F2582 = REPLACE(REPLACE(REPLACE(F2582, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ');
UPDATE SCL_TXT_TAB SET F297 = REPLACE(REPLACE(REPLACE(F297, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ');

/* Truncate PRICE_EVENT table for records older than 7 days */
IF OBJECT_ID('PRICE_EVENT','U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_EVENT','OBJECT','DELETE') = 1 DELETE FROM PRICE_EVENT WHERE F254 < DATEADD(DAY,-7,GETDATE());

/* Drop specific tables older than 30 days */
DECLARE @cmd varchar(4000);
DECLARE cmds CURSOR FOR
SELECT 'drop table [' + name + ']'
FROM sys.tables
WHERE (name LIKE 'TMP_%' OR name LIKE 'MSVHOST%' OR name LIKE 'MMPHOST%' OR name LIKE 'M$StoreNumber%' OR name LIKE 'R$StoreNumber%') AND DATEDIFF(DAY, create_date, GETDATE()) > 30;
OPEN cmds;
WHILE 1 = 1
BEGIN
    FETCH cmds INTO @cmd;
    IF @@fetch_status != 0 BREAK;
    EXEC(@cmd);
END;
CLOSE cmds;
DEALLOCATE cmds;

/* Rebuild indexes and update database statistics */
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD';
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN';

/* Set DB to highest compatibility level supported by this SQL Server version */
DECLARE @ProductVersion varchar(128), @MajorVersion int, @CompatLevel int, @sql nvarchar(4000);
SET @ProductVersion = CONVERT(varchar(128), SERVERPROPERTY('ProductVersion'));
SET @MajorVersion = CONVERT(int, LEFT(@ProductVersion, CHARINDEX('.', @ProductVersion + '.') - 1));
SET @CompatLevel = CASE @MajorVersion WHEN 17 THEN 170 WHEN 16 THEN 160 WHEN 15 THEN 150 WHEN 14 THEN 140 WHEN 13 THEN 130 WHEN 12 THEN 120 WHEN 11 THEN 110 WHEN 10 THEN 100 WHEN 9 THEN 90 ELSE NULL END;
IF @CompatLevel IS NOT NULL
BEGIN
    SET @sql = N'ALTER DATABASE ' + QUOTENAME(N'$storeDbName') + N' SET COMPATIBILITY_LEVEL = ' + CONVERT(varchar(3), @CompatLevel);
    EXEC(@sql);
END

/* Shrink the main database file */
DBCC SHRINKFILE ($storeDbName);

/* Shrink the database log file */
DBCC SHRINKFILE (${storeDbName}_Log);

/* Restrict the indefinite log file growth */
ALTER DATABASE $storeDbName SET RECOVERY SIMPLE;
"@
	
	$script:ServerSQLScript = $ServerSQLScript
	
	# ------------------ Scheduled maintenance server script ------------------
	$ScheduleServerScript = @"
/* Set a long timeout so the entire script runs */
@WIZRPL(DBASE_TIMEOUT=E);

/* Set memory configuration */
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', 8192;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;

/* Truncate unnecessary tables */
IF OBJECT_ID('COST_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('COST_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE COST_REV;
IF OBJECT_ID('POS_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('POS_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE POS_REV;
IF OBJECT_ID('OBJ_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('OBJ_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE OBJ_REV;
IF OBJECT_ID('PRICE_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE PRICE_REV;
IF OBJECT_ID('REV_HDR', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('REV_HDR', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE REV_HDR;
IF OBJECT_ID('SAL_REG_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_REG_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_REG_SAV;
IF OBJECT_ID('SAL_HDR_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_HDR_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_HDR_SAV;
IF OBJECT_ID('SAL_TTL_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_TTL_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_TTL_SAV;
IF OBJECT_ID('SAL_DET_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_DET_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_DET_SAV;

/* Drop specific tables older than 30 days */
DECLARE @cmd varchar(4000) 
DECLARE cmds CURSOR FOR 
SELECT 'drop table [' + name + ']' 
FROM sys.tables 
WHERE (name LIKE 'TMP_%' OR name LIKE 'MSVHOST%' OR name LIKE 'MMPHOST%' OR name LIKE 'M$StoreNumber%' OR name LIKE 'R$StoreNumber%') 
OPEN cmds 
WHILE 1 = 1 
BEGIN 
FETCH cmds INTO @cmd 
IF @@fetch_status != 0 BREAK 
EXEC(@cmd) 
END 
CLOSE cmds; 
DEALLOCATE cmds;

/* Shrink database and log files */
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD'
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN'
DBCC SHRINKFILE ($storeDbName)
DBCC SHRINKFILE (${storeDbName}_Log)
ALTER DATABASE $storeDbName SET RECOVERY SIMPLE

/* Clear the long database timeout */
@WIZCLR(DBASE_TIMEOUT);
"@
	
	$script:ScheduleServerScript = $ScheduleServerScript
	

}

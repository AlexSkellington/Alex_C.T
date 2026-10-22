# ===================================================================================================
#                              FUNCTION: Install_FUNCTIONS_Into_SMS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Installs the SMS deploy artifacts into the Office folder: DEPLOY_SYS.sql (with the
#   "Multiple Functions" menu entry) and DEPLOY_MULTI_FCT.sqm (CSV / ranges like 123,234,300-305).
#   Both files are written as ANSI (Windows-1252), CRLF line endings, no BOM.
#
# Parameters:
#   - StoreNumber  [string]: kept for signature compatibility (not used in content)
#   - OfficePath   [string]: Office folder (default: $script:OfficePath, else <BasePath>\Office)
#   - Silent       [switch]: suppress Write_Log output
#
# Writes:
#   - <OfficePath>\DEPLOY_SYS.sql and <OfficePath>\DEPLOY_MULTI_FCT.sqm
# ===================================================================================================

function Install_FUNCTIONS_Into_SMS
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$OfficePath,
		[switch]$Silent
	)

	$writeLog = {
		param (
			[string]$Message,
			[string]$Color = "white"
		)
		if (-not $Silent) { Write_Log $Message $Color }
	}

	& $writeLog "`r`n==================== Starting Install_FUNCTIONS_Into_SMS ====================`r`n" "blue"

	# Registered-sign macro marker (U+00AE) built from its code point so this source stays ASCII;
	# the files are written as Windows-1252, where it is the single byte 0xAE SMS expects.
	$reg = [char]0x00AE

	# --------------------------------------------------------------------------------------------------------------
	# Resolve/validate OfficePath (prefer explicit param; fallback to common script-scoped variables)
	# --------------------------------------------------------------------------------------------------------------
	if (-not $OfficePath -or [string]::IsNullOrWhiteSpace($OfficePath))
	{
		# (the storeman ROOT is not the Office folder - before v2.7.0 this fell back to BasePath itself)
		if ($script:OfficePath) { $OfficePath = $script:OfficePath }
		elseif ($script:BasePath) { $OfficePath = Join-Path $script:BasePath 'Office' }
	}
	if (-not $OfficePath -or -not (Test-Path -LiteralPath $OfficePath))
	{
		& $writeLog "Office path not found or not provided: '$OfficePath'." "red"
		return
	}

	# --- Destination paths ---
	$DeploySysPath = Join-Path -Path $OfficePath -ChildPath "DEPLOY_SYS.sql"
	$DeployMultiFctPath = Join-Path -Path $OfficePath -ChildPath "DEPLOY_MULTI_FCT.sqm"

	# --- Encoding: ANSI (Windows-1252), no BOM ---
	$ansi = [System.Text.Encoding]::GetEncoding(1252)

	# ===============================================================================================================
	#                                             DEPLOY_SYS.sql (ALWAYS includes "Multiple Functions")
	# ===============================================================================================================
	$DeploySysContent = @"
@FMT(CMP,@dbHot(FINDFIRST,UD_DEPLOY_SYS.SQL)=,$($reg)WIZRPL(UD_RUN=0));
@FMT(CMP,@WIZGET(UD_RUN)=,'$($reg)EXEC(SQL=UD_DEPLOY_SYS)$($reg)FMT(CHR,27)');

@FMT(CMP,@TOOLS(MESSAGEDLG,"!TO KEEP THE LANE'S REFERENCE SAMPLE UP TO DATE YOU SHOULD USE THE "REFERENCE SAMPLE MECHANISM". DO YOU WANT TO CONTINUE?",,NO,YES)=1,'$($reg)FMT(CHR,27)');

@EXEC(INI=HOST_OFFICE[DEPLOY_SYS]);

@WIZRPL(STYLE=SIL);
@WIZRPL(TARGET_FILTER=@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE));

@EXEC(sqi=USERB_DEPLOY_SYS);

@WIZINIT;
@WIZMENU(ONESQM=What do you want to send,
    Functions (NEW)=DEPLOY_MULTI_FCT,
    Function link=fcz_load,
    Totalizer=tlz_load,
    Drill down pages=dril_page_load,
    Drill down files=dril_file_load,
    All system=ALL);
@WIZDISPLAY;

@WIZINIT;
@WIZTARGET(TARGET=,@FMT(CMP,"@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)=","
SELECT F1000,F1018 FROM STO_TAB WHERE F1181=1 ORDER BY F1000","
SELECT DISTINCT STO.F1000,STO.F1018 
FROM LNK_TAB LN2 JOIN LNK_TAB LNK ON LN2.F1056=LNK.F1056 AND LN2.F1057=LNK.F1057
JOIN STO_TAB STO ON STO.F1000=LNK.F1000 
WHERE STO.F1181='1' AND LN2.F1000='@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)'
ORDER BY STO.F1000"));
@WIZDISPLAY;

@FMT(CMP,@dbSelect(select distinct 1 from lnk_tab where F1000='@Wizget(Target)' and f1056='999')=,,"$($reg)EXEC(msg=!*****_can_not_deploy_system_tables_to_a_host_****);$($reg)FMT(CHR,27);")

@WIZINIT;
@WIZMENU(ACTION=Action on the target database,Add or replace=ADDRPL,Add only=ADD,Replace only=UPDATE,Clean and load=LOAD);
@WIZDISPLAY;

/* SEND ONLY ONE / MULTI / OR OTHERS */

@FMT(CMP,@wizget(ONESQM)=tlz_load,$($reg)EXEC(SQM=tlz_load));
@FMT(CMP,@wizget(ONESQM)=fcz_load,$($reg)EXEC(SQM=fcz_load));
@FMT(CMP,@wizget(ONESQM)=fct_load,$($reg)EXEC(SQM=fct_load));
@FMT(CMP,@wizget(ONESQM)=dril_file_load,$($reg)EXEC(SQM=DRIL_FILE_LOAD));
@FMT(CMP,@wizget(ONESQM)=dril_page_load,$($reg)EXEC(SQM=DRIL_PAGE_LOAD));
@FMT(CMP,@wizget(ONESQM)=DEPLOY_MULTI_FCT,$($reg)EXEC(SQM=DEPLOY_MULTI_FCT));

@FMT(CMP,@WIZGET(ONESQM)=ALL,,'$($reg)EXEC(SQM=exe_activate_accept_sys)$($reg)fmt(chr,27)');

@FMT(CMP,@wizget(tlz_load)=0,,$($reg)EXEC(SQM=tlz_load));
@FMT(CMP,@wizget(fcz_load)=0,,$($reg)EXEC(SQM=fcz_load));
@FMT(CMP,@wizget(fct_load)=0,,$($reg)EXEC(SQM=fct_load));
@FMT(CMP,@wizget(DRIL_FILE_LOAD)=0,,$($reg)EXEC(SQM=DRIL_FILE_LOAD));
@FMT(CMP,@wizget(DRIL_PAGE_LOAD)=0,,$($reg)EXEC(SQM=DRIL_PAGE_LOAD));
@FMT(CMP,@wizget(DEPLOY_MULTI_FCT)=0,,$($reg)EXEC(SQM=DEPLOY_MULTI_FCT));

@FMT(CMP,@wizget(exe_activate_accept_all)=0,,$($reg)EXEC(SQM=exe_activate_accept_sys));
@FMT(CMP,@wizget(exe_refresh_menu)=1,$($reg)EXEC(SQM=exe_refresh_menu));

@EXEC(sqi=USERE_DEPLOY_SYS);
"@

	# ===============================================================================================================
	#                                      DEPLOY_MULTI_FCT.sqm (template)
	#   Supports CSV and ranges (e.g., 123,234,300-305).
	# ===============================================================================================================
	$DeployMultiFctContent = @"
INSERT INTO HEADER_DCT VALUES
('HC','00000001','001901','001001',,,1997001,0000,1997001,0001,,'LOAD','CREATE DCT',,,,,,'1/1.0','V1.0',,);

CREATE TABLE FCT_DCT(@MAP_FROM_QUERY);

INSERT INTO HEADER_DCT VALUES
('HM','00000001','001901','001001',,,1997001,0000,1997001,0001,,'@WIZGET(ACTION)','@WIZGET(ACTION) SELECTED FUNCTIONS',,,,,,'1/1.0','V1.0','F1063',);

CREATE VIEW FCT_CHG AS SELECT @FIELDS_FROM_QUERY FROM FCT_DCT;

INSERT INTO FCT_CHG VALUES

/* EXTRACT SECTION */

@DBHOT(HOT_WIZ,PARAMTOLINE,PARAMSAV_DEPLOY_MULTI_FCT);
@FMT(CMP,'@WIZGET(TARGET)<>','$($reg)WIZRPL(TARGET_FILTER=@WIZGET(TARGET))');

@WIZINIT;
@WIZEDIT(FCT_LIST=Type ''ALL'' to deploy all functions,'Function IDs/ranges (122,129;122-129)');
@WIZDISPLAY;

@WIZINIT;
@WIZTARGET(TARGET_FILTER=,@FMT(CMP,"@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)=","
SELECT F1000,F1018 FROM STO_TAB WHERE F1181=1","
SELECT DISTINCT STO.F1000,STO.F1018
FROM LNK_TAB LN2 JOIN LNK_TAB LNK ON LN2.F1056=LNK.F1056 AND LN2.F1057=LNK.F1057
JOIN STO_TAB STO ON STO.F1000=LNK.F1000
WHERE STO.F1181='1' AND LN2.F1000='@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)'"));
@WIZDISPLAY;

@WIZINIT;
@WIZMENU(ACTION=Action on the target database,Add or replace=ADDRPL,Add only=ADD,Replace only=UPDATE,Create and load=LOAD);
@WIZDISPLAY;

@WIZSET(STYLE=SIL);
@WIZCLR(TARGET);
@WIZSET(FORCE_F1000=@F1056);

@dbEXEC(SET ANSI_NULLS ON)
@dbEXEC(SET QUOTED_IDENTIFIER ON)
@dbEXEC(SET ANSI_PADDING ON)
@dbEXEC(SET ANSI_WARNINGS ON)
@dbEXEC(SET CONCAT_NULL_YIELDS_NULL ON)
@dbEXEC(SET NUMERIC_ROUNDABORT OFF)
@dbEXEC(SET ARITHABORT ON)

@MAP_DEPLOY
SELECT FCT.F1056,FCT.F1056+FCT.F1057 AS F1000,@dbFld(FCT_TAB,FCT.,F1000) FROM
    (SELECT LNI.F1056,LNI.F1057,FCT.*,ROW_NUMBER() OVER (PARTITION BY FCT.F1063,LNI.F1056,LNI.F1057 ORDER BY CASE WHEN FCT.F1000='PAL' THEN 1 ELSE 2 END DESC) AS F1301
    FROM FCT_TAB FCT
    JOIN LNK_TAB LNI ON FCT.F1000=LNI.F1000
    JOIN LNK_TAB LNO ON LNI.F1056=LNO.F1056 AND LNI.F1057=LNO.F1057
    WHERE LNO.F1000='@WIZGET(TARGET_FILTER)' AND (
          UPPER(LTRIM(RTRIM('@WIZGET(FCT_LIST)'))) = 'ALL'
          OR EXISTS
          (
            SELECT 1
            FROM
            (
              SELECT
                CASE WHEN CHARINDEX('-',token)=0 AND token IS NOT NULL AND PATINDEX('%[^0-9]%',token)=0 THEN TRY_CAST(token AS INT) ELSE NULL END AS SingleNum,
                CASE WHEN CHARINDEX('-',token)>0 AND PATINDEX('%[^0-9]%',LEFT(token,CHARINDEX('-',token)-1))=0 AND PATINDEX('%[^0-9]%',SUBSTRING(token,CHARINDEX('-',token)+1,8000))=0 THEN TRY_CAST(LEFT(token,CHARINDEX('-',token)-1) AS INT) ELSE NULL END AS StartNum,
                CASE WHEN CHARINDEX('-',token)>0 AND PATINDEX('%[^0-9]%',LEFT(token,CHARINDEX('-',token)-1))=0 AND PATINDEX('%[^0-9]%',SUBSTRING(token,CHARINDEX('-',token)+1,8000))=0 THEN TRY_CAST(SUBSTRING(token,CHARINDEX('-',token)+1,8000) AS INT) ELSE NULL END AS EndNum
              FROM
              (
                SELECT T.N.value('.','nvarchar(100)') AS token
                FROM
                (
                  SELECT CAST('<i>'+REPLACE(REPLACE(REPLACE(REPLACE('@WIZGET(FCT_LIST)',CHAR(13),''),CHAR(10),''),' ',''),',','</i><i>')+'</i>' AS XML) AS xdoc
                ) X
                CROSS APPLY xdoc.nodes('/i') AS T(N)
              ) S
              WHERE ISNULL(token,'')<>''
            ) FF
            WHERE
                   (FF.SingleNum IS NOT NULL AND FCT.F1063 IS NOT NULL AND ISNUMERIC(CONVERT(VARCHAR(50),FCT.F1063))=1 AND CAST(CONVERT(VARCHAR(50),FCT.F1063) AS INT)=FF.SingleNum)
                OR (FF.StartNum IS NOT NULL AND FF.EndNum IS NOT NULL AND FF.EndNum>=FF.StartNum AND FCT.F1063 IS NOT NULL AND ISNUMERIC(CONVERT(VARCHAR(50),FCT.F1063))=1 AND CAST(CONVERT(VARCHAR(50),FCT.F1063) AS INT) BETWEEN FF.StartNum AND FF.EndNum)
          )
        )
    ) FCT
WHERE FCT.F1301=1
ORDER BY F1000,F1063;

/* RESTORE INITITAL PARAMETER POOL */
@WIZRESET;
@DBHOT(HOT_WIZ,LINETOPARAM,PARAMSAV_DEPLOY_MULTI_FCT);
@DBHOT(HOT_WIZ,CLR,PARAMSAV_DEPLOY_MULTI_FCT);
"@

	# ===============================================================================================================
	#                                                  WRITE FILES
	#   normalize to CRLF -> write ANSI (no BOM) -> clear read-only/hidden attributes -> log
	# ===============================================================================================================

	# -- DEPLOY_SYS.sql --
	try
	{
		$norm = [regex]::Replace($DeploySysContent, "(`r)?`n", "`r`n") # normalize to CRLF
		[System.IO.File]::WriteAllText($DeploySysPath, $norm, $ansi) # write as ANSI, no BOM
		Set-ItemProperty -LiteralPath $DeploySysPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		& $writeLog "Updated 'DEPLOY_SYS.sql' in '$OfficePath'." "green"
	}
	catch
	{
		& $writeLog "Failed to write 'DEPLOY_SYS.sql'. Error: $_" "red"
	}

	# -- DEPLOY_MULTI_FCT.sqm --
	try
	{
		$norm = [regex]::Replace($DeployMultiFctContent, "(`r)?`n", "`r`n") # normalize to CRLF
		[System.IO.File]::WriteAllText($DeployMultiFctPath, $norm, $ansi) # write as ANSI, no BOM
		Set-ItemProperty -LiteralPath $DeployMultiFctPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal)
		& $writeLog "Wrote 'DEPLOY_MULTI_FCT.sqm' (All + CSV + ranges) to '$OfficePath'." "green"
	}
	catch
	{
		& $writeLog "Failed to write 'DEPLOY_MULTI_FCT.sqm'. Error: $_" "red"
	}

	& $writeLog "`r`n==================== Install_FUNCTIONS_Into_SMS Completed ====================`r`n" "blue"
}

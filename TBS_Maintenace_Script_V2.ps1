#######################################################################################################
#  TBS Maintenance Script - LAUNCHER
#
#  irm https://bit.ly/TBS_Maintenance_Script | iex  downloads THIS file (the short link cannot be
#  changed), so it stays here as a small launcher. It downloads the real script - the asset of the
#  latest GitHub Release - and runs it:
#    https://github.com/AlexSkellington/Alex_C.T/releases/latest/download/TBS_Maintenance_Script.ps1
#
#  Do not put code here. Sources: TBS_Maintenance\src (see TBS_Maintenance\README.md).
#  Stores get a new version when a GitHub Release is published, not when main changes.
#######################################################################################################
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
$__tbsUrl = 'https://github.com/AlexSkellington/Alex_C.T/releases/latest/download/TBS_Maintenance_Script.ps1'
$__tbsCode = $null
try
{
	$__tbsWeb = New-Object System.Net.WebClient
	$__tbsWeb.Headers['User-Agent'] = 'TBS-Maintenance-Launcher'
	$__tbsCode = $__tbsWeb.DownloadString($__tbsUrl)
	$__tbsWeb.Dispose()
}
catch { Write-Host "Could not download the TBS Maintenance Script from:`n  $__tbsUrl`n  $($_.Exception.Message)" -ForegroundColor Red }
if ($__tbsCode) { Invoke-Expression $__tbsCode }

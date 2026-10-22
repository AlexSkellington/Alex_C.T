#######################################################################################################
#  MiniGhost - LAUNCHER
#
#  irm https://bit.ly/MiniGhost | iex  downloads THIS file (the short link cannot be changed), so it
#  stays here as a small launcher. It downloads the real script from the MiniGhost folder on main
#  and runs it:
#    https://raw.githubusercontent.com/AlexSkellington/Alex_C.T/main/MiniGhost/MiniGhost.ps1
#
#  Do not put code here - edit MiniGhost\MiniGhost.ps1.
#######################################################################################################
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
$__mgUrl = 'https://raw.githubusercontent.com/AlexSkellington/Alex_C.T/main/MiniGhost/MiniGhost.ps1'
$__mgCode = $null
try
{
	$__mgWeb = New-Object System.Net.WebClient
	$__mgWeb.Headers['User-Agent'] = 'MiniGhost-Launcher'
	$__mgCode = $__mgWeb.DownloadString($__mgUrl)
	$__mgWeb.Dispose()
}
catch { Write-Host "Could not download MiniGhost from:`n  $__mgUrl`n  $($_.Exception.Message)" -ForegroundColor Red }
if ($__mgCode) { Invoke-Expression $__mgCode }
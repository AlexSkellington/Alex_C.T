#######################################################################################################
#  MiniGhost - LAUNCHER
#
#  irm https://bit.ly/MiniGhost | iex  downloads THIS file (the short link cannot be changed), so it
#  stays here as a small launcher. It saves the real script from the MiniGhost folder on main to
#  %TEMP%\MiniGhost.ps1 and starts that FILE as administrator in its own (minimized) PowerShell window:
#    https://raw.githubusercontent.com/AlexSkellington/Alex_C.T/main/MiniGhost/MiniGhost.ps1
#  MiniGhost is never run from memory (Invoke-Expression of downloaded code is what Microsoft
#  Defender blocks as Trojan:Win32/Commando), and it deletes the temp copy when it closes.
#
#  Do not put code here - edit MiniGhost\src.
#######################################################################################################
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
$__mgUrl = 'https://raw.githubusercontent.com/AlexSkellington/Alex_C.T/main/MiniGhost/MiniGhost.ps1'
$__mgFile = Join-Path ([System.IO.Path]::GetTempPath()) 'MiniGhost.ps1'
try
{
	$__mgWeb = New-Object System.Net.WebClient
	$__mgWeb.Headers['User-Agent'] = 'MiniGhost-Launcher'
	$__mgWeb.DownloadFile($__mgUrl, $__mgFile)
	$__mgWeb.Dispose()
	$__mgPs = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
	Start-Process -FilePath $__mgPs -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $__mgFile)) -Verb RunAs -WindowStyle Minimized
	Write-Host 'MiniGhost is starting in its own window (accept the administrator prompt).' -ForegroundColor Green
}
catch { Write-Host "Could not start MiniGhost from:`n  $__mgUrl`n  $($_.Exception.Message)" -ForegroundColor Red }

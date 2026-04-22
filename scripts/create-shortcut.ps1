param(
    [Parameter(Mandatory=$true)][string]$TrayLauncher,
    [Parameter(Mandatory=$true)][string]$WorkingDir,
    [Parameter(Mandatory=$true)][string]$IconPath
)

# Creates a per-user Startup-folder shortcut that launches the tray icon
# hidden at login. TrayLauncher points at scripts\start-tray.vbs.
#
# TargetPath must be an EXE for Windows Shell to resolve the shortcut
# reliably. We use wscript.exe to run the .vbs without a console window.

$ws = New-Object -ComObject WScript.Shell
$startupDir = [Environment]::GetFolderPath('Startup')
$lnkPath = Join-Path $startupDir 'Claude for Word RTL Tray.lnk'
$sc = $ws.CreateShortcut($lnkPath)
$sc.TargetPath = "$env:SystemRoot\System32\wscript.exe"
$sc.Arguments = "`"$TrayLauncher`""
$sc.WorkingDirectory = $WorkingDir
$sc.IconLocation = "$IconPath,0"
$sc.WindowStyle = 7
$sc.Description = 'Claude for Word RTL tray icon - right-click to connect / disconnect'
$sc.Save()
Write-Host "Startup shortcut created: $lnkPath"

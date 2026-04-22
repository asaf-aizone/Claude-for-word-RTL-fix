' Claude for Word RTL - run the Node injector silently in the background.
' Invoked by word-wrapper.bat. No visible window.
' Also launches the PowerShell tray-icon indicator (no window, no console).

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
injectJs  = scriptDir & "\scripts\inject.js"
trayPs1   = scriptDir & "\scripts\tray-icon.ps1"

Set shell = CreateObject("WScript.Shell")
' Run hidden (0), don't wait for completion (False)
shell.Run "cmd /c node """ & injectJs & """", 0, False

' Tray icon: only launch if not already running. We detect a prior tray by
' the presence of the status file being actively refreshed - but that's
' fragile, so we simply launch; the PS script is cheap and a second copy
' will just sit next to the first. Cleanup.bat can kill stragglers.
If fso.FileExists(trayPs1) Then
    shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & trayPs1 & """", 0, False
End If

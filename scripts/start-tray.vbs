' Claude for Word RTL - launch the system-tray status/controller icon.
'
' This is the user's only "entry point" after install: the Startup folder
' entry points here, so the tray icon comes up automatically at login.
' The tray provides Connect/Disconnect/Restart injector/Exit and is the
' primary UI for enabling or disabling the RTL fix on demand.
'
' Runs hidden (window style 0). The tray script itself enforces a
' singleton via a global mutex, so duplicate invocations are harmless.

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
trayPs1 = scriptDir & "\tray-icon.ps1"

If Not fso.FileExists(trayPs1) Then
    ' Install folder is broken or moved; nothing to do.
    WScript.Quit 0
End If

Set shell = CreateObject("WScript.Shell")
shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & trayPs1 & """", 0, False

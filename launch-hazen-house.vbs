' ============================================================
' Hazen House — Local Launcher
' Starts the local web server (if not already running) and opens
' the app in your default browser. Runs silently — no console window.
' ============================================================

Option Explicit

Dim WshShell, FSO, appPath, url, port
Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' This script lives inside the DBK-CM folder; serve from that folder
appPath = FSO.GetParentFolderName(WScript.ScriptFullName)
port = 8765
url = "http://localhost:" & port & "/"

' Check whether the server is already running on this port
Dim http, alreadyRunning
alreadyRunning = False
On Error Resume Next
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", url, False
http.SetRequestHeader "Cache-Control", "no-cache"
http.Send
If Err.Number = 0 And http.Status = 200 Then alreadyRunning = True
On Error Goto 0

If Not alreadyRunning Then
  ' Start Python's built-in HTTP server in the DBK-CM directory.
  ' Window state 0 = hidden. The server keeps running until the user
  ' shuts down or kills it from Task Manager (process: python.exe).
  WshShell.Run "cmd /c cd /d """ & appPath & """ && python -m http.server " & port, 0, False

  ' Give the server a moment to bind the port before opening the browser.
  WScript.Sleep 1500
End If

' Open the app in the default browser
WshShell.Run url, 1, False

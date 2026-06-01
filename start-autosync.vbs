Option Explicit

Const PROJECT_DIR = "C:\Copilot\RND-Dashboard"

Dim oShell, oWMI, oProcesses, oProcess
Set oShell = CreateObject("WScript.Shell")
Set oWMI   = GetObject("winmgmts:\\.\root\cimv2")

Set oProcesses = oWMI.ExecQuery( _
    "SELECT * FROM Win32_Process WHERE Name='py.exe' OR Name='python.exe' OR Name='python3.exe'")

Dim bRunning : bRunning = False
For Each oProcess In oProcesses
    If InStr(LCase(oProcess.CommandLine), "auto_sync") > 0 Then
        bRunning = True
        Exit For
    End If
Next

If Not bRunning Then
    oShell.CurrentDirectory = PROJECT_DIR
    oShell.Run "py auto_sync.py", 0, False
End If

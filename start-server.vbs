Option Explicit

Dim oShell, oWMI, oProcesses, oProcess
Set oShell = CreateObject("WScript.Shell")
Set oWMI   = GetObject("winmgmts:\\.\root\cimv2")

' py.exe / python.exe 중 http.server 가 이미 포함된 프로세스가 있으면 종료
Set oProcesses = oWMI.ExecQuery( _
    "SELECT * FROM Win32_Process WHERE Name='py.exe' OR Name='python.exe' OR Name='python3.exe'")

Dim bRunning : bRunning = False
For Each oProcess In oProcesses
    If InStr(LCase(oProcess.CommandLine), "http.server") > 0 Then
        bRunning = True
        Exit For
    End If
Next

If Not bRunning Then
    Dim sDir : sDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
    oShell.CurrentDirectory = Left(sDir, Len(sDir) - 1)
    ' 창 숨김(0)으로 서버 실행
    oShell.Run "py -m http.server 8080", 0, False
End If

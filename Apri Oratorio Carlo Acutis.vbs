Option Explicit

Dim shell, fso, projectRoot, localPythonw, bundledPythonw, pythonCommand
Dim appPath, browserUrl, fallbackBrowserUrl, healthcheckUrl, logDir, logPath, noBrowser
Dim edgePath64, edgePath86, chromePath86, chromePath64

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

projectRoot = fso.GetParentFolderName(WScript.ScriptFullName)
localPythonw = projectRoot & "\runtime\python\pythonw.exe"
bundledPythonw = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"
appPath = projectRoot & "\app.py"
healthcheckUrl = "http://127.0.0.1:8000/login"
browserUrl = "http://oratoriocarloacutis.don:8000/login"
fallbackBrowserUrl = healthcheckUrl
logDir = projectRoot & "\outputs"
logPath = logDir & "\desktop-launcher.log"
edgePath64 = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Microsoft\Edge\Application\msedge.exe"
edgePath86 = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Microsoft\Edge\Application\msedge.exe"
chromePath64 = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Google\Chrome\Application\chrome.exe"
chromePath86 = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Google\Chrome\Application\chrome.exe"
noBrowser = False

If WScript.Arguments.Count > 0 Then
    noBrowser = (LCase(Trim(WScript.Arguments(0))) = "--nobrowser")
End If

WriteLog "Avvio launcher. noBrowser=" & CStr(noBrowser)

If Not TestServerReady(healthcheckUrl) Then
    pythonCommand = ResolvePythonCommand()
    WriteLog "Server non attivo. Avvio con: " & pythonCommand
    shell.Run pythonCommand, 0, False
    WaitForServer healthcheckUrl, 20
Else
    WriteLog "Server gia attivo."
End If

If Not noBrowser Then
    OpenBrowserWithFallback
Else
    WriteLog "Avvio in background senza browser."
End If

Function ResolvePythonCommand()
    If fso.FileExists(localPythonw) Then
        ResolvePythonCommand = """" & localPythonw & """ """ & appPath & """"
    ElseIf fso.FileExists(bundledPythonw) Then
        ResolvePythonCommand = """" & bundledPythonw & """ """ & appPath & """"
    Else
        ResolvePythonCommand = "pythonw """ & appPath & """"
    End If
End Function

Function TestServerReady(url)
    On Error Resume Next
    Dim request, statusCode
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 2000, 2000, 2000, 2000
    request.Open "GET", url, False
    request.Send
    statusCode = 0
    If Err.Number = 0 Then
        statusCode = request.Status
    End If
    TestServerReady = (Err.Number = 0 And statusCode >= 200 And statusCode < 500)
    Err.Clear
    On Error GoTo 0
End Function

Sub WaitForServer(url, timeoutSeconds)
    Dim startedAt
    startedAt = Timer

    Do While TimerSafeElapsed(startedAt) < timeoutSeconds
        If TestServerReady(url) Then
            WriteLog "Server pronto dopo attesa."
            Exit Do
        End If
        WScript.Sleep 500
    Loop

    If Not TestServerReady(url) Then
        WriteLog "Server non raggiungibile entro il timeout."
    End If
End Sub

Function TimerSafeElapsed(startedAt)
    Dim currentValue
    currentValue = Timer
    If currentValue < startedAt Then
        TimerSafeElapsed = (86400 - startedAt) + currentValue
    Else
        TimerSafeElapsed = currentValue - startedAt
    End If
End Function

Sub OpenBrowserWithFallback()
    Dim resolvedUrl
    resolvedUrl = browserUrl

    If Not TestServerReady(browserUrl) Then
        resolvedUrl = fallbackBrowserUrl
    End If

    WriteLog "Apro il browser su: " & resolvedUrl

    If fso.FileExists(chromePath64) Then
        If TryRunBrowser("""" & chromePath64 & """ --new-window """ & resolvedUrl & """", 1) Then Exit Sub
    End If
    If fso.FileExists(chromePath86) Then
        If TryRunBrowser("""" & chromePath86 & """ --new-window """ & resolvedUrl & """", 1) Then Exit Sub
    End If
    If fso.FileExists(edgePath64) Then
        If TryRunBrowser("""" & edgePath64 & """ --new-window """ & resolvedUrl & """", 1) Then Exit Sub
    End If
    If fso.FileExists(edgePath86) Then
        If TryRunBrowser("""" & edgePath86 & """ --new-window """ & resolvedUrl & """", 1) Then Exit Sub
    End If
    If TryRunBrowser("explorer.exe """ & resolvedUrl & """", 1) Then Exit Sub

    WriteLog "Nessun metodo di apertura browser ha avuto esito positivo."
End Sub

Function TryRunBrowser(command, windowStyle)
    On Error Resume Next
    shell.Run command, windowStyle, False
    If Err.Number = 0 Then
        WriteLog "Comando browser eseguito: " & command
        TryRunBrowser = True
    Else
        WriteLog "Errore apertura browser: " & command & " | " & Err.Description
        Err.Clear
        TryRunBrowser = False
    End If
    On Error GoTo 0
End Function

Sub WriteLog(message)
    On Error Resume Next
    If Not fso.FolderExists(logDir) Then
        fso.CreateFolder(logDir)
    End If

    Dim stream, timestamp
    timestamp = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
        Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)
    Set stream = fso.OpenTextFile(logPath, 8, True)
    stream.WriteLine "[" & timestamp & "] " & message
    stream.Close
    On Error GoTo 0
End Sub

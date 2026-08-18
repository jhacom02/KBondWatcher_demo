Attribute VB_Name = "KBondWatcher"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const PYTHONW_PATH As String = "pythonw.exe"
Private Const PROJECT_DIR As String = "C:\mycode\KBondWatcher"
Private Const MAIN_PATH As String = "C:\mycode\KBondWatcher\main.py"
Private Const CONFIG_PATH As String = "C:\mycode\KBondWatcher\.env"
Private Const STOP_FLAG_PATH As String = "C:\temp\kbond_watcher.stop"
Private Const PID_PATH As String = "C:\temp\kbond_watcher.pid"
Private Const STOP_WAIT_MS As Long = 8000
Private Const STOP_POLL_MS As Long = 200

Private Const STATUS_CELL As String = "F2"
Private Const LOOKING_FOR_CELL As String = "G2"
Private Const LAST_QUOTE_CELL As String = "H2"
Private Const LAST_PNL_CELL As String = "I2"
Private Const LAST_ACTION_CELL As String = "J2"

Public Sub StartKBondWatcher()
    On Error GoTo Fail

    WriteStopFlag
    WaitForWatcherExit
    KillWatcherProcesses
    On Error GoTo Fail
    DeleteIfExists STOP_FLAG_PATH

    Range(STATUS_CELL).Value = ""
    Range(LOOKING_FOR_CELL).Value = ""
    Range(LAST_ACTION_CELL).Value = ""
    Range(LAST_QUOTE_CELL).Value = ""
    Range(LAST_PNL_CELL).Value = ""

    Dim cmd As String
    Dim sh As Object
    cmd = """" & PYTHONW_PATH & """ """ & MAIN_PATH & """ --config """ & CONFIG_PATH & """"
    Set sh = CreateObject("WScript.Shell")
    sh.CurrentDirectory = PROJECT_DIR
    sh.Run cmd, 0, False
    Exit Sub

Fail:
    Range(STATUS_CELL).Value = "ERROR"
    Range(LAST_ACTION_CELL).Value = "(" & Format(Now, "HH:nn:ss") & ") Error: " & Err.Description
End Sub

Public Sub StopKBondWatcher()
    On Error GoTo Fail

    WriteStopFlag
    WaitForWatcherExit
    KillWatcherProcesses
    On Error GoTo Fail

    Range(STATUS_CELL).Value = "STOPPED"
    Range(LAST_ACTION_CELL).Value = "(" & Format(Now, "HH:nn:ss") & ") Stopped"
    Exit Sub

Fail:
    Range(STATUS_CELL).Value = "ERROR"
    Range(LAST_ACTION_CELL).Value = "(" & Format(Now, "HH:nn:ss") & ") Error: " & Err.Description
End Sub

Private Function EnsureParentFolder(filePath As String) As Object
    Dim fso As Object
    Dim folderPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = Left$(filePath, InStrRev(filePath, "\") - 1)
    If Len(folderPath) > 0 Then
        If Not fso.FolderExists(folderPath) Then
            fso.CreateFolder folderPath
        End If
    End If
    Set EnsureParentFolder = fso
End Function

Private Sub WriteStopFlag()
    Dim fso As Object
    Dim ts As Object
    Set fso = EnsureParentFolder(STOP_FLAG_PATH)
    Set ts = fso.CreateTextFile(STOP_FLAG_PATH, True)
    ts.WriteLine "stop"
    ts.Close
End Sub

Private Sub DeleteIfExists(filePath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        fso.DeleteFile filePath, True
    End If
End Sub

Private Function ReadPidText(fso As Object) As String
    Dim ts As Object
    Dim pidText As String
    If Not fso.FileExists(PID_PATH) Then
        ReadPidText = ""
        Exit Function
    End If
    Set ts = fso.OpenTextFile(PID_PATH, 1)
    If Not ts.AtEndOfStream Then pidText = Trim$(ts.ReadAll)
    ts.Close
    ReadPidText = Replace(Replace(pidText, vbCr, ""), vbLf, "")
End Function

Private Function PidStillRunning(pidText As String, sh As Object) As Boolean
    Dim rc As Long
    If Not IsNumeric(pidText) Then
        PidStillRunning = False
        Exit Function
    End If
    rc = sh.Run( _
        "cmd /c tasklist /FI ""PID eq " & CLng(pidText) & """ | findstr /C:"" " & CLng(pidText) & " "">NUL", _
        0, True)
    PidStillRunning = (rc = 0)
End Function

Private Sub WaitForWatcherExit()
    Dim fso As Object
    Dim sh As Object
    Dim elapsed As Long
    Dim pidText As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sh = CreateObject("WScript.Shell")
    elapsed = 0
    Do While elapsed < STOP_WAIT_MS
        If Not fso.FileExists(PID_PATH) Then Exit Sub
        pidText = ReadPidText(fso)
        If Not PidStillRunning(pidText, sh) Then Exit Sub
        Sleep STOP_POLL_MS
        elapsed = elapsed + STOP_POLL_MS
        DoEvents
    Loop
End Sub

Private Sub KillWatcherProcesses()
    On Error GoTo KillFail
    Dim sh As Object
    Dim fso As Object
    Dim pidText As String
    Dim cmd As String
    Dim killRc As Long
    Dim psRc As Long

    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")

    pidText = ReadPidText(fso)
    If IsNumeric(pidText) Then
        killRc = sh.Run("taskkill /F /PID " & CLng(pidText), 0, True)
        If killRc <> 0 And killRc <> 128 Then
            Err.Raise vbObjectError + 1000, "KillWatcher", "taskkill failed: " & killRc
        End If
    End If
    If fso.FileExists(PID_PATH) Then fso.DeleteFile PID_PATH, True

    cmd = "powershell.exe -NoProfile -WindowStyle Hidden -Command " & _
          """Get-CimInstance Win32_Process | Where-Object { " & _
          "($_.Name -eq 'pythonw.exe' -or $_.Name -eq 'python.exe') -and " & _
          "$_.CommandLine -like '*" & MAIN_PATH & "*' } | " & _
          "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"""
    psRc = sh.Run(cmd, 0, True)
    If psRc <> 0 Then
        Err.Raise vbObjectError + 1001, "KillWatcher", "Stop-Process failed: " & psRc
    End If
    Exit Sub

KillFail:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Attribute VB_Name = "KBondWatcher"
Option Explicit

Private Const PYTHONW_PATH As String = "pythonw.exe"
Private Const PROJECT_DIR As String = "C:\mycode\KBondWatcher"
Private Const MAIN_PATH As String = "C:\mycode\KBondWatcher\main.py"
Private Const CONFIG_PATH As String = "C:\mycode\KBondWatcher\.env"
Private Const STOP_FLAG_PATH As String = "C:\temp\kbond_watcher.stop"
Private Const PID_PATH As String = "C:\temp\kbond_watcher.pid"

Private Const STATUS_CELL As String = "F2"
Private Const LOOKING_FOR_CELL As String = "G2"
Private Const LAST_QUOTE_CELL As String = "H2"
Private Const LAST_PNL_CELL As String = "I2"
Private Const LAST_ACTION_CELL As String = "J2"

Public Sub StartKBondWatcher()
    On Error GoTo Fail

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

Private Sub KillWatcherProcesses()
    On Error GoTo KillFail
    Dim sh As Object
    Dim fso As Object
    Dim ts As Object
    Dim pidText As String
    Dim cmd As String
    Dim killRc As Long
    Dim psRc As Long

    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(PID_PATH) Then
        Set ts = fso.OpenTextFile(PID_PATH, 1)
        If Not ts.AtEndOfStream Then pidText = Trim$(ts.ReadAll)
        ts.Close
        pidText = Replace(Replace(pidText, vbCr, ""), vbLf, "")
        If IsNumeric(pidText) Then
            killRc = sh.Run("taskkill /F /PID " & CLng(pidText), 0, True)
            ' 0 = killed, 128 = process already gone
            If killRc <> 0 And killRc <> 128 Then
                Err.Raise vbObjectError + 1000, "KillWatcher", "taskkill failed: " & killRc
            End If
        End If
        If fso.FileExists(PID_PATH) Then fso.DeleteFile PID_PATH, True
    End If

    ' Only python/pythonw: this PowerShell command line also contains MAIN_PATH,
    ' so matching every process by CommandLine would Stop-Process itself (Run → -1).
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

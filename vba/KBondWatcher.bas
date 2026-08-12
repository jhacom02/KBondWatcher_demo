Attribute VB_Name = "KBondWatcher"
Option Explicit

Private Const PYTHONW_PATH As String = "pythonw.exe"
Private Const PROJECT_DIR As String = "C:\mycode\KBondWatcher_kbond"
Private Const MAIN_PATH As String = "C:\mycode\KBondWatcher_kbond\main.py"
Private Const CONFIG_PATH As String = "C:\mycode\KBondWatcher_kbond\.env"
Private Const STOP_FLAG_PATH As String = "C:\temp\kbond_watcher.stop"

Private Const STATUS_CELL As String = "F2"
Private Const LOOKING_FOR_CELL As String = "G2"
Private Const LAST_QUOTE_CELL As String = "H2"
Private Const LAST_PNL_CELL As String = "I2"
Private Const LAST_ACTION_CELL As String = "J2"

Public Sub StartKBondWatcher()
    On Error GoTo Fail

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
    Range(LAST_ACTION_CELL).Value = "VBA Error: " & Err.Description
End Sub

Public Sub StopKBondWatcher()
    On Error GoTo Fail

    Dim fso As Object
    Dim folderPath As String
    Dim ts As Object
    folderPath = Left$(STOP_FLAG_PATH, InStrRev(STOP_FLAG_PATH, "\") - 1)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(folderPath) > 0 Then
        If Not fso.FolderExists(folderPath) Then
            fso.CreateFolder folderPath
        End If
    End If
    Set ts = fso.CreateTextFile(STOP_FLAG_PATH, True)
    ts.WriteLine "stop"
    ts.Close
    Exit Sub

Fail:
    Range(STATUS_CELL).Value = "ERROR"
    Range(LAST_ACTION_CELL).Value = "VBA Error: " & Err.Description
End Sub

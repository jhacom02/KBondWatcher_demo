Attribute VB_Name = "KBondWatcher"
' KBond → Excel → K-Bond watcher launcher
' Import this module into sample/sample.xlsx and assign macros to START/STOP buttons.
'
' Before first use, adjust the Const paths below to your machine.

Option Explicit

' === Edit these paths for your environment ===
Private Const PYTHONW_PATH As String = "pythonw.exe"
Private Const PROJECT_DIR As String = "C:\mycode\KBondHandler"
Private Const MAIN_PATH As String = "C:\mycode\KBondHandler\main.py"
Private Const CONFIG_PATH As String = "C:\mycode\KBondHandler\config.env"
Private Const STOP_FLAG_PATH As String = "C:\temp\KBond_watcher.stop"

' Status cells (must match config.env)
Private Const STATUS_CELL As String = "G2"
Private Const LAST_QUOTE_CELL As String = "H2"
Private Const LAST_PNL_CELL As String = "I2"
Private Const LAST_ACTION_CELL As String = "J2"

Public Sub StartKBondWatcher()
    On Error GoTo Fail

    Range(STATUS_CELL).Value = "STARTING"
    Range(LAST_ACTION_CELL).Value = ""
    Range(LAST_QUOTE_CELL).Value = ""
    Range(LAST_PNL_CELL).Value = ""

    Dim cmd As String
    Dim sh As Object

    ' Quote paths for spaces; run hidden & async (False = do not wait).
    cmd = """" & PYTHONW_PATH & """ """ & MAIN_PATH & """ --config """ & CONFIG_PATH & """"

    Set sh = CreateObject("WScript.Shell")
    ' WindowStyle 0 = hidden; WaitOnReturn False = async
    sh.CurrentDirectory = PROJECT_DIR
    sh.Run cmd, 0, False

    Range(LAST_ACTION_CELL).Value = "python main.py launched"
    Exit Sub

Fail:
    Range(STATUS_CELL).Value = "ERROR"
    Range(LAST_ACTION_CELL).Value = "VBA Start failed: " & Err.Description
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

    Range(LAST_ACTION_CELL).Value = "stop flag written"
    Exit Sub

Fail:
    Range(STATUS_CELL).Value = "ERROR"
    Range(LAST_ACTION_CELL).Value = "VBA Stop failed: " & Err.Description
End Sub

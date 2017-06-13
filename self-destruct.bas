' Excel-SelfDestruct
' @url https://github.com/hash-bang/Excel-SelfDestruct
' @author Matt Carter (https://github.com/hash-bang) <m@ttcarter.com>
' @date 2017-06-13

Dim RunWhen As Double
Const TimerWait As String = "00:10:00"
Const NUM_MINUTES = 10

Public Sub SaveAndClose()
    Dim Message As String
    Message = "The document " & ThisWorkbook.Name & " has been closed due to " & NUM_MINUTES & " minutes of inactivity"
    
    Dim WScriptShell As Object
    Set WScriptShell = CreateObject("WScript.Shell")
    WScriptShell.Run "mshta.exe vbscript:close(CreateObject(""WScript.Shell"").Popup(""" & Message & """," & 0 & ",""" & ThisWorkbook.Name & """," & 48 & "))"
    
    ThisWorkbook.Close savechanges:=True
End Sub

Public Sub RestartTimer()
    On Error Resume Next
    Application.OnTime RunWhen, "!ThisWorkbook.SaveAndClose", , False
    On Error GoTo 0
    RunWhen = Now + TimeValue(TimerWait)
    Application.OnTime RunWhen, "!ThisWorkbook.SaveAndClose", , True
End Sub

Public Sub StopTimer()
    On Error Resume Next
    Application.OnTime RunWhen, "!ThisWorkbook.SaveAndClose", , False
    On Error GoTo 0
End Sub

Private Sub Workbook_Open()
    Call RestartTimer
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Call RestartTimer
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Call RestartTimer
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call StopTimer
End Sub

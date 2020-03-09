Option Explicit

Sub Sample()
    Dim Ret

    Ret = IsWorkBookOpen("C:\myWork.xlsx")

    If Ret = True Then
        MsgBox "File is open"
    Else
        MsgBox "File is Closed"
    End If
End Sub

Function IsWorkBookOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function



Private Function WorkbookIsOpen(wbname) As Boolean
'   Returns TRUE if the workbook is open
    Dim x As Workbook
    On Error Resume Next
    Set x = Workbooks(wbname)
    If Err = 0 Then WorkbookIsOpen = True _
        Else WorkbookIsOpen = False
End Function
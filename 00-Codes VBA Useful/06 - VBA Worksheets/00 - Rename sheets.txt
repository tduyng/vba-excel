Sub RENAME_SHEET()

Dim ws As Worksheet
Dim ws1 As Worksheet
Dim strErr As String

On Error Resume Next
For Each ws In ActiveWorkbook.Sheets
Set ws1 = Sheets(ws.Name)
    If Right(ws1.Name, 3) <> "_v1" Then
        'Debug.Print "ok"
    'End If
    'If ws1 Is Nothing Then
        ws.Name = ws.Name & "_v1"
    Else
         strErr = strErr & ws.Name & " with the suffix " & "_v1" & vbNewLine
    End If
Set ws1 = Nothing
Next
On Error GoTo 0

If Len(strErr) > 0 Then MsgBox strErr, vbOKOnly, "these sheets already existed"

End Sub

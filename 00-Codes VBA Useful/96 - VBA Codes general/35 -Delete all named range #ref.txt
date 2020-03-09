Sub DELETE_NAMERANGE()
On Error Resume Next
    Dim nameRng As Name
    For Each nameRng In ThisWorkbook.Sheets("FAC.11x").Names
        nameRng.Delete
    Next nameRng
    
'    For Each Name In ActiveWorkbook.Names
'        Name.Delete
'    Next Name

'deletes all the names in the active workbook
 
'with a #REF error- confirms before running
'
'Dim N As Name
'
'If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Confirm macro") = vbNo Then Exit Sub
'
'For Each N In ActiveWorkbook.Names
'
'If InStr(N.Value, "#REF") Then N.Delete
'
'Next N
End Sub


Sub DeleteBadNames()
Dim nm As Excel.Name
Dim vTest As Variant
Dim i As Long
Dim ListSh As Worksheet
Set ListSh = Sheets("DELETED")
For Each nm In ActiveWorkbook.Names
    vTest = Empty
    On Error Resume Next
    vTest = Application.Evaluate(nm.RefersTo)
    On Error GoTo 0
    If TypeName(vTest) = "Error" Then
        i = i + 1
        ListSh.Cells(i, 1).Value = nm.Name
        If IsError(Application.Match(mn.Name, Sheets("DONT DELETE").Columns("A"), 0)) Then nm.Delete
    End If
Next nm
End Sub
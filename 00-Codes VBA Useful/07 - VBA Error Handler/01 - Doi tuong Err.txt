Option Explicit
 
Public Const COL_A = "A"
Public Const COL_B = "B"
 
Sub ErrorHanlding_Demo()
    Dim i As Integer
     
    On Error GoTo InvalidValue:
     
    For i = 1 To 5
        Cells(i, COL_B).Value = Sqr(Cells(i, COL_A).Value)
    Next i
     
    Exit Sub
     
InvalidValue:
    MsgBox "Error: " & Err.number & " : " & Err.Description
    Resume Next
End Sub

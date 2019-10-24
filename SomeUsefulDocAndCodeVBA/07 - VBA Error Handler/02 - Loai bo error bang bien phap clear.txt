Private Function FirstEmptyRow() As Long
    ' 9 Apr 2017

    Dim Clm As Long

    With ActiveSheet
        On Error GoTo ErrHandler:
        Clm = WorksheetFunction.Match("Style", .Rows(3), 0)
        FirstEmptyRow = .Cells(.Rows.Count, Clm).End(xlUp).Row + 1
    End With
ErrHandler:
    Err.Clear
End Function
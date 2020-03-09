Private Function SelectFirstEmptyRowInColumnWithGivenHeader(ByVal sheet As Worksheet, Optional ByVal header As String = "Style") As Long
    Dim col As Variant
    With sheet
        col = Application.Match(header, .Rows(1), 0)
        If Not IsError(col) Then
            .Activate '<--| you must select a sheet to activate a cell of it
            .Cells(.Rows.Count, col).End(xlUp).Offset(1).Select
        End If
    End With
End Function


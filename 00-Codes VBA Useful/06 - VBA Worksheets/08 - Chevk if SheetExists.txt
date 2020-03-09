Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each sheet In Worksheets
        If sheetToFind = sheet.name Then
            sheetExists = True
            Exit Function
        End If
    Next sheet
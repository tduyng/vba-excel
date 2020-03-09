Sub Filter_NewSheet()
    Dim wbBook As Workbook
    Dim wsSheet As Worksheet
    Dim rnStart As Range, rnData As Range
    Dim i As Long
    Set wbBook = ThisWorkbook
    Set wsSheet = wbBook.Worksheets("Sheet1")
    With wsSheet
        'Make sure that the first row contains headings.
        Set rnStart = .Range("A2")
        Set rnData = .Range(.Range("A2"), .Cells(.Rows.Count, 3).End(xlUp))
    End With
    Application.ScreenUpdating = True
    For i = 1 To 5
        'Here we filter the data with the first criterion.
        rnStart.AutoFilter Field:=1, Criteria1:="AA" & i
        'Copy the filtered list
        rnData.SpecialCells(xlCellTypeVisible).Copy
        'Add a new worksheet to the active workbook.
        Worksheets.Add Before:=wsSheet
        'Name the added new worksheets.
        ActiveSheet.Name = "AA" & i
        'Paste the filtered list.
        Range("A2").PasteSpecial xlPasteValues
    Next i
    'Reset the list to its original status.
    rnStart.AutoFilter Field:=1
    With Application
        .CutCopyMode = False
        .ScreenUpdating = False
    End With
End Sub
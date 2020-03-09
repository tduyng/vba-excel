Sub tach_du_lieu1()
    Dim rng As Range
    Dim wb As Workbook
    Dim lstRow&, lstCol&
    Dim dCriteria As Date
    Dim arr As Variant
    

    
    
    With Sheet1
        lstRow = .Cells(.Rows.Count, 2).End(xlUp).Row
        lstCol = .Cells(11, .Columns.Count).End(xlToLeft).Column
        Set rng = .Range("B11", .Cells(lstRow, lstCol))
        dCriteria = .Range("H9").Value

    'Test if using Advanced filter
'    If .FilterMode Then .ShowAllData

	'test if using autofilter
    If .AutoFilterMode Then rng.AutoFilter
    
    End With
    
    Debug.Print dCriteria
    rng.Select
'    arr = Array(0, dCriteria)
    rng.AutoFilter Field:=1, Criteria1:=Array(0, dCriteria)
    rng.Copy
    
    With Sheets.Add(After:=Sheets(Sheets.Count))
        .Range("A").PasteSpecial Paste:=xlPasteValues
        .Name = CStr(Sheet1.Range("H9"))
    End With
    
'    Application.CutCopyMode = False
    
    rng.AutoFilter
    
    ActiveWorkbook.Save
    
End Sub

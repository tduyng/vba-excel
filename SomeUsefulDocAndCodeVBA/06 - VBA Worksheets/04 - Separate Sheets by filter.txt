Function SHEET_EXISTE(shName As String, Optional wb As Workbook) As Boolean
    'Set wb = Workbooks"Transpose Tablo.xlsm")
    Set wb = Nothing
    If wb Is Nothing Then Set wb = ThisWorkbook
    Dim sht As Worksheets
    On Error Resume Next
    Set sht = Sheets(shName)
    On Error GoTo 0
    
    SHEET_EXIST Not sht Is Nothing
End Function


Sub SEPARATE_SHEETS_BY_FILTER()
    Dim wb As Workbook, shNew As Worksheet, wbNew As Workbook
    Dim coll As New Collection, item As Variant, itemName As String, filePath As String
    Dim lastRow As Long, rngKey As Range
     'Set wb = Nothing
    Set wb = ThisWorkbook
    Set Sh = wb.Sheets(Sheet3.Name)
    Set rng = Sh.Range("Data_filter")
    lastRow = Sh.Range("A" & Rows.count).End(xlUp).Row
    
    'get the path of the file excel
    filePath = Application.ActiveWorkbook.Path
    
    On Error Resume Next
    For Each rngKey In Sh.Range("B13:B" & lastRow)
        coll.Add rngKey.Value, rngKey.Value
    Next rngKey
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    'create new sheets and new workbook with the datas and the names which corespond the data filter
    For Each item In coll
        itemName = item
        If SHEET_EXIST(itemName, wb) Then
            Set shNew = wb.Sheets(item)
        Else
            Set shNew = Sheets.Add(after:=Worksheets(Worksheets.count))
            shNew.Name = item
        End If
        Set wbNew = Workbooks.Add
        'wbNew.Name = itemName
        rng.AutoFilter field:=2, Criteria1:=item
        
        'Copy datas to the new sheet in the same worbook (workbook active)
        rng.SpecialCells(xlCellTypeVisible).Copy shNew.Range("A1")
        
        'Copy Data to the new sheet of new workbook
        rng.SpecialCells(xlCellTypeVisible).Copy wbNew.Sheets("Feuil1").Range("A1")
        wbNew.Sheets("Feuil1").Name = itemName
        wbNew.SaveAs Filename:=filePath & "\" & itemName & ".xlsx"
    Next item
    
    'reset autofilter
    rng.AutoFilter
    
    'Re activate sheet data
    Sh.Activate
    
    
    Application.ScreenUpdating = True
End Sub

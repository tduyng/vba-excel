Sub TAO_FILE_OPTI()
    Dim Dic As Object
    Dim i&, j&, k&, n As Long, t As Single
    Dim rngKey As Range, wb As Workbook, rngi As Range, rngTotal As Range
    Dim arr(), keyDic()
    Dim strFilter As String
    
    FastWB True: t = Timer
    Set Dic = CreateObject("Scripting.Dictionary")
    
    Set rngTotal = Sheet4.Range("A1:F" & Sheet4.[A65536].End(xlUp).Row)
    Set rngKey = Sheet4.Range("F2:F" & Sheet4.[A65536].End(xlUp).Row)
    
    For Each rngi In rngKey
        If Not Dic.Exists(rngi.Value) Then Dic.Add rngi.Value, rngi.Value
    Next rngi
    
    keyDic = Dic.keys
    For i = 0 To UBound(keyDic)
        With rngTotal
            .AutoFilter field:=6, Criteria1:=keyDic(i)
            .Copy
        End With
        Set wb = Workbooks.Add
        With wb
            .Worksheets(1).Range("A2").PasteSpecial xlPasteColumnWidths
            .Worksheets(1).Range("A2").PasteSpecial xlPasteAll
            .Worksheets(1).Range("A2").Select
            .Worksheets(1).Range("A2").Copy
            'Sheet4.Range("A1:F1").Copy .Worksheets(1).Range("A1")
            .Worksheets(1).Columns("A:F").AutoFit
            .SaveAs ThisWorkbook.Path & "\" & keyDic(i) & ".xlsx"
            .Close
        End With

        rngTotal.AutoFilter
    Next i

    FastWB False: MsgBox "Done" & Round((Timer - t), 2) & "secondes"
End Sub

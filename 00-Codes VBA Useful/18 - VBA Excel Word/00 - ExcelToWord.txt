Sub EXCEL_TO_WORD_ADVANCED()
Dim oSourceDoc As Object, dataObj As Object, selectRng As Object

Dim srchTxt As String, sValue As String
Dim replaceRng As Range
Dim i&, j&, arr(), num_of_cust&, num_of_column&, t As Object, lTextAfter&

Application.ScreenUpdating = False
'Set dataObj = CreateObject("MSForms.DataObject")
Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    'So code chua ma code can dung
    num_of_column = 6
    
    'So luong cot chua du lieu, bo di hang du tien chua ma code
    num_of_cust = Sheet1.Cells(Rows.Count, "A").End(xlUp).Row - 1

        arr = Sheet1.Range("A2:F5").Value
        ReDim Preserve arr(1 To num_of_cust, 1 To num_of_column)
With CreateObject("word.application") ' late binding
        .Visible = True
        
        For i = 1 To num_of_cust
            Set oSourceDoc = .Documents.Open("C:\Users\tien-duy.nguyen\Desktop\VBA Word\template.docx")
            Set t = oSourceDoc.Content
            For j = 1 To num_of_column
'                arr(i, j) = Sheet1.Cells(i + 1, j).Value
                sValue = Sheet1.Cells(i + 1, j).Value
                If Len(sValue) <= 250 Then
                
                    t.Find.Execute _
                    FindText:=Sheet1.Cells(1, j).Value, _
                    ReplaceWith:=sValue, _
                    Replace:=wdReplaceAll
                    
                 Else
                    dataObj.SetText sValue
                    dataObj.PutinClipboard
                    Set selectRng = oSourceDoc.Range
                    selectRng.Select
                    lTextAfter = Len(sValue) - 250
                        
                    With oSourceDoc.Range.Find
'                        .ClearFormatting
                        .Text = Sheet1.Cells(1, j).Value
'                        .Replacement.ClearFormatting
                        .Forward = True
                        .Replacement.Text = "^c"
                        .Wrap = wdFindContinue
                        .Execute Replace:=wdReplaceAll
                    End With
                End If
            Next j
            oSourceDoc.SaveAs Filename:=ThisWorkbook.Path & Application.PathSeparator & i & "-giay_moi.docx"
        Next i
        .Quit
End With
   
Set t = Nothing
Set oSourceDoc = Nothing
Application.ScreenUpdating = True
End Sub
Public Sub Save_Workbook_As_PDF2()

    Dim PDFfileName As String
    
    With ActiveWorkbook
        PDFfileName = .Worksheets(1).Range("B4").Value & .Worksheets(1).Range("B5").Value & ".pdf"
    End With

    With Application.FileDialog(msoFileDialogSaveAs)

        .Title = "Save workbook as PDF"
        .InitialFileName = ThisWorkbook.Path & "\" & PDFfileName
        
        If .Show Then
            ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, fileName:=.SelectedItems(1), _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        End If
    
    End With
    
End Sub


Public Sub Save_Workbook_As_PDF()

    Dim i As Integer, PDFindex As Integer
    Dim PDFfileName As String
    
    With ActiveWorkbook
        PDFfileName = .Worksheets(1).Range("B4").Value & .Worksheets(1).Range("B5").Value & ".pdf"
    End With
    
    With Application.FileDialog(msoFileDialogSaveAs)
            
        PDFindex = 0
        For i = 1 To .Filters.Count
            If InStr(VBA.UCase(.Filters(i).Description), "PDF") > 0 Then PDFindex = i
        Next

        .Title = "Save workbook as PDF"
        .InitialFileName = PDFfileName
        .FilterIndex = PDFindex
        
        If .Show Then
            ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, fileName:=.SelectedItems(1), _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        End If
    
    End With
    
End Sub

Sub Test()
    Application.Dialogs(xlDialogSaveAs).Show , 46
End Su
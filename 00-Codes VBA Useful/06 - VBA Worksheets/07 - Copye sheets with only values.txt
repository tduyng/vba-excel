Sub COPY_SHEET_ONLY_VALUE()
    Dim wbCopy As Workbook, wbPaste As Workbook, sh As Worksheet, rngPrintArea As Range, nameSheet As String, firstCell As String
    Set wbCopy = ActiveWorkbook
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    ActiveSheet.Copy
    Set wbPaste = ActiveWorkbook
    Set rngPrintArea = wbPaste.Sheets(1).Range(wbPaste.Sheets(1).PageSetup.PrintArea)
    firstCell = rngPrintArea.Cells(1, 1).Address
    Debug.Print firstCell
    With wbPaste.Sheets(1).UsedRange
        .Cells.Copy
        .Cells.PasteSpecial xlPasteValues
    End With

    wbPaste.Sheets(1).Copy After:=wbPaste.Sheets(1)
    Set sh = ActiveSheet
    
    With sh.UsedRange.Cells
        .ClearContents
        .ClearFormats
    End With
    With wbPaste
'        On Error Resume Next
        rngPrintArea.Copy sh.Range(firstCell)
        nameSheet = .Sheets(1).Name
        .Sheets(1).Delete
        sh.Name = nameSheet
        
    End With
'    On Error GoTo 0
'    With sh
'        sh.Columns.AutoFit
'        sh.Rows.AutoFit
'    End With
        With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
End Sub


Sub COPY_SHEET_ONLY_VALUE_METHOD_2()
    Dim wbCopy As Workbook, wbPaste As Workbook, sh As Worksheet, rngPrintArea As Range, nameSheet As String, firstCell As String, rng As Range
    Set wbCopy = ActiveWorkbook
    With wbCopy.ActiveSheet
    Set rngPrintArea = wbCopy.ActiveSheet.Range(wbCopy.ActiveSheet.PageSetup.PrintArea)
    firstCell = rngPrintArea.Cells(1, 1).Address
    ActiveSheet.Copy
    End With
    Set wbPaste = ActiveWorkbook

    Debug.Print rngPrintArea.Column
    Debug.Print rngPrintArea.Columns.Count
    
    Debug.Print rngPrintArea.Row
    Debug.Print rngPrintArea.Rows.Count
    With wbPaste.Sheets(1).UsedRange.Cells
        .ClearContents
        .ClearFormats
        rngPrintArea.Copy wbPaste.Sheets(1).Range(firstCell)
        .Copy
        .PasteSpecial xlPasteValues
    End With
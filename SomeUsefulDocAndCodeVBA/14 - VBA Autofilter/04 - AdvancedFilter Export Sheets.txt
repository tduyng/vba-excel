Sub ADVANCED_EXPORT_SHEET()
    Dim sh As Worksheet, newWb As Workbook
    Dim rngFilter As Range, rngCriterial As Range, rngi As Range
    Dim i&
    Dim filePath As String, fileName As String
    Dim Dic As Object
    Dim lastRow&
    Dim sAllCriterial As Variant, sCrite As Variant
    Dim keyDic As Variant
    Dim fso As Object
    
    Application.ScreenUpdating = False
    Application.AlertBeforeOverwriting = False
    Application.DisplayAlerts = False
    
    Set Dic = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.filesystemObject")
    
    Set sh = ThisWorkbook.Sheets("DSHH")
    Set rngFilter = sh.Range("dataFilter")
    Set rngCriterial = sh.Range("CriterialFilter")
    
    filePath = ThisWorkbook.path & "\"
    fileName = fso.GetBaseName(ThisWorkbook.Name)
    
    lastRow = sh.Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 2 To lastRow
        sAllCriterial = sh.Cells(i, 4).value
        If Not Dic.Exists(sAllCriterial) Then
            Dic.Add sAllCriterial, vbNullString
        End If
    Next i
    keyDic = Dic.Keys
    ReDim Preserve keyDic(0 To Dic.Count - 1)
    For i = 0 To Dic.Count - 1
        sh.Range("J2") = keyDic(i)
        On Error Resume Next
        sh.ShowAllData
        rngFilter.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=rngCriterial, Unique:=False
        On Error GoTo 0
        Set newWb = Workbooks.Add
        With newWb
            rngFilter.Copy .Sheets(1).Range("A1")
            .Sheets(1).Columns("A:D").AutoFit
            .SaveAs filePath & fileName & "_" & keyDic(i) & 1 & ".xlsx"
            .Close
        End With
        
    Next i
    On Error Resume Next
    sh.ShowAllData
    On Error GoTo 0
    Application.DisplayAlerts = True
     Application.AlertBeforeOverwriting = True
    Application.ScreenUpdating = True
End Sub

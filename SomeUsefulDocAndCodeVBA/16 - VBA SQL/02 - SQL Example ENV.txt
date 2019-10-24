Sub TRAITER_SUIVI_QUOITIDIEN()
    Dim lstRwTotal&, lstColTotal&, lstRwSQl&
    Dim lstRwMember1&, lstRwMember2&, lstRwMember&
    Dim i&, j&
    Dim countAllData&, lineDepMember&, lineDepSuiviTotal&, lineField&
    Dim rngData As Range, rngValue As Range
    Dim arrData As Variant
    Dim cn As Object, rs As Object
    Dim sql As String, Ext As String, filePath As String
    Dim nameField1 As String, nameField2 As String, nameField3 As String, nameField4 As String, nameField5 As String
    Dim shSQL As Worksheet
    Dim DataSQL As String
    Dim Réalisateur As String, lastPrevious&
    Dim xName As Name
    
    Call OpenFile_Excel
    countSheetSuivi = 0
    
    If wbSQ Is Nothing Then Exit Sub
    
    'Count number sheet suivi
    With wbSQ
        ReDim arrSheets(1 To .Sheets.Count)
        For i = 1 To .Sheets.Count
            If .Sheets(i).Name <> ST And Left(.Sheets(i).Name, 5) = "Suivi" Then 'Pour toute les feuilles qui commence par "Suivi-" outre Suivi-total
                countSheetSuivi = countSheetSuivi + 1                                       'Stock le nombre de personne suivi
                arrSheets(countSheetSuivi) = .Sheets(i).Name                               'Stock le nom des personnes suivi
            End If
        Next
        ReDim Preserve arrSheets(1 To countSheetSuivi)
    Set shSuiviTotal = wbSQ.Sheets(ST)
    End With
    
    'Delete all names range
    For Each xName In wbSQ.Names
        If InStr(xName.Name, "DataSQL") <> 0 Then
            xName.Delete
        End If
    Next xName
    
    
    
    'Clear all old datas of sheets Suivi_Total
    lstRwTotal = getLastRowUsed(shSuiviTotal)
    lstColTotal = getLastColUsed(shSuiviTotal)
    If lstRwTotal = 1 Then lstRwTotal = 2
    With shSuiviTotal
        .Range("A2", .Cells(lstRwTotal, lstColTotal)).ClearContents
    End With
    
    'Start line of data useful in 2 sheets: SuiviMember and Suivi_Total
    lineField = 6
    lineDepMember = 8
    lineDepSuiviTotal = 2
    countAllData = 0
    
    
    'Start part using SQL
    filePath = wbSQ.FullName
    'Get extension of file
    Ext = Right(filePath, Len(filePath) - InStrRev(filePath, "."))
    
    '---Connecting to the Data Source---
    'File excel nhu xlsx,xlsm,xlsb, xls
    Set cn = CreateObject("ADODB.Connection")
    With cn
        If Ext = "xlsx" Or Ext = "xlsb" Or Ext = "xlsm" Then
            .provider = "Microsoft.ACE.OLEDB.12.0"
            .connectionstring = "Data Source=" & filePath & ";" & _
                                "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
        ElseIf Ext = "xls" Then
            .provider = "Microsoft.Jet.OLEDB.4.0"
            .connectionstring = "Data Source=" & filePath & ";" & _
                                "Extended Properties=""Excel 8.0;HDR=YES"";"
        End If
        On Error GoTo ERR_Handler
        .Open
    End With
    
   Set shSQL = Nothing
   On Error Resume Next
    Set shSQL = wbSQ.Sheets("Total_SQL")
   On Error GoTo 0
   If shSQL Is Nothing Then
        With wbSQ.Sheets.Add(After:=shSuiviTotal)
            .Name = "Total_SQL"
        End With
    End If
    
    Set shSQL = wbSQ.Sheets("Total_SQL")
    With shSQL
            .Cells.ClearContents
            .Cells.ClearFormats
            .Visible = xlSheetHidden
    End With
   
   
   'Search with SQL
    lstRwSQl = 1
   For i = 1 To countSheetSuivi 'Pour toute les personnes suivi

        With wbSQ.Sheets(arrSheets(i))
        
            lstRwMember1 = .Range("C" & .Rows.Count).End(xlUp).Row
            lstRwMember2 = .Range("F" & .Rows.Count).End(xlUp).Row
            If lstRwMember1 >= lstRwMember2 Then
                lstRwMember = lstRwMember1
            Else
                lstRwMember = lstRwMember2
            End If

            Set rngData = .Range("C" & lineField, "G" & lstRwMember)
            'Test if using Advanced filter
            If .FilterMode Then .ShowAllData
            'Test if using Autofilter
            If .AutoFilterMode Then rngData.AutoFilter

            wbSQ.Names.Add Name:="DataSQL_" & i, RefersTo:=rngData
            Réalisateur = .Range("C3").Value
            'Name info of SQL: field and  range
            DataSQL = "DataSQL_" & i
            nameField1 = "[" & .Range("C" & lineField).Value & "]"
            nameField2 = "[" & .Range("D" & lineField).Value & "]"
            nameField3 = "[" & .Range("E" & lineField).Value & "]"
            nameField4 = "[" & .Range("F" & lineField).Value & "]"
            nameField5 = "([" & .Range("G" & lineField).Value & " ] * 0 + 2)"
'            nameField4 = "[" & .Range("G" & lineField).Value & "]"
            'Lister les activité
            sql = "SELECT " & nameField1 & "," & nameField2 & "," & nameField3 & "," & nameField4 & _
            ", SUM(2) AS [NbHeure] FROM " & DataSQL & "  WHERE " & nameField1 & " <> """" AND " & nameField4 & " <> """"" & _
            " GROUP BY " & nameField1 & "," & nameField2 & "," & nameField3 & "," & nameField4
            
            Set rs = cn.Execute(sql)
            shSQL.Range("A" & lstRwSQl + 1).Interior.Color = vbGreen
            lastPrevious = lstRwSQl
            'lstPrevious: is a lastRow in colonne A of a previous loop
            
            'From fin of SQL, copy data grom recordset to sheet Total_SQL
            If Not rs.EOF Then
                shSQL.Range("A" & lstRwSQl + 1).CopyFromRecordset rs
                
                For j = 0 To rs.Fields.Count - 1
                    shSQL.Cells(1, j + 1).Value = rs.Fields(j).Name
                Next
            End If
            lstRwSQl = shSQL.Range("A" & shSQL.Rows.Count).End(xlUp).Row
            If lastPrevious < lstRwSQl Then
                shSQL.Range("F" & lastPrevious + 1, "F" & lstRwSQl).Value = Réalisateur
            Else
                shSQL.Range("F" & lastPrevious + 1).Value = Réalisateur
            End If
        End With
    Next i
    shSQL.Range("F1").Value = NOM
    'Fin search with  SQL
    
    
    '------------------------------------------------------------------
    'Format first Row of sheet TOTAL_SQL
        
    With shSQL.Range("A1:E1")
        .Font.Size = 14
        .Font.Bold = True
        .Font.ThemeColor = xlThemeColorDark1

        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End With
    With shSQL
        .Columns("A:E").EntireColumn.AutoFit
        Dim arr1 As Variant, arr2 As Variant
        arr1 = .Range("A2", "B" & lstRwSQl).Value
        arr2 = .Range("C2", "F" & lstRwSQl).Value
    End With
    
    'Copy Data from Total_SQL to SheetSuivi
    With shSuiviTotal
        .Range("A" & lineDepSuiviTotal).Resize(UBound(arr1) - LBound(arr1) + 1, 2) = arr1
        .Range("E" & lineDepSuiviTotal).Resize(UBound(arr1) - LBound(arr1) + 1, 4) = arr2
        Dim irng As Range
        On Error Resume Next
        Set rngData = wbSQ.Sheets("Liste des projets").Range("Listeprojets")
        rngData.Select
        For Each irng In .Range("B" & lineDepSuiviTotal, "B" & UBound(arr1) - LBound(arr1) + 1 + lineDepSuiviTotal)
            If irng.Value <> vbNullString Then
                irng.Offset(0, 1).Value = Application.WorksheetFunction.VLookup(CStr(irng.Value), rngData, 2, False)
                irng.Offset(0, 2).Value = Application.WorksheetFunction.VLookup(CStr(irng.Value), rngData, 3, False)
            End If
        
        Next irng
        
        Set irng = Nothing
        Set rngData = Nothing
        Erase arr1
        Erase arr2
        On Error GoTo 0
    End With
    
    
    
    shSuiviTotal.Columns("A:CC").EntireColumn.AutoFit
'    sbSQ.Save
    
    
'   Fin sql
    cn.Close
    Set cn = Nothing
    Set rs = Nothing

    
'    With Application
'        .ScreenUpdating = True
'        .EnableEvents = True
'        .Calculation = xlCalculationAutomatic
'    End With
'
    
END_Handler:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic

    End With
    Exit Sub
    
ERR_Handler:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    If Err.Number = 3704 Then Exit Sub
    If Err.Number = 91 Then Exit Sub
    MsgBox "Error " & Err.Number & "," & Err.Description
    Set rs = Nothing
    Resume END_Handler:
    
End Sub

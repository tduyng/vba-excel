Function RunSQL(Target As Range, sql As String, Optional FilePath As String = Empty, _
                Optional IncludeHeader As Boolean = True)
    If FilePath = Empty Then
        FilePath = ThisWorkbook.Path & "\" & ThisWorkbook.Name
    End If
    'Get extension of file
    Dim Ext As String
    Ext = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "."))
    
    Dim cn As Object, rs As Object
    
    '---Connecting to the Data Source---
    'Setting nay chi ap dung cho cac file excel nhu xlsx,xlsm,xlsb
    Set cn = CreateObject("ADODB.Connection")
    With cn
        If Ext = "xlsx" Or Ext = "xlsb" Or Ext = "xlsm" Then
            .provider = "Microsoft.ACE.OLEDB.12.0"
            .connectionstring = "Data Source=" & FilePath & ";" & _
                                "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
        ElseIf Ext = "xls" Then
            .provider = "Microsoft.Jet.OLEDB.4.0"
            .connectionstring = "Data Source=" & FilePath & ";" & _
                                "Extended Properties=""Excel 8.0;HDR=YES"";"
        End If
        On Error GoTo ERR_Handler
        .Open
    End With
    
    Dim blankPos As Integer
    blankPos = InStr(sql, " ")
    Dim SQLtype As String
    SQLtype = LCase(Left(sql, blankPos - 1))
    '---Run the SQL SELECT Query---
    If SQLtype = "update" Or SQLtype = "insert" Then
        Set rs = cn.Execute(sql, , adExecuteNoRecords)
    Else
        Set rs = cn.Execute(sql)
    End If
    
     ' Check we have data.
    If Not rs.EOF Then
        'Transfer result.
        'Worksheets("Result").Range("A1").CopyFromRecordset rs
        If IncludeHeader = True Then
            'Copy ca tieu de cua recordset
            Dim iCol As Integer
            For iCols = 0 To rs.Fields.count - 1
                Target.Cells(1, iCols + 1).value = rs.Fields(iCols).Name
            Next
            Target.Cells(2, 1).CopyFromRecordset rs
        Else
            Target.CopyFromRecordset rs
        End If
    ' Close the recordset
        rs.Close
    Else
        MsgBox "Error: No records returned.", vbCritical
    End If
    
    '---Clean up---
    cn.Close
    Set cn = Nothing
    Set rs = Nothing
END_Handler:

    Exit Function
    
ERR_Handler:
    If Err.Number = 3704 Then Exit Function
    If Err.Number = 91 Then Exit Function
    MsgBox "Error " & Err.Number & "," & Err.Description
    Set RunSQL = Nothing
    Resume END_Handler:
End Function
Function TaoChuoiKetNoi() As String
    Dim sAppPath As String
    ' Cac ban co the chinh sua tuy theo yeu cau ket noi cua minh
    sAppPath = ThisWorkbook.Path
    TaoChuoiKetNoi = "Driver={Microsoft Access Driver (*.mdb)}; Dbq=" & sAppPath & "\" & DBName & "; UID=Admin; PWD=;"

End Function




Function KetNoi() As Long

    On Error GoTo ErrorHandle

    Set gcnObj = CreateObject("ADODB.Connection")
    With gcnObj
        .Mode = 3                                     'i.e adModeReadWrite
        'Neu sau thoi gian nay ma khong ket noi duoc se bao loi
        .ConnectionTimeout = 30
        'CursorTypeEnum
        '--------------
        'adOpenDynamic     = 2
        'adOpenForwardOnly = 0
        'adOpenKeySet      = 1
        'adOpenStatic      = 3

        'The CursorLocationEnum:
        '-------------------------
        'adUseClient = 3 | Uses client-side cursors supplied by a local cursor library.
        'Local cursor services often will allow many features
        'that driver-supplied cursors may not,
        'so using this setting may provide an advantage
        'with respect to features that will be enabled.
        'For backward compatibility, the synonym adUseClientBatch is also supported.
        'adUseServer = 2 | Default. Uses cursors supplied by the data provider or driver.
        'These cursors are sometimes very flexible and allow
        'for additional sensitivity to changes others make to the data source.
        'However, some features of the The Microsoft Cursor Service for OLE DB,
        'such as disassociated
        'Recordset objects, cannot be simulated with server-side cursors
        'and these features will be unavailable with this setting.
        'adUseNone   = 1 | Does not use cursor services. (
        'This constant is obsolete and appears solely for the sake of backward compatibility.)
        .CursorLocation = 3
        .ConnectionString = TaoChuoiKetNoi
        .Open
    End With
    ' Neu thanh cong dat gia tri ham bang 1
    KetNoi = 1
    ' Dong ket noi lai
    gcnObj.Close

ErrorExit:

    Exit Function

ErrorHandle:
    ' Neu co loi thi dat ma loi cho ham KetNoi
    KetNoi = Err.Number
    Err.Clear
    Resume ErrorExit

End Function




Sub ExcelToAccess()

    Dim lLastRow As Long, i As Long
    Dim wsObj As Worksheet
    Dim arrFieldNames As Variant, arrValues As Variant, arrRecordvals As Variant
    Dim sTableName As String
    Dim lCon  As Long
    Dim oRs   As Object

    On Error GoTo ErrorHandle


    sTableName = "TB_NhanVien"

    Set wsObj = Application.ThisWorkbook.Worksheets("CSDL")
    lLastRow = FindLastRow(wsObj)

    If lLastRow = 1 Then
        ' Neu hang bang 1
        ' Co nghia la khong co du lieu nao
        MsgBox "Khong co du lieu nao de xuat." & vbCrLf & _
               "Xin kiem tra lai.", vbOKOnly + vbInformation, gcsAppName
        GoTo ErrorExit
    End If

    lCon = KetNoi
    If lCon = 1 Then
        gcnObj.Open
        arrFieldNames = Array("MaSo", "HoTen", _
                              "NgaySinh", "GioiTinh", _
                              "NgayNhap")             'Ban co the thay doi tuy theo so truong cua ban

        ' Tang toc
        SpeedOn
        ' Tao bien recordset
        Set oRs = CreateObject("ADODB.Recordset")
        oRs.CursorLocation = 3                        'adUseClient
        oRs.Open sTableName, gcnObj, 3, 4             'adOpenStatic, adLockBatchOptimistic
        arrValues = wsObj.Range("A2:E" & lLastRow)
        For i = 1 To UBound(arrValues, 1)
            If Len(arrValues(i, 1)) > 0 Then
                arrRecordvals = Array(arrValues(i, 1), arrValues(i, 2), _
                                      arrValues(i, 3), arrValues(i, 4), _
                                      arrValues(i, 5))
                oRs.AddNew arrFieldNames, arrRecordvals
                Application.StatusBar = "Ban dang nhap ban ghi " & i & "/" & (lLastRow - 1)
            End If
        Next i

        Application.StatusBar = "Dang cap nhat...Xin cho trong giay lat..."
        oRs.UpdateBatch
        MsgBox "Ban da xuat du lieu thanh cong.", vbOKOnly + vbInformation, gcsAppName

    Else
        MsgBox "Khong the ket noi voi CSDL." & vbCrLf & _
               "Xuat du lieu khong thanh cong.", vbOKOnly + vbInformation, gcsAppName
    End If

ErrorExit:

    ' Giai phong bien
    Set wsObj = Nothing
    Set oRs = Nothing
    SpeedOff
    If Not gcnObj Is Nothing Then
        If (gcnObj.State And adStateOpen) = adStateOpen Then
            gcnObj.Close
        End If
    End If
    Exit Sub

ErrorHandle:
    ' Xu ly loi tai day
    MsgBox "Co loi xay ra. Xin kiem tra lai.", vbOKOnly + vbInformation, gcsAppName
    Debug.Print Err.Number & ", " & Err.Description
    Resume ErrorExit

End Sub




Sub AccessToExcel()
    Dim lCon  As Long
    Dim sSQL  As String
    Dim adoCommand As Object                          'ADODB.Command
    Dim oRs   As Object                               'ADODB.Recordset
    Dim wsObj As Worksheet
    Dim lLastRow As Long


    Set wsObj = Application.ThisWorkbook.Worksheets("FilterFromAccess")
    lLastRow = FindLastRow(wsObj)
    lCon = KetNoi
    If lCon <> 1 Then
        MsgBox "Khong the ket noi voi CSDL." & vbCrLf & _
               "Xuat du lieu khong thanh cong.", vbOKOnly + vbInformation, gcsAppName
    Else
        ' Mo ket noi
        gcnObj.Open
        sSQL = "SELECT MaSo, HoTen, NgaySinh, GioiTinh, NgayNhap " & _
               "FROM TB_NhanVien " & _
               "WHERE MaSo='" & "HTN005" & "';"
        Set adoCommand = CreateObject("ADODB.Command")
        With adoCommand
            .CommandType = 1                          '1: adCmdText, 2: adCmdTable, 4: adCmdStoredProc
            .ActiveConnection = gcnObj
            .CommandText = sSQL
        End With
        Set oRs = CreateObject("ADODB.Recordset")
        oRs.Open adoCommand, , 3, 4
        If oRs.EOF Then
            MsgBox "Khong co record nao thoa dieu kien.", vbOKOnly + vbInformation, gcsAppName
        Else
            wsObj.Cells(lLastRow + 1, 1).CopyFromRecordset oRs
        End If

    End If

ErrorExit:
    Set adoCommand = Nothing
    Set oRs = Nothing
    Set wsObj = Nothing
    If Not gcnObj Is Nothing Then
        If (gcnObj.State And adStateOpen) = adStateOpen Then
            gcnObj.Close
        End If
    End If
    Exit Sub

ErrorHandle:
    ' Xu ly loi tai day
    MsgBox "Co loi xay ra. Xin kiem tra lai.", vbOKOnly + vbInformation, gcsAppName
    Debug.Print Err.Number & ", " & Err.Description
    Resume ErrorExit
End Sub
Sub Spinner_getData()
  Dim ws As Worksheet 'Thi?t l?p bi?n ws là Sheet dang m?
    Set ws = ActiveSheet
  Dim rngData As Range  'Thi?t l?p vùng b?ng d? li?u b?t d?u t? ô D3
    Set rngData = Range(ws.Range("D3").End(xlDown), ws.Range("D3").End(xlToRight))
  Dim Data As Variant 'Xét n?i dung trong b?ng D? li?u
    Data = rngData.Value
  Dim iCell As Integer  'V? trí k?t qu? Spinner
    iCell = ws.Range("M2").Value

On Error Resume Next
Dim tbl As ListObject
Set tbl = ws.ListObjects("TableTest")

tbl.DataBodyRange.Delete        'Xoa du lieu trong Table
For x = LBound(Data) To UBound(Data)    'Xet vong lap trong toan bo Data tu dong dau toi dong cuoi
    If iCell = Data(x, 3) Then              'Neu ket qua Spinner = ket qua Data cot F
        With tbl.ListRows.Add                   'Lay du lieu vao Table
            .Range(, 2).Value = Data(x, 1)       'Lay noi dung cot 2
            .Range(, 3).Value = Data(x, 2)       'Lay noi dung cot 3
        End With
    End If
Next
tbl.DataBodyRange.AutoFit       'Tu chinh do rong trong Table
End Sub




Sub XuatPDF()
  'Tìm dòng cu?i b?ng kê
  Dim maxR As Integer
  maxR = Sheet1.Range("F" & Rows.Count).End(xlUp).Value   'Luu ý c?t c?n xác d?nh ? dây là c?t F

  'Xác d?nh du?ng d?n t?i thu m?c luu k?t qu?
  Set xSht = ActiveSheet
  Set xFileDlg = Application.FileDialog(msoFileDialogFolderPicker)
 
If xFileDlg.Show = True Then
   xFolder = xFileDlg.SelectedItems(1)
Else
   MsgBox "You must specify a folder to save the PDF into." & vbCrLf & vbCrLf & "Press OK to exit this macro.", vbCritical, "Must Specify Destination Folder"
   Exit Sub
End If
    
Application.ScreenUpdating = False  'B? qua vi?c c?p nh?t màn hình

For x = 1 To maxR
    With ActiveSheet.Range("M2")  '<== V? trí ô k?t qu? c?a Spin Button
      .Value = x  
      Call Spinner_getData      '<== G?i l?i câu l?nh l?y d? li?u vào PXK sau m?i l?n thay d?i k?t qu? Spin
        
    xFile = xFolder + "\" + xSht.Range("K4").Value + ".pdf"     'Xác d?nh tên file s? du?c luu, tên file l?y theo v? trí ô K4
        
      'Ki?m tra n?u tên file dã có s?n, b? trùng tên
        If Len(Dir(xFile)) > 0 Then
            xYesorNo = MsgBox(xFile & " already exists." & vbCrLf & vbCrLf & "Do you want to overwrite it?", _
                              vbYesNo + vbQuestion, "File Exists")
            On Error Resume Next
            If xYesorNo = vbYes Then
                Kill xFile
            Else
                MsgBox "if you don't overwrite the existing PDF, I can't continue." _
                            & vbCrLf & vbCrLf & "Press OK to exit this macro.", vbCritical, "Exiting Macro"
                Exit Sub
            End If
            If Err.Number <> 0 Then
                MsgBox "Unable to delete existing file.  Please make sure the file is not open or write protected." _
                            & vbCrLf & vbCrLf & "Press OK to exit this macro.", vbCritical, "Unable to Delete File"
                Exit Sub
            End If
        End If
         
        Set xUsedRng = xSht.UsedRange
        
        If Application.WorksheetFunction.CountA(xUsedRng.Cells) <> 0 Then
          'Luu du?i d?nh d?ng file PDF
            xSht.ExportAsFixedFormat Type:=xlTypePDF, Filename:=xFile, Quality:=xlQualityStandard
        Else
          MsgBox "The active worksheet cannot be blank"     'báo l?i tru?ng h?p b?ng kê không có d? li?u
          Exit Sub
        End If
     End With
Next  
    Application.ScreenUpdating = True     'm? l?i ch? d? c?p nh?t màn hình sau khi hoàn thành vòng l?p

    MsgBox "Well Done!"     'Thông báo hoàn thành công vi?c

End Sub
Sub Mail_ActiveSheet()
'Huong dan: http://www.hocexcel.online/gui-email-tu-excel-bang-vba
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set Sourcewb = ActiveWorkbook

    ActiveSheet.Copy
    Set Destwb = ActiveWorkbook


    With Destwb
        If Val(Application.Version) < 12 Then
            
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
           
            Select Case Sourcewb.FileFormat
            Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If .HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56: FileExtStr = ".xls": FileFormatNum = 56
            Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
            End Select
        End If
    End With

    TempFilePath = Environ$("temp") & "\"
    TempFileName = "Part of " & Sourcewb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .to = "hocexcel@hocexcel.online"
            .CC = ""
            .BCC = ""
            .Subject = "Thay doi tieu de thu o day"
            .Body = "Thay doi noi dung thu o day"
            .Attachments.Add Destwb.FullName

            .Send
        End With
        On Error GoTo 0
        .Close savechanges:=False
    End With

    'Xoa file vua gui
    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

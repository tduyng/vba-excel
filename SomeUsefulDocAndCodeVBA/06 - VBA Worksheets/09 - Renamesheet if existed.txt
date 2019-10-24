Sub Test()

    Debug.Print RenameSheet("Sheet1")
    Debug.Print RenameSheet("Sheet2")
    Debug.Print RenameSheet("ABC")

    Dim wrkSht As Worksheet
    Set wrkSht = Worksheets.Add
    wrkSht.Name = RenameSheet("DEF")

End Sub

    Public Function RenameSheet(SheetName As String, Optional Book As Workbook) As String

        Dim lCounter As Long
        Dim wrkSht As Worksheet

        If Book Is Nothing Then
            Set Book = ThisWorkbook
        End If

        lCounter = 0
        On Error Resume Next
            Do
                'Try and set a reference to the worksheet.
                Set wrkSht = Book.Worksheets(SheetName & IIf(lCounter > 0, "_" & lCounter, ""))
                If Err.Number <> 0 Then
                    'If an error occurs then the sheet name doesn't exist and we can use it.
                    RenameSheet = SheetName & IIf(lCounter > 0, "_" & lCounter, "")
                    Exit Do
                End If
                Err.Clear
                'If the sheet name does exist increment the counter and try again.
                lCounter = lCounter + 1
            Loop
        On Error GoTo 0

    End Function  



Sub RenameSheet()

Dim Sht                 As Worksheet
Dim NewSht              As Worksheet
Dim VBA_BlankBidSheet   As Worksheet
Dim newShtName          As String

' modify to your sheet's name
Set VBA_BlankBidSheet = Sheets("Sheet1")

VBA_BlankBidSheet.Copy After:=ActiveSheet    
Set NewSht = ActiveSheet

' you can change it to your needs, or add an InputBox to select the Sheet's name
newShtName = "New Name"

For Each Sht In ThisWorkbook.Sheets
    If Sht.Name = "New Name" Then
        newShtName = "New Name" & "_" & ThisWorkbook.Sheets.Count               
    End If
Next Sht

NewSht.Name = newShtName

End Sub



Private Sub nameNewSheet(sheetName As String, newSheet As Worksheet)
    Dim named As Boolean, counter As Long
    On Error Resume Next
        'try to name the sheet. If name is already taken, start looping
        newSheet.Name = sheetName
        If Err Then
            If Err.Number = 1004 Then 'name already used
                Err.Clear
            Else 'unexpected error
                GoTo nameNewSheet_Error
            End If
        Else
            Exit Sub
        End If

        named = False
        counter = 1

        Do
            newSheet.Name = sheetName & counter
            If Err Then
                If Err.Number = 1004 Then 'name already used
                    Err.Clear
                    counter = counter + 1 'increment the number until the sheet can be named
                Else 'unexpected error
                    GoTo nameNewSheet_Error
                End If
            Else
                named = True
            End If
        Loop While Not named

        On Error GoTo 0
        Exit Sub

    nameNewSheet_Error:
    'add errorhandler here

End Sub
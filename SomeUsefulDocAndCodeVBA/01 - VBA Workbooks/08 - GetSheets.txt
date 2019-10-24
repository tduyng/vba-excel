Sub GetSheets()
Path = “C:\Users\karra\Desktop\Bai Tap”
Filename = Dir(Path & “*.xls”)
Do While Filename <> “”
Workbooks.Open Filename:=Path & Filename, ReadOnly:=True
For Each Sheet In ActiveWorkbook.Sheets
Sheet.Copy After:=ThisWorkbook.Sheets(1)
Next Sheet
Workbooks(Filename).Close
Filename = Dir()
Loop
End Sub
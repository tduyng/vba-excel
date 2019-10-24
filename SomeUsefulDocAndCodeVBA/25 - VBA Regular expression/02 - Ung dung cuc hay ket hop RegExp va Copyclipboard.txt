'[a1:b30000] = [{"ABC123-009",""}]: Dim t As Double: t = Timer ' used for testing

Dim r As Range, s As String
Set r = ThisWorkbook.Worksheets("Data").UsedRange.Resize(, 1) ' Data!A1:A30000
With New MSForms.DataObject ' needs reference to "Microsoft Forms 2.0 Object Library" or use a bit slower late binding - With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
   r.Copy
   .GetFromClipboard
    Application.CutCopyMode = False
    s = .GetText
    .Clear ' optional - clear the clipboard if using Range.PasteSpecial instead of Worksheet.PasteSpecial "Text"

    With New RegExp ' needs reference to "Microsoft VBScript Regular Expressions 5.5" or use a bit slower late binding - With CreateObject("VBScript.RegExp")
        .Global = True
        '.IgnoreCase = False ' .IgnoreCase is False by default
        .Pattern = "[^0-9A-Za-z\r\n]+" ' because "[^\w\r\n]+" also matches _ and Unicode letters
        s = .Replace(s, vbNullString)
    End With

    .SetText s
    .PutInClipboard
End With

' about 70% of the time is spent here in pasting the data 
r(, 2).PasteSpecial 'xlPasteValues ' paste the text from clipboard in B1

'Debug.Print Timer - t
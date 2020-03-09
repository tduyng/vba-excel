Sub ADD_SHEET()
Dim sh As Worksheet
Dim i As Variant
For Each i In Feuil1.Range("NameSheet")
    Set sh = Sheets.Add(After:=Worksheets(Worksheets.Count))
    sh.Name = i.Value
Next

End Sub
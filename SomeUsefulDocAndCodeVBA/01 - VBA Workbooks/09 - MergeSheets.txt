Sub MergeSheets()
Const NHR = 1
Dim MWS As Worksheet
Dim AWS As Worksheet
Dim FAR As Long
Dim LR As Long
Set AWS = ActiveSheet
For Each MWS In ActiveWindow.SelectedSheets
If Not MWS Is AWS Then
FAR = AWS.UsedRange.Cells(AWS.UsedRange.Cells.Count).Row + 1
LR = MWS.UsedRange.Cells(MWS.UsedRange.Cells.Count).Row
MWS.Range(MWS.Rows(NHR + 1), MWS.Rows(LR)).Copy AWS.Rows(FAR)
End If
Next MWS
End Sub
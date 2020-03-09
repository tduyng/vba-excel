Sub InsertPictures()
Dim PicList() As Variant, lLoop As Variant
Dim PicFormat As String, xColIndex As Integer, xRowIndex As Integer
Dim Rng As Range
Dim sShape As Shape
On Error Resume Next
PicList = Application.GetOpenFilename(PicFormat, MultiSelect:=True)
xColIndex = Application.ActiveCell.Column
If IsArray(PicList) Then
    xRowIndex = Application.ActiveCell.Row
    For lLoop = LBound(PicList) To UBound(PicList)
        Set Rng = Cells(xRowIndex, xColIndex)
        Set sShape = ActiveSheet.Shapes.AddPicture(PicList(lLoop), msoFalse, msoCTrue, Rng.Left, Rng.Top, Rng.Width, Rng.Height)
        xRowIndex = xRowIndex + 1
    Next
End If
End Sub

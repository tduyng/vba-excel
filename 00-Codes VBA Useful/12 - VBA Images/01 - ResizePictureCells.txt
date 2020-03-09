Sub ResizePictureCells()
'Dim Picture As Variant
For Each Picture In ActiveSheet.DrawingObjects
PictureTop = Picture.Top
PictureLeft = Picture.Left
PictureHeight = Picture.Height
PictureWidth = Picture.Width
For N = 2 To 256
If Columns(N).Left > PictureLeft Then
PictureColumn = N - 1
Exit For
End If
Next N
For N = 2 To 65536
If Rows(N).Top > PictureTop Then
PictureRow = N - 1
Exit For
End If
Next N
Rows(PictureRow).RowHeight = PictureHeight
Columns(PictureColumn).ColumnWidth = PictureWidth * (54.29 / 288)
Picture.Top = Cells(PictureRow, PictureColumn).Top
Picture.Left = Cells(PictureRow, PictureColumn).Left
Next Picture
End Sub
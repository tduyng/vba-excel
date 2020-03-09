
Sub DELETE_SHAPES()
Dim sh As Worksheet
Dim rng As Range, arr As Variant, count_Shape As Integer, i As Integer
Dim shpAll As Shape, shp As Shape
Dim lastRow, firstRow As Integer
Set sh = Worksheets("Tirage")
Set rng = sh.Range("XOA")

ReDim arr(1 To sh.Shapes.count)
count_Shape = 1
i = 0
'On Error Resume Next
    For Each shp In sh.Shapes
        If Not Intersect(shp.TopLeftCell, rng) Is Nothing Then
            Debug.Print count_Shape
            arr(count_Shape) = shp.Name
            Debug.Print shp.Name
            'shp.Delete
            i = i + 1
            count_Shape = count_Shape + 1
            
        End If
        'count_Shape = count_Shape + 1
    Next shp
    Debug.Print i, count_Shape
    ReDim Preserve arr(count_Shape - 1)
   sh.Range("M1").Resize(UBound(arr), 1).Value = Application.Transpose(arr)
    
    'sh.Range("M20").Resize(UBound(arr), 1).Value = arr
End Sub

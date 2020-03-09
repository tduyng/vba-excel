'// Truyen cac Items cua Collection vao Array (2 chieu)
Public Function CollectionToArray(myCol As Collection) As Variant
    Dim arr() As Variant, i As Long
    ReDim arr(1 To myCol.Count, 1 To 1)
    For i = 1 To myCol.Count
        arr(i, 1) = myCol.Item(i)
    Next i
    CollectionToArray = arr
End Function
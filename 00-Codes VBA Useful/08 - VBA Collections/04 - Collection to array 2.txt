Sub COLLECTION_TO_ARRAY()
    Dim coll As New Collection
    Dim arr As Variant
    Dim sh As Worksheet
    Dim rng As Range, rngi As Range, i As Integer
    
    Set sh = ActiveSheet
    Set rng = sh.Range("A6:A17")
    For Each rngi In rng
        coll.Add rngi.Value
    Next rngi
    
    'truong hop mang 1 chieu: 1 hang, nhieu cot
    'ReDim arr(1 To coll.Count)
    'For i = 1 To coll.Count
        'arr(i) = coll.Item(i)
    'Next i
    'sh.Range("B1").Resize(UBound(arr, 1)) = Application.Transpose(arr)
    
    'truong hop mang 2 chieu: n hang , 1 cot
    ReDim arr(1 To coll.Count, 1 To 1)
    For i = 1 To coll.Count
        arr(i, 1) = coll.Item(i)
    Next i
    sh.Range("B1").Resize(UBound(arr, 1), 1) = arr
End Sub

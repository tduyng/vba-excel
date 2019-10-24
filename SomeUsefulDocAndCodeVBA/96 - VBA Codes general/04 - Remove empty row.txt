Option Explicit

Sub REMOVE_EMPTY_ROW()
    Dim rRange As Range
    Dim firstRow As Long, lastRow As Long, usedRow As Long
    Dim firstCol As Long, lastCol As Long
    Dim firstColTxt As String, lastColTxt As String, usedCol As Long
    Dim i As Long, Target As Range, j As Long
    Set rRange = Application.InputBox(Prompt:="Select rows to remove", Title:="Remove empty rows", Type:=8)
    
    firstRow = rRange.Row
    usedRow = rRange.Rows.Count
    lastRow = usedRow + firstRow - 1
    firstColTxt = Split(rRange.Address, "$")(1)
    lastColTxt = Split(rRange.Address, "$")(3)
    firstCol = rRange.Column
    usedCol = rRange.Columns.Count
    lastCol = usedCol + firstCol - 1
    Application.ScreenUpdating = False
    For i = lastRow To firstRow Step -1
        For j = firstCol To lastCol
            Set Target = ActiveSheet.Cells(i, j)
            If Application.CountA(Target) = 0 Then
            Target.Delete Shift:=xlUp
            End If
        Next j
    Next i
    Application.ScreenUpdating = True
End Sub

Sub COLLECTION_TO_ARRAY()
    Dim coll As New Collection
    Dim arr As Variant
    Dim sh As Worksheet
    Dim rng As Range, rngi As Range, i As Integer
    
    Set sh = ActiveSheet
    Set rng = sh.Range("A6:A17")
    For Each rngi In rng
        coll.Add rngi.value
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

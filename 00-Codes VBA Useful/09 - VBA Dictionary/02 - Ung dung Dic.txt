Option Explicit

Sub UNIQUE_ITEMS()
    Dim sh1 As Worksheets, rng As Range
    Dim i As Integer, key As Variant, items As Variant
    
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.CompareMode = vbBinaryCompare

    Set rng = Sheet1.Range("A1:A" & Sheet1.Range("A" & Rows.Count).End(xlUp).Row)

    On Error Resume Next
    For Each key In rng
        Dic.Add key.Value, ""
    Next
    On Error GoTo 0

    Sheet1.Range("C1").Resize(Dic.Count, 1) = Application.Transpose(Dic.Keys)
    'items = Application.Transpose(Dic.Keys)
    
End Sub


Sub UNIQUE_ITEMS2()
    Dim sh1 As Worksheets, rng As Range
    Dim i As Integer, key As Variant, items As Variant
    
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.CompareMode = vbBinaryCompare

    Set rng = Sheet1.Range("A1:A" & Sheet1.Range("A" & Rows.Count).End(xlUp).Row)


    For Each key In rng
        Dic.Item(key.Value) = key.Value
    Next key
    items = Dic.items

    Sheet1.Range("D1").Resize(Dic.Count, 1) = Application.Transpose(Dic.Keys)
    'items = Application.Transpose(Dic.Keys)
    
End Sub



Sub UNIQUE_ITEMS3()
    Dim sh1 As Worksheets, rng As Range
    Dim i As Integer, key As Variant, items As Variant
    
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.CompareMode = vbBinaryCompare

    Set rng = Sheet1.Range("A1:A" & Sheet1.Range("A" & Rows.Count).End(xlUp).Row)

    On Error Resume Next
    For Each key In rng
        Dic.Add key.Value, ""
    Next
    On Error GoTo 0

    Sheet1.Range("C1").Resize(Dic.Count, 1) = Application.Transpose(Dic.Keys)
    'items = Application.Transpose(Dic.Keys)
    
End Sub


Sub TONG_HOP_EX_DIC()
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dim sh As Worksheet, rng As Range, arr As Variant, key As Variant
    Dim i As Integer, j As Integer
    Dim lastRow As Integer
    
    Set sh = Worksheets(Sheet2.Name)
    lastRow = sh.Range("A" & Rows.Count).End(xlUp).Row
    Set rng = sh.Range("A2:B" & lastRow)
    arr = rng.Value

    For i = 1 To UBound(arr, 1)
        If Not Dic.Exists(arr(i, 1)) Then
            Dic.Add arr(i, 1), arr(i, 2)
        Else
            Dic.Item(arr(i, 1)) = Dic.Item(arr(i, 1)) + arr(i, 2)
        End If
    Next i
    
    
    Sheet2.Range("C2").Resize(Dic.Count, 1) = Application.Transpose(Dic.Keys)
    Sheet2.Range("D2").Resize(Dic.Count, 1) = Application.Transpose(Dic.items)
    
    
End Sub

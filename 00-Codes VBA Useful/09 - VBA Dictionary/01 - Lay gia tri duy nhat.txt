'//Loc loai trung mot cot'
Function UniqueColumn1D(ByVal rng As Range) As Variant
    If rng.Count = 1 Then UniqueColumn1D = rng.Value: Exit Function
    Dim Dic As Object, i As Long, arr()
    arr = rng.Value
    Set Dic = CreateObject("Scripting.Dictionary")
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) <> "" And Dic.Exists(arr(i, 1)) = False Then
            Dic.Add arr(i, 1), ""
        End If
    Next i
    UniqueColumn1D = Dic.Keys
End Function

Sub FILTER_UNIQUE_ITEMS()
    Dim sh As Worksheet, arr As Variant, i As Integer
    Dim rng As Range
    Dim Dic As Object
    
    Set sh = ActiveSheet
    Set rng = sh.Range("I3:I27")
    arr = rng.Value
    Set Dic = CreateObject("Scripting.Dictionary")
     
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) <> "" And Dic.Exists(arr(i, 1)) = False Then
            Dic.Add arr(i, 1), ""
        End If
    Next i
    sh.Range("J3").Resize(5, 1) = Application.Transpose(Dic.Keys)
End Sub

Sub RESIZE_ARRAY()
    Dim Array1 As Variant, Array2 As Variant
    Dim ColExport As Variant, ColVar As Variant
    Dim I As Long
    Application.ScreenUpdating = False
    
    Array1 = Sheet2.Range("Array1").Value
    ColExport = [{1,3,5,7,6,2}]
    ReDim ColVar(1 To UBound(Array1, 1), 1 To 1)
    
    For I = LBound(Array1, 1) To UBound(Array1, 1)
        ColVar(I, 1) = I
    Next
    Sheet2.Range("K3").Resize(UBound(Array1, 1), 7).ClearContents
    Sheet2.Range("K3").Resize(UBound(Array1, 1), UBound(ColExport, 1)) = Application.Index(Array1, ColVar, ColExport)
    
    Application.ScreenUpdating = True
End Sub
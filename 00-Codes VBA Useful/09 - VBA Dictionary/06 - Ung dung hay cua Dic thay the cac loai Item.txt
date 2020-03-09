Option Explicit

Sub VBA101_MinhThao()
    Dim sh As Worksheet
    Dim Dic As Object
    Dim i&, j&, k&
    Dim rng As Range, rngVal As Range
    Dim arr(), arrTemps(), keyDic(), itemDic()
    Dim iItemNull As Variant, iKeyDic As Variant
    Dim val1 As String
    Set Dic = CreateObject("Scripting.Dictionary")
    
    Set sh = ThisWorkbook.Sheets("VBA101")
    Set rng = sh.Range("A2:D11")
    arr = rng.Value
    ReDim arrTemps(1 To UBound(arr), 1 To 4)
    

    For i = 1 To UBound(arr)
        If Not Dic.exists(arr(i, 1)) Then
            Dic.Add arr(i, 1), arr(i, 3)
        Else
            Dic.Item(arr(i, 1)) = Dic.Item(arr(i, 1)) + arr(i, 3)
        End If
        
    Next i
    
    sh.Range("G2").Resize(Dic.Count, 1) = Application.Transpose(Dic.keys)
    sh.Range("I2").Resize(Dic.Count, 1) = Application.Transpose(Dic.Items)

    For Each iKeyDic In Dic.keys
        k = 1
      For i = 1 To UBound(arr)
        If iKeyDic = arr(i, 1) Then
            If k = 1 Then
                Dic.Item(iKeyDic) = arr(i, 4)
                val1 = Dic.Item(iKeyDic)
            Else
                If InStr(Dic.Item(iKeyDic), arr(i, 4)) = False Then Dic.Item(iKeyDic) = Dic.Item(iKeyDic) & " & " & arr(i, 4)
                Debug.Print InStr(Dic.Item(iKeyDic), arr(i, 4)), Dic.Item(iKeyDic), arr(i, 4)
            End If
            k = k + 1
        End If
      Next i
      sh.Range("J2").Resize(Dic.Count, 1) = Application.Transpose(Dic.Items)
      k = 1
     For i = 1 To UBound(arr)
        If iKeyDic = arr(i, 1) Then
            If k = 1 Then
                Dic.Item(iKeyDic) = arr(i, 2)
                val1 = Dic.Item(iKeyDic)
            Else
                If InStr(Dic.Item(iKeyDic), arr(i, 2)) = False Then Dic.Item(iKeyDic) = Dic.Item(iKeyDic) & " & " & arr(i, 2)
                Debug.Print InStr(Dic.Item(iKeyDic), arr(i, 2)), Dic.Item(iKeyDic), arr(i, 2)
            End If
            k = k + 1
        End If
      Next i
    Next iKeyDic
    sh.Range("H2").Resize(Dic.Count, 1) = Application.Transpose(Dic.Items)
End Sub

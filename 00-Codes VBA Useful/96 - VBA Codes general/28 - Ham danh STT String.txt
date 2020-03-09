'========================================================================================
'                             Order String
'========================================================================================
Sub test_OrderChar()
  Dim i, j
  j = 16384
  For i = 1 To j
    If OrderChar$(i) <> reCapCol(i) Then Exit For
    Debug.Print OrderChar$(i)
  Next
End Sub
    Function OrderChar$(Optional ByVal StartArea = 1, _
                        Optional ByVal IsLower As Boolean = False, _
                        Optional ByVal StrChar$ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ")
        If StartArea < 1 And StrChar = "" Then Exit Function
        Dim LenFT%, reChar$, ModB, ModD
        LenFT = Len(StrChar)
        If StartArea <= LenFT Then reChar = Mid(StrChar, StartArea, 1): GoTo result
        ModB = StartArea Mod LenFT
        Do
          If StartArea <= LenFT Then
            reChar = reChar & Mid(StrChar, IIf(ModB = 0, LenFT, ModB), 1)
            Exit Do
          Else
            StartArea = Int((StartArea - 1) / LenFT)
            ModD = StartArea Mod LenFT
            reChar = Mid(StrChar, IIf(ModD = 0, LenFT, ModD), 1) & reChar
          End If
        Loop
result:
        OrderChar = IIf(IsLower, LCase(reChar), reChar)
    End Function
'------------------------------------------
Sub Test_LocalOrder()
    Debug.Print reCapCol(16384)
    Debug.Print areaSeriesStr(reCapCol(16384))
End Sub
    Function LocalOrder&(Optional ByVal StartArea)  'correction
        Dim StrChar, i&, numStr&
        StrChar = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        If isArray(StartArea) Then
            Dim A, Temp&
            For Each A In StartArea
                If LocalOrder(A) > Temp Then
                    Temp = LocalOrder(A)
                End If
            Next A
            LocalOrder = Temp: Exit Function
        Else
            If IsMissing(StartArea) Then GoTo reArea
            If StartArea = vbNullString Then GoTo reArea
            For i = 1 To Len(StartArea)
                If InStr(1, StrChar, Mid(StartArea, i, 1), vbTextCompare) = 0 Then GoTo reArea
                numStr = numStr * 26 + InStr(1, StrChar, Mid(StartArea, i, 1), vbTextCompare)
            Next i
            LocalOrder = numStr: Exit Function
        End If
reArea:
        LocalOrder = 0
    End Function
    'L?y d?u d? c?a Excel
    Function reCapCol(ByVal numCol As Long, Optional ByVal bLCase As Boolean = False) As String
        'Limit 16384 (XL 2016)
        reCapCol = Split(Columns(numCol).Address(True, False), ":")(0)
        reCapCol = IIf(bLCase, LCase(reCapCol), reCapCol)
    End Function
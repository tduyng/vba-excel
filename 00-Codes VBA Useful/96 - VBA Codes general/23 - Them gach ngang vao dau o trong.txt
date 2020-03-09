Sub GhiThietBiThiCongCanThiet()
 Dim Rng As Range, sRng As Range, Cls As Range
 Dim MyAdd As String
 Dim Rws As Long, W As Integer
 
 With Sheets("Diary").[I1]
    Rws = .CurrentRegion.Rows.Count
    Set Rng = .Resize(Rws)
 End With
 Sheets("May Thi Cong").Select
 For Each Cls In Range([B2], [B2].End(xlDown))
    Set sRng = Rng.Find(Cls.Value, , xlFormulas, xlPart)
    If Not sRng Is Nothing Then
        MyAdd = sRng.Address
        W = W + 1:                  If W = 12 Then W = 0
        Do
            sRng.Interior.ColorIndex = 34 + W
            Sheets("Diary").Cells(sRng.Row, "T").Value = Cls.Offset(, 1).Value
            Set sRng = Rng.FindNext(sRng)
        Loop While Not sRng Is Nothing And sRng.Address <> MyAdd
    End If
 Next Cls
End Sub

Sub GhiThietBiThiCongCanThiet()
 Dim Rng As Range, sRng As Range, Cls As Range
 Dim MyAdd As String
 Dim Rws As Long, W As Integer
 
 With Sheets("Diary").[I1]
    Rws = .CurrentRegion.Rows.Count
    Set Rng = .Resize(Rws)
 End With
 Sheets("May Thi Cong").Select
 For Each Cls In Range([B2], [B2].End(xlDown))
    Set sRng = Rng.Find(Cls.Value, , xlFormulas, xlPart)
    If Not sRng Is Nothing Then
        MyAdd = sRng.Address
        W = W + 1:                  If W = 12 Then W = 0
        Do
            sRng.Interior.ColorIndex = 34 + W
            Sheets("Diary").Cells(sRng.Row, "T").Value = Cls.Offset(, 1).Value
            Set sRng = Rng.FindNext(sRng)
        Loop While Not sRng Is Nothing And sRng.Address <> MyAdd
    End If
 Next Cls
End Sub


Public Sub DungCuThiCong()
Dim Dic As Object, sArr(), dArr(), tArr(), Tmp, Txt As String
Dim I As Long, J As Long, N As Long, R1 As Long, R2 As Long
Set Dic = CreateObject("Scripting.Dictionary")
tArr = Sheets("May thi cong").Range("B2", Sheets("May thi cong").Range("B2").End(xlDown)).Resize(, 2).Value
R2 = UBound(tArr)
With Sheets("DIARY")
    sArr = .Range("I2", .Range("I60000").End(xlUp)).Value
    R1 = UBound(sArr)
    ReDim dArr(1 To R1, 1 To 1)
    For I = 1 To R1
        Dic.RemoveAll
        Tmp = Split(sArr(I, 1), ChrW(10))
        For J = LBound(Tmp) To UBound(Tmp)
            Txt = Tmp(J)
            For N = 1 To R2
                If Txt Like tArr(N, 1) & "*" Then
                    If Not Dic.Exists(tArr(N, 2)) Then
                        Dic.Item(tArr(N, 2)) = ""
                        dArr(I, 1) = dArr(I, 1) & "; " & tArr(N, 2)
                    End If
                End If
            Next N
        Next J
    Next I
    For I = 1 To R1
        If Len(dArr(I, 1)) Then
            dArr(I, 1) = Mid(dArr(I, 1), 3)
        End If
    Next I
    .Range("T2").Resize(R1) = dArr
End With
Set Dic = Nothing
End Sub


Sub GhiThietBiThiCongCanThiet()
 Dim Rng As Range, sRng As Range, Cls As Range
 Dim MyAdd As String
 Dim Rws As Long, W As Integer, VTr As Integer                  '*          '
 
 With Sheets("Diary").[I1]
    Rws = .CurrentRegion.Rows.Count
    Set Rng = .Resize(Rws)
 End With
 Sheets("May Thi Cong").Select
 For Each Cls In Range([B2], [B2].End(xlDown))
    Set sRng = Rng.Find(Cls.Value, , xlFormulas, xlPart)
    If Not sRng Is Nothing Then
        MyAdd = sRng.Address
        W = W + 1:                  If W = 12 Then W = 0
        Do
            sRng.Interior.ColorIndex = 34 + W
            With Sheets("Diary").Cells(sRng.Row, "T")               '|=>    '
                VTr = InStr(.Value, Cls.Offset(, 1).Value)
                If VTr < 1 Then
                    .Value = .Value & ", " & Cls.Offset(, 1).Value
                End If
            End With                                                '<=|    '
            Set sRng = Rng.FindNext(sRng)
        Loop While Not sRng Is Nothing And sRng.Address <> MyAdd
    End If
 Next Cls
End Sub


Sub Do_Tim_May_Thi_Cong()
Dim sh As Worksheet, MayThiCong(), sArr(), i As Long, ii As Long, Res()
With Sheets("May Thi Cong")
   MayThiCong = .Range("A2", .[A65536].End(3)).Resize(, 3).Value
End With
For Each sh In ThisWorkbook.Worksheets
   If Replace(LCase(sh.Name), " ", "") <> "maythicong" Then
      sArr = sh.Range("A2", sh.[A65536].End(3)).Resize(, 9).Value
      ReDim Res(1 To UBound(sArr), 1 To 1)
      For i = 1 To UBound(sArr)
         If sArr(i, 9) <> Empty Then
            For ii = 1 To UBound(MayThiCong)
               If sArr(i, 9) Like "*" & MayThiCong(ii, 2) & "*" Then
                  Res(i, 1) = MayThiCong(ii, 3)
               End If
            Next
         End If
      Next
      sh.[T2].Resize(UBound(Res), UBound(Res, 2)) = Res
   End If
Next
End Sub

Sub An()
Union([E1], [F1], [K1:L1]).EntireColumn.Hidden = True
End Sub
Sub Hien()
Union([E1], [F1], [K1], [L1]).EntireColumn.Hidden = False
End Sub
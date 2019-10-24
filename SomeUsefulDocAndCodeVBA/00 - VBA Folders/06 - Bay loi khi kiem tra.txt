Sub Main_OpenFileName()
  Dim arr, vFile
  vFile = Application.GetOpenFilename("Excel Files, *.xls;*.xlsx;*.xlsm")
  If TypeName(vFile) = "String" Then
    arr = GetData(CStr(vFile))
    If IsArray(arr) Then
      ThisWorkbook.Sheets(1).Range("A1").Resize(UBound(arr, 1) + 1, UBound(arr, 2) + 1).Value = arr
      MsgBox "Data has been successfully imported!"
    End If
  End If
End Sub


Sub DoSomething()
    Dim fso As Object, re As Object, Match As Object, oMatches As Object
    Dim FileName, Arr(), i As Long, s As String
        
        FileName = Application.GetOpenFilename()
        If FileName <> "False" Then
            Set fso = CreateObject("Scripting.FileSystemObject")
            With fso.OpenTextFile(FileName)
                s = .ReadAll
                .Close
            End With
            Set fso = Nothing
            Set re = CreateObject("VBScript.RegExp")
            With re
                .Global = True
                .Pattern = "<(!.+;)(?:\n|.)+?(?=(?:<!.+;|$))"
                Set Match = .Execute(s)
                If Not Match Is Nothing Then
                    ReDim Arr(1 To Match.count, 1 To 4)
                    
                    For i = 0 To Match.count - 1
                        Arr(i + 1, 1) = Match(i).SubMatches(0)
                        
                        .Pattern = "\d{9}\s"
                        Set oMatches = .Execute(Match(i))
                        If Not oMatches Is Nothing Then Arr(i + 1, 2) = oMatches.count
                        
                        .Pattern = "\sTBO-1\s"
                        Set oMatches = .Execute(Match(i))
                        If Not oMatches Is Nothing Then Arr(i + 1, 3) = oMatches.count
                        
                        .Pattern = "\sTBO-2\s"
                        Set oMatches = .Execute(Match(i))
                        If Not oMatches Is Nothing Then Arr(i + 1, 4) = oMatches.count
                    Next
                    [A2].Resize(UBound(Arr), 4) = Arr
                End If
            End With
        End If
End Sub



Sub DoSomething()
    Dim fso As Object, re As Object, Match As Object, oMatches As Object, r As Long, count As Long
    Dim FileName, Arr() As String, i As Long, s As String, result() As String
        FileName = Application.GetOpenFilename()
        If FileName <> "False" Then
            Set fso = CreateObject("Scripting.FileSystemObject")
            With fso.OpenTextFile(FileName)
                s = .ReadAll
                .Close
            End With
            
            Set re = CreateObject("VBScript.RegExp")
            With re
                .Global = True
                .Pattern = "<(!.+;)(?:\n|.)+?END"
                Set Match = .Execute(s)
                If Not Match Is Nothing Then
                    [B][COLOR=#ff0000]s = .Replace(s, "")
                    With fso.CreateTextFile(FileName & "_error.log")
                        .Write s
                        .Close
                    End With[/COLOR][/B]
                    
                    ReDim Arr(1 To 1)
                    .Pattern = "\s\d{2}(\d{7})\s"
                    For i = 0 To Match.count - 1
                        Set oMatches = .Execute(Match(i))
                        If Not oMatches Is Nothing Then
                            ReDim Preserve Arr(1 To count + 2 + oMatches.count)
                            count = count + 1
                            Arr(count) = Match(i).SubMatches(0)
                            For r = 0 To oMatches.count - 1
                                count = count + 1
                                Arr(count) = "n=" & oMatches(r).SubMatches(0) & ";"
                            Next r
                            count = count + 1
                        End If
                        Set oMatches = Nothing
                    Next
                End If
                Set Match = Nothing
            End With
            Set re = Nothing
            [B][COLOR=#0000ff]Set fso = Nothing
            
            If count Then
                ReDim result(1 To count - 1, 1 To 1)
                For r = 1 To count - 1
                    result(r, 1) = Arr(r)
                Next r
                [A1].Resize(UBound(result)) = result
            End If
        End If
End Sub


'cod ephien ban 2 cho bai 19
Sub DoSomething()
    Dim fso As Object, re As Object, Match As Object, oMatches As Object
    Dim FileName, Arr(), i As Long, s As String
        
        FileName = Application.GetOpenFilename()
        If FileName <> "False" Then
            Set fso = CreateObject("Scripting.FileSystemObject")
            With fso.OpenTextFile(FileName)
                s = .ReadAll
                .Close
            End With
            Set fso = Nothing
            Set re = CreateObject("VBScript.RegExp")
            With re
                .Global = True
                .Pattern = "<(!.+;)(?:\n|.)+?(?=(?:<!.+;|$))"
                Set Match = .Execute(s)
                If Not Match Is Nothing Then
                    ReDim Arr(1 To Match.count, 1 To 4)
                    
                    For i = 0 To Match.count - 1
                        Arr(i + 1, 1) = Match(i).SubMatches(0)
                        
                        .Pattern = "<suscp:snb=(?:\n|.)+?END
                        Set oMatches = .Execute(Match(i))
                        If Not oMatches Is Nothing Then Arr(i + 1, 2) = oMatches.count
                        
                        .Pattern = "\sTBO-1\s"
                        Set oMatches = .Execute(Match(i))
                        If Not oMatches Is Nothing Then Arr(i + 1, 3) = oMatches.count
                        
                        .Pattern = "\sTBO-2\s"
                        Set oMatches = .Execute(Match(i))
                        If Not oMatches Is Nothing Then Arr(i + 1, 4) = oMatches.count
                    Next
                    [A2].Resize(UBound(Arr), 4) = Arr
                End If
            End With
        End If
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
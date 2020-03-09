Option Explicit
    Dim Dic As Object
    
    Dim rng As Range, rngi As Range, arr(), sh As Worksheet, KQ(), shi As Worksheet
    Dim key As Variant, i As Integer, j As Integer, lastRow As Long, n As Integer, k As Integer
Sub TACH_SHEET_ARRDIC()
    Dim countWs As Integer
    Set Dic = CreateObject("Scripting.Dictionary")
    
    
    Application.ScreenUpdating = False
    lastRow = Sheet1.Range("A" & Rows.Count).End(xlUp).Row
    arr = Sheet1.Range("A2:C" & lastRow).Value

    For i = 1 To lastRow - 1
        If Not Dic.exists(Sheet1.Range("B" & i + 1).Value) Then Dic.Add Sheet1.Range("B" & i + 1).Value, vbNullString
    Next i
    key = Dic.keys
    ReDim Preserve key(0 To Dic.Count - 1)
    
    For i = 0 To Dic.Count - 1

        k = 0
        ReDim KQ(1 To UBound(arr), 1 To 3)
        For j = 1 To UBound(arr)
            If arr(j, 2) = key(i) Then
                k = k + 1
                For n = 1 To 3
                    KQ(k, n) = arr(j, n)
                Next n
            End If
        Next j
        
        countWs = 0
        For Each shi In Worksheets()
            If Dic.exists(shi.Name) Then
                countWs = countWs + 1
            End If
        Next shi
        Select Case countWs
            Case Is = 0
                ADD_SHEETS
            Case Is > 0
                SET_NAME_SHEET
        End Select
    Next i
    Application.ScreenUpdating = True
End Sub

Sub ADD_SHEETS()
        With Worksheets.Add(after:=Worksheets(Worksheets.Count))
        .Name = key(i)
        Sheet1.Range("A1:C1").Copy .Range("A1")
        .Range("A2").Resize(k, 3) = KQ
        .Columns("A:C").AutoFit
        End With
End Sub

Sub SET_NAME_SHEET()

        Set sh = Worksheets(key(i))
        With sh
            Sheet1.Range("A1:C1").Copy .Range("A1")
            .Range("A2").Resize(k, 3) = KQ
                .Columns("A:C").AutoFit
        End With
        Erase KQ
End Sub


Sub TACH_SHEET_AUTOFILTER()
    Dim countWs As Integer
    Set Dic = CreateObject("Scripting.Dictionary")
    
    
    Application.ScreenUpdating = False

    lastRow = Sheet1.Range("A" & Rows.Count).End(xlUp).Row
    arr = Sheet1.Range("A2:C" & lastRow).Value
    Set rng = Sheet1.Range("A1:C" & lastRow)
    
    For i = 1 To lastRow - 1
        If Not Dic.exists(Sheet1.Range("B" & i + 1).Value) Then Dic.Add Sheet1.Range("B" & i + 1).Value, vbNullString
    Next i
    For Each key In Dic.keys
        rng.AutoFilter field:=2, Criteria1:=key
                countWs = 0
        For Each shi In Worksheets()
            If Dic.exists(shi.Name) Then
                countWs = countWs + 1
            End If
        Next shi
        Select Case countWs
            Case Is = 0
                With Worksheets.Add(after:=Worksheets(Worksheets.Count))
                    .Name = key
                End With
        End Select
        rng.EntireRow.Copy Worksheets(key).Range("A1")
        Worksheets(key).Columns("A:C").AutoFit
        Sheet1.AutoFilterMode = False
    Next key

    Application.ScreenUpdating = True
End Sub

Sub CREATE_SHEET_FROM_DATA_AVAILABLE()
    Dim i&, j&, k&, n As Long
    Dim arr(), keyDic(), KQ(), t As Single
    Dim Dic As Object
    Dim rng As Range, sh As Worksheet, wb As Workbook
    Set Dic = CreateObject("Scripting.Dictionary")
Option Explicit
Public Sub FastWB(Optional ByVal opt As Boolean = True)
    With Application
        .Calculation = IIf(opt, xlCalculationManual, xlCalculationAutomatic)
        .DisplayAlerts = Not opt
        .DisplayStatusBar = Not opt
        .EnableAnimations = Not opt
        .EnableEvents = Not opt
        .ScreenUpdating = Not opt
    End With
    FastWS , opt
End Sub

Public Sub FastWS(Optional ByVal ws As Worksheet = Nothing, _
                  Optional ByVal opt As Boolean = True)
    If ws Is Nothing Then
        For Each ws In Application.ActiveWorkbook.Sheets
            EnableWS ws, opt
        Next
    Else
        EnableWS ws, opt
    End If
End Sub

Private Sub EnableWS(ByVal ws As Worksheet, ByVal opt As Boolean)
    With ws
        .DisplayPageBreaks = False
        .EnableCalculation = Not opt
        .EnableFormatConditionsCalculation = Not opt
        .EnablePivotTable = Not opt
    End With
End Sub

Sub CREATE_SHEET_FROM_DATA_AVAILABLE()
    Dim i&, j&, k&, n As Long
    Dim arr(), keyDic(), KQ(), t As Single
    Dim Dic As Object
    Dim rng As Range, sh As Worksheet, wb As Workbook
    Set Dic = CreateObject("Scripting.Dictionary")
    Set rng = Sheet4.Range("A2:F" & Sheet4.[A65536].End(xlUp).Row)
    
    FastWB True: t = Timer
    
    arr = rng.Value
    For i = 1 To UBound(arr)
        If Not Dic.Exists(arr(i, 6)) Then
            Dic.Add arr(i, 6), arr(i, 6)
        End If
    Next i
    
    keyDic = Dic.keys
    Application.ScreenUpdating = False
    
    For i = 0 To UBound(keyDic)
        k = 0
        ReDim KQ(1 To UBound(arr), 1 To 6)
        For j = 1 To UBound(arr)
            If arr(j, 6) = keyDic(i) Then
                k = k + 1
                For n = 1 To 6
                    KQ(k, n) = arr(j, n)
                Next n
            End If
        Next j
        
        Set wb = Workbooks.Add
        With wb
            Sheet4.Range("A1:F1").Copy .Worksheets(1).Range("A1")
            .Worksheets(1).Range("A2").Resize(k, 6) = KQ
            .Worksheets(1).Columns("A:F").AutoFit
            .SaveAs ThisWorkbook.Path & "\" & keyDic(i) & ".xlsx"
            .Close
        End With
        Erase KQ
        Set wb = Nothing
    Next i
    
    Application.ScreenUpdating = True
    FastWB False:  MsgBox "done " & Round((Timer - t), 2) & "secondes"
    
End Sub

Sub TAO_FILE_OPTI()
    Dim Dic As Object
    Dim i&, j&, k&, n As Long, t As Single
    Dim rngKey As Range, wb As Workbook, rngi As Range, rngTotal As Range
    Dim arr(), keyDic()
    Dim strFilter As String
    
    FastWB True: t = Timer
    Set Dic = CreateObject("Scripting.Dictionary")
    
    Set rngTotal = Sheet4.Range("A1:F" & Sheet4.[A65536].End(xlUp).Row)
    Set rngKey = Sheet4.Range("F2:F" & Sheet4.[A65536].End(xlUp).Row)
    
    For Each rngi In rngKey
        If Not Dic.Exists(rngi.Value) Then Dic.Add rngi.Value, rngi.Value
    Next rngi
    
    keyDic = Dic.keys
    For i = 0 To UBound(keyDic)
        With rngTotal
            .AutoFilter field:=6, Criteria1:=keyDic(i)
            .Copy
        End With
        Set wb = Workbooks.Add
        With wb
            .Worksheets(1).Range("A2").PasteSpecial xlPasteColumnWidths
            .Worksheets(1).Range("A2").PasteSpecial xlPasteAll
            .Worksheets(1).Range("A2").Select
            .Worksheets(1).Range("A2").Copy
'            Sheet4.Range("A1:F1").Copy .Worksheets(1).Range("A1")
            .Worksheets(1).Columns("A:F").AutoFit
            .SaveAs ThisWorkbook.Path & "\" & keyDic(i) & ".xlsx"
            .Close
        End With

        rngTotal.AutoFilter
    Next i

    FastWB False: MsgBox "Done" & Round((Timer - t), 2) & "secondes"
End Sub



Option Explicit
Sub chi_nhanh()
    Dim Dic As Object, i&, j&, arr(), KQ(), ARrDic(), k&, wb As Workbook, n&
    Set Dic = CreateObject("Scripting.dictionary")
    Dim t As Single
    t = Timer
    arr = Sheet4.Range("A2:F" & Sheet4.[A65536].End(xlUp).Row).Value
    For i = 1 To UBound(arr)
        If Not Dic.exists(arr(i, 6)) Then Dic.Add arr(i, 6), arr(i, 6)
    Next
    ARrDic = Dic.Keys
    Application.ScreenUpdating = False
    For i = 0 To UBound(ARrDic)
        k = 0
        ReDim KQ(1 To UBound(arr), 1 To 6)
        For j = 1 To UBound(arr)
            If arr(j, 6) = ARrDic(i) Then
               k = k + 1
               For n = 1 To 6
                   KQ(k, n) = arr(j, n)
               Next
            End If
        Next
        Set wb = Workbooks.Add
        With wb
            Sheet4.Range("A1:F1").Copy .Worksheets(1).Range("A1")
            .Worksheets(1).Range("A2").Resize(k, 6) = KQ
            .Worksheets(1).Columns("A:F").AutoFit
            .SaveAs ThisWorkbook.Path & "\" & ARrDic(i) & ".xlsx"
            .Close
        End With
        Erase KQ
    Next
    Application.ScreenUpdating = True
    MsgBox "done " & Round((Timer - t), 2) & "secondes"
End Sub

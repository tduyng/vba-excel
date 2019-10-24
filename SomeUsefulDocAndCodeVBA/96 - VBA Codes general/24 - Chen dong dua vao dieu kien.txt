Option Explicit

Sub insertRowAndFillDownFomular()

Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Dim startDate As Date: startDate = ws.Range("C1").Value
Dim endDate As Date: endDate = ws.Range("D1").Value

Dim dateCounter As Date: dateCounter = startDate
Dim i As Long
Dim LR As Long
LR = ws.Range("A" & ws.Rows.Count).End(xlUp).Row

Dim dicMa As Object
Set dicMa = CreateObject("Scripting.Dictionary")
Dim xKey As Variant

For i = 4 To LR
    xKey = ws.Range("A" & i).Value2
    If Not dicMa.exists(xKey) Then
        dicMa.Add xKey, ""
    End If
Next i

i = 4
Dim k As Variant
For Each k In dicMa.Keys
    Do
        ws.Range("A" & i).Value2 = k
        ws.Range("B" & i).Value = dateCounter
        dateCounter = dateCounter + 1
        i = i + 1
    Loop Until dateCounter > endDate
    dateCounter = startDate
Next k

LR = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
ws.Range("C4", "D" & LR).FillDown

End Sub
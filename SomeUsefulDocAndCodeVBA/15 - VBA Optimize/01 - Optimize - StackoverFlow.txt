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
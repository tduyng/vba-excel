Option Explicit

Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

Sub OptimzeCode_Begin()
'=============================================================='
'Enbale mode Optimize
'=============================================================='

Application.ScreenUpdating = False

EventState = Application.EnableEvents
Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False

'=============================================================='

End Sub


Sub OptimizeCode_End()

'=============================================================='
'Return by defaut after the mode Optimize
'=============================================================='

    ActiveSheet.DisplayPageBreaks = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    
End Sub


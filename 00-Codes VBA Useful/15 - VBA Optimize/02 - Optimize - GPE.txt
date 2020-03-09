'To speedUp: Add this function to the first and to the end of your code
 
Public Function SpeedUp(Optional optSpeedUp As Boolean = True)
    With Application
        .Interactive = Not optSpeedUp
        .EnableEvents = Not optSpeedUp
        .ScreenUpdating = Not optSpeedUp
        .Calculation = IIf(optSpeedUp, xlCalculationManual, xlCalculationAutomatic)
 'Option:
'        .DisplayAlerts = Not optSpeedUp
'        .DisplayStatusBar = Not optSpeedUp
'        .EnableAnimations = Not optSpeedUp
    End With
End Function
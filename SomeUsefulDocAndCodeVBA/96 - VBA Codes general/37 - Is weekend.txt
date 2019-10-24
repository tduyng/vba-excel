Public Function IsWeekend(InputDate As Date) As Boolean
    Select Case Weekday(InputDate)
        Case vbSaturday, vbSunday
            IsWeekend = True
        Case Else
            IsWeekend = False
    End Select
End Function
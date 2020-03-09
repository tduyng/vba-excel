Sub test()

    f = 5

    On Error GoTo message

check:
    Do Until Cells(f, 1).Value = ""

        Cells.Find(what:=refnumber, After:=ActiveCell, LookIn:=xlFormulas, _
              lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
              MatchCase:=False, SearchFormat:=False).Activate
    Loop

    Exit Sub

message:
    MsgBox "There is an error"
    f = f + 1
    GoTo check

End Sub
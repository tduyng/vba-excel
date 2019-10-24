Function InRange(rng1, rng2) As Boolean
'   Returns True if rng1 is a subset of rng2
    InRange = False
    If rng1.Parent.Parent.Name = rng2.Parent.Parent.Name Then
        If rng1.Parent.Name = rng2.Parent.Name Then
            If Union(rng1, rng2).Address = rng2.Address Then
                InRange = True
            End If
        End If
    End If
End Function

Sub Test()
    Dim ValidRange As Range, UserRange As Range
    Dim SelectionOK As Boolean
  
    Set ValidRange = Range("A1:E20")
    SelectionOK = False
    On Error Resume Next

    Do Until SelectionOK = True
        Set UserRange = Application.InputBox(Prompt:="Select a range", Type:=8)
        If TypeName(UserRange) = "Empty" Then Exit Sub
        If InRange(UserRange, ValidRange) Then
            MsgBox "The range is valid."

            SelectionOK = True
        Else
            MsgBox "Select a range within " & ValidRange.Address
        End If
    Loop
End Sub
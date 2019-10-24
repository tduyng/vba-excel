
Private Sub cbxCountry_Change()
    MATCHING_CITY cbxCountry.ListIndex
End Sub

Sub MATCHING_CITY(li As Long)
    Dim city As Range
    Set city = Sheet2.Range("I14:K14").Offset(li, 0)
    With cbxCity
        .Clear
        .List = WorksheetFunction.Transpose(city)
        .ListIndex = 0
    End With
End Sub



Private Sub UserForm_Initialize()
    cbxCountry.List = Sheet2.Range("H14:H16").value
    cbxCountry.ListIndex = 0
    
    MATCHING_CITY cbxCountry.ListIndex
End Sub

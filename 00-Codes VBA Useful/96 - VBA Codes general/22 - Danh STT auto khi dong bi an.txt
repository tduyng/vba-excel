Sub AnDongKhiKhongCoDuLieu()
 Dim Rws As Long, Dem As Integer
 Dim WF As Object, Rng As Range, STT As Long
 
 Set WF = Application.WorksheetFunction
 Rws = [B2].CurrentRegion.Rows.Count
 Rows("1:" & Rws).Hidden = False
 Application.ScreenUpdating = False
 For J = 2 To Rws
    Set Rng = Range(Cells(J, "D"), Cells(J, "Y"))
    Dem = WF.CountA(Rng)
    If Dem Then
        STT = STT + 1:              Cells(J, "A").Value = STT
    Else
        Rows(J & ":" & J).Hidden = True
    End If
 Next J
 Application.ScreenUpdating = True
 MsgBox "Xong Rôi!", , "GPE.COM Xin Chào!"
End Sub
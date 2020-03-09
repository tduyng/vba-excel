Option Explicit
Dim i As Integer, iList2 As Integer, icount As Integer
Dim Dic As Object


Private Sub UserForm_Initialize()

    Set Dic = CreateObject("Scripting.dictionary")
        lbxList1.List = Array("Sales", "Production", "Logistics", "Human resources")
    obSelectType3.Value = True
    
End Sub


Private Sub cbAdd_Click()
    
    For i = 0 To lbxList2.ListCount - 1
        Dic.Add lbxList2.List(i), vbNullString
    Next i
    For i = 0 To lbxList1.ListCount - 1
        
        If lbxList1.Selected(i) And Not Dic.exists(lbxList1.List(i)) Then
            lbxList2.AddItem lbxList1.List(i)
        End If
    Next i
    
    Dic.RemoveAll
    For i = 0 To lbxList1.ListCount - 1
        lbxList1.Selected(i) = False
    Next i
    ckxList1.Value = False
End Sub

Private Sub lbxList1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    For i = 0 To lbxList2.ListCount - 1
        Dic.Add lbxList2.List(i), vbNullString
    Next i
    If Not Dic.exists(lbxList1.List(lbxList1.ListIndex)) Then
            lbxList2.AddItem lbxList1.List(lbxList1.ListIndex)
    End If
    Dic.RemoveAll
    For i = 0 To lbxList1.ListCount - 1
        lbxList1.Selected(i) = False
    Next i
    ckxList1.Value = False
End Sub

'Private Sub cbAdd_Click()
'
'    For i = 0 To lbxList2.ListCount - 1
'        Dic.Add lbxList2.List(i), vbNullString
'    Next i
'    For i = 0 To lbxList1.ListCount - 1
'        If lbxList1.Selected(i) And Not Dic.exists(lbxList1.List(i)) Then
'            lbxList2.AddItem addValueToListbox2(lbxList1.List(i))
'        End If
'    Next i
'    Dic.RemoveAll
'End Sub

Private Function addValueToListbox2(str As String) As String
    Dim valExists As Boolean
    valExists = False
    For i = 0 To lbxList2.ListCount - 1
        If lbxList2.List(i) = str Then valExists = True
    Next i
    If valExists Then
        MsgBox str & " has already added to the listbox"
        addValueToListbox2 = vbNullString
        Exit Function
    Else
        addValueToListbox2 = str
    End If
End Function


Private Sub cbRemove_Click()
    icount = 0
    For i = 0 To lbxList2.ListCount - 1
        If lbxList2.Selected(i - icount) Then
            lbxList2.RemoveItem (i - icount)
            icount = icount + 1
        End If
    Next i
    ckxList2.Value = False
End Sub
Private Sub lbxList2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    lbxList2.RemoveItem (lbxList2.ListIndex)
    ckxList2.Value = False
End Sub


Private Sub ckxList1_Change()
    Select Case ckxList1.Value
    Case True
        For i = 0 To lbxList1.ListCount - 1
            lbxList1.Selected(i) = True
        Next i
    Case False
        For i = 0 To lbxList1.ListCount - 1
                lbxList1.Selected(i) = False
        Next i
    End Select
End Sub



Private Sub ckxList2_Change()
    Select Case ckxList2.Value
    Case True
        For i = 0 To lbxList2.ListCount - 1
            lbxList2.Selected(i) = True
        Next i
    Case False
        For i = 0 To lbxList2.ListCount - 1
                lbxList2.Selected(i) = False
        Next i
    End Select

End Sub


Private Sub obSelectType3_Change()
    If obSelectType3.Value Then
        lbxList1.MultiSelect = fmMultiSelectExtended
        lbxList2.MultiSelect = fmMultiSelectExtended
    End If
End Sub

Private Sub obSelectType2_Change()
    If obSelectType2.Value Then
        lbxList1.MultiSelect = fmMultiSelectMulti
        lbxList2.MultiSelect = fmMultiSelectMulti
    End If
End Sub
Private Sub obSelectType1_Change()
    If obSelectType3.Value Then
        lbxList1.MultiSelect = fmMultiSelectSingle
        lbxList2.MultiSelect = fmMultiSelectSingle
    End If
End Sub

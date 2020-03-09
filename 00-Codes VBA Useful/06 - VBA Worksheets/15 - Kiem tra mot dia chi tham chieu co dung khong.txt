Option Explicit

Public Function IsValidRef(sRef As String) As Boolean
'-------------------------------------------------------------------------
' Procedure : IsValidRef Created by Jan Karel Pieterse
' Company   : JKP Application Development Services (c) 2005
' Author    : Jan Karel Pieterse
' Created   : 21-12-2005
' Purpose   : Checks of argument is a valid cell reference
'-------------------------------------------------------------------------
    Dim sTemp As String
    Dim oSh As Worksheet
    Dim oCell As Range
    '    On Error GoTo LocErr
    IsValidRef = False
    On Error Resume Next
    sTemp = Left(sRef, InStr(sRef, "!") - 1)
    sTemp = Replace(sTemp, "=", "")
    If Not IsIn(ActiveWorkbook.Worksheets, sTemp) Then
        IsValidRef = False
        Exit Function
    End If
    Set oSh = ActiveWorkbook.Worksheets(sTemp)
    If oSh Is Nothing Then
        Set oSh = ActiveWorkbook.Worksheets(Replace(sTemp, "'", ""))
    End If
    sTemp = Right(sRef, Len(sRef) - InStr(sRef, "!"))
    Set oCell = oSh.Range(sTemp)
    If oCell Is Nothing Then
        IsValidRef = False
    Else
        IsValidRef = True
    End If
End Function
Function IsIn(vCollection As Variant, ByVal sName As String) As Boolean
'-------------------------------------------------------------------------
' Procedure : funIsIn Created by Jan Karel Pieterse
' Company   : JKP Application Development Services (c) 2005
' Author    : Jan Karel Pieterse
' Created   : 28-12-2005
' Purpose   : Determines if object is in collection
'-------------------------------------------------------------------------
    Dim oObj As Object
    On Error Resume Next
    Set oObj = vCollection(sName)
    If oObj Is Nothing Then
        IsIn = False
    Else
        IsIn = True
    End If
    If IsIn = False Then
        sName = Replace(sName, "'", "")
        Set oObj = vCollection(sName)
        If oObj Is Nothing Then
            IsIn = False
        Else
            IsIn = True
        End If
    End If
End Function


Sub test1()
    Dim bTest As Boolean
    bTest = IsValidRef("hoa!A3")
    Debug.Print bTest
End Sub

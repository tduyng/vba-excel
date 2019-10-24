Option Explicit

Sub DELETE_CHARACTER_SPECIAL_1()
'   Using loop for each item in range data
    Dim sh As Worksheet
    Dim rng As Range, rngi As Range
    Dim RegEx As Object, RegMC As Object
    Dim strPattern As String, strValue As String
    Dim t As Double
    t = Timer
    Application.ScreenUpdating = False
    Set sh = ThisWorkbook.Sheets("Feuil2")
    Set rng = sh.Range("A1:A65536")
    Set RegEx = CreateObject("VBScript.RegExp")
    
    strPattern = "\W+" 'or = "[^a-zA-Z_0-9]" ' or = "[^0-9A-Za-z\r\n]+" ' because "[^\w\r\n]+" also matches _ and Unicode letters
        With RegEx
            .Pattern = strPattern
            .Global = True
            .ignorecase = False

        End With
    For Each rngi In rng
'                 If .test(rngi.Value) Then
'    '            RegMC = RegEx.Execute(rngi.Value)
        rngi.Offset(0, 1) = RegEx.Replace(rngi.Value, vbNullString)
'            Else
'                rngi.Offset(0, 1) = rngi.Value
'            End If
    Next rngi
    Application.ScreenUpdating = True
    MsgBox Round(Timer - t, 3) & "secondes"
    '90s - -> trop long
End Sub

Sub DELETE_CHARACTER_SPECIAL_2()
'   Using methode copy text from clipboard
    Dim sh As Worksheet
    Dim rng As Range, rngi As Range
    Dim RegEx As Object
    Dim strCopy As String
    Dim t As Double
    Application.ScreenUpdating = False
    t = Timer
   
    Set sh = ThisWorkbook.Sheets("Feuil2")
    Set rng = sh.Range("A1:A65536")
    Set RegEx = CreateObject("VBScript.RegExp")
    
    'Using the MS form data object
    'MSForms.DataObject ' needs reference to "Microsoft Forms 2.0 Object Library"
    'or using late binding: With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        rng.Copy
        .getFromClipboard
        Application.CutCopyMode = False
        strCopy = .GetText
        .Clear
        With RegEx
            .Pattern = "[^0-9A-Za-z\r\n]+" 'or = "[^a-zA-Z_0-9]" ' or = "[^0-9A-Za-z\r\n]+" ' because "[^\w\r\n]+" also matches _ and Unicode letters
            .Global = True
            .ignorecase = False
            strCopy = .Replace(strCopy, vbNullString)
        End With
        .SetText strCopy
        .PutinClipBoard
    End With
    sh.Range("C1").PasteSpecial
    Application.ScreenUpdating = True
    MsgBox Round(Timer - t, 3) & "secondes"
    '281 ms
End Sub
Sub DELETE_CHARACTER_SPECIAL_3()
'   Using loop + arr
    Dim sh As Worksheet
    Dim rng As Range
    Dim RegEx As Object
    Dim strPattern As String, strValue As String
    Dim arr As Variant
    Dim t As Double
    t = Timer
    Application.ScreenUpdating = False
    Set sh = ThisWorkbook.Sheets("Feuil2")
    Set rng = sh.Range("A1:A65536")
    arr = rng.Value
    Set RegEx = CreateObject("VBScript.RegExp")
    
    With RegEx
        .Pattern = "[^\w]+"
        .Global = True
        .ignorecase = False
    End With
        
    Dim i&
    For i = LBound(arr) To UBound(arr)
        arr(i, 1) = RegEx.Replace(arr(i, 1), vbNullString)
    Next i
    sh.Range("D1").Resize(UBound(arr)) = arr
    Application.ScreenUpdating = True
    MsgBox Round(Timer - t, 3) & "secondes"
    '375ms
End Sub

Sub test22()
[E1:E10] = ["ABC123-009"]
End Sub
Sub insert_car()
Range("A1:A65536").Formula = "=CHAR(RANDBETWEEN(65,90)) & CHAR(RANDBETWEEN(65,90)) & CHAR(RANDBETWEEN(65,90)) & CHAR(RANDBETWEEN(10,60)) & RANDBETWEEN(0,9) & RANDBETWEEN(0,9) & RANDBETWEEN(0,9)"
Range("A1:A65536") = Range("A1:A65536").Value
End Sub


Sub DELETE_CARACTER_SPECIAL_BY_InStr()
    'Methode INSTR
    Dim sh As Worksheet
    Dim rng As Range
    Dim arr As Variant
    Dim t As Double
    t = Timer
    Application.ScreenUpdating = False
    Set sh = ThisWorkbook.Sheets("Feuil2")
    Set rng = sh.Range("A1:A65536")

    arr = rng.Value

    Dim validValues As String
        validValues = "01234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" 'put numbers and capitals at the start as they are more likely'

    Dim i&, j&
    Dim result As String

        For i = LBound(arr) To UBound(arr)
            result = vbNullString
            For j = 1 To Len(arr(i, 1))
                If InStr(validValues, Mid(arr(i, 1), j, 1)) <> 0 Then
                    result = result & Mid(arr(i, 1), j, 1)
                End If
            Next j
        arr(i, 1) = result
        Next i

    sh.Range("E1").Resize(UBound(arr)) = arr
    Debug.Print Round(Timer - t, 3)
    '434 ms
End Sub

Sub Intervenant_Maitre_Ouvrage()
    'Déclaration
    Dim sh As Worksheet
    Dim picName As String, rowAffect As String
    Dim numberRowHide As Integer, countLine As Integer
    Dim rng As Range
    Dim nameShape As String, shp As Shape
    Set sh = ActiveSheet
    
    '==============================================================================================='
    'Donner numéro de ligne à afficher ou maquer
    
    numberRowHide = 12
    
    '==============================================================================================='
    




    On Error Resume Next
    countLine = 1000
    nameShape = Application.Caller
    
    Set shp = sh.Shapes(nameShape)
    
    shp.Rotation = 90
    
    
    countLine = shp.TopLeftCell.Row
    rowAffect = (countLine + 1) & ":" & (countLine + numberRowHide)
    
    Set rng = Rows(rowAffect)
    If rng.EntireRow.Hidden = True Then
        
        shp.Rotation = 180
        rng.EntireRow.Hidden = False
    Else:
        shp.Rotation = 90
        rng.EntireRow.Hidden = True
        
    End If

End Sub

Sub Interlocuteur2()
    'Déclaration
    Dim sh As Worksheet
    Dim picName As String, rowAffect As String
    Dim numberRowHide As Integer, countLine As Integer
    Dim rng As Range
    Dim nameShape As String, shp As Shape
    Set sh = ActiveSheet
    
    '==============================================================================================='
    'Donner numéro de ligne à afficher ou maquer
    
    numberRowHide = 5
    
    '==============================================================================================='
    

    On Error Resume Next
    countLine = 1000
    nameShape = Application.Caller
    
    Set shp = sh.Shapes(nameShape)
    
    shp.Rotation = 90
    
    
    countLine = shp.TopLeftCell.Row
    rowAffect = (countLine + 1) & ":" & (countLine + numberRowHide)
    
    Set rng = Rows(rowAffect)
    If rng.EntireRow.Hidden = True Then
        
        shp.Rotation = 180
        rng.EntireRow.Hidden = False
    Else:
        shp.Rotation = 90
        rng.EntireRow.Hidden = True
        
    End If

End Sub
Sub Intervenant_Autres()

    'Déclaration
    Dim sh As Worksheet
    Dim picName As String, rowAffect As String
    Dim numberRowHide As Integer, countLine As Integer
    Dim rng As Range
    Dim nameShape As String, shp As Shape
    Set sh = ActiveSheet
    
    '==============================================================================================='
    'Donner numéro de ligne à afficher ou maquer
    
    numberRowHide = 12
    
    '==============================================================================================='
    

    On Error Resume Next
    countLine = 1000
    nameShape = Application.Caller
    
    Set shp = sh.Shapes(nameShape)
    
    shp.Rotation = 90
    
    
    countLine = shp.TopLeftCell.Row
    rowAffect = (countLine + 1) & ":" & (countLine + numberRowHide)
    
    Set rng = Rows(rowAffect)
    If rng.EntireRow.Hidden = True Then
        
        shp.Rotation = 180
        rng.EntireRow.Hidden = False
    Else:
        shp.Rotation = 90
        rng.EntireRow.Hidden = True
        
    End If
    
End Sub

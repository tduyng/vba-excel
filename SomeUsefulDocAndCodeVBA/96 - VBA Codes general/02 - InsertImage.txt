Option Explicit


Sub Inserer_Logo()
Dim fNameAndPath As String, sh As Worksheet, rng As Range, rng1 As Range
Dim img As Object, zoomF As Integer
Dim shp As Shape, countLine As Integer, nameShape As String
Dim calcul As Double
Set sh = ActiveSheet
On Error Resume Next
Application.ScreenUpdating = False
nameShape = Application.Caller
Set shp = sh.Shapes(nameShape)
countLine = shp.TopLeftCell.Row

'========================================================================'
Set rng = sh.Range("B" & (countLine + 3))
Set rng1 = sh.Range("B1", "F1")
'zoomF = 100
'========================================================================'

rng.Select

fNameAndPath = Application.GetOpenFilename(Title:="Sélectioner une photo")
Set img = LoadPicture(fNameAndPath)
    calcul = img.Width / img.Height
    'Set img = sh.Pictures.Insert(fNameAndPath)
    Set shp = sh.Shapes.AddPicture(fNameAndPath, msoFalse, msoCTrue, rng.Left, rng.Top, rng1.Width, rng1.Width / calcul)
 
    Application.ScreenUpdating = True
End Sub


Sub Inserer_Capture_Critères_Selection()
Dim fNameAndPath As String, sh As Worksheet, rng As Range, rng1 As Range
Dim img As Object, zoomF As Integer
Dim shp As Shape, countLine As Integer, nameShape As String
Dim calcul As Double
Set sh = ActiveSheet
On Error Resume Next
Application.ScreenUpdating = False
nameShape = Application.Caller
Set shp = sh.Shapes(nameShape)
countLine = shp.TopLeftCell.Row
Debug.Print countLine
'========================================================================'
Set rng = sh.Range("C" & (countLine + 2))
Set rng1 = sh.Range("C1", "S1")
'zoomF = 100
'========================================================================'

rng.Select

fNameAndPath = Application.GetOpenFilename(Title:="Sélectioner une photo")
Set img = LoadPicture(fNameAndPath)
    calcul = img.Width / img.Height
    'Set img = sh.Pictures.Insert(fNameAndPath)
    Set shp = sh.Shapes.AddPicture(fNameAndPath, msoFalse, msoCTrue, rng.Left, rng.Top, rng1.Width, rng1.Width / calcul)
 
    Application.ScreenUpdating = True
End Sub


Sub Inserer_Capture_Documents_demandés()
Dim fNameAndPath As String, sh As Worksheet, rng As Range, rng1 As Range
Dim img As Object, zoomF As Integer
Dim shp As Shape, countLine As Integer, nameShape As String
Dim calcul As Double
Set sh = ActiveSheet
On Error Resume Next
Application.ScreenUpdating = False
nameShape = Application.Caller
Set shp = sh.Shapes(nameShape)
countLine = shp.TopLeftCell.Row

'========================================================================'
Set rng = sh.Range("W" & (countLine + 2))
Set rng1 = sh.Range("W1", "AM1")
'zoomF = 100
'========================================================================'

rng.Select

fNameAndPath = Application.GetOpenFilename(Title:="Sélectioner une photo")
Set img = LoadPicture(fNameAndPath)
    calcul = img.Width / img.Height
    'Set img = sh.Pictures.Insert(fNameAndPath)
    Set shp = sh.Shapes.AddPicture(fNameAndPath, msoFalse, msoCTrue, rng.Left, rng.Top, rng1.Width, rng1.Width / calcul)
 
    Application.ScreenUpdating = True
End Sub


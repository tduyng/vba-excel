Public Function LangueSysteme()
'par: https://excel-malin.com

On Error GoTo FunctionErreur

Dim CodeLangueSysteme As Single
'trouver le code de la langue
CodeLangueSysteme = Application.LanguageSettings.LanguageID(msoLanguageIDUI)

'associer le code à la langue (langues principales, les autres sont remplacées par l'anglais"
Select Case CodeLangueSysteme
Case 1036, 2060, 11276, 3084, 9228, 12300, 15372, 5132, 13324, 6156, 14348, 58380, 8204, 10252, 4108, 7180: LangueSysteme = "Français"
Case 1033, 2057, 3081, 10249, 4105, 9225, 15369, 16393, 14345, 6153, 8201, 17417, 5129, 13321, 18441, 7177, 11273, 12297: LangueSysteme = "Anglais"
Case 1031, 3079, 5127, 4103, 2055: LangueSysteme = "Allemand"
Case 2052, 4100, 1028, 3076, 5124: LangueSysteme = "Chinois"
Case 1043, 2067: LangueSysteme = "Néerlandais"
Case 1040, 2064: LangueSysteme = "Italien"
Case 3082, 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 22538, 2058, 19466, 6154, 15370, 10250, 20490, 21514, 14346, 8202: LangueSysteme = "Espagnol"
Case 1025, 5121, 15361, 3073, 2049, 11265, 13313, 12289, 4097, 6145, 8193, 16385, 10241, 7169, 14337, 9217: LangueSysteme = "Arabe"
Case Else: LangueSysteme = "Autre"
End Select

Exit Function

FunctionErreur:
LangueSysteme = ""
End Function
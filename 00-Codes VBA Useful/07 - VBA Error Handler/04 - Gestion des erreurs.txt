'gestion des erreurs
On Error Resume Next 'si il y une erreur passe à la ligne suivante
ActiveSheet.PivotTables("TDC_New").PivotFields("To to"). _
CurrentPage = "_INCONNU_"
If Err > 0 Then 'condition : si le code d'erreur est différent de 0 (donc si il y a eu plantage)
    MsgBox "ton message d'erreur." 'message (ou Exit Sub, éventuellement)
End If 'fin de la condition
On Error GoTo 0 'annule la gestion des erreurs
'... ton code si il n'y a pas plantage




On error resume next
    Sheets("Bridge").Range("W" & SumIfInt) = Sheets("Bridge").Range("AA" & 
    SumIfInt) / Sheets("Bridge").Range("D" & SumIfInt)

    if err <>0 then
       Sheets("Bridge").Range("W" & SumIfInt) = 0
    end if
    On error goto 0



If Not IsError("passage de la valeur sur le filtre") Then
    'Ton Code
Else
    Exit Sub 'Ou tout autre solution si le code plante
End If



Set QQCH = passage de la valeur sur le filtre
If Not QQCH Is Nothing Then
    'Ton Code
Else
    Exit Sub 'Ou tout autre solution si le code plante
End If




Dim Erreur As Byte
'S'il y a une erreur, alors on saute par-dessus la section. Sinon, tout va comme
'sur des roulettes
'On commence le code en disant qu'il y a une erreur
Erreur = 1
On Error GoTo Continuer


'*********************************************************************************

'S'il y a une erreur, alors ça ira à continuer sans remettre la variable Erreur à 0
ActiveSheet.PivotTables("TDC_New").PivotFields("to to").CurrentPage = "_INCONNU_"

With ActiveSheet.PivotTables("TDC_New").PivotFields("No m Client")
    .Orientation = xlRowField
    .Position = 2
End With

'Si tout c'est déroulé comme on voulait, alors on enlève l'erreur.
Erreur = 0

'*********************************************************************************

Continuer:
If Erreur = 1 Then
    'Code pour gérer l'erreur
End If





On Error GoTo Repmacro
ActiveSheet.PivotTables("TDC_New").PivotFields("soldto"). _
CurrentPage = "_INCONNU_"
With ActiveSheet.PivotTables("TDC_New").PivotFields("Nom Client")
.Orientation = xlRowField
.Position = 2
End With
.... tâches spécifiques

On Error GoTo 0
GoTo Repmacro


Repmacro:
... la suite des tâches


Si la valeur du filtre plante : Passage direct à Repmacro
Si la valeur du filtre ne plante pas, tâches spéciques puis passage à Repmacro

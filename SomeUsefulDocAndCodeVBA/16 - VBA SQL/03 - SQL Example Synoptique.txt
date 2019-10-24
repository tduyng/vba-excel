Sub FONCTION_MULTIPLE_SYNOPTIQUE(Optional typeFonction As String = vbNullString)

    '--------------------------------------------------------
    'Créer automatiquement synoptique
    'typeFonction = "MEFSynop" --> Pour la mise en forme
    'typeFonction = "CREATESynop" or = vbNullString --> Pour créer des tableaux synoptiques
    'typeFonction = "COLORSynop" --> Pour elever couleur de la colonne redondante
    '--------------------------------------------------------
    Dim sht_InfoComp As Worksheet, sht_Synoptique As Worksheet
    Dim iLigneDepInfoComp&, lstRowInfoComp&
    Dim rngSynop As Range, rngDataInfo As Range, rngModelSynop As Range
'    Dim colApp As String, colBatiment As String, colHall As String, colEtage As String, colNom As String
    Dim fieldApp As String, fieldBat As String, fieldEtage As String, fieldNom As String, fieldHall As String
    
    Dim cn As Object, rs As Object
    Dim arr(), arrEtage(), arrLog(), arrHall(), arrNbHall()
    Dim itemEtage As Variant, itemHall As Variant, maxColHall As Variant, itemLog As Variant
    Dim nbEtage&, nbHall&, i&, iEtage&, k&
    Dim filePath As String
    Dim sql As String, sqlEtage As String
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")

    ligneDepSyn = ThisWorkbook.Sheets("1 - PARAMETRES").Range("Ligne_Synop").value


   
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    'Actualiser Name range INFO_COMPLEMENT, INFO_SYNOPTIQUE, INFO_
    Call REFER_NAME_RANGE_INFO
            
    Set sht_InfoComp = ThisWorkbook.Sheets("4 - INFOS COMPLEMENT")
    Set sht_Synoptique = ThisWorkbook.Sheets("9 - SYNOPTIQUE")
    sht_Synoptique.Activate

    
    iLigneDepInfoComp = ThisWorkbook.Sheets("1 - PARAMETRES").Range("Ligne_Date")
    With sht_InfoComp
        
        lstRowInfoComp = .Range(colApp & .Rows.Count).End(xlUp).Row
        If lstRowInfoComp < iLigneDepInfoComp Then
            MsgBox "Pas de données du locatiaire dans la colonne N° Logement dans la feuille Info Complément." & vbNewLine & "Merci de les vérifier.", vbOKOnly + vbCritical, "Créer Synoptiques"
            Exit Sub
        End If
        Set rngDataInfo = .Range("INFO_SYNOP")
        
        fieldBat = .Range(colBatiment & iLigneDepInfoComp - 1).value
        fieldHall = .Range(colHall & iLigneDepInfoComp - 1).value
        fieldEtage = .Range(colEtage & iLigneDepInfoComp - 1).value
        fieldNom = .Range(colNom & iLigneDepInfoComp - 1).value
        fieldApp = .Range(colApp & iLigneDepInfoComp - 1).value
    End With
    Set rngModelSynop = ThisWorkbook.Sheets("1 - PARAMETRES").Range("Model_Synop")
    '****************************************************************************************
    'Part SQL
    filePath = ThisWorkbook.FullName
    'Get extension of file
    Dim Ext As String
    Ext = Right(filePath, Len(filePath) - InStrRev(filePath, "."))
    
    '---Connecting to the Data Source---
    'file excel nhu xlsx,xlsm,xlsb
    Set cn = CreateObject("ADODB.Connection")
    With cn
        If Ext = "xlsx" Or Ext = "xlsb" Or Ext = "xlsm" Then
            .provider = "Microsoft.ACE.OLEDB.12.0"
            .connectionstring = "Data Source=" & filePath & ";" & _
                                "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
        ElseIf Ext = "xls" Then
            .provider = "Microsoft.Jet.OLEDB.4.0"
            .connectionstring = "Data Source=" & filePath & ";" & _
                                "Extended Properties=""Excel 8.0;HDR=YES"";"
        End If
        On Error GoTo ERR_Handler
        .Open
    End With
    
'Test
'    sql = "SELECT COUNT (" & "[" & fieldApp & "]" & ")" _
'   & " FROM INFO_SYNOP WHERE " & "[" & fieldHall & "]" & " = ""2A"""
'   Set rs = cn.Execute(sql)
'   arrEtage = rs.getRows()
   
   
    'Lister tous les étages
    sql = "SELECT DISTINCT " & "([" & fieldEtage & "])" & " FROM INFO_SYNOP ORDER BY " & "[" & fieldEtage & "]" & " DESC"
    Set rs = cn.Execute(sql)
    arrEtage = rs.getRows()
    arrEtage = Application.Index(arrEtage, 0)
    nbEtage = UBound(arrEtage) - LBound(arrEtage) + 1
    
    'Lister tous les halls
    sql = "SELECT DISTINCT " & "([" & fieldHall & "])" & " FROM INFO_SYNOP " & " ORDER BY " & "[" & fieldHall & "]" & " ASC"
    Set rs = cn.Execute(sql)
    arrHall = rs.getRows()
    arrHall = Application.Index(arrHall, 0)
    i = 0
    For Each itemHall In arrHall
        If Not Dic.exists(itemHall) Then Dic.Add itemHall, i
        i = i + 1
    Next itemHall
'    nbHall = UBound(arrHall) - LBound(arrHall) + 1
   
'Calculer mal colonnes (nombre logement dans chaque Hall
'Cette nombre maximun est pour caculer la position de chaque model synoptique à poser
    i = 0
    For Each itemHall In arrHall
        If Not IsNumeric(itemHall) Then
            sql = "SELECT COUNT" & "([" & fieldApp & "])" & _
            " FROM INFO_SYNOP WHERE " & "[" & fieldHall & "]" & " = " & """" & itemHall & """" _
            & " GROUP BY " & "[" & fieldEtage & "]" _
            & " ORDER BY COUNT" & "([" & fieldApp & "])" & " DESC"
        Else
            sql = "SELECT COUNT" & "([" & fieldApp & "])" & _
            " FROM INFO_SYNOP WHERE " & "[" & fieldHall & "]" & " = " & itemHall _
            & " GROUP BY " & "[" & fieldEtage & "]" _
            & " ORDER BY COUNT" & "([" & fieldApp & "])" & " DESC"
        End If
            Set rs = cn.Execute(sql)
            arrNbHall = rs.getRows()
            maxColHall = arrNbHall(0, 0)
            
            'Dictionaire pour la list de hall et le maximum de logement dans chaque hall
            If (Dic.exists(itemHall) And (i > 0)) Then
                Dic.Item(itemHall) = 2 + Dic.Item(itemHall) * maxColHall * 6
            Else
                Dic.Item(itemHall) = 2
            End If
        i = i + 1
    Next itemHall
    arr = Dic.items

'----------------------------------------------------------------------------------------------------
    'Vérifier TypeFonction pour décider la fonction à faire: Soit Créer synoptique, soit Mise en form
    Select Case typeFonction
    Case vbNullString, "CREATESynop"

        
        'Supprimer toutes les valeurs dans la feuilles Synoptique
       Call VIDER_ShtSynop
        Dim iRun As Single, currentCount&
        Dim totalApp&
        '***********************************************'
        'La partie de céation des tableau synoptique
        'Lister tous les logements avec leur halle
        
        'Using this key for cancel loop with progress bar
        On Error GoTo ErrExeInterrup
        Application.EnableCancelKey = xlErrorHandler
    
        totalApp = lstRowInfoComp - iLigneDepInfoComp + 1
        iEtage = 0
        currentCount = 0
         For Each itemEtage In arrEtage
            'sql: string pour récupérer des logements et trier par batiment, puis par hall
            If Not IsNumeric(itemEtage) Then
                sql = "SELECT " & "[" & fieldApp & "]" & " FROM INFO_SYNOP WHERE " & "[" & fieldEtage & "]" & " = " & """" & itemEtage & """" _
                    & " ORDER BY " & "[" & fieldApp & "]" & " ASC , " & "[" & fieldHall & "]" & " ASC"
            Else
                sql = "SELECT " & "[" & fieldApp & "]" & ", [" & fieldHall & "]" & " FROM INFO_SYNOP WHERE " & "[" & fieldEtage & "]" & " = " & itemEtage _
                    & " ORDER BY " & "[" & fieldApp & "]" & " ASC , " & "[" & fieldHall & "]" & " ASC"
    '            sql = "SELECT " & "[" & fieldApp & "]" & ", [" & fieldHall & "]" & " FROM INFO_SYNOP WHERE " & "[" & fieldEtage & "]" & " = " & itemEtage & " GROUP BY " & "[" & fieldHall & "]"
            End If
    
            'Remplir nom de logement
            Set rs = cn.Execute(sql)
            'Array arr compose 2 list: list 1: num du logement, list 2
            arr = rs.getRows()
            arr = Application.Transpose(arr)
            arrLog = Application.Index(arr, , 1)
            arrHall = Application.Index(arr, , 2)
            
            'Terminer la partie search SQL, Maintenant on va remplir ces information pour construire des tableau synoptique
            'Avant de construire on va re-actualier Name rane INFO_COMP et INFO_GEN pour activer les formules vlookup et hlookup
            i = 0
            k = 0
            For Each itemLog In arrLog
                i = i + 1
                With sht_Synoptique
                    If Dic.exists(arrHall(i, 1)) Then
                        maxColHall = Dic.Item(arrHall(i, 1))
                        If .Cells((ligneDepSyn + iEtage * 14), maxColHall).value = "" Then
                            k = 0
    '                        .Cells((LigneDepSyn + iEtage * 14), maxColHall).Select
                            rngModelSynop.Copy .Cells((ligneDepSyn + iEtage * 14), maxColHall)
                            .Cells((ligneDepSyn + iEtage * 14), maxColHall).Offset(0, 1).value = itemLog
                        Else
                            k = k + 1
                            rngModelSynop.Copy .Cells((ligneDepSyn + iEtage * 14), maxColHall + (k * 6))
    '                        .Cells((LigneDepSyn + iEtage * 14), maxColHall + (k * 6)).Select
                            .Cells((ligneDepSyn + iEtage * 14), maxColHall + (k * 6)).Offset(0, 1).value = itemLog
                            
                            'Ajouter une fonction pour évier une boucle ilimité
                            'Hypothèse: max 200 logmements dans le même étage dans un Hall
                            If k = 100 Then
                                MsgBox "Il y a de problème de boucle par rapport à vos données. Le nombre de logement dans l'étage " & itemEtage _
                                & " - Hall " & Dic.Key(arrHall(i, 1)) & "est dépassé de 100 logements." _
                                & vbNewLine & "Veillez contacter le développeur pour résoudre ce problème ou  vérifier vos données.", vbCritical
                                
                                 With Application
                                    .ScreenUpdating = True
                                    .EnableEvents = True
                                    .Calculation = xlCalculationAutomatic
                                End With
                                Exit Sub
    
                            End If
                        End If
                    End If
                End With
                
                'Using progress bar
                currentCount = currentCount + 1
                iRun = currentCount / (totalApp)
                UF_PROGRESS_BAR_DOEVENTS iRun
            
            
            Next itemLog
            iEtage = iEtage + 1
        Next itemEtage
        
    Case "MEFSynop"
        Call FONCTION_MEF_Synoptique(arr)
    Case "COLORSynop"
        Call FONCTION_ENLEVER_COULEUR_DERNIERE_COLONNONE(sht_Synoptique, nbEtage, arr)
    End Select

   
'Fin SQL
    cn.Close
    Set cn = Nothing
    Set rs = Nothing
    Erase arrEtage
    Erase arrHall
    Erase arrNbHall
    Set Dic = Nothing
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    
END_Handler:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic

    End With
    Exit Sub
    
ERR_Handler:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    If Err.Number = 3704 Then Exit Sub
    If Err.Number = 91 Then Exit Sub
    MsgBox "Error " & Err.Number & "," & Err.Description
    Set rs = Nothing
    Resume END_Handler:
    
ErrExeInterrup:
    If Err.Number = 18 Then
        If MsgBox("Êtes-vous sûr de vouloir mettre fin à l'opération?", vbQuestion + vbYesNo, "Exécution interrompue") = vbNo Then
            Resume
        Else
            With Application
                .ScreenUpdating = True
                .EnableEvents = True
                .Calculation = xlCalculationAutomatic
            End With
            Unload ufProgress
        End If
    End If

End Sub
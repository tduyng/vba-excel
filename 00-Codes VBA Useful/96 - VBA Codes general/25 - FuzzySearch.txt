Option Explicit

Private Type PersonQuery
    firstName As String
    lastName As String
    fullName As String
    firstNames As Range
    lastNames As Range
    fullNames As Range
    namesSwapped As Boolean
End Type

Private Type PersonResult
    found As Boolean
    row As Integer
    lookupMethod As String
End Type

Const INTEGER_MIN As Integer = -32768
Private NOT_FOUND As PersonResult

Const EDIT_DIST_NOMATCH_THRESHOLD As Double = 0.5
Const LEN_RATIO_NOMATCH_THRESHOLD As Double = 0.25
Const STR_LOC_NOMATCH_THRESHOLD As Double = 0.25
Const UNORDERED_SIM_NOMATCH_THRESHOLD As Double = 0.3

Const ASC_A As Integer = 65
Const NUM_LETTERS As Integer = 26
Const NUM_LETTERS_AND_SPECIALS As Integer = 27
Const NON_LETTER_VAL As Integer = 26

' Searches for a person based on first name and last name
' Returns an array of (person's id, full name in database, method of matching name)
' If the person is not found, id and full name will be "N/A"
' The methods used are:
'  - search for exact full name,
'  - reverse first and last and search for that full name (sometimes first and last get confused)
'  - filter by first name search on similarity for last name (and vice versa and also reversing first/last)
'  - then search based on similarity for the full name and reversing first/last

Private Function FindPersonID(firstName As String, lastName As String, firstNameList As Range, _
    lastNameList As Range, fullNameList As Range, ids As Range)
    
    Dim pr As PersonResult
    Dim pq As PersonQuery
    pq.firstName = firstName
    pq.lastName = lastName
    pq.fullName = firstName + " " + lastName
    Set pq.firstNames = firstNameList
    Set pq.lastNames = lastNameList
    Set pq.fullNames = fullNameList
    pq.namesSwapped = False
    
    pr = LookupPersonInternal(pq)
    
    Dim fullName As String
    Dim id As String
    If pr.found Then
        id = ids.Cells(pr.row, 1).Value
        fullName = fullNameList.Cells(pr.row, 1).Value
    Else
        id = "N/A"
        fullName = "N/A"
    End If
    
    FindPersonID = Array(id, fullName, pr.lookupMethod)
End Function

Private Function LookupPersonInternal(pq As PersonQuery) As PersonResult
    Dim pr As PersonResult
    
    If Trim(pq.fullName) = "" Then
        GoTo NotFound
    End If
    
    pr = LookupExact(pq)
    If pr.found Then
        pr.lookupMethod = "EXACT_FULL"
        LookupPersonInternal = pr
        Exit Function
    End If
    
    pr = LookupExact(SwapFirstAndLast(pq))
    If pr.found Then
        pr.lookupMethod = "EXACT_FULL_REVERSED"
        LookupPersonInternal = pr
        Exit Function
    End If
    
    pr = LookupLastExactFirstSimilar(pq)
    If pr.found Then
        pr.lookupMethod = "LAST_EXACT_FIRST_SIMILAR"
        LookupPersonInternal = pr
        Exit Function
    End If
    
    pr = LookupLastExactFirstSimilar(SwapFirstAndLast(SwapFirstAndLastLists(pq)))
    If pr.found Then
        pr.lookupMethod = "FIRST_EXACT_LAST_SIMILAR"
        LookupPersonInternal = pr
        Exit Function
    End If
    
    pr = LookupLastExactFirstSimilar(SwapFirstAndLast(pq))
    If pr.found Then
        pr.lookupMethod = "LAST_EXACT_REVERSED_FIRST_SIMILAR_REVERSED"
        LookupPersonInternal = pr
        Exit Function
    End If
    
    pr = LookupLastExactFirstSimilar(SwapFirstAndLastLists(pq))
    If pr.found Then
        pr.lookupMethod = "FIRST_EXACT_REVERSED_LAST_SIMILAR_REVERSED"
        LookupPersonInternal = pr
        Exit Function
    End If
            
    pr = LookupFullSimilar(pq)
    If pr.found Then
        pr.lookupMethod = "SIMILAR_FULL"
        LookupPersonInternal = pr
        Exit Function
    End If
    
    pr = LookupFullSimilar(SwapFirstAndLast(pq))
    If pr.found Then
        pr.lookupMethod = "SIMILAR_FULL_REVERSED"
        LookupPersonInternal = pr
        Exit Function
    End If
    
NotFound:
    pr.lookupMethod = "NOT_FOUND"
    LookupPersonInternal = pr
End Function

Private Function LookupExact(pq As PersonQuery) As PersonResult
    Dim foundRange As Range
    Dim pr As PersonResult
    Set foundRange = pq.fullNames.Find(pq.fullName, LookIn:=xlValues, LookAt:=xlWhole)
    pr.found = Not foundRange Is Nothing
    
    If pr.found Then
        pr.row = foundRange.row
    End If
    
    LookupExact = pr
End Function

Private Function LookupLastExactFirstSimilar(pq As PersonQuery) As PersonResult
    Dim pr As PersonResult
    pr.found = False
    
    Dim lastNameRows As Collection
    Set lastNameRows = FindRowsForKey(pq.lastName, pq.lastNames)
    If lastNameRows.Count > 0 Then
        Dim firstNamesForLast As Collection
        Set firstNamesForLast = LookupValuesForRows(lastNameRows, pq.firstNames)
        
        Dim mostSimilar As Variant
        mostSimilar = MostSimilarIndexInCollection(pq.firstName, firstNamesForLast)
        
        If CDbl(mostSimilar(1)) / CDbl(Len(pq.firstName)) > EDIT_DIST_NOMATCH_THRESHOLD Then
            pr.found = True
            pr.row = lastNameRows(mostSimilar(0))
        End If
    End If
    
    LookupLastExactFirstSimilar = pr
End Function

Private Function LookupFullSimilar(pq As PersonQuery) As PersonResult
    Dim pr As PersonResult
    
    Dim rowAndSimilarity As Variant
    rowAndSimilarity = SimilarMatch(pq.fullName, pq.fullNames)
    
    pr.row = rowAndSimilarity(0)
    pr.found = (rowAndSimilarity(1) / CDbl(Len(pq.fullName))) > EDIT_DIST_NOMATCH_THRESHOLD
    
    LookupFullSimilar = pr
End Function

Private Function SwapFirstAndLast(pq As PersonQuery) As PersonQuery
    Dim pqSwapped As PersonQuery
    pqSwapped = pq
    pqSwapped.firstName = pq.lastName
    pqSwapped.lastName = pq.firstName
    pqSwapped.fullName = pqSwapped.firstName + " " + pqSwapped.lastName
    pqSwapped.namesSwapped = Not pq.namesSwapped
    SwapFirstAndLast = pqSwapped
End Function

Private Function SwapFirstAndLastLists(pq As PersonQuery) As PersonQuery
    Dim pqListsSwapped As PersonQuery
    pqListsSwapped = pq
    Set pqListsSwapped.firstNames = pq.lastNames
    Set pqListsSwapped.lastNames = pq.firstNames
    pqListsSwapped.namesSwapped = Not pq.namesSwapped
    SwapFirstAndLastLists = pqListsSwapped
End Function

Function MostSimilarIndexInCollection(str As String, col As Collection)
    Dim bestSimilarity As Variant
    Dim mostSimilarI As Integer
    Dim curSimilarity As Integer
    bestSimilarity = INTEGER_MIN
    mostSimilarI = -1
    
    Dim i As Integer
    For i = 1 To col.Count
        curSimilarity = LevenshteinSimilarity(str, col.Item(i))
        If curSimilarity > bestSimilarity Then
            mostSimilarI = i
            bestSimilarity = curSimilarity
        End If
    Next
    
    MostSimilarIndexInCollection = Array(mostSimilarI, bestSimilarity)
End Function

Private Function LookupValuesForRows(rows As Collection, rng As Range) As Collection
    Dim i As Integer
    Dim vals As New Collection
    For i = 1 To rows.Count
        vals.Add (rng.Cells(rows(i), 1).Value)
    Next
    Set LookupValuesForRows = vals
End Function

Private Function FindRowsForKey(key As Variant, keys As Range) As Collection
    Dim rows As New Collection
    
    Dim foundRange As Range
    Set foundRange = keys.Find(key, LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim firstAddress As Variant
    If Not foundRange Is Nothing Then
        firstAddress = foundRange.Address
        Do
            rows.Add (foundRange.row)
            Set foundRange = keys.Find(key, foundRange, LookIn:=xlValues, LookAt:=xlWhole)
        Loop While Not foundRange Is Nothing And foundRange.Address <> firstAddress
    End If
    
    Set FindRowsForKey = rows
End Function

Function SimilarMatch(search As Variant, list As Range) As Variant
    Dim maxSim As Integer
    maxSim = 0
    
    Dim i As Integer
    Dim maxSimI As Integer
    maxSimI = -1
    
    Dim lastRow As Integer
    lastRow = list.Cells.End(xlDown).row
    
    Dim firstRow As Integer
    firstRow = list.Cells.End(xlUp).row
    
    Dim searchStr As String
    
    If TypeOf search Is Range Then
        searchStr = search.Value
    Else
        searchStr = search
    End If
    
    Dim valInList As String
    Dim sim As Integer
    For i = firstRow To lastRow
        valInList = list.Cells(i, 1).Value
           
        'For performance sake check the length, space location and unordered similarity first
        'before running the edit distance algorithm
        If Abs(Len(valInList) - Len(searchStr)) / Len(searchStr) _
              < LEN_RATIO_NOMATCH_THRESHOLD Then
            If Abs(InStr(searchStr, " ") - InStr(valInList, " ")) / Len(searchStr) _
                  < STR_LOC_NOMATCH_THRESHOLD Then
                If UnorderedSimilarity(searchStr, valInList) / Len(searchStr) _
                      > UNORDERED_SIM_NOMATCH_THRESHOLD Then
                    sim = LevenshteinSimilarity(searchStr, valInList)
                    If sim > maxSim Then
                        maxSim = sim
                        maxSimI = i
                    End If
                End If
            End If
        End If
    Next
        
    SimilarMatch = Array(maxSimI, maxSim)
End Function

Private Function LetterValue(letter As String) As Integer
    Dim ascVal As Integer
    ascVal = Asc(letter)
    
    If ascVal < ASC_A Or ascVal > ASC_A + NUM_LETTERS Then
        LetterValue = NON_LETTER_VAL
    Else
        LetterValue = ascVal - ASC_A
    End If
End Function

Private Function UnorderedSimilarity(ByVal str1 As String, ByVal str2 As String) As Integer
    Dim i, j, ascVal As Integer
    Dim ascA As Integer
    Dim ascZ As Integer
    Dim letterCountDiffs(27) As Integer
    Dim lettersDiff As Integer
    Dim letterVal As Integer
    Dim len1 As Integer
    Dim len2 As Integer
    
    len1 = Len(str1)
    len2 = Len(str2)
    str1 = UCase(str1)
    str2 = UCase(str2)
    
    i = 0
    While i < len1
        letterVal = LetterValue(CharAt(str1, i))
        letterCountDiffs(letterVal) = letterCountDiffs(letterVal) + 1
        i = i + 1
    Wend
    
    i = 0
    While i < len2
        letterVal = LetterValue(CharAt(str2, i))
        letterCountDiffs(letterVal) = letterCountDiffs(letterVal) - 1
        i = i + 1
    Wend
    
    lettersDiff = 0
    For i = 0 To UBound(letterCountDiffs)
        lettersDiff = lettersDiff + Abs(letterCountDiffs(i))
    Next
    
    Dim minLen As Integer
    If len1 < len2 Then minLen = len1 Else minLen = len2
        
    UnorderedSimilarity = minLen - lettersDiff
End Function

'--------------------------------------------------------------------
' Calculates the edit distance between str1 and str2 using the
' Levenshtein distance dynamic programming algorithm
' This is really "edit similarity" as more similar strings have a
' larger score.
Private Function LevenshteinSimilarity(str1 As String, str2 As String)
    Dim len1, len2, i, j, score, charSim, gap1, gap2, matchVal As Integer
    
    len1 = Len(str1)
    len2 = Len(str2)
    
    Dim gap_score As Integer
    gap_score = -1

    Dim D As Variant
    ReDim D(0 To (len1 + 1), 0 To (len2 + 1)) As Integer
    D(0, 0) = 0
 
    For i = 0 To len1
        D(i, 0) = gap_score * i
    Next
    
    For j = 0 To len2
        D(0, j) = gap_score * j
    Next
    
    For i = 1 To len1
        For j = 1 To len2
            matchVal = D(i - 1, j - 1) + CharSimilarity(CharAt(str1, i - 1), CharAt(str2, j - 1))
            gap2 = D(i, j - 1) + gap_score
            gap1 = D(i - 1, j) + gap_score
            D(i, j) = Application.WorksheetFunction.Max(matchVal, gap2, gap1)
        Next
    Next
    
    'Dim alignment As String
    'alignment = ""
    
    i = len1
    j = len2
    score = 0
    
    Dim align As String
    
    While i > 0 And j > 0
        charSim = CharSimilarity(CharAt(str1, i - 1), CharAt(str2, j - 1))
        If D(i, j) - charSim = D(i - 1, j - 1) Then
            If charSim > 0 Then
                'align = "M"
            Else
                'align = "C"
            End If
            
            i = i - 1
            j = j - 1
            
            score = score + charSim
        ElseIf D(i, j) - gap_score = D(i, j - 1) Then
            'align = "A"
            j = j - 1
        ElseIf D(i, j) - gap_score = D(i - 1, j) Then
            'align = "D"
            i = i - 1
            score = score + gap_score
        Else
            MsgBox "Unexpected score in backtracking"
        End If
        'alignment = align + alignment
    Wend
    
    While j > 0
        'alignment = "A" + alignment
        j = j - 1
        score = score + gap_score
    Wend
    
    While i > 0
        'alignment = "D" + alignment
        i = i - 1
        score = score + gap_score
    Wend

    LevenshteinSimilarity = score
End Function

Function CharAt(str, zeroStartingIndex)
    CharAt = Mid(str, zeroStartingIndex + 1, 1)
End Function

Function CharSimilarity(chr1 As Variant, chr2 As Variant)
    If UCase(chr1) = UCase(chr2) Then
        CharSimilarity = 1
    Else
        CharSimilarity = -1
    End If
End Function
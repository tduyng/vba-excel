CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

ex:
With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        rng.Copy
        .getFromClipboard
        Application.CutCopyMode = False
        strCopy = .GetText
        .Clear
    strPattern = "[^0-9A-Za-z\r\n]" 'or = "[^a-zA-Z_0-9]" ' or = "[^0-9A-Za-z\r\n]+" ' because "[^\w\r\n]+" also matches _ and Unicode letters
        With RegEx
            .Pattern = strPattern
            .Global = True
            strCopy = .Replace(strCopy, vbNullString)
        End With
        .SetText strCopy
        .PutinClipBoard
    End With
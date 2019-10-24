Sub ListComments()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cmt As Comment
    Dim cmtCount As Long
    cmtCount = 2
    On Error Resume Next
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(xlWorksheet)
    With wb.Sheets(1)
        .Range("$A$1") = "Author"
        .Range("$B$1") = "Book"
        .Range("$C$1") = "Sheet"
        .Range("$D$1") = "Range"
        .Range("$E$1") = "Comment"
    End With
    For Each cmt In ws.Comments
        With wb.Sheets(1)
            .Cells(cmtCount, 1) = cmt.author
            .Cells(cmtCount, 2) = cmt.Parent.Parent.Parent.Name
            .Cells(cmtCount, 3) = cmt.Parent.Parent.Name
            .Cells(cmtCount, 4) = cmt.Parent.Address
            .Cells(cmtCount, 5) = CleanComment(cmt.author, cmt.Text)
        End With
        cmtCount = cmtCount + 1
    Next
    wb.Sheets(1).UsedRange.WrapText = False
    Application.ScreenUpdating = True
    Set ws = Nothing
    Set wb = Nothing
End Sub
Private Function CleanComment(author As String, cmt As String) As String
    Dim tmp As String
    tmp = Application.WorksheetFunction.Substitute(cmt, author & ":", "")
    tmp = Application.WorksheetFunction.Substitute(tmp, Chr(10), "")
    CleanComment = tmp
End Function
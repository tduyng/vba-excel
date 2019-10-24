Function ChangeSeed(strTbl As String, strCol As String, lngSeed As Long) As Boolean
'strTbl = Ten Table co chua truong kieu autonumber
'strCol = Ten truong co chua autonumber
'lngSeed = Con so ma ban muon no hien thi cho dong ke tiep.

Dim cnn As ADODB.Connection
Dim cat As New ADOX.Catalog
Dim col As ADOX.Column

Set cnn = CurrentProject.Connection
cat.ActiveConnection = cnn

Set col = cat.Tables(strTbl).Columns(strCol)

col.Properties("Seed") = lngSeed
cat.Tables(strTbl).Columns.Refresh
If col.Properties("seed") = lngSeed Then
    ChangeSeed = True
Else
    ChangeSeed = False
End If
Set col = Nothing
Set cat = Nothing
Set cnn = Nothing

End Function
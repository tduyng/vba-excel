Option Explicit
 
Public Const dbName = "student.accdb"
 
Sub ADODB_Connect()
    Dim dbPath As String
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim strConn As String
    Dim query As String
     
    On Error GoTo ErrorProcess
     
    ' create dpPath from current folder and dbName
    dbPath = Application.ActiveWorkbook.Path & "\" & dbName
    ' information to connect to 2007/2010 AccessDB
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & dbPath & ";" & _
        "User Id=admin;Password="       
    ' open connection
    conn.Open (strConn)    
    ' define query
    query = "SELECT * FROM student"  
    ' execute the query
    rs.Open query, conn, adOpenKeyset
    ' show number of records
    MsgBox rs.RecordCount    
    ' show data from AccessDB
    Do Until rs.EOF
        MsgBox rs.Fields.Item("name") & ", " & rs.Fields.Item("age")
        rs.MoveNext
    Loop
     
    GoTo EndSub
ErrorProcess:
    MsgBox Err.Number & ": " & Err.Description
EndSub:
    Set rs = Nothing
    Set conn = Nothing
End Sub

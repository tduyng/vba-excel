
Sub SQL_01(sourceSQL As String)
    Dim orderConn As Object, orderData As Object, orderField As Object
    Set orderConn = CreateObject("ADODB.Connection")
    Set orderData = CreateObject("ADODB.Recordset")
    Dim i As Integer
    
    On Error GoTo close_Connection
    With orderConn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
        .Open
    End With
    
    With orderData
        .activeConnection = orderConn
        .Source = sourceSQL
        .LockType = 1 'adLockReadyOnly
        .CursorType = 0 'adForwardOnly
        .Open
    End With
    i = 0
    
    On Error GoTo close_RecordSet
    Worksheets.Add after:=Worksheets(Worksheets.Count)
    For Each orderField In orderData.Fields
        i = i + 1
        Cells(1, i) = orderField.Name
    Next orderField
    Range("A2").CopyFromRecordset orderData
    
    'Dat ten lai cot
    'Select [Order Date] as [Date],Item as [test-Item],region as [Departement] from SaleOrders
    'Them 1 cot co gia tri giong nhau: Select [Order Date],'SQL That tuyet' from SalesOrders
    'Select [Unit Cost] as [Cost],Region as [Departement],'SQL That Tuyet' as [SQL That hehe] from SalesOrders
    'Co the them cot tinh toan giua tren gia tri da co:
    'Select Total,(Total*0.1) as [tax 10%] from SalesOrders
    'Select sum(Total) as Sum Total from SalesOrders
    'Select Sum (Units) as [Sum Units],Sum(Total) as [Sum Total], (Sum (Total)/Sum(Units)) as [Revenu per Unit] from SalesOrders
    'Select * from SalesOrders where [Unit Cost] = 1.99
    'Select * from SalesOrders where [Unit Cost] <= 50
    'Select * from SalesOrders where [orderdate] >= #25-06-2016#
    'Select * from SalesOrders where [rep] IN("morgan","parent")
    'Select rep, item,Total as [Tong doanh thu] from SalesOrders order by Total ASC
    'Select rep, item, sum(Total) as [Tong doanh thu] from SalesOrders group by rep,item
    
    'lay nam cua gia tri trong cot
'    Select year(OrderDate) as [Year], month(OrderDate) as [Month], sum(Total) as [Doanh thu tong] from SalesOrders group by year(OrderDate),month(OrderDate)
    
close_RecordSet:
    If Err.Number <> 0 Then MsgBox Err.Description
    orderData.Close
    
close_Connection:
    If Err.Number <> 0 Then MsgBox Err.Description
    orderConn.Close
    
    
    Set orderConn = Nothing
    Set orderData = Nothing
End Sub
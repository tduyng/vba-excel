Sub APPLY_FILTER_DATA()
    Set Sh = Sheet3
    Set rng = Sheet3.Range("Data_filter")
    
    'Filter 1 cot
    '
    '************************************************************************************************************
    'rng.AutoFilter field:=2, Criteria1:=Sheet3.Range("D9").Value, Criteria2:=Sheet3.Range("D10"), Operator:=xlOr
    '
    '************************************************************************************************************
    
    
    'Filter nhieu cot
    '
    '************************************************************************************************************
    'rng.AutoFilter field:=2, Criteria1:=Sheet3.Range("D9").Value
    'rng.AutoFilter field:=4, Criteria1:=Sheet3.Range("D10").Value
    'rng.AutoFilter field:=3, Criteria1:=Sheet3.Range("D10").Value & "*"
    '
    '************************************************************************************************************
    
    'Filter su dung mang
    rng.AutoFilter field:=2, Criteria1:=Array("West", "East"), Operator:=xlFilterValues
    
End Sub

Sub RESET_AUTOFILTER()
    Sheet3.Range("Data_filter").AutoFilter
End Sub

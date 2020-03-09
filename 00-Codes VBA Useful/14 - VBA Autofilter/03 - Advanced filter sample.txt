
'Advanced filter truc tiep trong bang du lieu
Application.CutCopyMode = False
    Range("A1:F331").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Range("I1:N2"), Unique:=False
    ActiveSheet.ShowAllData


'Advanced filter roi copy ra ngoai mot vi tri khac ngoai bang du lieu
Range("dataFilter").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:= _
        Range("I1:N2"), CopyToRange:=Range("I10"), Unique:=False
    Range("N2").Select


Option Explicit

Sub CollectionsVSdictionaries() ' Requires ref to "Microsoft Scripting Runtime" Library
    Dim c As Collection         ' TypeName 1-based indexed
    Dim d As Dictionary         ' 0-based arrays
    Set c = New Collection      ' or: "Dim c As New Collection"
    Set d = New Dictionary      ' or: "Dim d As New Dictionary"

    c.Add Key:="A", Item:="AA": c.Add Key:="B", Item:="BB": c.Add Key:="C", Item:="CC"
    d.Add Key:="A", Item:="AA": d.Add Key:="B", Item:="BB": d.Add Key:="C", Item:="CC"

    Debug.Print TypeName(c)    ' -> "Collection"
    Debug.Print TypeName(d)    ' -> "Dictionary"

    Debug.Print c(3)            ' -> "CC"
    Debug.Print c("C")          ' -> "CC"
    'Debug.Print c("CC")       ' --- Invalid ---

    Debug.Print d("C")          ' -> "CC"
    Debug.Print d("CC")        ' Adds Key:="CC", Item:=""
    Debug.Print d.Items(2)      ' -> "CC"
    Debug.Print d.Keys(2)       ' -> "C"
    Debug.Print d.Keys()(0)     ' -> "A"    - Not well known ***************************
    Debug.Print d.Items()(0)    ' -> "AA"   - Not well known ***************************

    'Collection methods:
    '    .Add                   ' c.Add Item, [Key], [Before], [After] (Key is optional)
    '    .Count
    '    .Item(Index)           ' Default property;   "c.Item(Index)" same as "c(Index)"
    '    .Remove(Index)
    'Dictionary methods:
    '    .Add                   ' d.Add Key, Item (Key is required, and must be unique)
    '    .CompareMode           ' 1. BinaryCompare     - case-sensitive   ("A" < "a")
    '    .CompareMode           ' 2. DatabaseCompare   - MS Access only
    '    .CompareMode           ' 3. TextCompare       - case-insensitive ("A" = "a")
    '    .Count
    '    .Exists(Key)           ' Boolean **********************************************
    '    .Item(Key)
    '    .Items                 ' Returns full array: .Items(0)(0)
    '    .Key(Key)
    '    .Keys                  ' Returns full array: .Keys(0)(0)
    '    .Remove(Key)
    '    .RemoveAll             ' ******************************************************
End Sub
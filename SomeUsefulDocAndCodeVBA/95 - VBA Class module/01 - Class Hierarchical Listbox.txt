Private Sub m_BuildNodeList()
'
' Fill listbox with valid nodes
'
    Dim intNodeIndex As Integer
    Dim intItemIndex As Integer
    Dim strStatus As String
    Dim strBuf As String
    
    m_objListbox.Clear
    For intNodeIndex = 0 To m_intNNodes - 1
        If m_intNodeLevel(intNodeIndex) <= 0 Then
            m_objListbox.AddItem m_MakeNode(intNodeIndex)
            intItemIndex = m_objListbox.ListCount - 1
            m_objListbox.List(intItemIndex, 1) = CStr(intNodeIndex)
        End If
    Next
End Sub

Public Property Get DisplayListCount() As Integer
' number of items currenlt showing
    DisplayListCount = m_objListbox.ListCount - 1
End Property
Public Property Let NodeClosedText(Text As String)
' user set Close Node text
    m_strNode_Closed = Text
    
    m_BuildDropNode
End Property
Public Property Get NodeClosedText() As String
    NodeClosedText = m_strNode_Closed
End Property
Public Property Get NodeOpenText() As String
    NodeOpenText = m_strNode_Open
End Property
Public Property Get NodeEndText() As String
    NodeEndText = m_strNode_End
End Property
Public Property Get NodeDropText() As String
    NodeDropText = m_strNode_Drop
End Property
Public Property Let NodeOpenText(Text As String)
' user set Open Node text
    m_strNode_Open = Text
    m_BuildDropNode
End Property
Public Property Let NodeDropText(Text As String)
' user set Drop Node text
    m_strNode_Drop = Text
    
    m_BuildDropNode
End Property
Public Property Let NodeEndText(Text As String)
' user set End Node text
    m_strNode_End = Text
    m_BuildDropNode

End Property
Public Property Get NodeCount() As Integer
' count of all items
    NodeCount = m_intNNodes - 2
End Property
Public Sub Refresh()
    m_CreateNodes
    m_BuildNodeList
End Sub

Public Property Set UseListbox(LBox As MSForms.ListBox)
    Set m_objListbox = LBox
End Property
Public Property Get Version() As String
    Version = HLIST_VERSION
End Property

Private Sub Class_Initialize()

    Dim intLen As Integer
    
    m_strNode_Closed = "..+."
    m_strNode_Open = "..-."
    m_strNode_End = "...."
    m_strNode_Drop = "!"

    m_BuildDropNode
    
End Sub

Private Sub Class_Terminate()
    
    Set m_objListbox = Nothing
    
End Sub


Private Sub m_objListbox_Click()

    Dim intIndex As Integer
    Dim intNodeIndex As Integer
    
    intIndex = m_objListbox.ListIndex
    intNodeIndex = CInt(m_objListbox.List(intIndex, 1))
    
    RaiseEvent ClickedNode(intNodeIndex, intIndex)
    
End Sub
Private Sub m_objListbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim intRowIndex As Integer
    Dim intNodeIndex As Integer
    
    intRowIndex = m_objListbox.ListIndex
    intNodeIndex = CInt(m_objListbox.List(intRowIndex, 1))
    
    m_UpdateNodes intNodeIndex, intRowIndex
    RaiseEvent DblClickedNode(intNodeIndex, intRowIndex)
    
End Sub
Private Sub m_objListbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Dim intRowIndex As Integer
    Dim intNodeIndex As Integer
    
    If KeyCode = vbKeyReturn Then
        intRowIndex = m_objListbox.ListIndex
        intNodeIndex = CInt(m_objListbox.List(intRowIndex, 1))
        m_UpdateNodes intNodeIndex, intRowIndex
    End If
    
End Sub

'Ð? s? d?ng class trên. B?n hãy thêm vào m?t form, trên form thêm vào m?t listbox có tên ListBox1 vào b?n hãy thêm vào do?n code này vào module c?a form dó nhu sau:





Private Sub UserForm_Initialize()
    
    ListBox1.Move 0, 0, 200, Me.InsideHeight - 2
    
   
    Set m_objHList = New ajpHList 'Khoi tao bien
    Set m_objHList.UseListbox = ListBox1 'ListBox1 la HList
        
    
    With m_objHList
    
    '
    ' The levels begin at 0 (zero)
    ' Increase in level should be incremental
    '
        .AddNode "One", 0
        .AddNode "Two", 1
        .AddNode "Three", 1
        .AddNode "Four", 1
        .AddNode "Five", 0
        .AddNode "Six", 1
        .AddNode "Seven", 2
        .AddNode "Eight", 2
        .AddNode "Nine", 2
        .AddNode "Ten", 1
        .AddNode "Eleven", 2
        .AddNode "Twelve", 2
        .AddNode "Thirteen", 2
        .AddNode "Fourteen", 2
        .AddNode "Fifteen", 0
        .AddNode "Sixteen", 1
        .AddNode "Seventeen", 2
        .AddNode "Eighteen", 3
        .AddNode "Nineteen", 4
        .AddNode "Twenty", 0
        .AddNode "Twenty One", 1
        .AddNode "Twenty Two", 1
        .AddNode "Twenty Three", 2
        .AddNode "Twenty Four", 3
        .AddNode "Twenty Five", 0
        .Create
    End With
    
End Sub

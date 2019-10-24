'--------------------------------------------------
' ajpHList Class Object v1.0
' Written by Andy Pope ©2003, all rights reserved.
' May be redistributed for free, but
' may not be sold without the author's explicit permission.
'--------------------------------------------------
'
Option Explicit

Const HLIST_NAME = "Hierarchical Listbox"
Const HLIST_VERSION = "v1.0"
Const HLIST_AUTHOR = "Andy Pope ©2003"

Private m_intNLevels As Integer
Private m_intNNodes As Integer
Private m_intNodeLevel() As Integer
Private m_intNodeStatus() As Integer
Private m_strNodeText() As String
Private m_strNodes() As String

Private m_strNode_Closed As String
Private m_strNode_Open As String
Private m_strNode_End As String
Private m_strNode_DropPad As String
Private m_strNode_Drop As String
Private m_strNode_Pad As String

Private WithEvents m_objListbox As MSForms.ListBox

Const NODEEND = 0
Const NODEOPEN = 1
Const NODECLOSED = 2

Public Event ClickedNode(ByVal NodeIndex As Integer, ByVal ListIndex As Integer)
Public Event DblClickedNode(ByVal NodeIndex As Integer, ByVal ListIndex As Integer)
Public Property Get Author() As String
    Author = HLIST_AUTHOR
End Property
Public Property Get Name() As String
    Name = HLIST_NAME
End Property

Public Property Get Node(Index As Integer) As String
    Node = m_strNodeText(Index)
End Property
Private Sub m_BuildDropNode()

    Dim intLen As Integer
    
    intLen = Len(m_strNode_Closed)
    If Len(m_strNode_Open) > intLen Then intLen = Len(m_strNode_Open)
    If Len(m_strNode_End) > intLen Then intLen = Len(m_strNode_End)
    
    m_strNode_DropPad = m_strNode_Drop & String(intLen - 1, " ")
    m_strNode_Pad = String(Len(m_strNode_DropPad) + 1, " ")
    
End Sub

Private Sub m_UpdateNodes(NodeIndex As Integer, RowIndex As Integer)
'
' Check whether to expand/collapse node
'
    If m_intNodeStatus(NodeIndex) = NODEEND Then
        ' do nothing
    ElseIf m_intNodeStatus(NodeIndex) = NODECLOSED Then
        m_OpenNode RowIndex, NodeIndex
    ElseIf m_intNodeStatus(NodeIndex) = NODEOPEN Then
        m_CloseNode RowIndex, NodeIndex
    End If
    m_objListbox.ListIndex = RowIndex
    
End Sub
Private Sub m_CloseNode(RowIndex As Integer, NodeIndex As Integer)
'
' Collapse Row item
'
    Dim intIndex As Integer
    Dim strBuf As String
    Dim intInsertIndex As Integer
    Dim intLevel As Integer
    Dim intRowItem As Integer
    Dim intLastNodeIndex As Integer
    Dim intNodeIndex As Integer
        
    m_strNodes(NodeIndex, m_intNodeLevel(NodeIndex)) = m_strNode_Drop & m_strNode_Closed
    m_intNodeStatus(NodeIndex) = NODECLOSED
    
    m_objListbox.List(RowIndex, 0) = m_MakeNode(NodeIndex)
    
    intLevel = m_intNodeLevel(NodeIndex)
    intLastNodeIndex = NodeIndex + 1
    Do While intLastNodeIndex < m_intNNodes - 1
        If m_intNodeLevel(intLastNodeIndex) <= intLevel Then
            ' collapse to this item
            Exit Do
        End If
        intLastNodeIndex = intLastNodeIndex + 1
    Loop
    
    intRowItem = RowIndex + 1
    intNodeIndex = m_objListbox.List(intRowItem, 1)
    Do While intNodeIndex < intLastNodeIndex
        If m_intNodeStatus(intNodeIndex) = NODEOPEN Then
            ' close open nodes as we collapse
            m_intNodeStatus(intNodeIndex) = NODECLOSED
            m_strNodes(intNodeIndex, m_intNodeLevel(intNodeIndex)) = m_strNode_Drop & m_strNode_Closed
        End If
        
        m_objListbox.RemoveItem intRowItem
        If intRowItem >= m_objListbox.ListCount Then Exit Do
        intNodeIndex = m_objListbox.List(intRowItem, 1)
    Loop
    
End Sub
Private Sub m_OpenNode(RowIndex As Integer, NodeIndex As Integer)
'
' Remove RowIndex Item and replace with next level down stuff
'
    Dim intLevel As Integer
    Dim intLastNodeIndex As Integer
    Dim intIndex As Integer
    Dim strBuf As String
    Dim intInsertIndex As Integer
    Dim strStatus As String
    Dim intItemIndex As Integer
    Dim blnLastIndent As Boolean
    
    If RowIndex = m_objListbox.ListCount - 1 Then
        ' last item in list
        intLastNodeIndex = m_intNNodes - 1
        blnLastIndent = True
    Else
        intLastNodeIndex = CInt(m_objListbox.List(RowIndex + 1, 1))
        blnLastIndent = False
    End If
    
    m_strNodes(NodeIndex, m_intNodeLevel(NodeIndex)) = m_strNode_Drop & m_strNode_Open
    m_intNodeStatus(NodeIndex) = NODEOPEN
    
    m_objListbox.List(RowIndex, 0) = m_MakeNode(NodeIndex)
    
    intInsertIndex = RowIndex + 1
    intLevel = m_intNodeLevel(NodeIndex) + 1
    
    For intIndex = NodeIndex + 1 To intLastNodeIndex - 1
        If m_intNodeLevel(intIndex) <= intLevel Then
            m_objListbox.AddItem m_MakeNode(intIndex), intInsertIndex
            m_objListbox.List(intInsertIndex, 1) = CStr(intIndex)
            intInsertIndex = intInsertIndex + 1
        End If
    Next
    
End Sub
Private Function m_MakeNode(NodeIndex As Integer) As String
'
' Create a node constructor list
'
    Dim intLevelIndex As Integer
    Dim strText As String
    
    For intLevelIndex = 0 To m_intNodeLevel(NodeIndex) - 1 'm_intNLevels
        strText = strText & m_strNodes(NodeIndex, intLevelIndex)
    Next
    strText = strText & m_strNodes(NodeIndex, m_intNodeLevel(NodeIndex)) & m_strNodeText(NodeIndex)
    
    m_MakeNode = strText

End Function
Public Sub AddNode(Text As String, Level As Integer)
'
'
    ReDim Preserve m_strNodeText(m_intNNodes) As String
    ReDim Preserve m_intNodeLevel(m_intNNodes) As Integer
    ReDim Preserve m_intNodeStatus(m_intNNodes) As Integer
    
    m_strNodeText(m_intNNodes) = Text
    m_intNodeLevel(m_intNNodes) = Level
    
    m_intNNodes = m_intNNodes + 1
    
    If Level > m_intNLevels Then m_intNLevels = Level   ' remember deepest level
    
End Sub
Public Sub Create()
'
' build up node constructors
'
    m_CreateNodes
    m_BuildNodeList
        
End Sub
Private Sub m_CreateNodes()
'
' Construct the node connectors
' Default starting position is Closed
'
    Dim intLevelIndex As Integer
    Dim intNodeIndex As Integer
    Dim intCurrentLevel As Integer
    Dim intStartDrop As Integer
    Dim intEndDrop As Integer
    Dim intNode As Integer
    Dim strText As String
    Dim intLevel As Integer
    
    ReDim m_strNodes(m_intNNodes, m_intNLevels) As String
    
    For intNodeIndex = 1 To m_intNNodes - 1
        For intLevel = 0 To m_intNodeLevel(intNodeIndex)
            m_strNodes(intNodeIndex, intLevel) = m_strNode_Pad
        Next
    Next
    
    For intNodeIndex = 0 To m_intNNodes - 1
    
        m_intNodeStatus(intNodeIndex) = NODEEND
        intCurrentLevel = m_intNodeLevel(intNodeIndex)
        If (intNodeIndex + 1) < m_intNNodes Then
            intStartDrop = intNodeIndex + 1
            intEndDrop = intStartDrop
            If intEndDrop < m_intNNodes Then
                Do While intEndDrop < m_intNNodes
                    If m_intNodeLevel(intEndDrop) <> intCurrentLevel Then
                        If m_intNodeLevel(intEndDrop) < intCurrentLevel Then
                            ' do not drop thru higher level node
                            intEndDrop = intStartDrop - 1
                            Exit Do
                        End If
                    ElseIf m_intNodeLevel(intEndDrop) = intCurrentLevel Then
                        ' do not drop thru higher level node
                        intEndDrop = intEndDrop - 1
                        Exit Do
                    End If
                    intEndDrop = intEndDrop + 1
                Loop
                If intEndDrop = m_intNNodes Then
                ' gone passes last node
                    If m_intNodeLevel(m_intNNodes - 1) <> intCurrentLevel Then
                        ' check last node against the level
                        intEndDrop = intStartDrop - 1
                    Else
                        intEndDrop = m_intNNodes
                    End If
                End If
            
            Else
                intEndDrop = intStartDrop - 1
            End If
            For intNode = intStartDrop To intEndDrop
                m_strNodes(intNode, intCurrentLevel) = m_strNode_DropPad
            Next
            
            If m_intNodeLevel(intNodeIndex + 1) > intCurrentLevel Then
                m_strNodes(intNodeIndex, intCurrentLevel) = m_strNode_Drop & m_strNode_Closed
                m_intNodeStatus(intNodeIndex) = NODECLOSED
            Else
                m_strNodes(intNodeIndex, intCurrentLevel) = m_strNode_Drop & m_strNode_End
                m_intNodeStatus(intNodeIndex) = NODEEND
            End If
        Else
            ' last item
            m_strNodes(intNodeIndex, intCurrentLevel) = m_strNode_Drop & m_strNode_End
            m_intNodeStatus(intNodeIndex) = NODEEND
        End If
    Next
    
End Sub
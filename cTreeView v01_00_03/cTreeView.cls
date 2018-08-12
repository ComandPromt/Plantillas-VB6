VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
'
' Class Name:   cTreeView         (requires mTreeView.BAS)
' Author:       Graeme Grant
' Date:         28/09/2001
' Version:      01.00.03
' Description:  Advanced TreeView Handler
' Edit History: 01.00.00 09/09/01 Initial Release
'               01.00.01 18/09/01 Added CodeScroll: Home/PgUp/Up/
'                                                   End/PgDn/Down
'               01.00.01 18/09/01 Added Node First/Last Viewable
'               01.00.01 18/09/01 Added Drag Auto-scrolling
'               01.00.02 21/09/01 Fixed Drag mode Icon
'               01.00.03 28/09/01 Added FlatBorder property
'               01.00.03 28/09/01 Add IsChildNode function
'               01.00.03 28/09/01 Added NodeCopy function, CancelCopy
'                                 property and CopyNode Event.
'               01.00.03 28/09/01 Added 'Cut Icon State' [Thanks to
'                                 Brad Martinez]
'               01.00.03 28/09/01 Fixed problem if TreeView adjusts
'                                 the display of nodes between MouseDown
'                                 and MouseUp events causing a VB Error
'                                 in MouseMove event.
'               01.00.03 28/09/01 Fixed Move node when node was not
'                                 Expanded and loaded - bug only when
'                                 loading branches on demand. Worked fine
'                                 if TreeView was prelaoded fully before
'                                 move operation was performed.
'
'===========================================================================

Option Explicit

'===========================================================================
' Debugging... Saves adding the debug statements to the form events
'
#Const DEBUGMODE = 1                    '## 0=No debug
                                        '   1=debug
#Const MOUSEEVENTS = 1                  '## 0=No mouse events
                                        '   1=Mouse Up & Mouse Down
                                        '   2=All Mouse events
#If DEBUGMODE = 1 Then
    Private dbgCtrlName  As String
#End If

'===========================================================================
' Public: Variables and Declarations
'
Public Enum eContextMenuMode
    [Before Click] = 0
    [After Click] = 1
End Enum

Public Enum eCodeScrollView         '@@ v01.00.01
    [Home] = 0
    [Page Up] = 1
    [Up] = 2
    [Down] = 3
    [Page Down] = 4
    [End] = 5
End Enum

Public Enum eNodeDepth              '@@ v01.00.03
    [First Branch] = 0
    [All Branches] = 1
End Enum

'===========================================================================
' Private: Variables and Declarations
'
Private WithEvents oTree  As MSComctlLib.TreeView
Attribute oTree.VB_VarHelpID = -1
Private meContextMenuMode As eContextMenuMode
Private mbDragEnabled     As Boolean
Private mbStartDrag       As Boolean
Private mbInDrag          As Boolean
Private mlNodeHeight      As Long               '## Physical node height (pixels) @@ v01.00.01
Private moFindNode        As MSComctlLib.Node
Private moDragNode        As MSComctlLib.Node

Private mbCancelCopy      As Boolean
Private moCopyNode        As MSComctlLib.Node

'===========================================================================
' Private: Used By pClearTreeviewNodes
'
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
                                        (ByVal hwnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Const WM_SETREDRAW    As Long = &HB
Private Const TV_FIRST        As Long = &H1100
Private Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Private Const TVM_DELETEITEM  As Long = (TV_FIRST + 1)
Private Const TVGN_ROOT       As Long = &H0

'===========================================================================
' Private: cTreeView Events
'
Public Event StartDrag(SourceNode As MSComctlLib.Node)
Public Event Dragging(SourceNode As MSComctlLib.Node, TargetParent As MSComctlLib.Node)
Public Event Dropped(SourceNode As MSComctlLib.Node, TargetParent As MSComctlLib.Node)

Public Event ContextMenu(Node As MSComctlLib.Node, x As Single, y As Single)
Public Event Selected(Node As MSComctlLib.Node)

Public Event CopyNode(DestNode As MSComctlLib.Node, SrcNode As MSComctlLib.Node)    '@@ v01.00.03

'===========================================================================
' TreeView: Events
'
Private Sub oTree_AfterLabelEdit(Cancel As Integer, NewString As String)
#If DEBUGMODE = 1 Then
    Debug.Print dbgCtrlName; "::AfterLabelEdit -> Cancel="; CStr(Cancel); "  NewString="; NewString
#End If
End Sub

Private Sub oTree_BeforeLabelEdit(Cancel As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::BeforeLabelEdit -> Cancel="; CStr(Cancel)
    #End If

End Sub

Private Sub oTree_Click()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Click"
    #End If

End Sub

Private Sub oTree_Collapse(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Collapse -> Node="; Node.Text
    #End If

End Sub

Private Sub oTree_DblClick()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::DblClick"
    #End If

End Sub

Private Sub oTree_DragDrop(Source As Control, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::DragDrop -> Source="; Source.Name; "  X="; CStr(x); "  Y="; CStr(y)
    #End If

    If mbDragEnabled Then
        With oTree
            If Not (.DropHighlight Is Nothing) Then                 '## Did we drop a node?
                If moDragNode <> .DropHighlight Then                '## Yes. Did we drag the node onto itself?
                    RaiseEvent Dropped(moDragNode, .DropHighlight)  '## Notify programmer & Reset
                End If
            End If
            '## Reset
            Set .DropHighlight = Nothing
            Set moDragNode = Nothing
            mbInDrag = False
            mbStartDrag = False
        End With
    End If

End Sub

Private Sub oTree_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::DragOver -> Source="; Source.Name; "  X="; CStr(x); "  Y="; CStr(y)
    #End If

    If mbDragEnabled Then
        With oTree
            If mbStartDrag = True Then
                If mbInDrag = True Then
                    '## Set DropHighlight to the mouse's coordinates.
                    Set .DropHighlight = .HitTest(x, y)
                    If y > (.GetVisibleCount - 2) * mlNodeHeight Then       '@@ v01.00.01
                        pScrollToNextNode(1).EnsureVisible                  '@@ Autoscroll during
                    ElseIf y < 2 * mlNodeHeight Then                        '@@ Dragging operation
                        pScrollToPrevNode(1).EnsureVisible                  '@@ if near top or
                    Else                                                    '@@ bottom of control
                        If Not (.DropHighlight Is Nothing) Then
                            '## We're over a node
                            RaiseEvent Dragging(moDragNode, .DropHighlight)
                        End If                                              '@@ v01.00.01
                    End If
                End If
            End If
        End With
    End If

End Sub

Private Sub oTree_Expand(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Expand -> Node="; Node.Text
    #End If

End Sub

Private Sub oTree_GotFocus()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::GotFocus"
    #End If

End Sub

Private Sub oTree_KeyDown(KeyCode As Integer, Shift As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::KeyDown -> KeyCode="; CStr(KeyCode); "  Shift="; CStr(Shift)
    #End If

End Sub

Private Sub oTree_KeyPress(KeyAscii As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::KeyPress -> KeyAscii="; CStr(KeyAscii)
    #End If

End Sub

Private Sub oTree_KeyUp(KeyCode As Integer, Shift As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::KeyUp -> KeyCode="; CStr(KeyCode); "  Shift="; CStr(Shift)
    #End If

End Sub

Private Sub oTree_LostFocus()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::LostFocus"
    #End If

End Sub

Private Sub oTree_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 1 Or MOUSEEVENTS = 2 Then
            Debug.Print dbgCtrlName; "::MouseDown -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If

    With oTree
        If mbDragEnabled Then                               '## Is drag'n'drop allowed?
            If Button = vbLeftButton Then
                Set moDragNode = .HitTest(x, y)             '## Capture the node to be dragged
            End If
        End If

        '## Fire ContextMenu event before click event?
        If meContextMenuMode = [Before Click] Then
            If Button = vbRightButton Then
                Set .SelectedItem = .HitTest(x, y)
                RaiseEvent ContextMenu(.SelectedItem, x, y)
            End If
        End If
    End With

End Sub

Private Sub oTree_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 2 Then
            Debug.Print dbgCtrlName; "::MouseMove -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If

    On Error GoTo ErrorHandler                              '@@ v01.00.03

    If mbDragEnabled Then                                   '## Is drag'n'drop allowed?
        If Button = vbLeftButton Then                       '## Yes. Signal a Drag operation.
            With oTree
                If Not (.HitTest(x, y) Is Nothing) Then     '## Do we have a node selected?
                    mbInDrag = True                         '## Yes. Set the flag to true.
                    '## Set the drag icon with the CreateDragImage method.
                    '.DragIcon = .SelectedItem.CreateDragImage
                    .DragIcon = moDragNode.CreateDragImage  '@@ v01.00.02
                    .Drag vbBeginDrag                       '## Signal VB to start drag operation.
                    If Not (moDragNode Is Nothing) Then
                        RaiseEvent StartDrag(moDragNode)    '## Notify programmer starting drag operation
                        mbStartDrag = True
                    End If
                End If
            End With
        End If
    End If
    Exit Sub

ErrorHandler:                                               '@@ v01.00.03
    mbInDrag = False                                        '@@
End Sub

Private Sub oTree_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 1 Or MOUSEEVENTS = 2 Then
            Debug.Print dbgCtrlName; "::MouseUp -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If

    '## Fire ContextMenu event after click event?
    If meContextMenuMode = [After Click] Then
        If Button = vbRightButton Then
            With oTree
                Set .SelectedItem = .HitTest(x, y)
                RaiseEvent ContextMenu(.SelectedItem, x, y)
            End With
        End If
    End If

End Sub

Private Sub oTree_NodeCheck(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::NodeCheck -> Node="; Node.Text
    #End If

End Sub

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::NodeClick -> Node="; Node.Text
    #End If

    If mbDragEnabled Then
        '## Cancel if dragging a node
        mbStartDrag = False
        mbInDrag = False
        Set oTree.DropHighlight = Nothing
    End If
    '## Node selected, raise event
    RaiseEvent Selected(Node)
End Sub

Private Sub oTree_OLECompleteDrag(Effect As Long)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLECompleteDrag -> Effect="; CStr(Effect)
    #End If

End Sub

Private Sub oTree_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEDragDrop -> Effect="; CStr(Effect); "  Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
    #End If

End Sub

Private Sub oTree_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEDragOver -> Effect="; CStr(Effect); "  Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y); "  State="; CStr(State)
    #End If

End Sub

Private Sub oTree_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEGiveFeedback -> Effect="; CStr(Effect); "  DefaultCursors="; CStr(DefaultCursors)
    #End If

End Sub

Private Sub oTree_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLESetData -> Effect="; CStr(DataFormat)
    #End If

End Sub

Private Sub oTree_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEStartDrag -> AllowedEffects="; CStr(AllowedEffects)
    #End If

End Sub

Private Sub oTree_Validate(Cancel As Boolean)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Validate -> Cancel="; CStr(Cancel)
    #End If

End Sub

'===========================================================================
'
Private Sub Class_Initialize()
    '## Set defaults...
    meContextMenuMode = [After Click]
    mbDragEnabled = False
End Sub

Private Sub Class_Terminate()
    '## release control
    UnHookCtrl
End Sub

'===========================================================================
' Public subroutines and functions
'
Public Function NodeAdd(Optional ByVal vRelative As Variant, _
                        Optional ByVal eRelationship As MSComctlLib.TreeRelationshipConstants, _
                        Optional ByVal sKey As String, _
                        Optional ByVal sText As String, _
                        Optional ByVal vImage As Variant, _
                        Optional ByVal vSelectedImage As Variant, _
                        Optional ByVal vTag As Variant, _
                        Optional ByVal bBold As Boolean = False, _
                        Optional ByVal bChecked As Boolean = False, _
                        Optional ByVal bEnsureVisible As Boolean = False, _
                        Optional ByVal bExpanded As Boolean = True, _
                        Optional ByVal bSelected As Boolean = False, _
                        Optional ByVal bVisible As Boolean = True, _
                        Optional ByVal lForeColor As OLE_COLOR, _
                        Optional ByVal lBackColor As OLE_COLOR, _
                        Optional ByVal vExpandedImage As Variant) As MSComctlLib.Node
Attribute NodeAdd.VB_Description = "Adds node to TreeView and sets properties in one call"

    Dim Node As MSComctlLib.Node

    On Error GoTo ErrorHandler

    Set Node = oTree.Nodes.Add(vRelative, eRelationship, sKey, sText, vImage, vSelectedImage)
    With Node
        #If DEBUGMODE = 1 Then
            Debug.Print "ADD-> Text = "; .Text; "   Index = "; CStr(.Index)
        #End If
        .Tag = vTag
        .Bold = bBold
        .Checked = bChecked
        If bEnsureVisible Then .EnsureVisible
        .Expanded = bExpanded
        .Selected = bSelected
        '.Visible = bVisible
        If lForeColor = 0 Then
            .ForeColor = vbWindowText
        Else
            .ForeColor = lForeColor
        End If
        If lBackColor = 0 Then
            .BackColor = vbWindowBackground
        Else
            .BackColor = lBackColor
        End If
        If Not IsMissing(vExpandedImage) Then
            .ExpandedImage = vExpandedImage
        End If
    End With
    
    Set NodeAdd = Node
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ClearTreeView()
Attribute ClearTreeView.VB_Description = "A very fast way of clearing all nodes"
    '## Clear the treeview

    Dim hItem As Long
    Dim hwnd  As Long

    Redraw False
    hwnd = oTree.hwnd
    Do
        hItem = SendMessageLong(hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
        If hItem <= 0 Then Exit Do
        SendMessageLong hwnd, TVM_DELETEITEM, &O0, hItem
    Loop
    Redraw True

End Sub

Public Sub Redraw(Enabled As Boolean)
Attribute Redraw.VB_Description = "Enable/disable screen updates."

    Dim hwnd  As Long

    If Not (oTree Is Nothing) Then
        hwnd = oTree.hwnd
        If Enabled Then
            '## Lock the window update to avoid flickering
            SendMessageLong hwnd, WM_SETREDRAW, False, &H0
        Else
            '## Unlock the window
            SendMessageLong hwnd, WM_SETREDRAW, True, &H0
        End If
    End If

End Sub
Public Sub CollapseAll()
Attribute CollapseAll.VB_Description = "Hide all children"
    '## Collapse all nodes

    Dim Node  As Node

    Redraw False
    For Each Node In oTree.Nodes
        With Node
            If .Children Then .Expanded = False
        End With
    Next
    Redraw True

End Sub

Public Sub CollapseChildNodes(ByVal Node As MSComctlLib.Node)
Attribute CollapseChildNodes.VB_Description = " Collapse all children that belong to Node"

    Dim lCount As Long
    Dim lLoop  As Long

    Redraw False
    With Node
        If Not Node Is Nothing Then
            .Expanded = False
        End If

        If .Children Then               '## Does node have children?
            lCount = .Children
            Set Node = .Child
            For lLoop = 1 To lCount
                CollapseChildNodes Node '## Yes, collapse them as well (recursive)
                If lLoop < lCount Then
                    Set Node = .Next
                End If
            Next
        End If
    End With
    Redraw True

End Sub

Public Property Get ContextMenuMode() As eContextMenuMode
Attribute ContextMenuMode.VB_Description = "Raise ContextMenu Event before or after Click Event [Default = After]"
    ContextMenuMode = meContextMenuMode
End Property

Public Property Let ContextMenuMode(ByVal Mode As eContextMenuMode)
    meContextMenuMode = Mode
End Property

Public Property Get Ctrl() As MSComctlLib.TreeView
Attribute Ctrl.VB_UserMemId = 0
Attribute Ctrl.VB_MemberFlags = "640"
    '## Default Object
    Set Ctrl = oTree
End Property

Public Sub CutIconState(SetState As Boolean)        '@@ v01.00.03
    '
    '## Sets or removes the slected (right-clicked) node's specified state.
    '
    Dim hItem As Long
  
    hItem = GetTVItemFromNode(oTree.hwnd, oTree.SelectedItem)
    If hItem Then
        SetTVItemState oTree.hwnd, hItem, TVIS_CUT, SetState
    End If

End Sub

Public Property Get DragEnabled() As Boolean
Attribute DragEnabled.VB_Description = "Enables/disables drag'n'drop"
    DragEnabled = mbDragEnabled
End Property

Public Property Let DragEnabled(ByVal Enabled As Boolean)
    mbDragEnabled = Enabled
End Property

Public Sub ExpandAll()
Attribute ExpandAll.VB_Description = "Show all children"
    '## Expand all nodes

    Dim Node  As Node

    On Error GoTo ErrorHandler                      '@@ v01.00.01
    Redraw False
    For Each Node In oTree.Nodes
        With Node
            If .Children Then .Expanded = True
        End With
    Next
    Redraw True
    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 35606: ExpandAll
    End Select
End Sub

Public Sub ExpandChildNodes(ByVal Node As MSComctlLib.Node)
Attribute ExpandChildNodes.VB_Description = "Expand all children that belong to Node"

    Dim lCount As Long
    Dim lLoop  As Long

    Redraw False
    With Node
        If Not Node Is Nothing Then
            .Expanded = True
        End If

        If .Children Then               '## Does node have children?
            lCount = .Children
            Set Node = .Child
            For lLoop = 1 To lCount
                ExpandChildNodes Node   '## Yes, expand them as well (recursive)
                If lLoop < lCount Then
                    Set Node = .Next
                End If
            Next
        End If
    End With
    Redraw True

End Sub

Public Sub HookCtrl(Ctrl As MSComctlLib.TreeView)

    Dim bState As Boolean

    Set oTree = Ctrl                                '## Capture the Treeview control.
                                                    '   Class won't work if this isn't called first.
    With oTree
        #If DEBUGMODE = 1 Then
            dbgCtrlName = .Parent.Name + "." + .Name
        #End If
        '
        '## calculate the height (pixels) of a node '@@ v01.00.01
        '
        bState = .Scroll                            '## Remember if scrollbars are enabled
        .Scroll = False                             '## Turn them off
        mlNodeHeight = .Height \ .GetVisibleCount   '## calculate height
        .Scroll = bState                            '## restore scrollbar state
        FlatBorder .hwnd                            '@@ v01.00.03 (Forced flat border)
    End With

End Sub

Public Property Get ParentNodeText(Node As MSComctlLib.Node) As String
Attribute ParentNodeText.VB_Description = "Returns the parent node text (Safe routine to stop error condition if no parent node)"
    '## Returns the parent text
    '   NOTE: If a Node has no parent then VB raises an error.
    '         This is a much safer Node.Parent.Text routine.
    With Node
        ParentNodeText = IIf(.FullPath = .Text, "", .Parent)
    End With
End Property

Public Function IsChildNode(TestNode As MSComctlLib.Node, _
                            ParentNode As MSComctlLib.Node, _
                   Optional NodeDepth As eNodeDepth = [First Branch]) As Boolean    '@@ v01.00.03

    If NodeDepth = [First Branch] Then
        If ParentNode.Children Then
            If Not IsRootNode(TestNode) Then
                IsChildNode = (TestNode.Parent = ParentNode)
            End If
        End If
    Else
        IsChildNode = IsParentNode(TestNode, ParentNode)
    End If

End Function

Public Function IsParentNode(ChildNode As MSComctlLib.Node, _
                             ParentNode As MSComctlLib.Node) As Boolean
    '## Checks if one node is the parent of another.
    '   This is a recursive routine that steps down through
    '   the branches of the parent node.

    Dim lNDX As Long

    If ParentNode.Children Then             '## Does the parent node have children?
        lNDX = ParentNode.Child.Index       '## Yes, remember the first child
        Do                                  '## Step through all child nodes
            If lNDX = ChildNode.Index Then  '## is ChildNode the test node?
                IsParentNode = True         '## ParentNode is the parent of ChildNode.
                Exit Do
            End If
            If IsParentNode(ChildNode, oTree.Nodes(lNDX)) Then  '## Step down through the branches
                IsParentNode = True         '## ParentNode is the parent of ChildNode.
                Exit Do
            End If
            If lNDX <> ParentNode.Child.LastSibling.Index Then  '## Have we tested the last child node?
                lNDX = oTree.Nodes(lNDX).Next.Index             '## No. Point to the next child node
            Else
                Exit Do                                         '## Yes.
            End If
        Loop
    End If

End Function

Public Function IsRootNode(Node As MSComctlLib.Node) As Boolean
Attribute IsRootNode.VB_Description = "Check is selected node is a root node."
    '## Check is selected node is a root node.
    With Node
        IsRootNode = (.FullPath = .Text)
    End With
End Function

Public Sub UnHookCtrl()
Attribute UnHookCtrl.VB_Description = "Releases the hooked control."
    '## Release the hooked control
    Set oTree = Nothing
End Sub

Public Property Let CancelCopy(IsCancelled As Boolean)                  '@@ v01.00.03
    '## Cancel the node copy operation
    mbCancelCopy = IsCancelled
End Property

Public Function NodeCopy(DestNode As MSComctlLib.Node, _
                         SrcNode As MSComctlLib.Node, _
          Optional ByVal IncludeChildren As Boolean = True) As Boolean  '@@ v01.00.03

    Dim oSNode As MSComctlLib.Node
    Dim oDNode As MSComctlLib.Node

    '
    '## Ensure that we don't overwrite the original nodes values
    '
    Set oSNode = SrcNode
    Set oDNode = DestNode

    If oSNode = oDNode Then
        '## Same node - therefore no point trying to copy
        mbCancelCopy = True
        Exit Function
    End If
    '
    '## Force the Nodes to be expanded before the copy operation is performed. This
    '   will allow child nodes to be loaded if done on demand
    '
    oDNode.Expanded = True
    oSNode.Expanded = True
    '
    '## Let the actual process happen externally to cTREEVIEW - a must if working with a database
    '   or multiple types of data.
    '
    RaiseEvent CopyNode(oDNode, oSNode)     '## Get the calling routine to copy the node.
    If mbCancelCopy Then                    '## Was the process cancelled?
        Exit Function                       '## Yes.
    End If
    If IncludeChildren Then                                     '## Include all child nodes?
        If oSNode.Children Then                                 '## Yes. Are there child nodes?
            Set oSNode = oSNode.Child.FirstSibling              '## Yes. Select the first child.
            Do
                If NodeCopy(oDNode.Child.LastSibling, oSNode, IncludeChildren) Then '## Copy child node
                    If Not (oSNode = oSNode.LastSibling) Then   '## Are there more child nodes?
                        Set oSNode = oSNode.Next                '## Yes.
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop
        End If
    End If
    NodeCopy = True                                         '## Success!

End Function

Public Function NodeDelete(Node As MSComctlLib.Node, _
            Optional ByVal bSelect As Boolean = True) As Boolean

    With Node
        If Not IsRootNode(Node) Then    '## Is this a root node?
            With .Parent                '## No.
                .EnsureVisible
                .Selected = bSelect
            End With
        End If
        '## Delete node
        oTree.Nodes.Remove .Index       '## Delete the node
    End With
    NodeDelete = True

End Function

Public Function NodeFind(oResultNode As MSComctlLib.Node, _
                   ByVal NodeText As String, _
          Optional ByVal NodeKey As String, _
          Optional ByVal bSelect As Boolean = True) As Boolean
Attribute NodeFind.VB_Description = "Finds a Node and optionally selects & ensures visibility"

    Dim Node     As Node
    Dim bCheckKey As Boolean
    Dim bKeyOK    As Boolean

    bCheckKey = (Len(Trim$(NodeKey)) > 0)
    bKeyOK = (Not bCheckKey)
    For Each Node In oTree.Nodes
        With Node
            If UCase$(.Text) = UCase$(NodeText) Then
                If bCheckKey Then bKeyOK = (UCase$(.Key) = UCase$(NodeKey))
                If bKeyOK Then
                    If bSelect Then
                        .Selected = True
                        .EnsureVisible
                    End If
                    Set oResultNode = Node
                    NodeFind = True
                    Exit For
                End If
            End If
        End With
    Next

End Function

Public Function NodeMove(ParentNode As MSComctlLib.Node, _
                         ChildNode As MSComctlLib.Node, _
          Optional ByVal bSelect As Boolean = True) As Boolean
Attribute NodeMove.VB_Description = "Move a node from one parent to another  and optionally selects & ensures visibility"

    Dim lNDX   As Long
    Dim lCount As Long
    Dim lLoop  As Long
    Dim bRoot  As Boolean

    With ChildNode
        If ParentNode = ChildNode Then
            '## Same node - therefore no point
            Exit Function
        End If
        If IsParentNode(ParentNode, ChildNode) Then '## Are we moving a parent node?
            If IsRootNode(ChildNode) Then           '## Yes. Is it a root node?
                Exit Function                       '## Yes. Can't move a root node.
            End If
            '## move the children before moving the designated node
            lCount = .Children
            For lLoop = 1 To lCount
                lNDX = .Child.Index
                Set oTree.Nodes(lNDX).Parent = .Parent
            Next
        End If
        '## Force the ParentNode to be expanded before the move
        ParentNode.Expanded = True                  '@@ v01.00.03
        '## Give the child a new parent
        Set .Parent = ParentNode
        If bSelect Then
            .EnsureVisible
            .Selected = bSelect
        End If
    End With
    NodeMove = True

End Function

Public Function NodeRename(Node As MSComctlLib.Node, _
                     ByVal NewNodeText As String, _
            Optional ByVal bSelect As Boolean = True) As Boolean
Attribute NodeRename.VB_Description = "Changes the text of a node and optionally selects & ensures visibility"

    With Node
        If Len(NewNodeText) Then            '## Check if we have new text
            .Text = NewNodeText
            If Not IsRootNode(Node) Then
                .EnsureVisible
                .Selected = bSelect
            End If
            NodeRename = True
        End If
    End With

End Function


Public Sub ScrollView(Dir As eCodeScrollView)                                       '@@ v01.00.01
Attribute ScrollView.VB_Description = "Scrolls the treview using code"
    '
    '## Scrolls the treview using code
    '
    Dim lPageSize As Long

    With oTree
        lPageSize = .GetVisibleCount - 1        '## Number of viewable nodes less 1
        Select Case Dir
            Case [Home]
                .Nodes(1).Root.EnsureVisible
            Case [Page Up]
                pScrollToPrevNode(lPageSize).EnsureVisible
            Case [Up]
                pScrollToPrevNode(1).EnsureVisible
            Case [Down]
                pScrollToNextNode(1).EnsureVisible
            Case [Page Down]
                pScrollToNextNode(lPageSize).EnsureVisible
            Case [End]
                pScrollToLastNode(.Nodes(1).Root.LastSibling).EnsureVisible
        End Select
    End With

End Sub

Public Function NodeFirstViewable() As MSComctlLib.Node                             '@@ v01.00.01
Attribute NodeFirstViewable.VB_Description = "Returns the first viewable node"
    '
    '## Returns the first viewable node
    '
    Dim yPos As Long
    Dim xPos As Long
    
    With oTree
        yPos = mlNodeHeight \ 2 '## Halfway down the first viewable node
        Do
            xPos = xPos + 100   '## move across the control gradually
            If Not (oTree.HitTest(xPos, 5) Is Nothing) Then
                '
                '## Find the node from left to right until we get a hit
                '   and return the node found
                '
                Set NodeFirstViewable = oTree.HitTest(xPos, 5)
                Exit Do
            End If
        Loop Until xPos > oTree.Width '## Have we hit the RHS?
    End With

End Function

Public Function NodeLastViewable() As MSComctlLib.Node  '@@ v01.00.01
Attribute NodeLastViewable.VB_Description = "Returns the last visible node"
    '
    '## Returns the last visible node
    '
    Set NodeLastViewable = pScrollToNextNode
End Function

'===========================================================================
' Private subroutines and functions
'
Private Function pScrollToLastNode(Node As MSComctlLib.Node) As MSComctlLib.Node    '@@ v01.00.01
    '
    '## Recursive find and return the last expanded and visible node
    '
    Dim oNode  As MSComctlLib.Node

    With Node
        Debug.Print Node.Text
        If .Children Then
            '
            '## We have child nodes
            '
            If .Expanded Then
                '
                '## And expanded. Therefore traverse the next branch
                '
                Set pScrollToLastNode = pScrollToLastNode(Node.Child.LastSibling)
            Else
                '
                '## we've hit the end
                '
                Set pScrollToLastNode = Node
            End If
        Else
            '
            '## we've hit the end
            '
            Set pScrollToLastNode = Node
        End If
    End With

End Function

Private Function pScrollToNextNode(Optional ByVal NumNodes As Long = 0) As MSComctlLib.Node '@@ v01.00.01

    Dim oVNode As MSComctlLib.Node          '## First viewable node
    Dim oSNode As MSComctlLib.Node          '## Start node
    Dim lCount As Long                      '## bean counter
    Dim lMax   As Long                      '## Maximum beans
    Dim bStart As Boolean                   '## Set true if we've reached first viewable node

    Set oVNode = NodeFirstViewable          '## Get first viewable node
    Set oSNode = oTree.Nodes(1).Root        '## record start node as first root node
    lMax = oTree.GetVisibleCount + NumNodes '## Set to maximum nodes viewable + number of
                                            '   nodes to scroll down
    Do                                      '## Step through each root node
        '
        '## Step through every node under each root node looking counting each
        '   node after the first viewable node until we've hit either the last
        '   node for found the last viewable.
        '
        Set pScrollToNextNode = pTraverseDown(oSNode, oVNode, lMax, lCount, bStart)
        If (Not (lCount = lMax)) And (Not (oSNode = oSNode.LastSibling)) Then
            '
            '## We still haven't found the node and we haven't exhausted
            '   every root node
            '
            Set oSNode = oSNode.Next
        Else
            '
            '## We've either hit the end or found the last viewable node
            '
            Exit Do
        End If
    Loop

End Function

Private Function pScrollToPrevNode(Optional ByVal NumNodes As Long = 0) As MSComctlLib.Node '@@ v01.00.01

    Dim lCount As Long                          '## bean counter
    Dim oNode As MSComctlLib.Node               '## Start node

    Set oNode = NodeFirstViewable               '## record start node as first root node
    '
    '## Loop backwards and forwards unit the node is found
    '
    Do
        '
        '## Backwards
        '
        If oNode = oNode.FirstSibling Then      '## Is the it the first child node of parent?
            If IsRootNode(oNode) Then           '## Yes. Is it a root node?
                If oNode.Root = oNode Then      '## Yes. Is it the primary root node?
                    Exit Do                     '## we have found the node wanted
                Else
                    lCount = lCount + 1         '## Increment bean counter
                    If lCount = NumNodes Then
                        Exit Do                 '## we have found the node wanted
                    End If
                End If
            Else
                Set oNode = oNode.Parent        '## step back up the branch
                lCount = lCount + 1             '## Increment bean counter
                If lCount = NumNodes Then
                    Exit Do                     '## we have found the node wanted
                End If
            End If
        Else
            Set oNode = oNode.Previous          '## move up a node for the same parent
            If oNode.Children Then              '## Does the new node have child nodes?
                '
                '## Yes. Step down (Forward) through the branches to the last node
                '
                Do
                    If oNode.Expanded Then      '## Are the child nodes expanded?
                        '
                        '## Yes. Get the next last node of child branch
                        '
                        Set oNode = oNode.Child.LastSibling
                    Else
                        '
                        '## we've found the previous node
                        '
                        Set pScrollToPrevNode = oNode
                        Exit Do                 '## we have found the node wanted
                    End If
                Loop While oNode.Children       '## are there more child nodes?
                lCount = lCount + 1             '## Increment bean counter
                If lCount = NumNodes Then
                    Exit Do                     '## we have found the node wanted
                End If
            Else
                '
                '## we've found the previous node
                '
                lCount = lCount + 1             '## Increment bean counter
                If lCount = NumNodes Then
                    Exit Do                     '## we have found the node wanted
                End If
            End If
        End If
    Loop
    Set pScrollToPrevNode = oNode               '## Return the requested previous node

End Function

Private Function pTraverseDown(Node As MSComctlLib.Node, _
                               StartNode As MSComctlLib.Node, _
                               Max As Long, _
                               Count As Long, _
                               Start As Boolean) As MSComctlLib.Node    '@@ v01.00.01
    '
    '## This will recursively step through every node from the start node
    '   looking for the first viewable. When found will count every following
    '   node until either the last node is hit or max is reached.
    '
    Dim oNode  As MSComctlLib.Node
    Dim lLoop As Long

    With Node
        'Debug.Print .Text, .FullPath, Count
        If .Key = StartNode.Key Then
            '
            '## we've reached the first viewable node. Start counting beans.
            '
            Start = True
        End If
        If Start Then
           Count = Count + 1                    '## Count the node
            If Count = Max Then                 '## Have we reached our quota?
                Set pTraverseDown = Node        '## Yes. Return the Last viewable node
                Exit Function
            End If
        End If
        If .Children Then                       '## Node has children?
            If .Expanded Then                   '## Yes. But have we expanded?
                Set oNode = .Child.FirstSibling '## Yes. Lets start work on the children
                Do
                    '
                    '## Lets start the process for every child node (recursively)
                    '
                    Set pTraverseDown = pTraverseDown(oNode, StartNode, Max, Count, Start)
                    If Count = Max Then         '## Have we reached our quota?
                        Exit Do                 '## Yes. Let's stop looking
                    End If
                    If Not (oNode = oNode.LastSibling) Then
                        '
                        '## We still haven't found the node and we haven't exhausted
                        '   every root node
                        '
                        Set oNode = oNode.Next
                    Else
                        '
                        '## No more child node left. Let's cut or losses and return
                        '   the last child node
                        '
                        Exit Do
                    End If
                Loop
            Else
                '
                '## Node's not expanded. Therefore return the same node
                '
                Set pTraverseDown = Node
            End If
        Else
            '
            '## Node's has no children. Therefore return the same node.
            '
            Set pTraverseDown = Node
        End If
    End With

End Function

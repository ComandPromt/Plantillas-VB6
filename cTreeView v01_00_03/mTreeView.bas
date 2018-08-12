Attribute VB_Name = "mTreeView"
Option Explicit
'
' Brad Martinez, http://www.mvps.org/ccrp/
'

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            wParam As Any, _
                            lParam As Any) As Long   ' <---

' ===========================================================================
' treeview definitions defined in Commctrl.h at:
' http://premium.microsoft.com/msdn/library/sdkdoc/c67_4c8m.htm

Public Type TVITEM   ' was TV_ITEM
    mask           As Long
    hItem          As Long
    State          As TVITEM_state
    stateMask      As Long
    pszText        As String   ' Long   ' pointer
    cchTextMax     As Long
    iImage         As Long
    iSelectedImage As Long
    cChildren      As Long
    lParam         As Long
End Type

Public Enum TVITEM_mask
    TVIF_TEXT = &H1
    TVIF_IMAGE = &H2
    TVIF_PARAM = &H4
    TVIF_STATE = &H8
    TVIF_HANDLE = &H10
    TVIF_SELECTEDIMAGE = &H20
    TVIF_CHILDREN = &H40
#If (WIN32_IE >= &H400) Then   ' WIN32_IE = 1024 (>= Comctl32.dll v4.71)
    TVIF_INTEGRAL = &H80
#End If
    TVIF_DI_SETITEM = &H1000   ' Notification
End Enum

Public Enum TVITEM_state
    TVIS_SELECTED = &H2
    TVIS_CUT = &H4
    TVIS_DROPHILITED = &H8
    TVIS_BOLD = &H10
    TVIS_EXPANDED = &H20
    TVIS_EXPANDEDONCE = &H40
#If (WIN32_IE >= &H300) Then
    TVIS_EXPANDPARTIAL = &H80
#End If
    
    TVIS_OVERLAYMASK = &HF00
    TVIS_STATEIMAGEMASK = &HF000
    TVIS_USERMASK = &HF000
End Enum

' messages
Public Const TV_FIRST = &H1100
Public Const TVM_GETNEXTITEM = (TV_FIRST + 10)
Public Const TVM_GETITEM = (TV_FIRST + 12)
Public Const TVM_SETITEM = (TV_FIRST + 13)

' TVM_GETNEXTITEM wParam values
Public Enum TVGN_Flags
    TVGN_ROOT = &H0
    TVGN_NEXT = &H1
    TVGN_PREVIOUS = &H2
    TVGN_PARENT = &H3
    TVGN_CHILD = &H4
    TVGN_FIRSTVISIBLE = &H5
    TVGN_NEXTVISIBLE = &H6
    TVGN_PREVIOUSVISIBLE = &H7
    TVGN_DROPHILITE = &H8
    TVGN_CARET = &H9
#If (WIN32_IE >= &H400) Then   ' >= Comctl32.dll v4.71
    TVGN_LASTVISIBLE = &HA
#End If
End Enum


'/////////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200

Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000


' If successful, returns the treeview item handle represented by
' the specified Node, returns 0 otherwise.

Public Function GetTVItemFromNode(hwndTV As Long, nod As Node) As Long
  Dim nod1 As Node
  Dim anSiblingPos() As Integer  ' contains the sibling position of the node and all it's parents
  Dim nLevel As Integer              ' hierarchical level of the node
  Dim hItem As Long
  Dim i As Integer
  Dim nPos As Integer

  Set nod1 = nod

  ' Continually work backwards from the current node to the current node's
  ' first sibling, caching the current node's sibling position in the one-based
  ' array. Then get the first sibling's parent node and start over. Keep going
  ' until the postion of the specified node's top level parent item is obtained...
  Do While (nod1 Is Nothing) = False
    nLevel = nLevel + 1
    ReDim Preserve anSiblingPos(nLevel)
    anSiblingPos(nLevel) = GetNodeSiblingPos(nod1)
    Set nod1 = nod1.Parent
  Loop

  ' Get the hItem of the first item in the treeview
  hItem = TreeView_GetRoot(hwndTV)
  If hItem Then

    ' Now work backwards through the cached node positions in the array
    ' (from the first treeview node to the specified node), obtaining the respective
    ' item handle for each node at the cached position. When we get to the
    ' specified node's position (the value of the first element in the array), we
    ' got it's hItem...
    For i = nLevel To 1 Step -1
      nPos = anSiblingPos(i)
      
      Do While nPos > 1
        hItem = TreeView_GetNextSibling(hwndTV, hItem)
        nPos = nPos - 1
      Loop
      
      If (i > 1) Then hItem = TreeView_GetChild(hwndTV, hItem)
    Next

    GetTVItemFromNode = hItem

  End If   ' hItem

End Function

' Returns the one-base position of the specified node
' with respect to it's sibling order.

Public Function GetNodeSiblingPos(nod As Node) As Integer
  Dim nod1 As Node
  Dim nPos As Integer
  
  Set nod1 = nod
  
  ' Keep counting up from one until the node has no more previous siblings
  Do While (nod1 Is Nothing) = False
    nPos = nPos + 1
    Set nod1 = nod1.Previous
  Loop
  
  GetNodeSiblingPos = nPos
  
End Function

' Sets the specified item state of the specified item.
' (TVIF_STATE acts only on the item's "state", not the item itself).

'   hwndTV          - treeview's window handle
'   hItem               - handle of the item
'   dwNewState   - state bits to set
'   fAdd                - flag indicating whether the state bits are added or removed,

' Returns True if the new state was successfully set, False otherwise.
' (some of the time...??)

Public Function SetTVItemState(hwndTV As Long, _
                               hItem As Long, _
                               dwNewState As TVITEM_state, _
                               fAdd As Boolean) As Boolean
    Dim tvi As TVITEM

    tvi.hItem = hItem
    tvi.mask = TVIF_HANDLE Or TVIF_STATE
    tvi.State = fAdd And dwNewState
    ' Indicate what state bits we're changing
    tvi.stateMask = dwNewState
    ' Old docs say returns 0 on success, -1 on failure. New docs say
    ' returns TRUE if successful, or FALSE otherwise we'll go new...
    SetTVItemState = TreeView_SetItem(hwndTV, tvi) '= 0

End Function

' ===========================================================================
' Treeview macros defined in Commctrl.h

' Retrieves some or all of a tree-view item's attributes.
' Returns TRUE if successful or FALSE otherwise.

Public Function TreeView_GetItem(hwnd As Long, pitem As TVITEM) As Boolean
    TreeView_GetItem = SendMessage(hwnd, TVM_GETITEM, 0, pitem)
End Function

' Sets some or all of a tree-view item's attributes.
' Old docs say returns zero if successful or - 1 otherwise.
' New docs say returns TRUE if successful, or FALSE otherwise

Public Function TreeView_SetItem(hwnd As Long, pitem As TVITEM) As Boolean
    TreeView_SetItem = SendMessage(hwnd, TVM_SETITEM, 0, pitem)
End Function

' TreeView_GetNextItem

' Retrieves the tree-view item that bears the specified relationship to a specified item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetNextItem(hwnd As Long, hItem As Long, flag As Long) As Long
    TreeView_GetNextItem = SendMessage(hwnd, TVM_GETNEXTITEM, ByVal flag, ByVal hItem)
End Function

' Retrieves the first child item. The hitem parameter must be NULL.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetChild(hwnd As Long, hItem As Long) As Long
    TreeView_GetChild = TreeView_GetNextItem(hwnd, hItem, TVGN_CHILD)
End Function

' Retrieves the next sibling item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetNextSibling(hwnd As Long, hItem As Long) As Long
    TreeView_GetNextSibling = TreeView_GetNextItem(hwnd, hItem, TVGN_NEXT)
End Function

' Retrieves the previous sibling item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetPrevSibling(hwnd As Long, hItem As Long) As Long
    TreeView_GetPrevSibling = TreeView_GetNextItem(hwnd, hItem, TVGN_PREVIOUS)
End Function

' Retrieves the parent of the specified item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetParent(hwnd As Long, hItem As Long) As Long
    TreeView_GetParent = TreeView_GetNextItem(hwnd, hItem, TVGN_PARENT)
End Function

' Retrieves the first visible item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetFirstVisible(hwnd As Long) As Long
    TreeView_GetFirstVisible = TreeView_GetNextItem(hwnd, 0, TVGN_FIRSTVISIBLE)
End Function

' Retrieves the next visible item that follows the specified item. The specified item must be visible.
' Use the TVM_GETITEMRECT message to determine whether an item is visible.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetNextVisible(hwnd As Long, hItem As Long) As Long
    TreeView_GetNextVisible = TreeView_GetNextItem(hwnd, hItem, TVGN_NEXTVISIBLE)
End Function

' Retrieves the first visible item that precedes the specified item. The specified item must be visible.
' Use the TVM_GETITEMRECT message to determine whether an item is visible.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetPrevVisible(hwnd As Long, hItem As Long) As Long
    TreeView_GetPrevVisible = TreeView_GetNextItem(hwnd, hItem, TVGN_PREVIOUSVISIBLE)
End Function

' Retrieves the currently selected item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetSelection(hwnd As Long) As Long
    TreeView_GetSelection = TreeView_GetNextItem(hwnd, 0, TVGN_CARET)
End Function

' Retrieves the item that is the target of a drag-and-drop operation.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetDropHilight(hwnd As Long) As Long
    TreeView_GetDropHilight = TreeView_GetNextItem(hwnd, 0, TVGN_DROPHILITE)
End Function

' Retrieves the topmost or very first item of the tree-view control.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetRoot(hwnd As Long) As Long
    TreeView_GetRoot = TreeView_GetNextItem(hwnd, 0, TVGN_ROOT)
End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////////
'## Control Routines
'
Public Sub FlatBorder(ByVal hwnd As Long)

    Dim TFlat As Long

    TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
    TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    SetWindowLong hwnd, GWL_EXSTYLE, TFlat
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or _
                                      SWP_NOZORDER Or _
                                      SWP_FRAMECHANGED Or _
                                      SWP_NOSIZE Or _
                                      SWP_NOMOVE

End Sub

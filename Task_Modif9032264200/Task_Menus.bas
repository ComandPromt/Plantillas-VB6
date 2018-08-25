Attribute VB_Name = "Task_Menus"
'Im still adding things to this procedure as i discover them.
'--------------------------------------------------
'some things i would like to do with this later:

'i need to get picture/icon data from each menu item and use it in the treeview as well
'--------------------------------------------------


Option Explicit
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Const MF_GRAYED As Long = &H1&
Private Const MF_DISABLED As Long = &H2&
Private Const MF_BITMAP As Long = &H4&
Private Const MF_CHECKED As Long = &H8&
Private Const MF_POPUP As Long = &H10&
Private Const MF_MENUBARBREAK As Long = &H20&
Private Const MF_MENUBREAK As Long = &H40&
Private Const MF_HILITE As Long = &H80&
Private Const MF_OWNERDRAW As Long = &H100&
Private Const MF_USECHECKBITMAPS As Long = &H200&
Private Const MF_BYPOSITION As Long = &H400& ' I dont think this goes in a STATE, but its the only value i have that fits.
Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_DEFAULT As Long = &H1000&
Private Const MF_SYSMENU As Long = &H2000&
Private Const MF_HELP As Long = &H4000&
Private Const MF_MOUSESELECT As Long = &H8000&
Private Const MF_NotKnown As Long = &HFF0000
Private Const MF_HSZ_INFO As Long = &H1000000
Private Const MF_SENDMSGS As Long = &H2000000
Private Const MF_POSTMSGS As Long = &H4000000
Private Const MF_CALLBACKS As Long = &H8000000
Private Const MF_ERRORS As Long = &H10000000
Private Const MF_LINKS As Long = &H20000000
Private Const MF_CONV As Long = &H40000000
Private Const MF_MASK As Long = &HFF000000
Public Const MF_REMOVE As Long = &H1000&
Private Const WM_COMMAND As Long = &H111

Private Sub AddItems2list(mylist As ListBox, ParamArray item())

  Dim X As Long

    For X = LBound(item) To UBound(item)
        mylist.AddItem item(X)
    Next X

End Sub

Public Sub CheckItem(MenuHwnd As Long, ItemID As Long, Check As Boolean)

    CheckMenuItem MenuHwnd, ItemID, Check

End Sub

Public Sub CheckMenuStats(mylist As ListBox, statedata As Long)

    mnuSetData mylist, statedata, MF_MASK, 0
    mnuSetData mylist, statedata, MF_CONV, 1
    mnuSetData mylist, statedata, MF_LINKS, 2
    mnuSetData mylist, statedata, MF_ERRORS, 3
    mnuSetData mylist, statedata, MF_CALLBACKS, 4
    mnuSetData mylist, statedata, MF_POSTMSGS, 5
    mnuSetData mylist, statedata, MF_SENDMSGS, 6
    mnuSetData mylist, statedata, MF_HSZ_INFO, 7
    mnuSetData mylist, statedata, MF_NotKnown, 8
    mnuSetData mylist, statedata, MF_MOUSESELECT, 9
    mnuSetData mylist, statedata, MF_HELP, 10
    mnuSetData mylist, statedata, MF_SYSMENU, 11
    mnuSetData mylist, statedata, MF_DEFAULT, 12
    mnuSetData mylist, statedata, MF_SEPARATOR, 13
    mnuSetData mylist, statedata, MF_BYPOSITION, 14
    mnuSetData mylist, statedata, MF_USECHECKBITMAPS, 15
    mnuSetData mylist, statedata, MF_OWNERDRAW, 16
    mnuSetData mylist, statedata, MF_HILITE, 17
    mnuSetData mylist, statedata, MF_MENUBREAK, 18
    mnuSetData mylist, statedata, MF_MENUBARBREAK, 19
    mnuSetData mylist, statedata, MF_POPUP, 20
    mnuSetData mylist, statedata, MF_CHECKED, 21
    mnuSetData mylist, statedata, MF_BITMAP, 22
    mnuSetData mylist, statedata, MF_DISABLED, 23
    mnuSetData mylist, statedata, MF_GRAYED, 24

End Sub

Public Sub EnableItem(MenuHwnd As Long, ItemID As Long, Enable As Boolean)

    EnableMenuItem MenuHwnd, ItemID, Enable

End Sub

'i use this to fill my list with items i hardcoded
Public Sub FillListWithMenuItems(mylist As ListBox)

    mylist.Clear
    AddItems2list mylist, "MF_MASK", "MF_CONV", "MF_LINKS", "MF_ERRORS", "MF_CALLBACKS", "MF_POSTMSGS", _
                  "MF_SENDMSGS", "MF_HSZ_INFO", "MF_&HFF0000", "MF_MOUSESELECT", "MF_HELP", _
                  "MF_SYSMENU", "MF_DEFAULT", "MF_SEPARATOR", "MF_BYPOSITION", "MF_USECHECKBITMAPS", _
                  "MF_OWNERDRAW", "MF_HILITE", "MF_MENUBREAK", "MF_MENUBARBREAK", "MF_POPUP", _
                  "MF_CHECKED", "MF_BITMAP", "MF_DISABLED", "MF_GRAYED"

End Sub

Private Function IsItemDisabled(MnuState As Long) As Boolean

    IsItemDisabled = ((MnuState And MF_DISABLED) Or (MnuState And MF_GRAYED)) And (MnuState <> -1)

End Function

Private Function IsItemSeparator(MnuState As Long) As Boolean

    IsItemSeparator = (MnuState And MF_SEPARATOR) And (MnuState <> -1)

End Function

Public Function IsMenu(hwnd As Long) As Boolean

  Dim MenuHwnd As Long, sysmenuhwnd As Long

    MenuHwnd = GetMenu(hwnd)
    sysmenuhwnd = GetSystemMenu(hwnd, 0)
    IsMenu = MenuHwnd Or sysmenuhwnd

End Function

'a loop i created that will most definately be used by others on psc
'it gets a list of menu items and puts them in a treeview
Public Function mchild(mTree As TreeView, hwnd As Long, MenuType As String, NodeIdentifier As String, Optional iCount As Long)

  ' this sub gets the menu List Children
  
  Dim mCount As Long, LookFor As Long, SubMenu As Long, SubMenuID As Long, TmpStr As String, ParentItem As String, ThisText As String
  Dim ThisItem As String
  Dim Nodx As Node
  Dim MnuState As Long

    ParentItem = MenuType
    mCount = GetMenuItemCount(hwnd)
    For LookFor = 0 To mCount - 1
        SubMenu = GetSubMenu(hwnd, LookFor)
        SubMenuID = GetMenuItemID(hwnd, LookFor)
        TmpStr = String$(255, " ")
        GetMenuString hwnd, LookFor, TmpStr, 255, MF_BYPOSITION
        ThisText = Left$(TmpStr, InStr(TmpStr, Chr$(0)) - 1)
        MnuState = GetMenuState(hwnd, SubMenuID, 0&)
        If IsItemSeparator(MnuState) Then
            ThisText = "{Seperator Bar}"
        End If
        If ThisText = "" Then
            ThisText = "{Various Uses}"
        End If
        iCount = iCount + 1
        ThisItem = NodeIdentifier & CStr(SubMenuID) & ":" & CStr(MnuState) & ":" & CStr(iCount)
        '''''also note:  if SubMenuID = -1 then it has Branches
        Set Nodx = mTree.Nodes.Add(ParentItem, tvwChild, ThisItem, ThisText)
        If IsItemDisabled(MnuState) Then
            Nodx.ForeColor = RGB(127, 127, 127)
        End If
        If (mCount > 0) Or (MnuState = -1) Or (SubMenuID = -1) Then
            MenuType = ThisItem
            mchild mTree, SubMenu, MenuType, NodeIdentifier, iCount
        End If
    Next LookFor
    Set Nodx = Nothing

End Function

Private Sub mnuSetData(mylist As ListBox, statedata As Long, MF_Flag As Long, ItemNum As Long)

    If statedata And MF_Flag Then
        mylist.Selected(ItemNum) = True
        statedata = statedata - MF_Flag
      Else
        mylist.Selected(ItemNum) = False
    End If

End Sub

Public Sub RemoveMenuItem(OwnerHwnd As Long, IsSystemMenu As Boolean, MenuID As Long)

  'still working on this Procedure
  
  Dim MenuHwnd As Long

    If IsSystemMenu Then
        MenuHwnd = GetSystemMenu(OwnerHwnd, 0)
      Else
        MenuHwnd = GetMenu(OwnerHwnd)
    End If
    Call RemoveMenu(MenuHwnd, MenuID, MF_REMOVE)
    DrawMenuBar OwnerHwnd

End Sub

Public Sub RunMenuItem(hwnd As Long, mnID As Long)

  'currently runs window menu items, but not system menu items

    SendMessageLong hwnd, WM_COMMAND, mnID, 0&

End Sub

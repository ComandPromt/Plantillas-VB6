Attribute VB_Name = "OwnerDrawn"
Option Explicit

DefLng A-Z

Const MFT_STRING = 0

' ***************************
' Owner drawn menus procedures.
'
' Please keep this notice with the code.
' Written by Bi Hai, 1998, thriller@163.net
' If you find this code useful, please let me know at thriller@163.net
'
' ***************************

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type Size
    cx As Long
    cy As Long
End Type

'MENUITEMINFO
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

' MEASUREITEMSTRUCT for ownerdraw
Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    itemData As Long
End Type

' DRAWITEMSTRUCT for ownerdraw
Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    itemData As Long
End Type

Public Declare Function GetMenu Lib "user32" _
   (ByVal hwnd As Long) As Long

Public Declare Function GetSubMenu Lib "user32" _
   (ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Declare Function GetMenuItemCount Lib "user32" _
   (ByVal hMenu As Long) As Long

Public Declare Function GetMenuItemInfo Lib "user32" _
    Alias "GetMenuItemInfoA" _
   (ByVal hMenu As Long, ByVal un As Long, _
    ByVal b As Boolean, lpmii As MENUITEMINFO) As Long

Declare Function GetMenuItemID Lib "user32" _
    (ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Declare Function SetMenuItemInfo Lib "user32" _
    Alias "SetMenuItemInfoA" _
   (ByVal hMenu As Long, ByVal uItem As Long, _
    ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long

Declare Function AppendMenu Lib "user32" _
    Alias "AppendMenuA" (ByVal hMenu As Long, _
    ByVal wFlags As Long, ByVal wIDNewItem As Long, _
    ByVal lpNewItem As Any) As Long

Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long

Declare Function CreateFont Lib "gdi32" _
    Alias "CreateFontA" (ByVal H As Long, _
    ByVal W As Long, ByVal E As Long, ByVal O As Long, _
    ByVal W As Long, ByVal I As Long, ByVal U As Long, _
    ByVal S As Long, ByVal C As Long, ByVal OP As Long, _
    ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, _
    ByVal F As String) As Long

Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long

'MENUITEMINFO
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_CHECKMARKS = &H8
Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA = &H20

'menustyle
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&

Public Const MF_STRING = &H0&
Public Const MF_BITMAP = &H4&
Public Const MF_OWNERDRAW = &H100&

'textout style
Public Const ETO_OPAQUE = 2

' Owner draw state
Public Const ODS_SELECTED = &H1
Public Const ODS_GRAYED = &H2
Public Const ODS_DISABLED = &H4
Public Const ODS_CHECKED = &H8
Public Const ODS_FOCUS = &H10

'messages:
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_MENUSELECT = &H11F
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONUP = &H205
Public Const WM_USER = &H400
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_SYSCOLORCHANGE = &H15

Declare Sub MemCopy Lib "kernel32" Alias _
        "RtlMoveMemory" (dest As Any, src As Any, _
        ByVal numbytes As Long)

Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)

Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal msg As Long, _
    ByVal wparam As Long, ByVal lparam As Long) As Long

Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function TextOut Lib "gdi32" Alias "TextOutA" _
    (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
    ByVal lpString As String, ByVal nCount As Long) As Long

Declare Function ExtTextOut Lib "gdi32" Alias _
    "ExtTextOutA" (ByVal hdc As Long, ByVal x As _
    Long, ByVal y As Long, ByVal wOptions As Long, _
    lpRect As RECT, ByVal lpString As String, _
    ByVal nCount As Long, lpDx As Long) As Long

Declare Function GetDC Lib "user32" _
    (ByVal hwnd As Long) As Long

Declare Function ReleaseDC Lib "user32" _
    (ByVal hwnd As Long, ByVal hdc As Long) As Long

Declare Function SelectObject Lib "gdi32" _
    (ByVal hdc As Long, ByVal hObject As Long) As Long

Declare Function SetBkColor Lib "gdi32" _
    (ByVal hdc As Long, ByVal crColor As Long) As Long

Declare Function SetTextColor Lib "gdi32" _
    (ByVal hdc As Long, ByVal crColor As Long) As Long

Declare Function GetSysColor Lib "user32" _
    (ByVal nIndex As Long) As Long

Declare Function GetTextExtentPoint Lib "gdi32" _
    Alias "GetTextExtentPointA" (ByVal hdc As Long, _
    ByVal lpszString As String, ByVal cbString As Long, _
    lpSize As Size) As Long

Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_GRAYTEXT = 17

'consts MenuItem IDs.
Public Const IDM_CHARACTER = 10
Public Const IDM_REGULAR = 11
Public Const IDM_BOLD = 12
Public Const IDM_ITALIC = 13
Public Const IDM_UNDERLINE = 14

Type myItemType
    hFont As Long
    cchItemText As Integer
    szItemText As String * 32
End Type

Public OldWindowProc
Public hMenu, hSubMenu
Public mnuItemCount, MyItem() As myItemType
Public clrPrevText, clrPrevBkgnd
Public hfntPrev

Public Function NewWindowProc(ByVal hwnd As Long, _
    ByVal msg As Long, ByVal wparam As Long, _
    lparam As Long) As Long

    Dim mM As MEASUREITEMSTRUCT
    Dim dM As DRAWITEMSTRUCT

    Select Case msg

        Case WM_DRAWITEM

            MemCopy dM, lparam, Len(dM)
            OnDrawMenuItem hwnd, dM

        Case WM_MEASUREITEM

            MemCopy mM, lparam, Len(mM)
            mM = OnMeasureItem(hwnd, mM)
            MemCopy lparam, mM, Len(mM)

        Case WM_COMMAND

           'Put your Menu Command here.

        Case WM_SYSCOLORCHANGE

           'Put your code here.

        Case Else


    End Select

NewWindowProc = CallWindowProc(OldWindowProc, _
hwnd, msg, wparam, VarPtr(lparam))

End Function

Sub CreateMenus(hwnd As Long)

    'get Menus
    hMenu = GetMenu(hwnd)
    hSubMenu = GetSubMenu(hMenu, 0)

    'remove original menu item
    RemoveMenu hSubMenu, 0, MF_BYPOSITION

    'creates string menus
    AppendMenu hSubMenu, MF_STRING, IDM_REGULAR, "Regular"
    AppendMenu hSubMenu, MF_STRING, IDM_BOLD, "Bold"
    AppendMenu hSubMenu, MF_STRING, IDM_ITALIC, "Italic"
    AppendMenu hSubMenu, MF_STRING, IDM_UNDERLINE, "Underline"

    'call to make OwnerDrawMenus
    CreateOwnerDrawMenus

End Sub
Sub CreateOwnerDrawMenus()

   Dim minfo As MENUITEMINFO, id As Integer

  'get the menuitem handle
   hSubMenu = GetSubMenu(GetMenu(Form1.hwnd), 0)
   mnuItemCount = GetMenuItemCount(hSubMenu)

   'ReDim usertype array for menuitems
   ReDim MyItem(0 To mnuItemCount - 1) As myItemType
   Dim r As Long

   'loop to fill array
   For id = 0 To mnuItemCount - 1
    minfo.cbSize = Len(minfo)
    minfo.fMask = MIIM_TYPE
    minfo.fType = MFT_STRING
    minfo.dwTypeData = Space$(256)
    minfo.cch = Len(minfo.dwTypeData)

    'get menuitem data
    r = GetMenuItemInfo(hSubMenu, id, True, minfo)

    'and save into user array
    MyItem(id).cchItemText = minfo.cch 'menuitem length
    MyItem(id).szItemText = Trim(minfo.dwTypeData) 'text
    MyItem(id).hFont = CreateMenuItemFont(id) 'font

    'change menu type
    minfo.fType = MF_OWNERDRAW
    minfo.fMask = MIIM_TYPE Or MIIM_DATA
    minfo.dwItemData = id

    'into MF_OWNERDRAW
    r = SetMenuItemInfo(hSubMenu, id, True, minfo)

   Next id


End Sub


Function OnMeasureItem(hwnd As Long, lpmis As MEASUREITEMSTRUCT) As MEASUREITEMSTRUCT

    Dim xM As MEASUREITEMSTRUCT, hfntOld As Long
    Dim S As Size, hdc As Long

    'find DC
    hdc = GetDC(hwnd)

    hfntOld = SelectObject(hdc, MyItem(lpmis.itemData).hFont)

    GetTextExtentPoint hdc, MyItem(lpmis.itemData).szItemText, _
            MyItem(lpmis.itemData).cchItemText, S

    'set menu item rect
    xM.itemWidth = S.cx + 10
    xM.itemHeight = S.cy

    SelectObject hdc, hfntOld
    ReleaseDC hwnd, hdc

    LSet OnMeasureItem = xM

End Function

Sub OnDrawMenuItem(hwnd As Long, lpdis As DRAWITEMSTRUCT)

    Dim x, y

    'set the menuitem colors
    If (lpdis.itemState And ODS_SELECTED) Then 'if selected
        clrPrevText = SetTextColor(lpdis.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT))
        clrPrevBkgnd = SetBkColor(lpdis.hdc, GetSysColor(COLOR_HIGHLIGHT))
    Else
        clrPrevText = SetTextColor(lpdis.hdc, GetSysColor(COLOR_MENUTEXT))
        clrPrevBkgnd = SetBkColor(lpdis.hdc, GetSysColor(COLOR_MENU))
    End If

    'leave space for checkmark
    'may use GetMenuCheckMarkDimensions
    x = lpdis.rcItem.Left + 20
    y = lpdis.rcItem.Top

    hfntPrev = SelectObject(lpdis.hdc, MyItem(lpdis.itemData).hFont)

    ExtTextOut lpdis.hdc, x, y, ETO_OPAQUE, _
        lpdis.rcItem, Trim(" "), 1&, 0&

    TextOut lpdis.hdc, x, y, MyItem(lpdis.itemData).szItemText, MyItem(lpdis.itemData).cchItemText

    'may put some bitblt function here also.

    SelectObject lpdis.hdc, hfntPrev
    SetTextColor lpdis.hdc, clrPrevText
    SetBkColor lpdis.hdc, clrPrevBkgnd

End Sub
Function CreateMenuItemFont(uID As Integer) As Long
Dim Weight As Long
Dim use_italic As Long
Dim use_underline As Long
Dim use_strikethrough As Long

   Select Case uID + 11

        Case IDM_BOLD

            Weight = 700

        Case IDM_ITALIC

            use_italic = True

        Case IDM_UNDERLINE

            use_underline = True

     End Select

CreateMenuItemFont = CreateFont(18, 0, _
        0, 0, Weight, _
        use_italic, use_underline, _
        use_strikethrough, 136, 0, _
        16, 0, 0, 0)

End Function

Sub OnDestroy()
Dim r As Long

   'do some clean works
   Dim minfo As MENUITEMINFO, id As Integer

   hSubMenu = GetSubMenu(GetMenu(Form1.hwnd), 0)
   mnuItemCount = GetMenuItemCount(hSubMenu)

  For id = 0 To mnuItemCount - 1
   minfo.fMask = MIIM_DATA

   r = GetMenuItemInfo(hSubMenu, id, True, minfo)

   DeleteObject minfo.dwItemData

   r = SetMenuItemInfo(hSubMenu, id, True, minfo)
  Next

  Erase MyItem

End Sub





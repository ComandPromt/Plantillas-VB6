Attribute VB_Name = "Declarations"
Option Private Module
Option Explicit
#If Win32 Then
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long

Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal wndrpcPrev As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Public Const BS_SOLID = 0

Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = -4

'Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
'                    (ByVal dwExStyle As Long, _
'                     ByVal lpClassName As String, _
'                     ByVal lpWindowName As String, _
'                     ByVal dwStyle As Long, _
'                     ByVal x As Long, ByVal y As Long, _
'                     ByVal nWidth As Long, ByVal nHeight As Long, _
'                     ByVal hWndParent As Long, _
'                     ByVal hMenu As Long, _
'                     ByVal hInstance As Long, lpParam As Any) As Long

Declare Function GetFocus Lib "user32" () As Long

Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpstr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Public Const COLOR_BACKGROUND = 1
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15

Public Const NULL_PEN = 8
Public Const PS_SOLID = 0

Public Const DT_LEFT = &H0
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800

Public Const LB_ADDSTRING = &H180
Public Const LB_INSERTSTRING = &H181
Public Const LB_GETITEMDATA = &H199
Public Const LB_SETITEMDATA = &H19A
Public Const LB_GETCARETINDEX = &H19F

'Owner draw action
Public Const ODA_DRAWENTIRE = &H1
Public Const ODA_SELECT = &H2
Public Const ODA_FOCUS = &H4

' Owner draw state
Public Const ODS_SELECTED = &H1
Public Const ODS_FOCUS = &H10

Public Const WM_DRAWITEM = &H2B
Public Const WM_SYSCOLORCHANGE = &H15

Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    ItemData As Long
End Type

#If USEGETPROP Then
Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As Any, ByVal hData As Long) As Long
Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As Any) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As Any) As Long
Public Atomizer As New Atomizer
#End If
#End If

'Subclassed WindowProc:
Public Function ODLCtlProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ODLCtlProc = ODLCtl(hWnd).WindowProc(hWnd, uMsg, wParam, lParam)
End Function

'Get a reference to the control associated with hWnd
Private Function ODLCtl(hWnd As Long) As OwnerDrawListBox
Dim Obj As Object
Dim pObj As Long
#If USEGETPROP Then
    pObj = GetProp(hWnd, Atomizer)
#Else
    pObj = GetWindowLong(hWnd, GWL_USERDATA)
#End If
    CopyMemory Obj, pObj, 4
    Set ODLCtl = Obj
    CopyMemory Obj, 0&, 4
End Function



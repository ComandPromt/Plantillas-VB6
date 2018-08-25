Attribute VB_Name = "Task_MouseOver"
' this Module is used to allow the user to find an Object/control on the treeview
'by using the mouse
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function HiliteMenuItem Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Private Const MF_HILITE As Long = &H80&
Public Const Evnt_FindHwndFromMouseOver As Long = 2
Public Const Evnt_Countdown As Long = 3
Public Function FindHwndFromMouseOver() As Long
    Dim Pt As POINTAPI
    'Get the current cursor position
    GetCursorPos Pt
    'Get the window under the cursor
    FindHwndFromMouseOver = WindowFromPoint(Pt.X, Pt.Y)
End Function

Public Sub Proc_FindHwndFromMouseOver(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)

'i made it very easy to find the node keyname
SelectedNodeKey = "t" & CStr(FindHwndFromMouseOver)
'lock treeview so it can search faster without the GUI involvement while its scrolling through the names
LockWindowUpdate frmMain.TaskTree.hwnd
For Each n In frmMain.TaskTree.Nodes
    'In other words If CurrentNodeInTheLoop=TheSelectedWindowFromMouse
    If n.Key = SelectedNodeKey Then
        n.Selected = True 'found it; Select the node.
        Exit For
    End If
    'If it gets to here then window doesn't exist on the treeview
    'maybe i'll make a future addition on this part.
Next n
LockWindowUpdate (0) 'unlock the treeview Gui

End Sub
Public Sub Proc_FHFMO_CountDown(ByVal Timerhwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)

Static CountD As Long
Dim MainMenu As Long, SubMenu As Long, lngID As Long
'if Count is 0 then either its first time running the sub,
'  or has ran the sub and it counted down before so start over.
If CountD = 0 Then
    CountD = 9
End If
CountD = CountD - 1
MainMenu = GetMenu(Timerhwnd)
lngID = GetMenuItemID(MainMenu, 2&)
If CountD Then
    'update the menu item so that it shows the correct Count
    frmMain.mnuFindObj.Caption = CStr(CountD) & " Seconds left..."
    HiliteMenuItem Timerhwnd, MainMenu, lngID, MF_HILITE 'hilite menu item
Else
    frmMain.mnuFindObj.Caption = "Find object with mouse"
    'we are done. Kill our timers
    KillTimer Timerhwnd, Evnt_Countdown
    KillTimer Timerhwnd, Evnt_FindHwndFromMouseOver
    HiliteMenuItem Timerhwnd, MainMenu, lngID, 0& 'unhilite
End If
'refresh the menu gui
DrawMenuBar Timerhwnd
End Sub

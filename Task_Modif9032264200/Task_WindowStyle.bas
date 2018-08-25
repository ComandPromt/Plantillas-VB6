Attribute VB_Name = "Task_WindowStyle"
'this function is used to get/set the windowstyles of an object
'this is very useful in a app that modifies form/control appearance
'i added the Unknown styles for future reference. they are Unknown to me, but
' i see them used in many Custom controls and Listviews and other controls.
Option Explicit
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const GWL_ID As Long = -12
Public Const GWW_HINSTANCE As Long = -6
Public Const GWL_EXSTYLE As Long = -20
Public Const GWL_STYLE As Long = -16
Private Const SWP_NOSIZE As Long = 1
Private Const SWP_NOMOVE As Long = 2
Private Const SWP_NOZORDER As Long = 4

' Window Styles
Public Const WS_ACTIVECAPTION As Long = &H1
Public Const WS_UnknownH2 As Long = &H2
Public Const WS_UnknownH4 As Long = &H4
Public Const WS_UnknownH8 As Long = &H8
Public Const WS_UnknownH10 As Long = &H10
Public Const WS_UnknownH20 As Long = &H20
Public Const WS_UnknownH40 As Long = &H40
Public Const WS_UnknownH80 As Long = &H80
Public Const WS_UnknownH100 As Long = &H100
Public Const WS_UnknownH200 As Long = &H200
Public Const WS_UnknownH400 As Long = &H400
Public Const WS_UnknownH800 As Long = &H800
Public Const WS_UnknownH1000 As Long = &H1000
Public Const WS_UnknownH2000 As Long = &H2000
Public Const WS_UnknownH4000 As Long = &H4000
Public Const WS_UnknownH8000 As Long = &H8000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_SYSMENU As Long = &H80000
Public Const WS_HSCROLL As Long = &H100000
Public Const WS_VSCROLL As Long = &H200000
Public Const WS_DLGFRAME As Long = &H400000
Public Const WS_BORDER As Long = &H800000
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const WS_CLIPCHILDREN As Long = &H2000000
Public Const WS_CLIPSIBLINGS As Long = &H4000000
Public Const WS_DISABLED As Long = &H8000000
Public Const WS_VISIBLE As Long = &H10000000
Public Const WS_MINIMIZE As Long = &H20000000
Public Const WS_CHILD As Long = &H40000000
Public Const WS_POPUP As Long = &H80000000

'Extended Window Styles
Public Const WS_EX_DLGMODALFRAME As Long = &H1
Public Const WS_EX_UnknownH2 As Long = &H2
Public Const WS_EX_NOPARENTNOTIFY As Long = &H4
Public Const WS_EX_TOPMOST As Long = &H8
Public Const WS_EX_ACCEPTFILES As Long = &H10
Public Const WS_EX_TRANSPARENT As Long = &H20
Public Const WS_EX_MDICHILD As Long = &H40
Public Const WS_EX_TOOLWINDOW As Long = &H80
Public Const WS_EX_WINDOWEDGE As Long = &H100
Public Const WS_EX_CLIENTEDGE As Long = &H200
Public Const WS_EX_CONTEXTHELP As Long = &H400
Public Const WS_EX_UnknownH800 As Long = &H800
Public Const WS_EX_RIGHT As Long = &H1000
Public Const WS_EX_RTLREADING As Long = &H2000
Public Const WS_EX_LEFTSCROLLBAR As Long = &H4000
Public Const WS_EX_UnknownH8000 As Long = &H8000
Public Const WS_EX_CONTROLPARENT As Long = &H10000
Public Const WS_EX_STATICEDGE As Long = &H20000
Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_EX_LAYERED As Long = &H80000
Public Const WS_EX_NOINHERITLAYOUT As Long = &H100000
Public Const WS_EX_UnknownH200000 As Long = &H200000
Public Const WS_EX_LAYOUTRTL As Long = &H400000
Public Const WS_EX_NOACTIVATE As Long = &H8000000

Private Sub AddItems2list(mylist As ListBox, ParamArray item())
'this is called by AddToList & AddToListX
' it adds a ton of items to a listbox with small amount of code
  Dim X As Long

    For X = LBound(item) To UBound(item)
        mylist.AddItem item(X)
    Next X

End Sub

Public Sub AddToList(Dalist As ListBox)
'add items to list
    AddItems2list Dalist, "WS_POPUP", "WS_CHILD", "WS_VISIBLE", "WS_DISABLED", "WS_MINIMIZE", "WS_MAXIMIZE", _
                  "WS_MINIMIZEBOX", "WS_MAXIMIZEBOX", "WS_THICKFRAME", "WS_BORDER", "WS_DLGFRAME", _
                  "WS_SYSMENU", "WS_VSCROLL", "WS_HSCROLL", "WS_CLIPSIBLINGS", "WS_CLIPCHILDREN", _
                  "WS_ACTIVECAPTION", "WS_UnknownH2", "WS_UnknownH4", "WS_UnknownH8", _
                  "WS_UnknownH10", "WS_UnknownH20", "WS_UnknownH40", "WS_UnknownH80", "WS_UnknownH100", _
                  "WS_UnknownH200", "WS_UnknownH400", "WS_UnknownH800", "WS_UnknownH1000", "WS_UnknownH2000", _
                  "WS_UnknownH4000", "WS_UnknownH8000"

End Sub

Public Sub AddToListX(Dalist As ListBox)
'add items to list
    AddItems2list Dalist, "WS_EX_DLGMODALFRAME", "WS_EX_UnknownH2", "WS_EX_NOPARENTNOTIFY", "WS_EX_TOPMOST", "WS_EX_ACCEPTFILES", _
                  "WS_EX_TRANSPARENT", "WS_EX_MDICHILD", "WS_EX_TOOLWINDOW", "WS_EX_WINDOWEDGE", _
                  "WS_EX_CLIENTEDGE", "WS_EX_CONTEXTHELP", "WS_EX_RIGHT", "WS_EX_RTLREADING", _
                  "WS_EX_LEFTSCROLLBAR", "WS_EX_CONTROLPARENT", "WS_EX_STATICEDGE", "WS_EX_APPWINDOW", _
                  "WS_EX_LAYERED", "WS_EX_LAYOUTRTL", "WS_EX_NOACTIVATE", "WS_EX_NOINHERITLAYOUT", _
                  "WS_EX_UnknownH800", "WS_EX_UnknownH8000", "WS_EX_UnknownH200000"

End Sub

Public Function GetWndStyle(wnd As Long, StyleType As Long, TheStyle As Long) As Byte

  Dim af As Long

    GetWndStyle = 0
    ' Get style
    af = GetWindowLong(wnd, StyleType&)
    If af And TheStyle& Then
        GetWndStyle = 1
    End If

End Function

Public Function GetWndTypeVal(hwnd As Long, StyleType As Long) As Long

    GetWndTypeVal = GetWindowLong(hwnd, StyleType)

End Function

Public Sub ListGetStyles(Dalist As ListBox, mainhwnd As Long)

  Dim istyle As Long

    istyle = GetWindowLong(mainhwnd, GWL_STYLE)
    Dalist.Selected(0) = StyleFuncHelper(istyle, WS_POPUP)
    Dalist.Selected(1) = StyleFuncHelper(istyle, WS_CHILD)
    Dalist.Selected(2) = StyleFuncHelper(istyle, WS_VISIBLE)
    Dalist.Selected(3) = StyleFuncHelper(istyle, WS_DISABLED)
    Dalist.Selected(4) = StyleFuncHelper(istyle, WS_MINIMIZE)
    Dalist.Selected(5) = StyleFuncHelper(istyle, WS_MAXIMIZE)
    Dalist.Selected(6) = StyleFuncHelper(istyle, WS_MINIMIZEBOX)
    Dalist.Selected(7) = StyleFuncHelper(istyle, WS_MAXIMIZEBOX)
    Dalist.Selected(8) = StyleFuncHelper(istyle, WS_THICKFRAME)
    Dalist.Selected(9) = StyleFuncHelper(istyle, WS_BORDER)
    Dalist.Selected(10) = StyleFuncHelper(istyle, WS_DLGFRAME)
    Dalist.Selected(11) = StyleFuncHelper(istyle, WS_SYSMENU)
    Dalist.Selected(12) = StyleFuncHelper(istyle, WS_VSCROLL)
    Dalist.Selected(13) = StyleFuncHelper(istyle, WS_HSCROLL)
    Dalist.Selected(14) = StyleFuncHelper(istyle, WS_CLIPSIBLINGS)
    Dalist.Selected(15) = StyleFuncHelper(istyle, WS_CLIPCHILDREN)
    Dalist.Selected(16) = StyleFuncHelper(istyle, WS_ACTIVECAPTION)
    Dalist.Selected(17) = StyleFuncHelper(istyle, WS_UnknownH2)
    Dalist.Selected(18) = StyleFuncHelper(istyle, WS_UnknownH4)
    Dalist.Selected(19) = StyleFuncHelper(istyle, WS_UnknownH8)
    Dalist.Selected(20) = StyleFuncHelper(istyle, WS_UnknownH10)
    Dalist.Selected(21) = StyleFuncHelper(istyle, WS_UnknownH20)
    Dalist.Selected(22) = StyleFuncHelper(istyle, WS_UnknownH40)
    Dalist.Selected(23) = StyleFuncHelper(istyle, WS_UnknownH80)
    Dalist.Selected(24) = StyleFuncHelper(istyle, WS_UnknownH100)
    Dalist.Selected(25) = StyleFuncHelper(istyle, WS_UnknownH200)
    Dalist.Selected(26) = StyleFuncHelper(istyle, WS_UnknownH400)
    Dalist.Selected(27) = StyleFuncHelper(istyle, WS_UnknownH800)
    Dalist.Selected(28) = StyleFuncHelper(istyle, WS_UnknownH1000)
    Dalist.Selected(29) = StyleFuncHelper(istyle, WS_UnknownH2000)
    Dalist.Selected(30) = StyleFuncHelper(istyle, WS_UnknownH4000)
    Dalist.Selected(31) = StyleFuncHelper(istyle, WS_UnknownH8000)

End Sub

Public Sub ListGetStylesX(DalistEx As ListBox, mainhwnd As Long)

  Dim istyle As Long

    istyle = GetWindowLong(mainhwnd, GWL_EXSTYLE)
    DalistEx.Selected(0) = StyleFuncHelper(istyle, WS_EX_DLGMODALFRAME)
    DalistEx.Selected(1) = StyleFuncHelper(istyle, WS_EX_UnknownH2)
    DalistEx.Selected(2) = StyleFuncHelper(istyle, WS_EX_NOPARENTNOTIFY)
    DalistEx.Selected(3) = StyleFuncHelper(istyle, WS_EX_TOPMOST)
    DalistEx.Selected(4) = StyleFuncHelper(istyle, WS_EX_ACCEPTFILES)
    DalistEx.Selected(5) = StyleFuncHelper(istyle, WS_EX_TRANSPARENT)
    DalistEx.Selected(6) = StyleFuncHelper(istyle, WS_EX_MDICHILD)
    DalistEx.Selected(7) = StyleFuncHelper(istyle, WS_EX_TOOLWINDOW)
    DalistEx.Selected(8) = StyleFuncHelper(istyle, WS_EX_WINDOWEDGE)
    DalistEx.Selected(9) = StyleFuncHelper(istyle, WS_EX_CLIENTEDGE)
    DalistEx.Selected(10) = StyleFuncHelper(istyle, WS_EX_CONTEXTHELP)
    DalistEx.Selected(11) = StyleFuncHelper(istyle, WS_EX_RIGHT)
    DalistEx.Selected(12) = StyleFuncHelper(istyle, WS_EX_RTLREADING)
    DalistEx.Selected(13) = StyleFuncHelper(istyle, WS_EX_LEFTSCROLLBAR)
    DalistEx.Selected(14) = StyleFuncHelper(istyle, WS_EX_CONTROLPARENT)
    DalistEx.Selected(15) = StyleFuncHelper(istyle, WS_EX_STATICEDGE)
    DalistEx.Selected(16) = StyleFuncHelper(istyle, WS_EX_APPWINDOW)
    DalistEx.Selected(17) = StyleFuncHelper(istyle, WS_EX_LAYERED)
    DalistEx.Selected(18) = StyleFuncHelper(istyle, WS_EX_LAYOUTRTL)
    DalistEx.Selected(19) = StyleFuncHelper(istyle, WS_EX_NOACTIVATE)
    DalistEx.Selected(20) = StyleFuncHelper(istyle, WS_EX_NOINHERITLAYOUT)
    DalistEx.Selected(21) = StyleFuncHelper(istyle, WS_EX_UnknownH800)
    DalistEx.Selected(22) = StyleFuncHelper(istyle, WS_EX_UnknownH8000)
    DalistEx.Selected(23) = StyleFuncHelper(istyle, WS_EX_UnknownH200000)

End Sub

Public Function SetWindowStyle(wnd As Long, GWLStyle As Long, dwNewStyle As Long, fAdd As Boolean, Optional fRedraw As Boolean = True) As Boolean

  Dim dwCurStyle As Long
  Dim dwStyleType As Long

    dwCurStyle = GetWindowLong(wnd, GWLStyle)
    If Err.LastDllError = 0 Then
        If fAdd And (dwCurStyle And dwNewStyle) = 0 Then
            ' Setting the new style and it is not already set...
            dwCurStyle = dwCurStyle Or dwNewStyle

          ElseIf (Not fAdd) And (dwCurStyle And dwNewStyle) Then
            ' Removing the new style and it's already set...
            dwCurStyle = dwCurStyle And (Not dwNewStyle)
        End If

        SetWindowLong wnd, GWLStyle, dwCurStyle
        SetWindowStyle = (Err.LastDllError = 0)

        If fRedraw Then
            SetWindowPos wnd, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOMOVE Or SWP_NOSIZE
        End If
    End If

End Function

Public Sub SetWS(mainhwnd As Long, item As Integer, IsSelected As Boolean)
'WS_Flag is the Window Style
'find which one and set it true
  Dim WS_Flag As Long
    Select Case item
      Case 0:        WS_Flag = WS_POPUP
      Case 1:        WS_Flag = WS_CHILD
      Case 2:        WS_Flag = WS_VISIBLE
      Case 3:        WS_Flag = WS_DISABLED
      Case 4:        WS_Flag = WS_MINIMIZE
      Case 5:        WS_Flag = WS_MAXIMIZE
      Case 6:        WS_Flag = WS_MINIMIZEBOX
      Case 7:        WS_Flag = WS_MAXIMIZEBOX
      Case 8:        WS_Flag = WS_THICKFRAME
      Case 9:        WS_Flag = WS_BORDER
      Case 10:        WS_Flag = WS_DLGFRAME
      Case 11:        WS_Flag = WS_SYSMENU
      Case 12:        WS_Flag = WS_VSCROLL
      Case 13:        WS_Flag = WS_HSCROLL
      Case 14:        WS_Flag = WS_CLIPSIBLINGS
      Case 15:        WS_Flag = WS_CLIPCHILDREN
      Case 16:        WS_Flag = WS_ACTIVECAPTION
      Case 17:        WS_Flag = WS_UnknownH2
      Case 18:        WS_Flag = WS_UnknownH4
      Case 19:        WS_Flag = WS_UnknownH8
      Case 20:        WS_Flag = WS_UnknownH10
      Case 21:        WS_Flag = WS_UnknownH20
      Case 22:        WS_Flag = WS_UnknownH40
      Case 23:        WS_Flag = WS_UnknownH80
      Case 24:        WS_Flag = WS_UnknownH100
      Case 25:        WS_Flag = WS_UnknownH200
      Case 26:        WS_Flag = WS_UnknownH400
      Case 27:        WS_Flag = WS_UnknownH800
      Case 28:        WS_Flag = WS_UnknownH1000
      Case 29:        WS_Flag = WS_UnknownH2000
      Case 30:        WS_Flag = WS_UnknownH4000
      Case 31:        WS_Flag = WS_UnknownH8000
    End Select
    SetWindowStyle mainhwnd, GWL_STYLE, WS_Flag, IsSelected, True

End Sub

'WS_EX_Flag is the Extended Windows Style
'find which one and set it true
Public Sub SetWSX(mainhwnd As Long, item As Integer, IsSelected As Boolean)

  Dim WS_EX_Flag As Long

    Select Case item
      Case 0:        WS_EX_Flag = WS_EX_DLGMODALFRAME
      Case 1:        WS_EX_Flag = WS_EX_UnknownH2
      Case 2:        WS_EX_Flag = WS_EX_NOPARENTNOTIFY
      Case 3:        WS_EX_Flag = WS_EX_TOPMOST
      Case 4:        WS_EX_Flag = WS_EX_ACCEPTFILES
      Case 5:        WS_EX_Flag = WS_EX_TRANSPARENT
      Case 6:        WS_EX_Flag = WS_EX_MDICHILD
      Case 7:        WS_EX_Flag = WS_EX_TOOLWINDOW
      Case 8:        WS_EX_Flag = WS_EX_WINDOWEDGE
      Case 9:        WS_EX_Flag = WS_EX_CLIENTEDGE
      Case 10:        WS_EX_Flag = WS_EX_CONTEXTHELP
      Case 11:        WS_EX_Flag = WS_EX_RIGHT
      Case 12:        WS_EX_Flag = WS_EX_RTLREADING
      Case 13:        WS_EX_Flag = WS_EX_LEFTSCROLLBAR
      Case 14:        WS_EX_Flag = WS_EX_CONTROLPARENT
      Case 15:        WS_EX_Flag = WS_EX_STATICEDGE
      Case 16:        WS_EX_Flag = WS_EX_APPWINDOW
      Case 17:        WS_EX_Flag = WS_EX_LAYERED
      Case 18:        WS_EX_Flag = WS_EX_LAYOUTRTL
      Case 19:        WS_EX_Flag = WS_EX_NOACTIVATE
      Case 20:        WS_EX_Flag = WS_EX_NOINHERITLAYOUT
      Case 21:        WS_EX_Flag = WS_EX_UnknownH800
      Case 22:        WS_EX_Flag = WS_EX_UnknownH8000
      Case 23:        WS_EX_Flag = WS_EX_UnknownH200000
    End Select
    SetWindowStyle mainhwnd, GWL_EXSTYLE, WS_EX_Flag, IsSelected, True

End Sub
'this recursive function is used to find style of controls.
Private Function StyleFuncHelper(wStyle As Long, wConst As Long) As Boolean

    If wStyle And wConst Then 'if wConst style is active
        StyleFuncHelper = CBool(wStyle And wConst) 'helper returns the style
        wStyle = wStyle - wConst 'delete the Style found from wStyle
    End If

End Function

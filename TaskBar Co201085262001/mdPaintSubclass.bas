Attribute VB_Name = "mdPaintSubclass"
Option Explicit

'=========================================================================
' Constants and variables
'=========================================================================

Private Const STR_OLD_PROC      As String = "PAINT_OLDPROC"
Private Const TIMER_ID          As Long = 1
Private Const TIMER_TIMEOUT     As Long = 1000

Public Const STR_SHORTTIME      As String = "Short Time"
Public Const STR_LONGDATE       As String = "Long Date"

Public UpdateRect           As RECT

'=========================================================================
' Functions
'=========================================================================

Public Function PaintSubclass(ByVal hWnd As Long, ByVal lTimeout As Long)
    Dim lOldProc        As Long
    
    On Error Resume Next
    lOldProc = GetWindowLong(hWnd, GWL_WNDPROC)
    SetProp hWnd, STR_OLD_PROC, lOldProc
    SetWindowLong hWnd, GWL_WNDPROC, AddressOf PaintWndProc
    If lTimeout > 0 Then
        SetTimer hWnd, TIMER_ID, lTimeout, 0
    End If
End Function

Public Function PaintUnsubclass(ByVal hWnd As Long)
    Dim lOldProc As Long
    
    On Error Resume Next
    lOldProc = GetProp(hWnd, STR_OLD_PROC)
    If lOldProc <> 0 Then
        SetWindowLong hWnd, GWL_WNDPROC, lOldProc
        RemoveProp hWnd, STR_OLD_PROC
    End If
    KillTimer hWnd, TIMER_ID
End Function

Private Function PaintWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static sPrevTime    As String
    Dim ps              As PAINTSTRUCT
    Dim lOldProc        As Long
    
    On Error Resume Next
    Select Case uMsg
    Case WM_PAINT
        If GetUpdateRect(hWnd, UpdateRect, 0) <> 0 Then
            BeginPaint hWnd, ps
            SendMessage hWnd, WM_KEYUP, 0, ByVal 0
            EndPaint hWnd, ps
        End If
        Exit Function
    Case WM_CANCELMODE
        SendMessage hWnd, WM_KEYDOWN, 0, ByVal 0
    Case WM_TIMER
        '--- repaint clock if necessary
        If sPrevTime <> Format(Now, STR_SHORTTIME) Then
            sPrevTime = Format(Now, STR_SHORTTIME)
            RefreshDC hWnd
        End If
    End Select
    lOldProc = GetProp(hWnd, STR_OLD_PROC)
    PaintWndProc = CallWindowProc(lOldProc, hWnd, uMsg, wParam, lParam)
End Function


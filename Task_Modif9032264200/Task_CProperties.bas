Attribute VB_Name = "Task_CProperties"

Option Explicit
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal aBOOL As Long) As Long

Private Const EM_SETPASSWORDCHAR As Long = &HCC
Private Const EM_GETPASSWORDCHAR As Long = &HD2
Private Const WM_SETTEXT As Long = &HC
Private Const WM_GETTEXT As Long = &HD
Private Const WM_GETTEXTLENGTH As Long = &HE

'sets the window to a new parent
Public Function AssignParent(SourceHwnd As Long, DestHwnd As Long) As Long

    If DestHwnd <> SourceHwnd Then
        SetParent SourceHwnd, DestHwnd
    End If
    AssignParent = GetParent(SourceHwnd)

End Function

'retrieves the boundries of an object
Public Sub GetControlRect(hwnd, ByRef rTop As Long, ByRef rLeft As Long, ByRef rWidth As Long, ByRef rHeight As Long)

  Dim DaRect As RECT

    GetWindowRect hwnd, DaRect
    rTop = DaRect.Top
    rLeft = DaRect.Left
    rWidth = DaRect.Right - DaRect.Left
    rHeight = DaRect.Bottom - DaRect.Top

End Sub

'Grabs the password character being used by a textbox
Public Function GetPassWordChar(hwnd As Long) As Byte

    GetPassWordChar = SendMessageLong(hwnd, EM_GETPASSWORDCHAR, 0, 0)

End Function

'gets the text of a textbox or a caption of most controls
Public Function GetText(WindowHandle As Long) As String

  Dim Buffer As String, TextLength As Long

    TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    'Just In Case Of OverFlow in textbox
    If TextLength& > 32565 Then
        TextLength& = 32565
    End If
    Buffer$ = String$(TextLength&, 0&)
    SendMessageStr WindowHandle&, WM_GETTEXT, TextLength& + 1, Buffer$
    GetText$ = Buffer$

End Function

'returns the state of the window :normal,maximized,minimized
Public Function GetWindowState(hwnd As Long) As Long

  'Finds The Windowstate
  
  Dim h As Long

    h = 2 'Normal
    If IsIconic(hwnd) Then 'Minimized
        h = 0
    End If
    If IsZoomed(hwnd) Then 'Maximized
        h = 1
    End If
    GetWindowState = h

End Function

'this sub moves an object
Public Sub MoveControl(hwnd As Long, Top As Long, Left As Long, Width As Long, Height As Long)

  'Note: When you Resize a child control you have to remember its not using Screen
  '       Coordinates. Its using its parents borders a left/top boundries.
  '       So a form will use the desktop. Controls will use the form/frame/etc
  '       as its container.

    On Local Error Resume Next
      MoveWindow hwnd, Left, Top, Width, Height, True
    On Local Error GoTo 0

End Sub

'Enables/Disables an object
Public Function SetControlEnabled(hwnd As Long, DoEnable As Boolean)

    EnableWindow hwnd, DoEnable

End Function

'changes the password character that a textbow is using
Public Function SetNewPwChar(hwnd As Long, NewChar As Byte) As Byte

  'Sets and then attempts to retrieve the new Pwchar.

    SendMessageLong hwnd, EM_SETPASSWORDCHAR, CLng(NewChar), 0
    SetNewPwChar = CByte(SendMessageLong(hwnd, EM_GETPASSWORDCHAR, 0, 0))

End Function

'sets the text of a textbox or a caption of most controls
Public Function SetText(hwnd As Long, Msg As String) As String

    SendMessageStr hwnd, WM_SETTEXT, 0, Msg
    SetText = GetText(hwnd)

End Function

'sets the state of an object :Normal,minimized,maximized
Public Function SetWindowState(hwnd As Long, NewState As Long) As Long   'Returns NewState

    ShowWindow hwnd, NewState
    SetWindowState = GetWindowState(hwnd)

End Function


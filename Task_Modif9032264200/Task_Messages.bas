Attribute VB_Name = "Task_StopChanges"
'A Simple Way To Stop Caption Changing For Your Form/Control
'Also a Way to stop sizing

'Made By Billy Conner in 1999

'The nice thing about this method is that your standard Vb code Object.Caption=""
'will change the caption. but calling from SendMessageA api function using "WM_SETTEXT"
'will have no effect on the caption.

Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const WM_DESTROY As Long = &H2
Private Const WM_SIZE As Long = &H5
Private Const WM_SETTEXT As Long = &HC
Private Const GWL_WNDPROC As Long = (-4)
Private Const MF_BYPOSITION As Long = &H400&
Public Current_Read As Long

Public Sub AllowCaptionChange(Hwnd As Long, Flag As Boolean)

  Dim My_proc As Long

    If Flag = False Then
        My_proc = GetAddressOf(AddressOf MyWndProc)
        Current_Read = SetWindowLong(Hwnd, GWL_WNDPROC, My_proc)
      Else
        SetWindowLong Hwnd, GWL_WNDPROC, Current_Read
    End If

End Sub

'a simple addressof function
Private Function GetAddressOf(ProcAddress As Long) As Long

    GetAddressOf = ProcAddress

End Function

'a callback used to stop caption changing
Private Function MyWndProc(ByVal Hwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case message
      Case WM_SETTEXT
        Exit Function
    End Select
    ' calls the default window procedure
    MyWndProc = CallWindowProc(Current_Read, Hwnd, message, wParam, lParam)

End Function

Public Sub NoSizing(Hwnd As Long)

  'The benefit of this sub is to DELETE the "Size" out of the system Menu list.
  'Even if you start ur app with Fixed window(vb only hides the Size, API can make
  'it visible again).

    RemoveMenu GetSystemMenu(Hwnd, 0), 2, MF_BYPOSITION Or MF_REMOVE

End Sub

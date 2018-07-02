Attribute VB_Name = "ModMouse"
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32.dll" (ByVal x As Long, ByVal y As Long) As Long

Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10

Public Declare Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const BT_LEFT = &H30
Public Const BT_RIGHT = &H40

Public Type CURPOS
    x As Long                 ' Current Mouse position
    y As Long                 ' Current Mouse Position
    click_x As Long           ' X Coor where clicked
    click_y As Long           ' Y Coor Where clicked
    Button As Long            ' Right or left button
    dblClicked As Boolean     ' True if double clicked
    last_idx As Long          ' the last index reached in the array
End Type

Public cp() As CURPOS
Public LastMouseButton As String

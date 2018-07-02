Attribute VB_Name = "Windows"
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

'place this code in the MouseMove event of the form
Public Sub MoveForm(Form As Object, ButtonDown As Integer)
   Dim lngReturnValue As Long
   If ButtonDown = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Form.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub


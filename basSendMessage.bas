Attribute VB_Name = "basSendMessage"
Option Explicit

' SendMessage API functions.
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" (ByVal hwnd As Long, _
   ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private Const WM_SETREDRAW = &HB

'
' Turns redraw off or on for any object with an hwnd
' Handy for the resize event.
'
' Note: If .ClipControls is not False then .Refresh will
' not update things completely.
'
' If someone knows how to force a form to do a complete redraw
' when .ClipControls = True please let me know.
' tarheit@alpha.wcoil.com
'
Public Sub SetRedraw(ob As Object, ByVal b As Boolean)
   Call SendMessage(ob.hwnd, WM_SETREDRAW, IIf(b, 1, 0), 0)
End Sub



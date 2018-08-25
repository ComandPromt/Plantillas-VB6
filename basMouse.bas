Attribute VB_Name = "basMouse"
Option Explicit

'  Mouse/cursor functions.

Private lShowCursor As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'
'  Hides the mouse cursor.
'
Public Sub HideMouse()
   Dim result As Integer
   
   Do
      lShowCursor = lShowCursor - 1
      result = ShowCursor(False)
   Loop Until result < 0
   
End Sub

'
'  Restores the mouse cursor to it's previous state regardless
'  if HideMouse and ShowMouse were called.
'
Public Sub RestoreMouse()
   If lShowCursor > 0 Then
      Do While lShowCursor <> 0
         ShowCursor (False)
         lShowCursor = lShowCursor - 1
      Loop
   ElseIf lShowCursor < 0 Then
      Do While lShowCursor <> 0
         ShowCursor (True)
         lShowCursor = lShowCursor + 1
      Loop
   End If
End Sub


'
'  Show's the mouse cursor.
'
Public Sub ShowMouse()
   Dim result
   
   Do
      lShowCursor = lShowCursor - 1
      result = ShowCursor(True)
   Loop Until result >= 0

End Sub


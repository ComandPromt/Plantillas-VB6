Attribute VB_Name = "General"
Option Explicit

Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long

Global Const GWW_HWNDPARENT = (-8)



Sub Wait(WaitSeconds As Single)

Dim StartTime As Single

StartTime = Timer

Do While Timer < StartTime + WaitSeconds
DoEvents
Loop

End Sub

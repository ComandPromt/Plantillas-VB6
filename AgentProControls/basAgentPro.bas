Attribute VB_Name = "basAgentPro"
Option Explicit

Public MyControl As AgentPro

Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


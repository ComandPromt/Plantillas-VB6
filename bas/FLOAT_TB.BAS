Attribute VB_Name = "Module1"
Option Explicit

Type RECT
    left As Integer
    Top As Integer
    right As Integer
    bottom As Integer
End Type

Type POINTAPI
    x As Integer
    y As Integer
End Type

#If Win32 Then
  Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
#Else
  Declare Function GetDC Lib "User" (ByVal hWnd As Integer) As Integer
#End If
#If Win32 Then
  Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
#Else
  Declare Function ReleaseDC Lib "User" (ByVal hWnd As Integer, ByVal hDC As Integer) As Integer
#End If
#If Win32 Then
'TODO: No Win32 API known for subroutine DRAWFOCUSRECT.  Convert to different API subroutine.
#Else
  Declare Sub DrawFocusRect Lib "User" (ByVal hDC As Integer, lpRect As RECT)
#End If
#If Win32 Then
  Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
#Else
  Declare Function GetSysColor Lib "User" (ByVal nIndex As Integer) As Long
#End If
#If Win32 Then
  Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
#Else
  Declare Sub SetWindowWord Lib "User" (ByVal hWnd%, ByVal nCmd%, ByVal nVal%)
#End If

Global Const SWW_HPARENT = -8
Global Const COLOR_ACTIVECAPTION = 2
Global Const COLOR_INACTIVECAPTION = 3

'Global variable to indicate if floating toolbar is visible
Global toolbarVisible As Integer

'These variables save the toolbar location while unloaded
Global toolbarX As Integer
Global toolbarY As Integer

'Indicates which tool is currently selected
Global currTool As Integer


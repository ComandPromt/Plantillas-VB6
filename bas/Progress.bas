Attribute VB_Name = "Module1"
Option Explicit

' Brought to you by:
'   Brad Martinez
'   btmtz@aol.com
'   http://members.aol.com/btmtz/vb

' If the function succeeds, the return value is the
' previous value of the specified 32-bit integer.
' If the function fails, the return value is zero.
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                            (ByVal hwnd As Long, ByVal nIndex As Long) As Long

' If the function succeeds, the return value is the
' previous value of the specified 32-bit integer.
' If the function fails, the return value is zero.
' The SetWindowLong function fails if the window
' specified by the hWnd parameter does not belong
' to the same process as the calling thread.
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                              (ByVal hwnd As Long, ByVal nIndex As Long, _
                              ByVal dwNewLong As Long) As Long

' GetWindowLong(), SetWindowLong() nIndex param
Public Const GWL_STYLE = (-16)

' Restricts input to the edit control to digits only
Public Const ES_NUMBER = &H2000

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Declare Function MoveWindow Lib "user32" _
                              (ByVal hwnd As Long, _
                              ByVal x As Long, ByVal y As Long, _
                              ByVal nWidth As Long, ByVal nHeight As Long, _
                              ByVal bRepaint As Long) As Long

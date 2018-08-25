Attribute VB_Name = "MProgressBarDefs"
Option Explicit

' Brought to you by:
'   Brad Martinez
'   btmtz@aol.com
'   http:' //members.aol.com/btmtz/vb

' //====== PROGRESS BAR CONTROL ====================================

Public Const PROGRESS_CLASS = "msctls_progress32"

Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
                          (ByVal dwExStyle As Long, ByVal lpClassName As String, _
                           ByVal lpWindowName As String, ByVal dwStyle As Long, _
                           ByVal x As Long, ByVal y As Long, _
                           ByVal nWidth As Long, ByVal nHeight As Long, _
                           ByVal hWndParent As Long, ByVal hMenu As Long, _
                           ByVal hInstance As Long, lpParam As Any) As Long
' Styles
Public Const PBS_SMOOTH = &H1    ' IE3 and later
Public Const PBS_VERTICAL = &H4   ' IE3 and later
Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000

Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                          (ByVal hwnd As Long, ByVal wMsg As Long, _
                          wParam As Any, lParam As Any) As Long

' =============================================
' Messages

Public Const WM_USER = &H400

' The PBM_SETRANGE message sets the minimum and maximum values
' for a progress bar and redraws the bar to reflect the new range.
' wParam = 0
' lParam =nMinRange Or (nMaxRange * &H10000)
' nMinRange = Minimum range value specified in the lParam low word.
'                     Default value is zero.
' nMaxRange = Maximum range value specified in the lParam high word.
'                      Default value is 100.
' Returns the previous range values if successful, or zero otherwise. The
' low-order word specifies the previous minimum value, and the high-order
' word specifies the previous maximum value.
Public Const PBM_SETRANGE = (WM_USER + 1)

' The PBM_SETPOS message sets the current position for a progress bar
' and redraws the bar to reflect the new position.
' wParam = New position.
' lParam = 0
' Returns the previous position.
Public Const PBM_SETPOS = (WM_USER + 2)

 'The PBM_DELTAPOS message advances the current position of a progress
' bar by a specified increment and redraws the bar to reflect the new position.
' wParam = Amount to advance the position.
' lParam = 0
' Returns the previous position.
Public Const PBM_DELTAPOS = (WM_USER + 3)

' The PBM_SETSTEP message specifies the step increment for a progress bar.
' The step increment is the amount by which the progress bar increases its
' current position whenever it receives a PBM_STEPIT message. By default, the
' step increment is set to 10.
' wParam = New step increment.
' lParam = 0
' Returns the previous step increment.
Public Const PBM_SETSTEP = (WM_USER + 4)

' The PBM_STEPIT message advances the current position for a progress bar by
' the step increment and redraws the bar to reflect the new position. An application
' sets the step increment by sending the PBM_SETSTEP message.
' wParam = 0
' lParam = 0
' Returns the previous position.
' When the position exceeds the maximum range value, this message resets the
' current position so that the progress indicator starts over again from the beginning.
Public Const PBM_STEPIT = (WM_USER + 5)

'====================================
' IE3 and later...

' The PBM_SETRANGE32 message sets the range of a progress bar control to
' a 32-bit value.
' wParam = A 32-bit value that represents the low limit to be set for the progress
'                 bar control.
' lParam = A 32-bit value that represents the high limit to be set for the progress
'               bar control.
' Returns a DWORD that holds the previous 16-bit low limit in its low word, and
' the previous 16-bit high limit in its high word. If the previous ranges were 32-bit
' values, the return value consists of the low words of both 32-bit limits.
' To retrieve the entire high and low 32-bit values, use the PBM_GETRANGE
' message.
Public Const PBM_SETRANGE32 = (WM_USER + 6)

' The PBM_GETRANGE message retrieves information about the current high
' and low limits of a given progress bar control.
' wParam = Flag value specifying which limit value is to be used as the message's
'                 return value.This parameter can be one of the following values.
'                 Value       Meaning
'                 TRUE       Return the low limit.
'                 FALSE     Return the high limit.
' lParam = Address of a PBRANGE structure that is to be filled with the high and
'               low limits of the progress bar control. If this parameter is set to NULL,
'               the control will return only the limit specified by wParam
' Returns an INT that represents the limit value specified by wParam. If
' lParam is not NULL, lParam must point to a PBRANGE structure that is to
' be filled with both limit values.
Public Const PBM_GETRANGE = (WM_USER + 7)

Type PPBRANGE
   iLow As Integer
   iHigh As Integer
End Type

' The PBM_GETPOS message retrieves the current position of the progress bar.
' wParam = 0
' lParam = 0
' Returns a UINT that represents the current position of the progress bar.
Public Const PBM_GETPOS = (WM_USER + 8)


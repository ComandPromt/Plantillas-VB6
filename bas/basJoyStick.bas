Attribute VB_Name = "basJoyStick"
Option Explicit

' Unfortunately VB 4.0 doesn't support capturing messages
' (and there's no way to trick it as with the TrayIcon)
' So we are left with poling which still can work quite
' well if written correctly.

' Public defines and structures
Public Const JOY_BUTTON1 = &H1
Public Const JOY_BUTTON2 = &H2
Public Const JOY_BUTTON3 = &H4
Public Const JOY_BUTTON4 = &H8

Public Type JOYINFO
   x As Long
   Y As Long
   Z As Long
   Buttons As Long
End Type


' Private defs
Private Const JOYERR_BASE = 160
Private Const JOYERR_NOERROR = (0)
Private Const JOYERR_NOCANDO = (JOYERR_BASE + 6)
Private Const JOYERR_PARMS = (JOYERR_BASE + 5)
Private Const JOYERR_UNPLUGGED = (JOYERR_BASE + 7)

Private Const MAXPNAMELEN = 32

Private Type JOYCAPS
   wMid As Integer
   wPid As Integer
   szPname As String * MAXPNAMELEN
   wXmin As Long
   wXmax As Long
   wYmin As Long
   wYmax As Long
   wZmin As Long
   wZmax As Long
   wNumButtons As Long
   wPeriodMin As Long
   wPeriodMax As Long
 End Type

Private Declare Function joyGetDevCaps Lib "winmm.dll" _
   Alias "joyGetDevCapsA" (ByVal id As Long, _
   lpCaps As JOYCAPS, ByVal uSize As Long) As Long
   
Private Declare Function joyGetNumDevs Lib "winmm.dll" _
   () As Long
   
Private Declare Function joyGetPos Lib "winmm.dll" _
   (ByVal uJoyID As Long, pji As JOYINFO) As Long
'
'  Fills the ji structure with the minimum x, y, and z
'  coordinates.  Buttons is filled with the number of
'  buttons.
'
Public Function GetJoyMin(ByVal joy As Integer, ji As JOYINFO) As Boolean
   Dim jc As JOYCAPS
   
   If joyGetDevCaps(joy, jc, Len(jc)) <> JOYERR_NOERROR Then
      GetJoyMin = False
      
   Else
      ji.x = jc.wXmin
      ji.Y = jc.wYmin
      ji.Z = jc.wZmin
      ji.Buttons = jc.wNumButtons
   
      GetJoyMin = True
   End If
End Function
'
'  Fills the ji structure with the maximum x, y, and z
'  coordinates.  Buttons is filled with the number of
'  buttons.
'
Public Function GetJoyMax(ByVal joy As Integer, ji As JOYINFO) As Boolean
   Dim jc As JOYCAPS
   
   If joyGetDevCaps(joy, jc, Len(jc)) <> JOYERR_NOERROR Then
      GetJoyMax = False
      
   Else
      ji.x = jc.wXmax
      ji.Y = jc.wYmax
      ji.Z = jc.wZmax
      ji.Buttons = jc.wNumButtons
   
      GetJoyMax = True
   End If
End Function
Public Function GetJoystick(ByVal joy As Integer, ji As JOYINFO) As Boolean
   If joyGetPos(joy, ji) <> JOYERR_NOERROR Then
      GetJoystick = False
   Else
      GetJoystick = True
   End If
End Function

'
'  If IsConnected is False then it returns the number of
'  joysticks the driver supports. (But may not be connected)
'
'  If IsConnected is True the it returns the number of
'  joysticks present and connected.
'
'  IsConnected is true by default.
'
Public Function IsJoyPresent(Optional IsConnected As Variant) As Long
   Dim ic As Boolean
   Dim i As Long
   Dim j As Long
   Dim ret As Long
   Dim ji As JOYINFO
   
   ic = IIf(IsMissing(IsConnected), True, CBool(IsConnected))

   i = joyGetNumDevs
   
   If ic Then
      j = 0
      Do While i > 0
         i = i - 1   'Joysticks id's are 0 and 1
         If joyGetPos(i, ji) = JOYERR_NOERROR Then
            j = j + 1
         End If
      Loop
   
      IsJoyPresent = j
   Else
      IsJoyPresent = i
   End If
   
End Function

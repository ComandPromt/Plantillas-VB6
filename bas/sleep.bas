Attribute VB_Name = "Wait"
'DECLARATIONS FOR WinTime.cls
Private Declare Function GetTickCount& Lib "kernel32" ()

Private Function WinTime() As String
Dim lngReturn As Long
Dim tmp, tmpTime, tmpHours, tmpMinutes, tmpSeconds, tmpMilliSeconds As String
lngReturn = GetTickCount()
WinTime = lngReturn
End Function

Public Function Sleep(Time As Integer, Optional Freeze As Boolean) As Boolean
Dim StartSleepTime As Long
Dim StopSleepTime As Long
StopSleepTime = WinTime + Time

Do
If Freeze = False Then DoEvents
If WinTime >= StopSleepTime Then GoTo 10
Loop

10 Sleep = True
End Function

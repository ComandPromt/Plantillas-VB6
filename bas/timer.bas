'DECLARATIONS FOR WinTime.cls
Private Declare Function GetTickCount& Lib "kernel32" ()
Dim StartTime As Currency
Dim TimerOn As Boolean
Dim EndTime As Currency

Private Function WinTime() As String
Dim lngReturn As Long
Dim tmp, tmpTime, tmpHours, tmpMinutes, tmpSeconds, tmpMilliSeconds As String
lngReturn = GetTickCount()
WinTime = lngReturn
End Function


Public Sub StartTimer()
StartTime = WinTime
TimerOn = True
EndTime = 0

End Sub

Public Sub StopTimer()
EndTime = TimerTime
TimerOn = False
End Sub

Public Sub ResetTimer()
EndTime = 0
StartTime = WinTime

End Sub

Public Property Get TimerActive() As Boolean
If TimerOn = True Then TimerActive = True Else TimerActive = False
End Property

Public Property Get TimerTime() As Currency
If TimerOn = False Then TimerTime = EndTime: Exit Property
TimerTime = WinTime - StartTime
End Property


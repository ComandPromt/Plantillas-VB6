Private Declare Function GetTickCount& Lib "kernel32" ()

Private Function WinTime() As String
Dim lngReturn As Long
Dim tmp, tmpTime, tmpHours, tmpMinutes, tmpSeconds, tmpMilliSeconds As String
lngReturn = GetTickCount()
WinTime = lngReturn
End Function
Private Function Sleep(Time As Integer, Optional Freeze As Boolean) As Boolean
Dim StartSleepTime As Long
Dim StopSleepTime As Long
StopSleepTime = WinTime + Time
Do
If Freeze = False Then DoEvents
If WinTime >= StopSleepTime Then GoTo 10
Loop
10 Sleep = True
End Function

Public Sub TypeText(TextBox As Object, Text As String, Optional Speed As Integer = 100)
Dim X As Long
Dim TxtLen As Long

X = 0

On Error Resume Next

Do
DoEvents
X = X + 1
TextBox = Mid(Text, 1, X)
TextBox.SelStart = Len(TextBox)
Call Sleep(Speed)
If X = Len(Text) Then GoTo 10
Loop

10
End Sub

Public Sub TypeTextAndDelete(TextBox As Object, Text As String, Optional Speed As Integer = 100, Optional PauseTime As Integer = 100, Optional ComeBackSpeed As Integer)
Dim X As Long
Dim TxtLen As Long

X = 0
On Error Resume Next

ComeBackSpeed = Speed
If ComeBackSpeed = 0 Then ComeBackSpeed = Speed / 2



Do
DoEvents
X = X + 1
TextBox = Mid(Text, 1, X)
TextBox.SelStart = Len(TextBox)
Call Sleep(Speed)
If X = Len(Text) Then GoTo 10
Loop

10

Call Sleep(PauseTime)

X = X + 1

Do
DoEvents
X = X - 1
TextBox = Mid(Text, 1, X)
TextBox.SelStart = Len(TextBox)
Call Sleep(ComeBackSpeed)
If X = 0 Then GoTo 20
Loop

20
End Sub



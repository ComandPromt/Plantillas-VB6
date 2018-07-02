Attribute VB_Name = "ModMain"
Declare Function GetTickCount Lib "kernel32" () As Long

Public Const DefaultIPAddress = "127.0.0.1"
Public Const DefaultPort = "10666"
Public Const DefaultTimerInterval = 10

Public Const DefaultScreenShotName = "temp.jpg"

Public XModifier
Public YModifier

Public Function GetVersion()
    GetVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Sub UpdateStatus(Status)
    FRMMain.Status.Text = FRMMain.Status.Text + Status + Chr$(13) + Chr$(10)
    FRMMain.Status.SelStart = Len(FRMMain.Status.Text)
End Sub

Sub Pause(HowLong As Long)
    Dim u%, tick As Long
    tick = GetTickCount()
    Do
        u% = DoEvents
    Loop Until tick + HowLong < GetTickCount
End Sub

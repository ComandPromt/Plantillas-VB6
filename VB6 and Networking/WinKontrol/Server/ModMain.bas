Attribute VB_Name = "ModMain"
Declare Function GetTickCount Lib "kernel32" () As Long

Public Const DefaultScreenShotName = "temp.jpg"
Public Const DefaultScreenShotSize = "20"
Public Const DefaultScreenShotBlockSize = 4169





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

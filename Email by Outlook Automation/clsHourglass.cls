VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHourglass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlOldMouse As Long


Private Sub Class_Initialize()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : Class_Initialize
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for Class_Initialize and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________

    On Error GoTo Proc_Err
    Const csProcName As String = "Class_Initialize"


    
    mlOldMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "clsHourglass->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Sub
 
End Sub


Private Sub Class_Terminate()

    '__________________________________________________
    ' Scope  : Private
    ' Type   : Sub
    ' Name   : Class_Terminate
    ' Params :
    ' Returns: Nothing
    ' Desc   : The Sub uses parameters  for Class_Terminate and returns Nothing.
    '__________________________________________________
    ' History
    ' CDK: 20010105: Added Error Trapping & Comments
    '__________________________________________________



    On Error Resume Next
    Screen.MousePointer = mlOldMouse

End Sub



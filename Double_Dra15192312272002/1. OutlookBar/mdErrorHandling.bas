Attribute VB_Name = "mdErrorHandling"
'=========================================================================
'
'   You are free to use this source as long as this copyright message
'     appears on your program's "About" dialog:
'
'   Outlook Bar Project
'   Copyright (c) 2002 Vlad Vissoultchev (wqweto@myrealbox.com)
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "mdErrorHandling"

'=========================================================================
' API
'=========================================================================

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'=========================================================================
' Constants and variables
'=========================================================================

Public g_lErrNumber             As Long
Public g_sErrDescription        As String
Public g_sErrSource             As String

'=========================================================================
' Error handling
'=========================================================================

'Private Sub RaiseError(sFunc As String)
'    PushError sFunc, MODULE_NAME
'    PopRaiseError
'End Sub
'
'Private Function ShowError(sFunc As String) As VbMsgBoxResult
'    PushError sFunc, MODULE_NAME
'    ShowError = PopShowError(CAP_MSG)
'End Function

'=========================================================================
' Functions
'=========================================================================

Public Sub PushError(Optional sFunc As String, Optional sModule As String)
    g_lErrNumber = Err.Number
    g_sErrDescription = Err.Description
    g_sErrSource = IIf(Len(sModule) > 0, _
                            "[\\" & ErrComputerName() & "] " & _
                            LIB_NAME & "." & _
                            sModule & "." & _
                            sFunc & _
                            IIf(Erl <> 0, "(" & Erl & ")", "") & vbCrLf, _
                            "") & _
                        Err.Source
End Sub

Public Sub PopError()
    Err.Number = g_lErrNumber
    Err.Description = g_sErrDescription
    Err.Source = g_sErrSource
End Sub

Public Sub PopRaiseError()
    PopError
    Err.Raise Err.Number
End Sub

Public Function PopShowError(sCaption As String) As VbMsgBoxResult
    PopShowError = MsgBox( _
            g_sErrDescription & vbCrLf & vbCrLf & _
            "Error: 0x" & Hex(g_lErrNumber) & vbCrLf & vbCrLf & _
            "Call stack:" & vbCrLf & _
            g_sErrSource, vbCritical Or vbAbortRetryIgnore, sCaption)
End Function

Public Function ErrComputerName() As String
    Static sName        As String
        
    If Len(sName) = 0 Then
        sName = String(256, 0)
        GetComputerName sName, Len(sName)
        sName = Left$(sName, InStr(sName, Chr(0)) - 1)
    End If
    ErrComputerName = sName
End Function

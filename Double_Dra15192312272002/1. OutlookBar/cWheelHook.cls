VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWheelHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Private Const MODULE_NAME As String = "cWheelHook"
Implements IHookingSink

'==============================================================================
' API
'==============================================================================

Private Const WM_MOUSEWHEEL             As Long = &H20A

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x                   As Long
    y                   As Long
End Type

'==============================================================================
' Constants and member variables
'==============================================================================

Private m_oHook             As cHookingThunk

'==============================================================================
' Base class events
'==============================================================================

Private Sub Class_Initialize()
    '--- install hook
    Set m_oHook = New cHookingThunk
    Call m_oHook.Hook(WH_GETMESSAGE, Me)
End Sub

'==============================================================================
' IHookingSink interface
'==============================================================================

Private Sub IHookingSink_Before(bHandled As Boolean, lReturn As Long, nCode As SubclassingSink.HookCode, wParam As Long, lParam As Long)
    Dim ms              As MSG
    Dim pt              As POINTAPI
    Dim hFwdWnd         As Long
    
    If nCode = HC_ACTION Then
        ms = m_oHook.MSG(lParam)
        If ms.message = WM_MOUSEWHEEL Then
            Call GetCursorPos(pt)
            hFwdWnd = WindowFromPoint(pt.x, pt.y)
            If hFwdWnd <> 0 Then
                Call PostMessage(hFwdWnd, MouseWheelFwdMsg, ms.wParam, ms.lParam)
            End If
        End If
    End If
End Sub

Private Sub IHookingSink_After(lReturn As Long, ByVal nCode As SubclassingSink.HookCode, ByVal wParam As Long, ByVal lParam As Long)

End Sub


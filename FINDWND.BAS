'---------------------------------------------------------
' This module contains Windows APIs delcares and functions
' utilize the API.
'---------------------------------------------------------
Option Explicit

'----------------------------------------------------
' Windows API Function Prototypes
'----------------------------------------------------
Declare Sub BringWindowToTop Lib "User" (ByVal hWnd)
Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function GetLastActivePopup Lib "User" (ByVal hwndOwnder) As Integer
Declare Function GetModuleHandle Lib "Kernel" (ByVal lpProgramName$) As Integer
Declare Function GetNextWindow% Lib "User" (ByVal hWnd, ByVal wFlag)
Declare Function GetWindow Lib "User" (ByVal hWnd, ByVal wCmd) As Integer
Declare Function GetWindowText Lib "User" (ByVal hWnd, ByVal lpString$, ByVal aint) As Integer
Declare Function IsIconic Lib "User" (ByVal hWnd) As Integer
Declare Function ShowWindow% Lib "User" (ByVal hWnd%, ByVal nCmdShow%)

'--------------------------------------------------------
' 3D Control Declarations
'--------------------------------------------------------
Declare Function Ctl3dRegister Lib "CTL3D.DLL" (ByVal hInstance As Integer) As Integer
Declare Function Ctl3dAutoSubClass Lib "CTL3D.DLL" (ByVal hInstance As Integer) As Integer
Declare Function Ctl3dUnRegister Lib "CTL3D.DLL" (ByVal hInstance As Integer) As Integer

' ShowWindow() Commands
Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_NORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_MAXIMIZE = 3
Global Const SW_SHOWNOACTIVATE = 4
Global Const SW_SHOW = 5
Global Const SW_MINIMIZE = 6
Global Const SW_SHOWMINNOACTIVE = 7
Global Const SW_SHOWNA = 8
Global Const SW_RESTORE = 9

' GetWindow() Constants
Global Const GW_HWNDFIRST = 0
Global Const GW_HWNDLAST = 1
Global Const GW_HWNDNEXT = 2
Global Const GW_HWNDPREV = 3
Global Const GW_OWNER = 4
Global Const GW_CHILD = 5

'--------------------------------------------------------
' Looks for a Window based on the WinName$ string.
' ExcludeWin$ is used to exclude a top-level invisible
' window that VB creates for the application
'--------------------------------------------------------
Function FindWindowByName (WinName$, ExcludeWin$) As Integer
    Dim xChars As Integer
    Dim Wnd As Integer, PopWnd  As Integer
    Dim Source$
    Dim WinCaption$
    
    WinCaption$ = String$(256, 0)
    WinName$ = UCase$(WinName$)
    
    Wnd = FindWindow(ByVal 0&, ByVal 0&)
    Wnd = GetWindow(Wnd, GW_HWNDFIRST)
    While Wnd <> 0
	xChars = GetWindowText(Wnd, WinCaption$, 256)
	If xChars > 0 Then
	    Source$ = UCase$(Left$(WinCaption$, xChars))
	    If InStr(Source$, WinName$) > 0 Then
		If ExcludeWin$ = "" Then
		    xChars = 0
		Else
		    xChars = InStr(Source$, ExcludeWin$)
		End If
		If xChars = 0 Then
		    PopWnd = GetLastActivePopup(Wnd)
		    If PopWnd Then
			Wnd = PopWnd
		    End If
		    FindWindowByName = Wnd
		    Exit Function
		End If
	    End If
	End If
	Wnd = GetNextWindow(Wnd, GW_HWNDNEXT)
    Wend
    FindWindowByName = 0
End Function


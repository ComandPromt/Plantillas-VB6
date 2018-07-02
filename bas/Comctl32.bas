Attribute VB_Name = "MInitComctl"
Option Explicit

' Brought to you by:
'   Brad Martinez
'   btmtz@aol.com
'   http:' //members.aol.com/btmtz/vb

' Ensures that the common control dynamic-link library (DLL) is loaded.
Declare Sub InitCommonControls Lib "comctl32.dll" ()

' IE3 & later
' Returns TRUE (non-zero) if successful, or FALSE otherwise.
Declare Function InitCommonControlsEx Lib "comctl32.dll" _
                            (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Type tagINITCOMMONCONTROLSEX   ' icc
    dwSize As Long    ' // size of this structure
    dwICC As Long    ' // flags indicating which classes to be initialized
End Type

' Set of bit flags that indicate which common control classes will be
' loaded from the DLL. This value can be a combination of the following:
Public Const ICC_LISTVIEW_CLASSES = &H1    ' // listview, header
Public Const ICC_TREEVIEW_CLASSES = &H2  ' // treeview, tooltips
Public Const ICC_BAR_CLASSES = &H4            ' // toolbar, statusbar, trackbar, tooltips
Public Const ICC_TAB_CLASSES = &H8             ' // tab, tooltips
Public Const ICC_UPDOWN_CLASS = &H10       ' // updown
Public Const ICC_PROGRESS_CLASS = &H20   ' // progress
Public Const ICC_HOTKEY_CLASS = &H40         ' // hotkey
Public Const ICC_ANIMATE_CLASS = &H80        ' // animate
Public Const ICC_WIN95_CLASSES = &HFF        ' loads everything above
Public Const ICC_DATE_CLASSES = &H100        ' // month picker, date picker, time picker, updown
Public Const ICC_USEREX_CLASSES = &H200   ' // comboex
Public Const ICC_COOL_CLASSES = &H400       ' // rebar (coolbar) control
 

Public Function IsNewComctl32() As Boolean
' Rtns True if the current working version of Comctl32.dll
' supports the new IE3 sytles & msgs. Rtns False if old version.
' Also ensures that the Comctl32.dll library is loaded for use.
' This hack is much easier than checking the file version...
' VB4 (& VB5 ?) resolves API function names only
' when they're called, not when it compiles code!

  Dim icc As tagINITCOMMONCONTROLSEX
  On Error GoTo OldVersion
  
  icc.dwSize = Len(icc)
  icc.dwICC = ICC_PROGRESS_CLASS
  
  ' VB will generate error 453 "Specified DLL function not found"
  ' here if the new version isn't installed.
  IsNewComctl32 = InitCommonControlsEx(icc)
  Exit Function

OldVersion:
  InitCommonControls

End Function
 


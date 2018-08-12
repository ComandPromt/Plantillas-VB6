Attribute VB_Name = "mdPPGlobal"
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
Private Const MODULE_NAME As String = "mdPPGlobal"

'=========================================================================
' Constants and variables
'=========================================================================

Public Const LIB_NAME               As String = "OBPropPages"
Public Const LNG_UNDEFINED          As Long = &H80000000

'=========================================================================
' Functions
'=========================================================================

Public Sub CopyFont(ByVal oDest As StdFont, ByVal oSrc As StdFont)
    With oDest
        .Bold = oSrc.Bold
        .Charset = oSrc.Charset
        .Italic = oSrc.Italic
        .Name = oSrc.Name
        .Size = oSrc.Size
        .Strikethrough = oSrc.Strikethrough
        .Underline = oSrc.Underline
        .Weight = oSrc.Weight
    End With
End Sub

Public Function C2Lng(v) As Long
    On Error Resume Next
    C2Lng = CLng(v)
End Function

Public Function C2Str(v) As String
    On Error Resume Next
    C2Str = CStr(v)
End Function



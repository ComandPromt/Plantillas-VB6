Attribute VB_Name = "mdGlobals"
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
Private Const MODULE_NAME As String = "mdGlobals"

'=========================================================================
' API
'=========================================================================

Private Const WM_MOUSEWHEEL             As Long = &H20A

Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Type ICONINFO
    fIcon               As Long
    xHotspot            As Long
    yHotspot            As Long
    hbmMask             As Long
    hbmColor            As Long
End Type

'=========================================================================
' Constants and variables
'=========================================================================

Public Const LIB_NAME               As String = "OutlookBar"
Public Const PROPPAGES_LIB_NAME     As String = "OBPropPages"
Public Const LNG_UNDEFINED          As Long = &H80000000

Private m_lMouseWheelFwdMsg     As Long
Public g_oWheelHook             As cWheelHook
#If DebugMode Then
    Public g_lObjCount          As Long
#End If

'=========================================================================
' Error handling
'=========================================================================

Private Sub RaiseError(sFunc As String)
    PushError sFunc, MODULE_NAME
    PopRaiseError
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get MouseWheelFwdMsg() As Long
    If m_lMouseWheelFwdMsg = 0 Then
        m_lMouseWheelFwdMsg = RegisterWindowMessage("MouseWheelFwdMsg")
    End If
    MouseWheelFwdMsg = m_lMouseWheelFwdMsg
'    MouseWheelFwdMsg = WM_MOUSEWHEEL
End Property

'=========================================================================
' Functions
'=========================================================================

Public Function CloneFont(ByVal oSrc As StdFont) As StdFont
    Dim oFont As IFont
    
    Set oFont = oSrc
    oFont.Clone CloneFont
End Function

Public Function C2Lng(v) As Long
    On Error Resume Next
    C2Lng = CLng(v)
End Function

Public Function C2Str(v) As String
    On Error Resume Next
    C2Str = CStr(v)
End Function

Public Sub WritePictureProperty( _
            oBag As PropertyBag, _
            sPropName As String, _
            oPic As StdPicture, _
            Optional DefaultValue As Variant)
    Const FUNC_NAME     As String = "WritePictureProperty"
    Dim ii              As ICONINFO
    Dim hr              As Long
    Dim oMemDC          As cMemDC
    
    On Error GoTo EH
    If Not oPic Is Nothing Then
        If oPic.Type = vbPicTypeIcon Then
            Set oMemDC = New cMemDC
            hr = GetIconInfo(oPic.handle, ii)
            With New PropertyBag
                Call .WriteProperty("c", oMemDC.BitmapToPicture(ii.hbmColor))
                Call .WriteProperty("m", oMemDC.BitmapToPicture(ii.hbmMask))
                Call oBag.WriteProperty(sPropName, .Contents, DefaultValue)
            End With
            Exit Sub
        End If
    End If
    '--- else
    Call oBag.WriteProperty(sPropName, oPic, DefaultValue)
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Function ReadPictureProperty( _
            oBag As PropertyBag, _
            sPropName As String, _
            Optional DefaultValue As Variant) As StdPicture
    Const FUNC_NAME     As String = "ReadPictureProperty"
    Dim ii              As ICONINFO
    Dim hr              As Long
    Dim imgColor        As StdPicture
    Dim imgMask         As StdPicture
    
    On Error GoTo EH
    If IsArray(oBag.ReadProperty(sPropName, DefaultValue)) Then
        With New PropertyBag
            .Contents = oBag.ReadProperty(sPropName, DefaultValue)
            Set imgColor = .ReadProperty("c")
            Set imgMask = .ReadProperty("m")
        End With
        ii.fIcon = 1
        ii.hbmColor = imgColor.handle
        ii.hbmMask = imgMask.handle
        With New cMemDC
            Set ReadPictureProperty = .IconToPicture(CreateIconIndirect(ii))
        End With
    Else
        Set ReadPictureProperty = oBag.ReadProperty(sPropName, DefaultValue)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

#If DebugMode Then
    Public Sub DebugInit(sDebugID As String, sModule As String)
        g_lObjCount = g_lObjCount + 1
        sDebugID = g_lObjCount & " " & sModule & " " & Timer
    End Sub
    
    Public Sub DebugTerm(sDebugID As String)
        Debug.Print "DebugTerm: " & sDebugID
    End Sub
#End If



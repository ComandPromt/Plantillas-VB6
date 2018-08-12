VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFontDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private Const MODULE_NAME As String = "cFontDef"

'=========================================================================
' Events
'=========================================================================

'Purpose: Occurs when the contents of the current <b>cFontDef</b> have changed.
Event Changed()
Attribute Changed.VB_HelpID = 3003

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsTriState
    ucsTri_Undefined = -2       ' Undefined value. Usually value is inherited from parent definition.
    ucsTriFalse = 0             ' (Default) Represents False.
    ucsTriTrue = -1             ' Represents True.
End Enum

'=========================================================================
' Private constants and variables
'=========================================================================

Private Const DEF_NAME              As String = ""
Private Const DEF_SIZE              As Long = LNG_UNDEFINED
Private Const DEF_BOLD              As Long = ucsTri_Undefined
Private Const DEF_ITALIC            As Long = ucsTri_Undefined
Private Const DEF_UNDERLINE         As Long = ucsTri_Undefined
Private Const DEF_STRIKETROUGH      As Long = ucsTri_Undefined

Private m_sName                 As String
Private m_dblSize               As Double
Private m_eBold                 As UcsTriState
Private m_eItalic               As UcsTriState
Private m_eUnderline            As UcsTriState
Private m_eStrikethrough        As UcsTriState
#If DebugMode Then
    Private m_sDebugID          As String
#End If

'=========================================================================
' Error handling
'=========================================================================

Private Sub RaiseError(sFunc As String)
    PushError sFunc, MODULE_NAME
    PopRaiseError
End Sub

'=========================================================================
' Propeties
'=========================================================================

'Purpose: Returns or sets the name of a font used in the <b>cFontDef</b> object.
Property Get Name() As String
Attribute Name.VB_HelpID = 3008
    Name = m_sName
End Property

Property Let Name(sValue As String)
    m_sName = sValue
    RaiseEvent Changed
End Property

'Purpose: Returns or sets the font size used in the <b>cFontDef</b> object.
Property Get Size() As Double
Attribute Size.VB_HelpID = 3009
    Size = m_dblSize
End Property

Property Let Size(ByVal dblValue As Double)
    m_dblSize = dblValue
    RaiseEvent Changed
End Property

'Purpose: Returns or sets the font style of the <b>cFontDef</b> object to either bold, nonbold or undefined.
Property Get Bold() As UcsTriState
Attribute Bold.VB_HelpID = 3002
    Bold = m_eBold
End Property

Property Let Bold(ByVal eValue As UcsTriState)
    m_eBold = eValue
    RaiseEvent Changed
End Property

'Purpose: Returns or sets the font style of the <b>cFontDef</b> object to either italic, nonitalic or undefined.
Property Get Italic() As UcsTriState
Attribute Italic.VB_HelpID = 3007
    Italic = m_eItalic
End Property

Property Let Italic(ByVal eValue As UcsTriState)
    m_eItalic = eValue
    RaiseEvent Changed
End Property

'Purpose: Returns or sets the font style of the <b>cFontDef</b> object to either underlined, nonunderlined or undefined.
Property Get Underline() As UcsTriState
Attribute Underline.VB_HelpID = 3011
    Underline = m_eUnderline
End Property

Property Let Underline(ByVal eValue As UcsTriState)
    m_eUnderline = eValue
    RaiseEvent Changed
End Property

'Purpose: Returns or sets the font style of the <b>cFontDef</b> object to either strkithrough, nonstrkithrough or undefined.
Property Get Strikethrough() As UcsTriState
Attribute Strikethrough.VB_HelpID = 3010
    Strikethrough = m_eStrikethrough
End Property

Property Let Strikethrough(ByVal eValue As UcsTriState)
    m_eStrikethrough = eValue
    RaiseEvent Changed
End Property

'Purpose: Returns or sets the contents of the <b>cFontDef</b> as a serialized byte array.
Property Get Contents() As Byte()
Attribute Contents.VB_HelpID = 3004
    Const FUNC_NAME As String = "Contents [get]"
    
    On Error GoTo EH
    With New PropertyBag
        Call .WriteProperty("N", Name, DEF_NAME)
        Call .WriteProperty("S", Size, DEF_SIZE)
        Call .WriteProperty("B", Bold, DEF_BOLD)
        Call .WriteProperty("I", Italic, DEF_ITALIC)
        Call .WriteProperty("U", Underline, DEF_UNDERLINE)
        Call .WriteProperty("ST", Strikethrough, DEF_STRIKETROUGH)
        Contents = .Contents
    End With
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Property Let Contents(aValue() As Byte)
    Const FUNC_NAME As String = "Contents [let]"
    
    On Error GoTo EH
    With New PropertyBag
        .Contents = aValue
        Name = .ReadProperty("N", DEF_NAME)
        Size = .ReadProperty("S", DEF_SIZE)
        Bold = .ReadProperty("B", DEF_BOLD)
        Italic = .ReadProperty("I", DEF_ITALIC)
        Underline = .ReadProperty("U", DEF_UNDERLINE)
        Strikethrough = .ReadProperty("ST", DEF_STRIKETROUGH)
    End With
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

'Purpose: Returns a string that is describing the <b>cFontDef</b> object.
Property Get Description() As String
Attribute Description.VB_HelpID = 3005
    Const FUNC_NAME As String = "Description"
    
    On Error GoTo EH
    If Name <> "" Then
        Description = IIf(Len(Description) > 0, Description & ", ", "") & Name
    End If
    If Size <> LNG_UNDEFINED Then
        Description = IIf(Len(Description) > 0, Description & ", ", "") & Size & "pt"
    End If
    If Bold <> ucsTri_Undefined Then
        Description = IIf(Len(Description) > 0, Description & ", ", "") & _
                IIf(Bold = ucsTriFalse, "Not ", "") & "Bold"
    End If
    If Italic <> ucsTri_Undefined Then
        Description = IIf(Len(Description) > 0, Description & ", ", "") & _
                IIf(Italic = ucsTriFalse, "Not ", "") & "Italic"
    End If
    If Underline <> ucsTri_Undefined Then
        Description = IIf(Len(Description) > 0, Description & ", ", "") & _
                IIf(Underline = ucsTriFalse, "Not ", "") & "Underlined"
    End If
    If Strikethrough <> ucsTri_Undefined Then
        Description = IIf(Len(Description) > 0, Description & ", ", "") & _
                IIf(Strikethrough = ucsTriFalse, "Not ", "") & "Strikethrough"
    End If
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

'=========================================================================
' Methods
'=========================================================================

'Purpose: Applies current <b>cFontDef</b> to a <b>StdFont</b> object and returns a new <b>StdFont</b> object.
Public Function ApplyTo(ByVal oSrc As StdFont) As StdFont
Attribute ApplyTo.VB_HelpID = 3001
    Const FUNC_NAME     As String = "ApplyTo"
    
    On Error GoTo EH
    If Name = "" And Size = LNG_UNDEFINED And Bold = ucsTri_Undefined And Italic = ucsTri_Undefined And Underline = ucsTri_Undefined And Strikethrough = ucsTri_Undefined Then
        Set ApplyTo = oSrc
    Else
        Set ApplyTo = CloneFont(oSrc)
        With ApplyTo
            If Name <> "" Then
                .Name = Name
            End If
            If Size <> LNG_UNDEFINED Then
                .Size = Size
            End If
            If Bold <> ucsTri_Undefined Then
                .Bold = Bold
            End If
            If Italic <> ucsTri_Undefined Then
                .Italic = Italic
            End If
            If Underline <> ucsTri_Undefined Then
                .Underline = Underline
            End If
            If Strikethrough <> ucsTri_Undefined Then
                .Strikethrough = Strikethrough
            End If
        End With
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'Purpose: Finds the difference between two <b>StdFont</b> objects and stores them in the <b>cFontDef</b> object.
Public Sub ExtractDelta(ByVal oDest As StdFont, ByVal oSrc As StdFont)
Attribute ExtractDelta.VB_HelpID = 3006
    Const FUNC_NAME     As String = "ExtractDelta"
    
    On Error GoTo EH
    If Not oSrc Is Nothing Then
        Name = IIf(oSrc.Name <> oDest.Name, oDest.Name, "")
        Size = IIf(oSrc.Size <> oDest.Size, oDest.Size, LNG_UNDEFINED)
        Bold = IIf(oSrc.Bold <> oDest.Bold, oDest.Bold, ucsTri_Undefined)
        Italic = IIf(oSrc.Italic <> oDest.Italic, oDest.Italic, ucsTri_Undefined)
        Underline = IIf(oSrc.Underline <> oDest.Underline, oDest.Underline, ucsTri_Undefined)
        Strikethrough = IIf(oSrc.Strikethrough <> oDest.Strikethrough, oDest.Strikethrough, ucsTri_Undefined)
    Else
        Name = oDest.Name
        Size = oDest.Size
        Bold = IIf(oDest.Bold, ucsTriTrue, ucsTri_Undefined)
        Italic = IIf(oDest.Italic, ucsTriTrue, ucsTri_Undefined)
        Underline = IIf(oDest.Underline, ucsTriTrue, ucsTri_Undefined)
        Strikethrough = IIf(oDest.Strikethrough, ucsTriTrue, ucsTri_Undefined)
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'=========================================================================
' Class events
'=========================================================================

Private Sub Class_Initialize()
    On Error Resume Next
    Name = DEF_NAME
    Size = DEF_SIZE
    Bold = DEF_BOLD
    Italic = DEF_ITALIC
    Underline = DEF_UNDERLINE
    Strikethrough = DEF_STRIKETROUGH
    #If DebugMode Then
        DebugInit m_sDebugID, MODULE_NAME
    #End If
End Sub

#If DebugMode Then
    Private Sub Class_Terminate()
        DebugTerm m_sDebugID
    End Sub
#End If


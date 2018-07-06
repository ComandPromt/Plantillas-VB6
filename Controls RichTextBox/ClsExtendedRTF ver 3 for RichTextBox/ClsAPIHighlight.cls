VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAPIHighlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'copyright 2002 Roger Gilchrist
'based on RtfbackColourTestProject1.vbp
'By oigres P (Sergio Perciballi) Email:oigres@postmaster.co.uk
'at Planet-Source-Code
'found this only after I'd built the textRTF manipulation method
'This is much, much better for highlighting
'If I'd found this first I'd never have got round to clsExtendedRTF
'so thanks to Sergio for the code but I'm glad I missed it

'------------VERY IMPORTANT INFORMATION-------------------------------
'based on RtfbackColourTestProject1.vbp
'By oigres P (Sergio Perciballi) Email:oigres@postmaster.co.uk
'at Planet-Source-Code
'who supplied the following
'------------IMPORTANT INFORMATION--------------
'This sample requires the Riched20.dll (version 3) look at the file properties
'This also requires Riched32.dll (5.00.2008.1)
'and probably the Richx32.ocx control
'Based in part on the work of Steve McMahon (www.vbaccelarator.com)
'Delete/overwrite the riched20 and riched32.dll files
'You need to have vb6/5 not running to replace these files.
'new richtextbox and dll files can be found at www.vbaccelerator.com

'___________Roger Again _____________________________________________
'I'm a bit compulsive about modularity so each module has its own
'Private calls to to the API. This means if you copy just one module
'you don't have to keep coming back to find that missing Public variable
'in another module. Once you have inculde a module in your program you can
'remove the Private Declare statements if you have a Public one in your project

Private m_RTB As RichTextBox
Private m_CDlg As CommonDialog
Private Const WM_USER As Long = &H400
Private Const CFM_BACKCOLOR = &H4000000
Private Const EM_GETCHARFORMAT As Long = (WM_USER + 58)
Private Const EM_SETCHARFORMAT As Long = (WM_USER + 68)
Private Const SCF_SELECTION = &H1&
Private Const LF_FACESIZE As Integer = 32
Private m_busy As Boolean
Public Enum SlugType
    ColourRGB
    ColourLong
    Description
End Enum
Rem Mark Off
'stops Ulli's Code formatter from noticing these as Duplicated Name without Scope or TypeCasting
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim ColourRGB
Dim ColourLong
Dim Description
#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Rem Mark On

Private Type CHARFORMAT2
    cbSize As Integer '2
    wPad1 As Integer  '4
    dwMask As Long    '8
    dwEffects As Long '12
    yHeight As Long   '16
    yOffset As Long   '20
    crTextColor As Long '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58
    wPad2 As Integer ' 60

    ' Additional stuff supported by RICHEDIT20
    wWeight As Integer            ' /* Font weight (LOGFONT value)      */
    sSpacing As Integer           ' /* Amount to space between letters  */
    crBackColor As Long        ' /* Background color                 */
    lLCID As Long               ' /* Locale ID                        */
    dwReserved As Long         ' /* Reserved. Must be 0              */
    sStyle As Integer            ' /* Style handle                     */
    wKerning As Integer            ' /* Twip size above which to kern char pair*/
    bUnderlineType As Byte     ' /* Underline type                   */
    bAnimation As Byte         ' /* Animated text like marching ants */
    bRevAuthor As Byte         ' /* Revision author index            */
    bReserved1 As Byte
End Type
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal Clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Sub About(Optional EndUser As Boolean = True)

  Dim MsgProg As String, MsgUser As String

    'Dual About Structure
    'copyright 2002 Roger Gilchrist

    MsgProg = "ClsAPIHighlight offers * public methods and functions " & vbNewLine & _
              "About(T|F): About gives 1 of 2 messages Help(default) for use by end-user or this one Programmer(F)." & vbNewLine & _
              "APIHighlight: Private one colour handler." & vbNewLine & _
              "APIHighlight2: Private two colour handler" & vbNewLine & _
              "APIHighlightHard: Change the colour behind text." & vbNewLine & _
              "APIHighlightHardHard: Two colours hardcoded in program" & vbNewLine & _
              "APIHighlightAuto: Program assigns back colour, Class chooses Text colour" & vbNewLine & _
              "APIHighlightUser: User selects Back colour fore colour unchanged" & vbNewLine & _
              "APIHighlightUserAuto: User selects Back colour, Class chooses Text colour" & vbNewLine & _
              "APIHighlightUserUser: User selects Back and Text colour" & vbNewLine & _
              "AssignControls: Provide RichTextBox and CommonDialog for class to use." & vbNewLine & _
              "IsHighlight: Boolean test of highlight status of cursor position." & vbNewLine & _
              "UserColor:Routine to allow user to select a colour(can be used outside class)." & vbNewLine & _
              vbNewLine & _
              "ClsAPIHIghlight is brought to you by Roger Gilchrist (Email:rojagilkrist@hotmail.com)" & vbNewLine & _
              "Based on the work of Sergio Perciballi (Email:oigres@postmaster.co.uk)"

    'these are the basics an end user will need to know about but pretty it up for release
    ' this is a very basic end user help you may not want or need some parts of it so edit it to suit
    'For example you would want to rename the routines to whatever you captions them
    MsgUser = "HighlightHard: Hard coded highlight colour." & vbNewLine & _
              "HighlightHardHard: Hard coded 2 colour handler." & vbNewLine & _
              "HighlightAuto: Program assigns back colour, Program chooses Text colour" & vbNewLine & _
              "HighlightUser: User selects Back colour fore colour unchanged" & vbNewLine & _
              "HighlightUserAuto: You select a Back colour, Program chooses Text colour" & vbNewLine & _
              "HighlightUserUser: You select Back and Text colours" & vbNewLine & _
              vbNewLine & _
              "Brought to you by Roger Gilchrist (Email:rojagilkrist@hotmail.com)" & vbNewLine & _
              "Based on the work of Sergio Perciballi (Email:oigres@postmaster.co.uk)"
    'I know Iif is slow but this is not time critical
    MsgBox IIf(EndUser, MsgUser, MsgProg), , IIf(EndUser, "Highlighting", "ClsAPIHighlight")

End Sub

Private Function AddSlug(ByVal MainStr$, LBrkt, Data$, RBrkt$, Optional AtStart As Boolean = False) As String

  'copyright 2002 Roger Gilchrist
  'A Slug is a delimited string which can be added and removed from Mainstr
  'MainStr
  '       string to insert or remove slug
  'Data
  '       The added string. If it is an empty string "" the slug is removed
  'LBrkt & RBrkt
  '       can be anything you like as long as they do not appear in MainStr for any other reason.
  '       It is recommended that you use a couple of punctuation characters
  '       only use a single character if you are very sure they will not appear.
  '       LBrkt and RBrkt can be the same character(s).
  '       If you want spaces add them to the lBrkt and RBrkt, the search will find them.
  '       Remember Long filenames can contain most punctuation so if you slug a caption with a filename
  '       don't use single character LBrkt & RBrkt as you cannot know what will be in a filename.
  'AtStart
  '       is optional for where to place a slug. You only need use it once for
  '       new slugs, after that the slug will replace in the same location, until it is deleted.
  '       so if you use it at all use it all the time for that slug value
  '       If you use multiple slugs make sure none of the slug brackets overlap.

  'NOTE I designed this to add a small slug to Form.Captions to report variables while debugging
  'so it does not contain any tests for multiple occurances of slugs
  'it will replace all slugs with the same LBrkt and RBrkt value
  'This is also why it just adds new slugs to the end or start of the string.

  'USAGE : Place a line like this whereever you want the detection to take place
  'The example would go in the Button's mouse_move event
  'Form1.Caption = AddSlug(Form1.Caption, "*|", "This is the Cancel button", "|*")
  'don't forget to place a matching line in the Form Mouse_move
  'Form1.Caption = AddSlug(Form1.Caption, "*|", "", "|*") to clear it.
  'I don't recommend this usage after all there is ToolTip but this is a simple demo
  '

  Dim PStart As Long, Pend As Long

    PStart = InStr(MainStr, LBrkt)
    If PStart Then  'slug already exists replace it
        ' V--this allows Brkts to be same character
        Pend = InStr(PStart + 1, MainStr, RBrkt)
        MainStr = Left$(MainStr, PStart - 1) & LBrkt & Data$ & RBrkt & Mid$(MainStr, Pend + Len(RBrkt))
        If InStr(MainStr, LBrkt & RBrkt) Then 'delete it if it is an empty insert
            MainStr = Replace(MainStr, LBrkt & RBrkt, "")
        End If
      Else 'PSTART = FALSE
        If Len(Data) Then 'The Slug is new and data exists just add it to end of string
            If AtStart Then
                MainStr = LBrkt & Data$ & RBrkt & MainStr
              Else 'ATSTART = FALSE
                MainStr = MainStr & LBrkt & Data$ & RBrkt
            End If
        End If
    End If
    AddSlug = MainStr

End Function

Private Function APIerrorDescription(ByVal code As Long) As String

  'GEWare Software

  Dim Msg As String, R As Long

    Msg = String$(256, 0)
    R = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
        FORMAT_MESSAGE_IGNORE_INSERTS, ByVal 0, code, 0, Msg, 256, ByVal 0)
    If R Then
        APIerrorDescription = Left$(Msg, R)
    End If

End Function

Private Sub APIHighlight(colr As Long)

  'based on RtfbackColourTestProject1.vbp
  'By oigres P (Sergio Perciballi) Email:oigres@postmaster.co.uk
  'at Planet-Source-Code

  Dim Ret As Long, tmpcolr As Long
  Dim cf As CHARFORMAT2

    'if you use a Public Variable to store this you have to reset any change or
    'it will still be set when you make a different call. EG APIHighLight followed by APIBold
    ' both Bolds and Highlights the second instance, If you want to build up a style with one clcik
    'this might be useful but mostly not such a good thing.
    On Error GoTo oops
    With cf
        .cbSize = LenB(cf) 'setup the size of the character format
        .dwMask = CFM_BACKCOLOR 'what to test
        If .crBackColor = 0 Then
            .crBackColor = colr
          Else 'NOT .CRBACKCOLOR...
            .crBackColor = m_RTB.BackColor
        End If
        Ret = SendMessage(m_RTB.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, cf)
    End With 'CF

Exit Sub

oops:
    ErrorMessage Err

End Sub

Private Sub APIHighlight2(BColr As Long, FColr As Long)

  'creator Roger Gilchrist
  'private access for all 2 colour modes

    APIHighlight BColr
    m_RTB.SelColor = FColr

End Sub

Public Sub APIHighlightAuto(BColr As Long)

  'creator Roger Gilchrist Modification of Sergio Perciballi's work  (see APIHihglight)
  'User sets highlight colour Program selects suitable(?) font colour

    APIHighlight2 BColr, InvertColor(BColr)

End Sub

Private Function APIHighlightColour() As Long

  'Copyright 2002 Roger Gilchrist
  'This function is used by the Public Functions GetHighLightColorLong and GetHighLightColorRGB
  'you could call it direct but it is safer wrapped in the class
  'PROGRAMMER'S NOTE
  'this does not work perfectly as .crBackColor is filled at the mouse release point in the selection
  'if the selection is partially off a highlight it returns the colour at the mouse release point(usually VBBlack)

  Dim Ret As Long
  Dim cf As CHARFORMAT2

    On Error GoTo oops
    With cf
        .cbSize = LenB(cf) 'setup the size of the character format
        .dwMask = CFM_BACKCOLOR
        Ret = SendMessage(m_RTB.hwnd, EM_GETCHARFORMAT, SCF_SELECTION, cf)
        If (CFM_BACKCOLOR Xor (.dwEffects - CFM_BACKCOLOR)) <> .dwEffects Then
            APIHighlightColour = .crBackColor
          Else 'NOT (CFM_BACKCOLOR...
            APIHighlightColour = 0
        End If
    End With 'CF

Exit Function

oops:
    ErrorMessage Err

End Function

Public Sub APIHighlightHard(BColr As Long)

  'creator Roger Gilchrist
  'Hard means hardcoded colour selection

    APIHighlight BColr

End Sub

Public Sub APIHighlightHardHard(BColr As Long, FColr As Long)

  'creator Roger Gilchrist
  'Hard means hardcoded colour selection

    APIHighlight BColr
    m_RTB.SelColor = FColr

End Sub

Public Sub APIHighlightRemove()

    APIHighlight TranslateSysColor(vbWindowBackground)

End Sub

Public Sub APIHighlightUser()

  'creator Roger Gilchrist Modification of Sergio Perciballi's work  (see APIHihglight)
  'User sets highlight colour Program selects suitable(?) font colour

    APIHighlight UserColor

End Sub

Public Sub APIHighlightUserAuto()

  'creator Roger Gilchrist Modification of Sergio Perciballi's work  (see APIHihglight)
  'User sets highlight colour Program selects suitable(?) font colour

  Dim UColr As Long

    UColr = UserColor
    APIHighlight2 UColr, InvertColor(UColr)

End Sub

Public Sub APIHighlightUserUser()

  'creator Roger Gilchrist Modification of Sergio Perciballi's work  (see APIHihglight)
  'User sets highlight colour Program selects suitable(?) font colour

    APIHighlight2 UserColor, UserColor

End Sub

Public Sub AssignControls(R As RichTextBox, C As CommonDialog)

  'Copyright 2002 Roger Gilchrist
  'Place the Call to this in Form_Load or Sub Main
  'RTBLooks.AssignControls RichTextBox1, CommonDialog1

    m_busy = False
    Set m_CDlg = C
    Set m_RTB = R

End Sub

Public Property Get Busy() As Variant

  'Copyright 2002 Roger Gilchrist
  ' allows program to signal that its working hard and should not be bothered

    Busy = m_busy

End Property

Private Sub ErrorMessage(ErrVal)

  Dim Msg As String, RTBM As String, CdlM As String

    Msg = ErrVal.Description
    If ErrVal.Description = "Object variable or With block variable not set" Then
        If m_RTB Is Nothing Then
            RTBM$ = "RichTextBox"
        End If
        If m_CDlg Is Nothing Then
            CdlM$ = "CommonDialog"
        End If
        'I know Iif is slow but this is not time critical
        Msg = ErrVal.Description & vbNewLine & "You need to assign a " & RTBM & IIf(Len(RTBM) And Len(CdlM), " and a ", "") & CdlM & " using" & vbNewLine & _
              "'MyHigh.AssignControls RichTextBox1, CommonDialog1'"
    End If
    MsgBox Msg, , "ClsAPIHighlight"

End Sub

Public Function GetHighLightColorLong() As Long

  'Copyright 2002 Roger Gilchrist
  'this returns the long number of the background colour of
  'the highlight you are over
  'probably not something you want to show end users but
  'can lead to a way of letting you program respond to
  'different highlights
  'RETURN 0 if not over highlight  or black
  'See function IsHighLightColour to detect all colors

    GetHighLightColorLong = APIHighlightColour

End Function

Public Function GetHighLightColorRGB() As String

  'Copyright 2002 Roger Gilchrist
  'this returns a string containing the background colour of
  'the highlight you are over
  'probably not something you want to show end users but
  'can lead to a way of letting you program respond to
  'different highlights
  'RETURN "" if not over highlight

  Dim R As Integer, G As Integer, B As Integer

    If IsHighLightColour(APIHighlightColour) Then
        GetRGB APIHighlightColour, R, G, B
        If R = 0 Then
            R = "0"
        End If
        If G = 0 Then
            G = "0"
        End If
        If B = 0 Then
            B = "0"
        End If

        GetHighLightColorRGB = R & ":" & G & ":" & B
    End If

End Function

Private Sub GetRGB(ByVal col As Long, Optional ByRef Red As Integer, Optional ByRef Green As Integer, Optional ByRef Blue As Integer)

  'don't know where I got this but there are a million variations around

    Red = col Mod 256
    Green = ((col And &HFF00FF00) / 256)
    Blue = (col And &HFF0000) / (65536)

End Sub

Public Function HighLightDescription() As String

  'Copyright 2002 Roger Gilchrist
  'MODIFY this so that you can use strings related to specific Highlights
  'use GetHighLightColorRGB to get RGB style colours
  'PROGRAMMER'S NOTE
  'this does not work perfectly as the colour feed is at the selection point
  'if the selection is partially off a highlight it returns the colour at the
  'mouse release point

    Select Case APIHighlightColour
      Case NamedColour("beige")
        HighLightDescription = "beige"
      Case vbGreen
        HighLightDescription = "Green"
      Case RGB(0, 0, 0)
        HighLightDescription = "Black Highlight"
      Case RGB(255, 0, 0)
        HighLightDescription = "New Addition"
      Case RGB(192, 192, 192)
        HighLightDescription = "Button on Demo"
      Case Else
        'HighLightDescription = "unassigned"
    End Select

End Function

Public Function HighLightSlug(MainStr$, LBrkt$, Mode As SlugType, RBrkt$) As String

  Dim TMp As String

    If IsHighLightColour(APIHighlightColour) Then
        Select Case Mode
          Case ColourRGB
            TMp = GetHighLightColorRGB
          Case ColourLong

            TMp = APIHighlightColour
          Case Description

            TMp = HighLightDescription
        End Select
    End If
    HighLightSlug = AddSlug(MainStr, LBrkt, TMp, RBrkt)

End Function

Private Function InvertColor(col As Long) As Long

  Dim R As Integer, G As Integer, B As Integer, GREYTEST As Long

    GetRGB col, R, G, B
    'the inversion is mine
    ' the different test should keep you from getting unusable colours
    ' if RGB = 128,128,128 then the usual invert will just return to itself
    'only if R,G & B are close in value (after all (0,0,255) Dark Blue is fine but would trigger the GREYTEST
    If Abs(R - G) < 10 And Abs(G - B) < 10 And Abs(R - B) < 10 Then
        'if RGB are close they are usually some sort of grey
        GREYTEST = (R + G + B) / 3
        'and the average is near dark gray
        If GREYTEST >= 120 And GREYTEST <= 135 Then
            R = R + 128 ' shift the colour half the RGB spectrum away
            B = B + 128
            G = G + 128
        End If
    End If
    InvertColor = RGB(Abs(R - 255), Abs(G - 255), Abs(B - 255))
    If InvertColor = col Then ' just a bit of paranoia
        InvertColor = RGB(Abs(R - 128), Abs(G - 128), Abs(B - 128))
    End If

End Function

'Public Function IsHighlight() As Boolean
'
'  Dim Return As Long, tmpcolr As Long
'  Dim cf As CHARFORMAT2
'
'    'if you use a Public Variable to store this you have to reset any change or
'    'it will still be set when you make a different call. EG APIHighLight followed by APIBold
'    ' both Bolds and Highlights the second instance, If you want to build up a style with one clcik
'    'this might be useful but mostly not such a good thing.
'    On Error GoTo oops
'    With cf
'        .cbSize = LenB(cf) 'setup the size of the character format
'        .dwMask = CFM_BACKCOLOR
'        '        If .crBackColor = 0 Then
'        '            .crBackColor = colr
'        '          Else 'NOT .CRBACKCOLOR...
'        '            .crBackColor = RTB.BackColor
'        '        End If
'        Return = SendMessage(m_RTB.hwnd, EM_GETCHARFORMAT, SCF_SELECTION, cf)
'        IsHighlight = (CFM_BACKCOLOR Xor (.dwEffects - CFM_BACKCOLOR)) <> .dwEffects
'    End With 'CF
'
'Exit Function
'
'oops:
'    ErrorMessage Err
'
'End Function

Public Function IsHighlight() As OLE_TRISTATE

  'copyright 2002 Roger Gilchrist
  'test for highlight of a selection or at selection point if no selection made
  'Returns one of the following
  '   0 - Unchecked   not a highlight
  '   1 - Checked     is a highlight
  '   2 - Gray        not all highlighted

  Dim cf As CHARFORMAT2, TestVal As Boolean, ReturnVal As Long, ORTest As Long

    On Error GoTo oops
    With cf
        .cbSize = LenB(cf) 'setup the size of the character format
        .dwMask = CFM_BACKCOLOR
        ReturnVal = SendMessage(m_RTB.hwnd, EM_GETCHARFORMAT, SCF_SELECTION, cf)
        TestVal = (CFM_BACKCOLOR Xor (.dwEffects - CFM_BACKCOLOR)) <> .dwEffects
        ORTest = .dwMask Or .dwEffects
        If ORTest = ReturnVal And ReturnVal = -16777217 Then 'not a mixed selection or selpoint only
            If TestVal Then
                IsHighlight = Checked
              Else 'TESTVAL = FALSE
                IsHighlight = Unchecked
            End If
          Else 'NOT ORTEST...
            IsHighlight = Gray
        End If
    End With 'CF

Exit Function

oops:
    ErrorMessage Err

End Function

Private Function IsHighLightColour(HColr As Long) As Boolean

    IsHighLightColour = IsHighlight And GetHighLightColorLong = HColr

End Function

Public Function NamedColour(ColStr$) As Long

  'Copyright 2002 Roger Gilchrist
  'MODIFY this so that you can use named colours in your program
  '

    Select Case LCase$(ColStr)
      Case "beige"
        NamedColour = RGB(250, 200, 200)
      Case "fadedpink"
        NamedColour = RGB(250, 200, 200)
      Case "fred"

        NamedColour = RGB(23, 200, 200)
      Case "random"
        NamedColour = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
    End Select

End Function

Private Function TranslateSysColor(ByVal lColor As Long) As Long

    lColor = lColor And (Not &H80000000)
    'SysColor2RGB = GetSysColor(lColor)
    TranslateSysColor = GetSysColor(lColor)

End Function

'
'Public Function TranslateSysColorX(ByVal clr As OLE_COLOR) As Long
'
''Convert system colors
''Color values are specified in Visual Basic using the OLE_COLOR type. The colors
''come in two flavors; either true RGB values or the system colors with values
''above &H80000000. When using Visual Basic functions that take color parameters,
''VB will convert the system colors to their real value for you. However, Windows
''API functions don't understand the system values, so you will have to convert
''them manually. One way to do this is to use the OleTranslateColor function. If
''you convert a real color, i.e. below &H80000000, the function simply returns the
''same value, but for system colors, it will return a value that can be used as
''color parameter in an API call.
'  Dim R As Long
'  R = TranslateColor(clr, 0, TranslateSysColorX)
'  If R <> 0 Then 'raise an error
'  End If
'End Function

Public Function UserColor() As Long

  'creator(?) Roger Gilchrist
  'don't know where I got this, may have worked it out myself,
  'have been using it for ages when I need user's color input
  'This  has been left Public so that the programmers can access
  'to get colours elsewhere in their programs

    On Error GoTo error_cancel
    With m_CDlg
        '.DialogTitle = NO it would be nice to set it to
        '   "Highlight Colour|Font Colour"
        ' so of course it is not supported with ShowColor
        .Flags = cdlCCRGBInit
        .Color = vbRed
        .CancelError = True
        .ShowColor
        UserColor = .Color
    End With 'CDLG1'M_CDLG
error_cancel:
    'This error trap just falls through returning VbBlack, 0, as user colour

End Function

':) Ulli's VB Code Formatter V2.13.6 (28/08/2002 2:38:23 PM) 86 + 563 = 649 Lines

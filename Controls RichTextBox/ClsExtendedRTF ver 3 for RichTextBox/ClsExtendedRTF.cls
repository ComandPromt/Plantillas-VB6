VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsExtendedRTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright 2002 Roger Gilchrist
'
'------------IMPORTANT INFORMATION-------------------------------
'there is a problem with supporting highlighting
'This requires the Riched20.dll (version 3)
'This also requires Riched32.dll (5.00.2008.1)
'and probably the Richx32.ocx control
' VB6 SP5 (what I use) has no difficult with this
' thanks to Thomas Görtler (maniac@ colorarts.de) who let me know of the problem then solved it himself
'and Sergio Perciballi (oigres@postmaster.co.uk) whose code contained the detailed version stuff above

'------------OTHER IMPORTANT INFORMATION-------------------------------
'some of the routines in this class MUST be modified when you use them in your work
'some of them are ones you might like to modify or remove
'just search for
'*---PROGRAMMER MODIFICATION POINT---*
'to find them
'if you want to remove or fiddle with routines without this mark make a copy and comment it out first

Option Explicit
'Search Help for "Supported RTF Codes" article for more ideas
Private Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long

Private m_RTB As RichTextBox
Private m_CDlg As CommonDialog
Private m_Cmbo As ComboBox
Private m_FileNameDef As String
Private m_FileName As String
Private m_FileNameShort As String
Private m_FilePath As String
Private m_FileInCaption As Boolean
Private m_ColorArray() As Long
Private m_lastUndeline As String
Private m_lastHiColor As Long
Private m_lastHiForeColor As Long
Private m_SaveColor As Long
Private m_MouseX As Single
Private m_MouseY As Single
Private m_Start() As Long
Private m_Len() As Long
Private m_button As Integer
Private m_Shift As Integer
Private m_busy As Boolean
Private m_Dirty As Boolean
Private m_DefFormat As String
Private m_DocPercent As Long
Private m_End_OF_RTF As Long
Private Const EM_CHARFROMPOS As Long = &HD7
Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Const DQ As String = """"      ' DB=DoubleQuote NL=NewLine
Private Const NL As String = vbNewLine ' I use these for long MsgBoxes layout.
' for clarity of coding and ease of cut-n-paste
' I use VbNewLine and chr$(34) in code

Private Const SafeJoinSkip As String = "~!~!~!*~!"

Public Enum FrmtTypes
    RTFHeadBit
    FontTbl
    ColorTbl
    StyleSheet
    RTFTailer
    WorkOnThis
    TextBlock
    RTFendBit
End Enum

Public Enum RippleStyle
    BaseLine
    BaseLine1
    THeight
End Enum
Private Enum LIFO
    Push = True
    Pop = False
End Enum
Private Enum Cases
    LoCase
    HiCase
    PropCase
    MixCase
End Enum

Private Enum RTFPartsType
    Blank
    RTFOnly
    TextOnly
    Mixed
End Enum
Private Enum Shifts
    NoShift = 0
    KShift = 1
    KCtrl = 2
    KShiftCtrl = 3
    KAlt = 4
    KShiftAlt = 5
    KCtrlAlt = 6
    KShiftCrtlAlt = 7
End Enum

'Public Enum spectrum
'    s1RedYellow
'    s2YellowGreen
'    s3GreenCyan
'    s4CyanBlue
'    s5BlueMagenta
'    s6MagentaRed
'End Enum

Private Enum RTFCodeLocs
    NotPresent
    EmbedOr1st
    LastWithBlank
    LastInSelection
    LastBeforeText
    EndOfString
End Enum
Rem Mark Off
''stops Ulli 's Code formatter from noticing these as Duplicated Name without Scope or TypeCasting
#If False Then '    'Enforce Case for Enums (does not compile)
Dim RTFHeadBit       'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Dim FontTbl
Dim ColorTbl
Dim StyleSheet
Dim RTFTailer
Dim WorkOnThis
Dim TextBlock
Dim RTFendBit
#End If
#If False Then
Dim Push
Dim Pop
#End If
#If False Then
Dim LoCase
Dim HiCase
Dim PropCase
Dim MixCase
#End If
#If False Then
Dim Blank
Dim RTFOnly
Dim TextOnly
Dim Mixed
#End If
#If False Then
Dim KShift
Dim KCtrl
Dim KShiftCtrl
Dim KAlt
Dim KShiftAlt
Dim KCtrlAlt
Dim KShiftCrtlAlt
#End If
#If False Then
Dim NotPresent
Dim EmbedOr1st
Dim LastWithBlank
Dim LastInSelection
Dim LastBeforeText
Dim EndOfString
#End If
Rem Mark Off

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
'Adding line numbers to the front of text in a rich text box
'Example by HeSaidJoe
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Private Const EM_GETLINE As Long = &HC4
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_LINELENGTH As Long = &HC1

Public Sub About()

  Dim Msg As String

    Msg = "Not settled enough for a hard coded description yet. You'll have to read the code itself. This was the original submission, I have moved a lot out of it and into other classes. If you want you can do a bit of code archeology by comparing how this class does things with how it is done elsewhere."
    MsgBox Msg, , "clsExtendedRTF"

End Sub

Public Sub ActOnTag(StartTag$, EndTag$, RTFCode$, Optional DeleteTag As Boolean = False)

  'Copyright 2002 Roger Gilchrist
  'find tags , apply code and optionally remove the tags

  Dim RTF As String
  Dim StartPos As Long, EndPos As Long

    With m_RTB
        RTF$ = .Text
        StartPos = InStr(RTF$, StartTag$)
        Do While StartPos > 0
            EndPos = InStr(StartPos + 1, RTF$, EndTag$)
            If EndPos = 0 Then
                EndPos = Len(RTF$)
            End If
            .SelStart = StartPos - 1
            .SelLength = EndPos - StartPos + Len(EndTag$)
            SelRTFToggle RTFCode$
            If DeleteTag Then
                RTF$ = .SelRTF
                RTF$ = Replace(RTF$, RTFise(StartTag), "")
                RTF$ = Replace(RTF$, RTFise(EndTag), "")
                .SelRTF = RTF$
            End If
            StartPos = InStr(EndPos + 1, .Text, StartTag)
        Loop
    End With 'M_RTB

End Sub

Public Function AddSlug(ByVal MainStr$, LBrkt, Data$, RBrkt$, Optional AtStart As Boolean = False) As String

  'Cut'n'Paste for other project explains enormous notes
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
  '       . " / \ [ ] : ; = , are illegal characters in file names so use them
  'AtStart
  '       is optional for where to place a slug. You only need use it once for
  '       new slugs, after that the slug will replace in the same location, until it is deleted.
  '       so if you use it at all use it all the time for that slug value
  '       If you use multiple slugs make sure none of the slug brackets overlap.

  'NOTE I designed this to add a small slug to Form.Captions to report variables while debugging
  'so it does not contain any tests for multiple occurances of slugs
  'it will replace all slugs with the same LBrkt and RBrkt value.
  'This is also why it just adds new slugs to the end or start of the string.

  'USAGE : Place a line like this where ever you want the detection to take place
  'The example would go in the Button's mouse_move event
  'Form1.Caption = AddSlug(Form1.Caption, "*|", "This is the Cancel button", "|*")
  'don't forget to place a matching line in the Form Mouse_move
  'Form1.Caption = AddSlug(Form1.Caption, "*|", "", "|*") to clear it.
  'I don't recommend this usage after all there is ToolTip. This is a simple demo

  Dim PStart As Long, Pend As Long

    'don't need to update the Slug
    If InStr(MainStr$, LBrkt & Data$ & RBrkt$) Then
        AddSlug = MainStr
        Exit Function '>---> Bottom
    End If
    PStart = InStr(MainStr, LBrkt)
    If PStart Then  'Slug exists so replace it
        ' V--this allows Brkts to be same character
        Pend = InStr(PStart + 1, MainStr, RBrkt)
        MainStr = Left$(MainStr, PStart - 1) & LBrkt & Data$ & RBrkt & Mid$(MainStr, Pend + Len(RBrkt))
        If InStr(MainStr, LBrkt & RBrkt) Then 'delete it if Slug is empty
            MainStr = Replace(MainStr, LBrkt & RBrkt, "")
        End If
      Else 'PSTART = FALSE
        If Len(Data) Then 'Slug is new and data exists just add it to MainStr
            If AtStart Then
                MainStr = LBrkt & Data$ & RBrkt & MainStr
              Else 'ATSTART = FALSE
                MainStr = MainStr & LBrkt & Data$ & RBrkt
            End If
        End If
    End If
    AddSlug = MainStr

End Function

Public Function APIerrorDescription(ByVal code As Long) As String

  'GEWare Software
  'haven't implemented it yet but its here if you want it

  Dim Msg As String, R As Long

    Msg = String$(256, 0)
    R = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
        FORMAT_MESSAGE_IGNORE_INSERTS, ByVal 0, code, 0, Msg, 256, ByVal 0)
    If R Then
        APIerrorDescription = Left$(Msg, R)
    End If

End Function

Public Sub ApplyTag(StartTag$, EndTag$)

  'Copyright 2002 Roger Gilchrist
  'place tags around the selection
  'see demo for use
  'I'll write better description someday

  Dim TmpRTF As String
  Dim fsLoc As Long, RTFend As Long, LastCurly As Long

    With m_RTB
        TmpRTF$ = .SelRTF
        fsLoc = InStr(TmpRTF$, "\fs")       '.SelRTF has a standard format in which \fsN comes last
        RTFend = InStr(fsLoc, TmpRTF$, " ") 'so find it then find the next space(End_of_RTFcode Token)
        TmpRTF$ = Left$(TmpRTF$, RTFend) & StartTag$ & Mid$(TmpRTF$, RTFend + 1) 'and insert starttag
        LastCurly = Len(TmpRTF$) - InStr(StrReverse(TmpRTF$), "}")                  'then find last "}"
        TmpRTF$ = Left$(TmpRTF$, LastCurly) & EndTag$ & Mid$(TmpRTF$, LastCurly + 1)  'and insert endtag just before it.
        .SelRTF = TmpRTF$
    End With 'M_RTB

End Sub

Public Sub AssignControls(R As RichTextBox, C As CommonDialog, Optional FullAssign As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  'Place the Call to this in Form_Load or Sub Main
  '   AssignControls RichTextBox1, CommonDialog1, Combo1
  '    Form1.Show

    m_busy = False
    Set m_CDlg = C
    Set m_RTB = R
    'This allows you to do basic setup for demo on textlook panel

    If FullAssign = False Then
        Exit Sub '>---> Bottom
    End If
    '*---PROGRAMMER MODIFICATION POINT---*
    'VERY BAD DESIGN IDEA (except in demos (: )
    'This is really own relevant to RTB with a preloaded document
    'but does no harm if initial doc is blank
    'You might use this if you want end-users
    'to always have to read something before proceeding.
    'They'll soon delete the document and then get annoyed
    'by the message box that is triggered in FileLoad.

    m_FileName = R.FileName                     'establish the filename for the IDE loaded Document
    R.FileName = ""                             'clear the existing document from RichTextBox
    m_DefFormat = GetFormatString(RTFTailer)    'Get the default format of RichTextBox
    FileLoad m_FileName                             'FileReLoad the IDE loaded Document
    '''    NB if there is no filename there is no problem with this

    '''    Combo1.SetFocus ' use this if you want Zooming active as soon as you start.
    '                      ' Probably only useful in a Demo of this class :)

End Sub

Private Sub BreakUpRTFString(RTFHead As String, Parts)

  'Copyright 2002 Roger Gilchrist
  'separates the RTFCode into 1.RTFHead =the unchanging head of the string
  '                           2. Parts  = an array of the pieces that can be changed

  Dim FntTblStart As Long, ClrTblStart As Long ' Cutpoints for various pieces
  Dim ClrTblEnd As Long, StyleSheetStart As Long
  Dim WorkThis As String, TmpRTF As String, FntTbl As String, clrTbl As String, RTFTail As String
  Dim StyleSht As String, RTFHeader As String, txt As String

    TmpRTF = WorkingStringRead
    RTFTail$ = Left$(TmpRTF, m_End_OF_RTF)
    ' the order and placement and existence of fonttbl and colortbl are auto set by RTF engine
    'so the assumptions that ClrTblStart = End_of_FontTbl is automatically correct
    FntTblStart = InStr(RTFTail$, "{\fonttbl")
    ClrTblStart = InStr(RTFTail$, "{\colortbl")      'gets end of fonttbl
    If ClrTblStart = 0 Then                           'blank documents use default color tabel which is not present
        ClrTblStart = InStr(FntTblStart, RTFTail$, ";}}") + 3
    End If
    StyleSheetStart = InStr(RTFTail, "\stylesheet{") ' Stylesheet only occurs if selected text is whole document
    If FntTblStart Then
        FntTbl$ = Mid$(RTFTail$, FntTblStart, ClrTblStart - FntTblStart)
        RTFHeader$ = Left$(RTFTail$, FntTblStart - 1)
        RTFTail$ = Mid$(RTFTail$, ClrTblStart)

    End If
    ClrTblStart = InStr(RTFTail$, "{\colortbl") 'retest as RTFTail has changed
    If ClrTblStart Then
        ClrTblEnd = InStr(ClrTblStart, RTFTail$, ";}") 'find certain end of colortbl not a multiple color table has ";}}"
        ClrTblEnd = InStr(ClrTblEnd, RTFTail, "\") - ClrTblStart    'so distance to next lump of code varies so get certain start at next "\"
        clrTbl$ = Mid$(RTFTail$, ClrTblStart, ClrTblEnd)
        RTFTail$ = Mid$(RTFTail$, ClrTblEnd)

    End If
    If StyleSheetStart Then ' selected text is whole document
        StyleSht = Left$(RTFTail, InStr(RTFTail, ";}}") + 4)
        RTFTail = Mid$(RTFTail, InStr(RTFTail, ";}}") + 5)
    End If
    txt$ = Mid$(TmpRTF, m_End_OF_RTF + 1)
    WorkThis = RTFTail$ & Mid$(TmpRTF, m_End_OF_RTF + 1)
    Parts = Split(WorkThis)
    RTFHead = RTFHeader & FntTbl & clrTbl & StyleSht

End Sub

Public Property Get Busy() As Variant

  'Copyright 2002 Roger Gilchrist
  ' allows program to signal that its working hard and should not be bothered

    Busy = m_busy

End Property

Private Sub Class_Initialize()

    m_lastHiColor = RGB(127, 127, 127)
    m_lastHiForeColor = vbBlack
    m_lastUndeline = "\ul"
    m_FilePath = App.Path
    '*---PROGRAMMER MODIFICATION POINT---*
    'you might want to change this to something else
    'but this reflects MSWord's behaviour and most
    'end-users will know what is happening when they see it
    m_FileNameDef = "Document.rtf"
    'you might want to set this to false and
    'explicitly set it with Property ShowFileInCaption
    m_FileInCaption = True
    m_busy = True

End Sub

Private Sub ColourArray(FillTPaintF As Boolean)

  'Copyright 2002 Roger Gilchrist
  'maintains a array of colours used on each character in string
  'Support routine for Sub ColourStore

  Dim i As Integer, OrigStart As Long, OrigLen As Long

    With m_RTB
        PositionStore Push
        OrigStart = .SelStart
        OrigLen = .SelLength
        If FillTPaintF Then
            ReDim m_ColorArray(OrigStart To OrigStart + OrigLen)
        End If
        For i = LBound(m_ColorArray) To UBound(m_ColorArray)
            .SelStart = i
            .SelLength = 1
            If FillTPaintF Then
                m_ColorArray(i) = .SelColor
              Else 'FILLTPAINTF = FALSE
                .SelColor = m_ColorArray(i)
            End If
        Next i
        PositionStore Pop
    End With 'M_RTB

End Sub

Private Function ColourInvert(col As Long) As Long

  'Copyright 2002 Roger Gilchrist

  Dim R As Integer, G As Integer, B As Integer, GREYTEST As Long

    ColourLong2RGB col, R, G, B

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
    ColourInvert = RGB(Abs(R - 255), Abs(G - 255), Abs(B - 255))
    If ColourInvert = col Then
        ColourInvert = RGB(Abs(R - 128), Abs(G - 128), Abs(B - 128))
    End If

End Function

Private Sub ColourLong2RGB(ByVal col As Long, Optional ByRef Red As Integer, Optional ByRef Green As Integer, Optional ByRef Blue As Integer)

  'don't know where I got this but there are a million variations around

    Red = col Mod 256
    Green = ((col And &HFF00FF00) / 256)
    Blue = (col And &HFF0000) / (65536)

End Sub

Private Function ColourMultiple() As Boolean

  'Copyright 2002 Roger Gilchrist
  'detect multicolour in selection

    ColourMultiple = IsNull(m_RTB.SelColor)

End Function

Public Function ColourNamed(ColStr$) As Long

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'If you want to use specific colours for highlighting in you program
  'give them names and set them here.
  'the name can of course be anything you like
  'Descriptive: "spellingEerror" or "SpellingErrorBack" and "SpellingErrorFore"
  'Colour:"beige"' if you do this make it a believable one especially if end-users will see it
  'Anything that makes sense to you:"fred","Wilma","98o9876r65","!*&#@$"
  '--------------IMPORTANT-----------------------
  'Remember this routine is designed to be case insensitive so use lower case when you add your colours
  '----------------------------------------------

    Select Case LCase$(ColStr$) ' this means that if you mistype capitals the routine doesn't care
      Case "beige"
        ColourNamed = RGB(255, 200, 100)
      Case "mauve"
        ColourNamed = RGB(255, 200, 255)
      Case "grey", "gray" ' so both English and American spellers get the same thing!
        ColourNamed = RGB(175, 175, 175)
      Case "pink"
        ColourNamed = RGB(255, 200, 175)
      Case "beige"
        ColourNamed = RGB(255, 200, 100)
      Case "marigold"
        ColourNamed = RGB(255, 255, 200)
      Case "palegreen"
        ColourNamed = RGB(200, 255, 200)
      Case "limegreen"
        ColourNamed = RGB(200, 255, 100)
      Case "random", "don't care", "dumb luck"
        ColourNamed = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
      Case Else ' for programmers only; end-user's should never see this
        MsgBox "Specified ColourNamed " & DQ & ColStr & DQ & " does not exist or is misspelled.", , "ColourNamed Error"
        End     ' and this is why
    End Select

End Function

Private Sub ColourStore(SaveTSetF As Boolean)

  'Copyright 2002 Roger Gilchrist
  'Store and reset colour of characters in a string
  'used by Highlighter to preserve colours while selcolour forces
  'colours into fonttable
  'this will not be needed in next version!

  Static ColorMode As Boolean

    If SaveTSetF Then
        ColorMode = ColourMultiple
    End If
    With m_RTB
        If SaveTSetF Then
            If ColorMode Then
                ColourArray True
              Else 'COLORMODE = FALSE
                m_SaveColor = .SelColor
            End If

          Else 'SAVETSETF = FALSE
            If ColorMode Then
                ColourArray False
              Else 'COLORMODE = FALSE
                .SelColor = m_SaveColor
            End If
        End If

    End With 'M_RTB

End Sub

Public Function ColoursUsed(Optional FullDoc As Boolean = True) As Long

  'Copyright 2002 Roger Gilchrist
  'count colors in Whole Doc or Selection depending on selection
  'FullDoc can force it to count Whole Doc even if there is a selection
  '*---PROGRAMMER MODIFICATION POINT---*
  'this is a bit of a time waster remove it if you don't need it

  Dim t As String

    If Not m_busy Then
        If FullDoc = False And m_RTB.SelLength = 0 Then
            ColoursUsed = 0
            Exit Function '>---> Bottom
        End If
        ColoursUsed = CountOccurances(GetColorTable(FullDoc), "\red") + 1
    End If

End Function

Public Function ColourUser() As Long

  'creator(?) Roger Gilchrist
  'don't know where I got this, may have worked it out myself,
  'have been using it for ages when I need user's color input
  'This  has been left Public so that the programmers can access
  'to get colours elsewhere in their programs

    On Error GoTo error_cancel
    With m_CDlg
        '        .ShowColor
        '        .Color
        '        .Flags
        '               cdlCCFullOpen           &H2 Entire dialog box is displayed, including the Define Custom Colors section.
        '               cdlCCHelpButton         &H8 Causes the dialog box to display a Help button.
        '               cdlCCPreventFullOpen    &H4 Disables the Define Custom Colors command button and prevents the user from defining custom colors.
        '               cdlCCRGBInit            &H1 Sets the initial color value for the dialog box.

        '.DialogTitle 'NO! It would be nice to set it to
        ' "Highlight Colour|Font Colour"
        ' so of course it is not supported with ShowColor
        .Flags = cdlCCRGBInit
        .CancelError = True
        .ShowColor
        ColourUser = .Color
    End With 'CDLG1'M_CDLG
error_cancel:
    'This error trap just falls through returning VbBlack, 0, as user colour

End Function

Private Function CountOccurances(x, Item$) As Long

  ' not mine but don't remember where I got it
  'this is a cheap and nasty way to count sub strings

    If Len(x) Then ' remove this test if you would like a return of -1 for zero len strings
        CountOccurances = UBound(Split(x, Item))
    End If

End Function

Private Property Let CurChar(ByVal CChr As String)

  'Copyright 2002 Roger Gilchrist
  'reset character at selection point
  'used by Shift+Tab macro
  'not fully tested

    With m_RTB
  Dim CurCharEnd As Long
        If Len(CChr) Then
            CurCharEnd = m_RTB.SelLength
          Else 'LEN(CCHR) = FALSE
            CurCharEnd = 1
        End If
        .SelStart = .SelStart - 1
        .SelLength = Len(CChr) + CurCharEnd
        .SelText = CChr
    End With 'M_RTB

End Property

Private Property Get CurChar() As String

  'Copyright 2002 Roger Gilchrist
  ' get the character at the selection point or first character of selection
  'used by Shift+Tab macro
  'not fully tested

  Dim CurCharEnd As Long

    If m_RTB.SelLength Then
        CurCharEnd = m_RTB.SelLength
      Else 'M_RTB.SELLENGTH = FALSE
        CurCharEnd = 1
    End If

    CurChar = Mid$(m_RTB.Text, m_RTB.SelStart, CurCharEnd)

End Property

Private Function CurrentCase(ByVal a$) As Cases

  'Copyright 1999 Roger Gilchrist
  '*PURPOSE: return a value indicating the case of A$ ('0=lcase 1= UCASE,2=Proper, 3=mixed)
  '*CALLED FROM: CycleCase
  '*---PROGRAMMER MODIFICATION POINT---*
  'if you remove CycleCase you may as well remove this too

    Select Case a$
      Case LCase$(a$)
        CurrentCase = LoCase
      Case UCase$(a$)
        CurrentCase = HiCase
      Case StrConv(a$, vbProperCase)
        CurrentCase = PropCase
      Case Else
        CurrentCase = MixCase
    End Select

End Function

Public Property Let CurrentUnderlineStyle(ByVal UlStyle As String)

  'Copyright 2002 Roger Gilchrist
  'force class to specific underline style

    m_lastUndeline = UlStyle

End Property

Public Property Get CurrentUnderlineStyle() As String

  'Copyright 2002 Roger Gilchrist
  'allow program read last used underline style

    CurrentUnderlineStyle = m_lastUndeline

End Property

Private Sub CycleCase()

  'Copyright 2002 Roger Gilchrist
  'Duplicates MSWord's Shift-F3 macro
  'Selected text case cycles Lcase>Ucase>ProperCase>Lcase
  '*---PROGRAMMER MODIFICATION POINT---*
  'remove if you don't want it

  Dim a As String

    With m_RTB
        If Len(.SelText) Then
            PositionStore Push
            a$ = .SelText
            Select Case CurrentCase(a$)
                '0=lcase 1= UCASE,2Proper
              Case 0
                .SelText = StrConv(a$, vbProperCase) 'lcase>Pcase
              Case 2
                .SelText = UCase$(a$) 'Pcase>UCASE
              Case 1
                .SelText = LCase$(a$)
            End Select
            PositionStore Pop
        End If
    End With 'CURRENTTEXT'M_RTB

End Sub

Public Function CycleSelColor(CColour As Long) As Long

  'Copyright 1999 Roger Gilchrist
  '*PURPOSE: cycle selection colours through the 8 VB colours
  '*---PROGRAMMER MODIFICATION POINT---*
  'this is a demo macro if you don't want it remove
  'don't forget to remove call in KeysFunction too

    Select Case CColour
      Case vbBlack
        CycleSelColor = vbRed
      Case vbRed
        CycleSelColor = vbGreen
      Case vbGreen
        CycleSelColor = vbYellow
      Case vbYellow
        CycleSelColor = vbBlue
      Case vbBlue
        CycleSelColor = vbMagenta
      Case vbMagenta
        CycleSelColor = vbCyan
      Case vbCyan
        CycleSelColor = vbWhite
      Case vbWhite
        CycleSelColor = vbBlack
      Case Else
        CycleSelColor = vbBlack
    End Select

End Function

Private Sub DeleteRTFCode(TmpRTF$, RTFCode$)

  'Copyright 2002 Roger Gilchrist

  Dim Looper As RTFCodeLocs

    'how you delete RTFCode depends on where the code occurs
    'do NOT change the structure of these Replace structures
    'Remember the extra space in the LastXXXXX Cases
    'is the End_Of_RTFCode Token NOT a space character
    'it also stops a short RTFcode like "\v"
    'hitting a long one like "\viewkind4"

    Looper = RTFCodeLoc(TmpRTF$, RTFCode$)

    Do While Looper <> NotPresent
        Select Case Looper
          Case NotPresent
          Case EmbedOr1st
            TmpRTF$ = Replace(TmpRTF$, RTFCode$, "")
          Case LastWithBlank
            TmpRTF$ = Replace(TmpRTF$, RTFCode$ & "  ", " ")
          Case LastInSelection
            TmpRTF$ = Replace(TmpRTF$, RTFCode$ & " ", "")
          Case LastBeforeText
            TmpRTF$ = Replace(TmpRTF$, RTFCode$ & " ", " ")
        End Select
        Looper = RTFCodeLoc(TmpRTF$, RTFCode$)
    Loop

End Sub

Private Sub DeleteRTFCodeX(Part$, CS$)

  'Copyright 2002 Roger Gilchrist
  'slightly different version of previous routine
  ' don't have time/inclination to work out which is better
  'both are used in different places

  Dim Parts2 As Variant, i As Integer, tpart As String, bt As String

    Parts2 = Split(Part$, "\")
    For i = LBound(Parts2) To UBound(Parts2)
        bt$ = Parts2(i)
        ' guards against short matching start of long
        If Left$("\" & bt$ & " ", Len(CS$) + 1) = CS$ & " " Then
            bt$ = Replace(bt$, Mid$(CS$, 2), "")
            Parts2(i) = SafeJoinTest(Parts2(i), bt$)
        End If
    Next i
    Part$ = SafeJoin(Parts2, "\")

End Sub

Public Property Let Dirty(ByVal vNewValue As Boolean)

    m_Dirty = vNewValue

End Property

Public Property Get Dirty() As Boolean

    Dirty = m_Dirty

End Property

Public Property Get DocPercent() As Variant

  'Copyright 2002 Roger Gilchrist
  'returns the percentage of document at selection point
  '                                 V---this may be unnecessary but should stop hits while working on selections

    If m_busy = False Then
        If Len(m_RTB.Text) > 0 And m_RTB.SelLength = 0 Then
            m_DocPercent = m_RTB.SelStart / Len(m_RTB.Text) * 100
            DocPercent = m_DocPercent
          Else 'NOT LEN(M_RTB.TEXT)...
            DocPercent = 0
        End If
    End If

End Property

Public Property Let DocPercent(ByVal vNewValue As Variant)

    m_RTB.SelStart = CLng(vNewValue / 100 * Len(m_RTB.Text))

End Property

Public Function ExtendedQuoteToggle(code$) As Boolean

  'Copyright 2002 Roger Gilchrist

  Dim RTFHead As String     'Initial section of RTF probably never useful
  Dim Parts As Variant      'Split MixedBit for application of code
  Dim FntTblStart As Long, ClrTblStart As Long, RTF_End As Long ' Cutpoints for various pieces
  Dim i As Integer          'For variable to manipulate Parts
  Dim TMp As String          'tmp for testing whether "" is original or edited
  Dim StartAdjust As Integer, LenAdjust As Integer
  Dim LQ As String, RQ As String

    With m_RTB
        If .SelLength > 0 Then
            PositionStore Push
            BreakUpRTFString RTFHead, Parts
            If code = "\dbquote" Then
                LQ = "\ldblquote "
                RQ = "\rdblquote "
              Else 'NOT CODE...
                LQ = "\lquote "
                RQ = "\rquote "
            End If

            Parts(LBound(Parts) + 1) = LQ & Parts(LBound(Parts) + 1)
            If Parts(LBound(Parts) + 1) = LQ Then
                Parts(LBound(Parts) + 2) = LQ & Parts(LBound(Parts) + 2)
                Parts(LBound(Parts) + 1) = ""
                StartAdjust = 1
            End If
            Parts(UBound(Parts)) = Left$(Parts(UBound(Parts)), InStr(Parts(UBound(Parts)), "}") - 1) & RQ & Mid$(Parts(UBound(Parts)), InStr(Parts(UBound(Parts)), "}"))

            If Left$(Parts(UBound(Parts)), 2) = "\r" Then
                Parts(UBound(Parts) - 1) = Parts(UBound(Parts) - 1) & RQ
                Parts(UBound(Parts)) = Mid$(Parts(UBound(Parts)), InStr(Parts(UBound(Parts)), "quote") + 6)
                LenAdjust = 1
            End If
            WorkingStringWrite RTFHead & SafeJoin(Parts, " ")
            PositionStore Pop
            LenAdjust = LenAdjust + .SelLength + 1
            .SelStart = .SelStart - StartAdjust
            .SelLength = LenAdjust
        End If
    End With 'M_RTB

End Function

Public Sub FileLoad(fname$)

  'Copyright 2002 Roger Gilchrist

    On Error GoTo oops
    m_RTB.LoadFile fname$
    m_FileName = fname$

    FileValues

Exit Sub

oops:
    MsgBox "The file you attempted to FileLoad has been moved, renamed or deleted.", , App.EXEName
    fname = ""
    m_FileName = ""
    m_Dirty = False
    'This error only really occurs in the Demo Program for the class which
    'attempts to reload the document that has been set in the IDE if you have deleted that document
    'In real world it is very unlikely that anything could happen to a file between getting its name
    'and reaching this routine.

End Sub

Public Property Get FileName() As String

    FileName = m_FileName

End Property

Public Property Let FileName(ByVal vNewValue As String)

    m_FileName = vNewValue

End Property

Public Property Let FileNameDefault(ByVal DefFileName As String)

    m_FileNameDef = DefFileName

End Property

Public Property Get FileNameDefault() As String

    FileNameDefault = m_FileNameDef

End Property

Public Property Get FileNameShort() As String

    FileNameShort = m_FileNameShort

End Property

Public Sub FileNew()

    FileSafeSave
    m_RTB.Text = ""
    m_FileName = m_FilePath & "\" & m_FileNameDef '"Document.rtf"
    FileValues

End Sub

Public Sub FileOpen()

    FileSafeSave
    With m_CDlg
        '.DialogTitle = "Open" Automatic with .ShowOpen
        .Filter = "All files(*.*)|Rich Text Format$(*.rtf)"
        .FilterIndex = 2
        .FileName = "*.rtf"
        .InitDir = m_FilePath
        .DefaultExt = ".rtf"
        .ShowOpen
        If Len(.FileName) > 4 And Left$(.FileName, 2) <> "*." Then
            FileLoad .FileName
        End If
    End With 'COMMONDIALOG1'M_CDLG
    '*---PROGRAMMER MODIFICATION POINT---*
    'These are just here for my own use when constructing
    'the CommonDialog handlers delete them immediately
    'I left them here for when(if) I add Font,Help and Printer handling to the class
    ' and to remind me to add proper error handling

    'I don't use them and ClsExtendedRTF is not set up for multiple files processing(yet:)) but just in case your interested I pasted them in.
    '            .MaxFileSize ' use with .Flags = cdlOFNAllowMultiselect Range is 1 — 32K. The default is 256.
    '           .Flags =
    '                Constant                   Value   Description
    '                cdlOFNAllowMultiselect     &H200       Specifies that the File Namelist box allows multiple selections. The user can select more than one file at run time by pressing the SHIFT key and using the UP ARROW and DOWN ARROW keys to select the desired files. When this is done, the FileName property returns a string containing the names of all selected files. The names in the string are delimited by spaces.
    '                cdlOFNCreatePrompt         &H2000      Specifies that the dialog box prompts the user to create a file that doesn't currently exist. This flag automatically sets the cdlOFNPathMustExist and cdlOFNFileMustExist flags.
    '                cdlOFNExplorer             &H80000     Use the Explorer-like Open A File dialog box template. Works with Windows 95 and Windows NT 4.0.
    '                CdlOFNExtensionDifferent   &H400       Indicates that the extension of the returned filename is different from the extension specified by the DefaultExt property. This flag isn't set if the DefaultExt property is Null, if the extensions match, or if the file has no extension. This flag value can be checked upon closing the dialog box.
    '                cdlOFNFileMustExist        &H1000      Specifies that the user can enter only names of existing files in the File Name text box. If this flag is set and the user enters an invalid filename, a warning is displayed. This flag automatically sets the cdlOFNPathMustExist flag.
    '                cdlOFNHelpButton           &H10        Causes the dialog box to display the Help button.
    '                cdlOFNHideReadOnly         &H4         Hides the Read Onlycheck box.
    '                cdlOFNLongNames            &H200000    Use long filenames.
    '                cdlOFNNoChangeDir          &H8         Forces the dialog box to set the current directory to what it was when the dialog box was opened.
    '                CdlOFNNoDereferenceLinks   &H100000    Do not dereference shell links (also known as shortcuts). By default, choosing a shell link causes it to be dereferenced by the shell.
    '                cdlOFNNoLongNames          &H40000     No long file names.
    '                CdlOFNNoReadOnlyReturn     &H8000      Specifies that the returned file won't have the Read Only attribute set and won't be in a write-protected directory.
    '                cdlOFNNoValidate           &H100       Specifies that the common dialog box allows invalid characters in the returned filename.
    '                cdlOFNOverwritePrompt      &H2         Causes the FileSave As dialog box to generate a message box if the selected file already exists. The user must confirm whether to overwrite the file.
    '                cdlOFNPathMustExist        &H800       Specifies that the user can enter only valid paths. If this flag is set and the user enters an invalid path, a warning message is displayed.
    '                cdlOFNReadOnly             &H1         Causes the Read Only check box to be initially checked when the dialog box is created. This flag also indicates the state of the Read Only check box when the dialog box is closed.
    '                cdlOFNShareAware           &H4000      Specifies that sharing violation errors will be ignored.
    '
    '-------------------------------------------------------------------------------------------------------

    '        USED or USEFUL for all SHOWXXXX method
    '        .CancelError
    '        .Flags
    '
    '        NOT USED BY clsExtendedRTF just here in case you think of a use'
    '        '.Action DONT USE .SHOWXXXX is better
    '        .hDC
    '        .Index
    '        .Name
    '        .Object
    '        .Parent
    '        .Tag
    '
    '        .ShowFont
    '        .FontBold
    '        .FontItalic
    '        .FontName
    '        .FontSize
    '        .FontStrikethru
    '        .FontUnderline
    '        .Color
    '       .Flags
    '                cdlCFANSIOnly          &H400 Specifies that the dialog box allows only a selection of the fonts that use the Windows character set. If this flag is set, the user won't be able to select a font that contains only symbols.
    '                cdlCFApply             &H200 Enables the Apply button on the dialog box.
    '                cdlCFBoth              &H3 Causes the dialog box to list the available printer and screen fonts. The hDC property identifies thedevice context associated with the printer.
    '                cdlCFEffects           &H100 Specifies that the dialog box enables strikethrough, underline, and color effects.
    '                cdlCFFixedPitchOnly    &H4000 Specifies that the dialog box selects only fixed-pitch fonts.
    '                cdlCFForceFontExist    &H10000 Specifies that an error message box is displayed if the user attempts to select a font or style that doesn't exist.
    '                cdlCFHelpButton        &H4 Causes the dialog box to display a Help button.
    '                cdlCFLimitSize         &H2000 Specifies that the dialog box selects only font sizes within the range specified by the Min and Max properties.
    '                cdlCFNoFaceSel         &H80000 No font name selected.
    '                cdlCFNoSimulations     &H1000 Specifies that the dialog box doesn't allow graphic device interface (GDI) font simulations.
    '                cdlCFNoSizeSel         &H200000 No font size selected.
    '                cdlCFNoStyleSel        &H100000 No style was selected.
    '                cdlCFNoVectorFonts     &H800 Specifies that the dialog box doesn't allow vector-font selections.
    '                cdlCFPrinterFonts      &H2 Causes the dialog box to list only the fonts supported by the printer, specified by the hDC property.
    '                cdlCFScalableOnly      &H20000 Specifies that the dialog box allows only the selection of fonts that can be scaled.
    '                cdlCFScreenFonts       &H1 Causes the dialog box to list only the screen fonts supported by the system.
    '                cdlCFTTOnly            &H40000 Specifies that the dialog box allows only the selection of TrueType fonts.
    '                cdlCFWYSIWYG           &H8000 Specifies that the dialog box allows only the selection of fonts that are available on both the printer and on screen. If this flag is set, the cdlCFBoth and cdlCFScalableOnly flags should also be set.
    '
    '        .ShowHelp
    '        .HelpCommand
    '        .HelpContext
    '        .HelpFile
    '        .HelpKey
    '

    '        .ShowPrinter
    '        .Copies
    '        .FromPage
    '        .Max
    '        .Min
    '        .Orientation
    '        .PrinterDefault
    '        .ToPage

End Sub

Public Sub FileReLoad()

  '*---PROGRAMMER MODIFICATION POINT---*
  'This is mainly for demo use.
  'It is the ultimate undo.
  'It is a bit dangerous in real world programs.
  'If you use it be sure to put a MsgBox with a way out in front of it.
  'You should also probably hide it on a submenu so that it is hard to click by mistake,
  'rather than in the middle of the standard file New/Open/Save/SaveAs as the demo program does.
  'Make a big deal about it in documentation so users are aware of it and don't click-and-try in a bored moment.

  Dim SelPoint As Long ' store SelStart of old so that new will be displayed at the same place

    ' or very near it unless you have done a lot of editing since last save.

    If MsgBox("WARNING!" & NL & _
       "This will re-load the current saved document." & NL & _
       "All edits since last save will be lost." & NL & _
       NL & _
       "Continue?" _
       , vbCritical + vbOKCancel + vbDefaultButton1, "WARNING!") = vbOK Then
        SelPoint = m_RTB.SelStart
        FileLoad m_FileName
        m_RTB.SelStart = SelPoint
    End If

End Sub

Public Sub FileSafeSave()

  'Copyright 2002 Roger Gilchrist
  'call from Form_Unload and you will be warned if you try to quit without saving
  'you could also place it just before you FileLoad a doucument or start a new one for further safety

    If m_Dirty Then
        If Len(m_FileName) Then ' ignore blank page
            If MsgBox("Your file: " & NL & _
               m_FileName & NL & _
               "has been edited." & NL & _
               "Save?", vbDefaultButton1 + vbExclamation + vbYesNo, "Unsaved Edits") = vbYes Then
                FileSave
            End If
        End If
    End If

End Sub

Public Sub FileSave()

  'Copyright 2002 Roger Gilchrist

    If Len(m_FileName$) Then
        If m_FileNameShort = m_FileNameDef Then
            FileSaveAs
          Else 'NOT M_FILENAMESHORT...
            m_RTB.SaveFile m_FileName$
            m_Dirty = False
        End If
      Else 'LEN(M_FILENAME$) = FALSE
        MsgBox "File has no name."
    End If

End Sub

Public Sub FileSaveAs()

  '*---PROGRAMMER MODIFICATION POINT---*
  'these are just here to remind me to think about using them
  '        .MaxFileSize ' use with .Flags = cdlOFNAllowMultiselect Range is 1 — 32K. The default is 256.
  '           .Flags =
  '                Constant                   Value   Description
  '                cdlOFNAllowMultiselect     &H200       Specifies that the File Namelist box allows multiple selections. The user can select more than one file at run time by pressing the SHIFT key and using the UP ARROW and DOWN ARROW keys to select the desired files. When this is done, the FileName property returns a string containing the names of all selected files. The names in the string are delimited by spaces.
  '                cdlOFNCreatePrompt         &H2000      Specifies that the dialog box prompts the user to create a file that doesn't currently exist. This flag automatically sets the cdlOFNPathMustExist and cdlOFNFileMustExist flags.
  '                cdlOFNExplorer             &H80000     Use the Explorer-like Open A File dialog box template. Works with Windows 95 and Windows NT 4.0.
  '                CdlOFNExtensionDifferent   &H400       Indicates that the extension of the returned filename is different from the extension specified by the DefaultExt property. This flag isn't set if the DefaultExt property is Null, if the extensions match, or if the file has no extension. This flag value can be checked upon closing the dialog box.
  '                cdlOFNFileMustExist        &H1000      Specifies that the user can enter only names of existing files in the File Name text box. If this flag is set and the user enters an invalid filename, a warning is displayed. This flag automatically sets the cdlOFNPathMustExist flag.
  '                cdlOFNHelpButton           &H10        Causes the dialog box to display the Help button.
  '                cdlOFNHideReadOnly         &H4         Hides the Read Onlycheck box.
  '                cdlOFNLongNames            &H200000    Use long filenames.
  '                cdlOFNNoChangeDir          &H8         Forces the dialog box to set the current directory to what it was when the dialog box was opened.
  '                CdlOFNNoDereferenceLinks   &H100000    Do not dereference shell links (also known as shortcuts). By default, choosing a shell link causes it to be dereferenced by the shell.
  '                cdlOFNNoLongNames          &H40000     No long file names.
  '                CdlOFNNoReadOnlyReturn     &H8000      Specifies that the returned file won't have the Read Only attribute set and won't be in a write-protected directory.
  '                cdlOFNNoValidate           &H100       Specifies that the common dialog box allows invalid characters in the returned filename.
  '                cdlOFNOverwritePrompt      &H2         Causes the FileSave As dialog box to generate a message box if the selected file already exists. The user must confirm whether to overwrite the file.
  '                cdlOFNPathMustExist        &H800       Specifies that the user can enter only valid paths. If this flag is set and the user enters an invalid path, a warning message is displayed.
  '                cdlOFNReadOnly             &H1         Causes the Read Only check box to be initially checked when the dialog box is created. This flag also indicates the state of the Read Only check box when the dialog box is closed.
  '                cdlOFNShareAware           &H4000      Specifies that sharing violation errors will be ignored.

    With m_CDlg
        .Filter = "All files(*.*)|Rich Text Format$(*.rtf)"
        .Flags = cdlOFNOverwritePrompt
        .FilterIndex = 2
        .InitDir = m_FilePath
        .DefaultExt = ".rtf"
        'I know Iif is slow but this is not time critical
        .FileName = IIf(m_FileNameShort = m_FileNameDef, "", m_FileName)
        .ShowSave
        If Len(.FileName) Then
            m_RTB.FileName = .FileName

            FileValues

            m_RTB.SaveFile m_FileName$
            m_Dirty = False
        End If

    End With 'COMMONDIALOG1'M_CDLG

End Sub

Private Sub FileValues()

  '*---PROGRAMMER MODIFICATION POINT---*
  'if you add more modern file/directory support
  'modify this (I'll get round to it sometime)
  'OK there are more modern ways to do this but this is Q&D demo-stuff

    m_FileNameShort = StrReverse(Left$(StrReverse(m_FileName), InStr(StrReverse(m_FileName), "\") - 1))
    If Left$(m_FileName, Len(m_FilePath)) <> m_FilePath Then
        m_FilePath = StrReverse(Mid$(StrReverse(m_FileName), InStr(StrReverse(m_FileName), "\") + 1))
    End If
    m_Dirty = False
    If m_FileInCaption Then
        m_RTB.Parent.Caption = AddSlug(m_RTB.Parent.Caption, " [", m_FileNameShort, "]")
    End If

End Sub

Private Function FontSizeExtract(str$) As Long

  'Copyright 2002 Roger Gilchrist

  Dim fsLoc As Long, RTFend As Long

    fsLoc = InStr(str$, "\fs")       '.SelRTF has a standard format in which \fsN comes last
    If fsLoc Then
        FontSizeExtract = Mid$(str$, fsLoc + 3)
      Else 'FSLOC = FALSE
        PositionStore Push
        With m_RTB
            .SelStart = 1
            .SelLength = 1
            FontSizeExtract = .SelFontSize
        End With 'M_RTB
        PositionStore Pop
    End If

End Function

Private Function GetColorTable(Optional ForceDoc As Boolean = False) As String

  'Copyright 2002 Roger Gilchrist
  'Specialised version of GetFormatString
  'used by ColoursUsed

  Dim ClrTblStart As Long ' Cutpoints for various pieces
  Dim ClrTblEnd As Long
  Dim RTFTail As String

    RTFTail$ = Left$(WorkingStringRead(ForceDoc), m_End_OF_RTF)
    ClrTblStart = InStr(RTFTail$, "{\colortbl")      'gets end of fonttbl
    If ClrTblStart = 0 Then                           'blank documents use default color table which is not present
        GetColorTable = ""
        Exit Function '>---> Bottom
      Else 'NOT CLRTBLSTART...
        If ClrTblStart Then
            ClrTblEnd = InStr(ClrTblStart, RTFTail$, ";}") 'find certain end of colortbl not a multiple color table has ";}}"
            ClrTblEnd = InStr(ClrTblEnd, RTFTail, "\") - ClrTblStart    'so distance to next lump of code varies so get certain start at next "\"
            GetColorTable = Mid$(RTFTail$, ClrTblStart, ClrTblEnd)
        End If
    End If

End Function

Private Function GetFormatString(FrmtPart As FrmtTypes, Optional ForceDoc As Boolean = False) As String

  'Copyright 2002 Roger Gilchrist
  'this routine allows you access various parts of the TextRTF code
  'it is not used by the nearly identical BreakUpRTFString (on which it is based) because it would
  'have too high a performance hit. If you find you are referencing many of the format strings consider
  'making a task specific version rather than multiple calls here

  Dim FntTblStart As Long, ClrTblStart As Long ', RTF_End As Long  ' Cutpoints for various pieces
  Dim ClrTblEnd As Long, StyleSheetStart As Long
  Dim WorkThis As String, TmpRTF As String, FntTbl As String, clrTbl As String, RTFTail As String
  Dim SyleSht As String, RTFHead As String, txt As String

    TmpRTF = WorkingStringRead(ForceDoc)
    ' RTF_End = End_OF_RTF(ForceDoc)
    RTFTail$ = Left$(TmpRTF, m_End_OF_RTF)
    ' the order and placement and existence of fonttbl and colortbl are auto set by RTF engine
    'so the assumptions that ClrTblStart = End_of_FontTbl is automatically correct
    FntTblStart = InStr(RTFTail$, "{\fonttbl")
    ClrTblStart = InStr(RTFTail$, "{\colortbl")      'gets end of fonttbl
    If ClrTblStart = 0 Then                           'blank documents use default color table which is not present
        ClrTblStart = InStr(FntTblStart, RTFTail$, ";}}") + 3
    End If
    StyleSheetStart = InStr(RTFTail, "\stylesheet{") ' Stylesheet only occurs if selected text is whole document
    If FntTblStart Then
        FntTbl$ = Mid$(RTFTail$, FntTblStart, ClrTblStart - FntTblStart)
        RTFHead$ = Left$(RTFTail$, FntTblStart - 1)
        RTFTail$ = Mid$(RTFTail$, ClrTblStart)

    End If
    ClrTblStart = InStr(RTFTail$, "{\colortbl") 'retest as RTFTail has changed
    If ClrTblStart Then
        ClrTblEnd = InStr(ClrTblStart, RTFTail$, ";}") 'find certain end of colortbl not a multiple color table has ";}}"
        ClrTblEnd = InStr(ClrTblEnd, RTFTail, "\") - ClrTblStart    'so distance to next lump of code varies so get certain start at next "\"
        clrTbl$ = Mid$(RTFTail$, ClrTblStart, ClrTblEnd)
        RTFTail$ = Mid$(RTFTail$, ClrTblEnd)

    End If
    If StyleSheetStart Then ' selected text is whole document
        SyleSht = Left$(RTFTail, InStr(RTFTail, ";}}") + 4)
        RTFTail = Mid$(RTFTail, InStr(RTFTail, ";}}") + 5)
    End If
    txt$ = Mid$(TmpRTF, m_End_OF_RTF + 1)
    WorkThis = RTFTail$ & Mid$(TmpRTF, m_End_OF_RTF + 1)
    Select Case FrmtPart
      Case RTFHeadBit
        GetFormatString = RTFHead ' RTF establishment code should probably never be touched
      Case FontTbl
        GetFormatString = FntTbl  ' you could use CountOccurances(GetFormatString(FontTbl),"\fchar") + 1 to count fonts in use
      Case ColorTbl
        GetFormatString = clrTbl   ' you could use CountOccurances(GetFormatString(ColorTbl),"\red") + 1 to count colors in use
      Case StyleSheet
        GetFormatString = SyleSht ' you could use CountOccurances(GetFormatString(StyleSheet),"\s") + 1 to count styles in use
      Case RTFTailer
        GetFormatString = RTFTail 'This contains the initial RTF paragraph and format details for a selection
      Case TextBlock
        GetFormatString = Mid$(TmpRTF, m_End_OF_RTF + 1) ' this contains the rest of the selection not used directly
      Case WorkOnThis
        GetFormatString = WorkThis ' this contains RTFTailer and the TxtBlock this is what the class does most work on

      Case RTFendBit
        GetFormatString = WorkThis ' this contains RTFTailer and the rest of the selection
    End Select
    'RTFHead & FntTbl & ClrTbl & SyleSht & RTFTail & WorkThis

End Function

Private Function GetGetValue(RTFCode$) As Variant

  'Copyright 2002 Roger Gilchrist
  'support for the Get SelXXXX properties
  '*---PROGRAMMER MODIFICATION POINT---*
  'this is provisional until I find/work out how to make it behave properly

    With m_RTB
        If .SelLength Then
            GetGetValue = Not (RTFCodeLoc(.SelRTF, RTFCode$) <> NotPresent)
          Else '.SELLENGTH = FALSE
            GetGetValue = HasRTFCode(RTFCode$)
        End If
    End With 'M_RTB

End Function

Public Function HasHiddenText() As Boolean

  'Copyright 2002 Roger Gilchrist
  'Use this as model for RTF code which does not have a numeric component
  'WARNING WARNING *---PROGRAMMER MODIFICATION POINT---*  WARNING WARNING
  'THIS IS REALLY slow don't use it or anything that calls it in real world work

    HasHiddenText = HasRTFCode("\v")

End Function

Public Function HasHighlitText() As Boolean

  'Copyright 2002 Roger Gilchrist
  'Use this as model for RTF code which does have a numeric component
  'WARNING WARNING *---PROGRAMMER MODIFICATION POINT---*  WARNING WARNING
  'THIS IS REALLY slow don't use it or anything that calls it in real world work

    HasHighlitText = HasRTFCodeNum("\highlight")

End Function

Private Function HasRTFCode(RTFCode$) As Boolean

  'Copyright 2002 Roger Gilchrist
  'Use this if the RTF tag you are interested in does not
  'have a numerical component OR numeric component is known
  'WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING
  'WARNING WARNING *---PROGRAMMER MODIFICATION POINT---*  WARNING WARNING
  'WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING
  'THIS IS REALLY slow don't use it or anything that calls it in real world work
  'I 'm hoping to find a fast way to do this using API but only partial success so far

    With m_RTB
        HasRTFCode = InStr(.TextRTF, RTFCode$ & "\")
        If HasRTFCode = False Then
            HasRTFCode = InStr(.TextRTF, RTFCode$ & " ")
        End If
    End With 'M_RTB
    '                     End if RTF Code group                    embedded in RTF Code group

End Function

Private Function HasRTFCodeNum(RTFCode$) As Boolean

  'Copyright 2002 Roger Gilchrist
  'Use this if the RTF tag you are interested in has an unknown number on the end
  'WARNING WARNING *---PROGRAMMER MODIFICATION POINT---*  WARNING WARNING
  'THIS IS REALLY slow don't use it or anything that calls it in real world work

  Dim i As Integer

    For i = 0 To 100  '  this is probably larger than needed but
        ' check documentation for limit of rtf code
        If HasRTFCode(RTFCode$ & i) Then
            HasRTFCodeNum = True
            Exit For '>---> Next
        End If
    Next i

End Function

Public Sub HiddenTextShow()

  'Copyright 2002 Roger Gilchrist
  'If you know where the hidden text is you can select over that
  'piece of text and let the SelRTFToggle turn it off,
  'BUT if you can't remember where it is use this to reveal all hidden code
  ' You will have to hide it all again your self OR use ActOnTag
  '*---PROGRAMMER MODIFICATION POINT---*
  'if you are not using Hidden text or
  'don't want your end-user finding out about it(naughty?)
  'delete/modify access to this routine

  Dim RTF As String

    With m_RTB
        PositionStore Push
        RTF$ = .TextRTF
        DeleteRTFCode RTF$, "\v"
        If RTF$ <> .TextRTF Then '  eliminates unnecessary flicker
            .TextRTF = RTF$
        End If
        PositionStore Pop
    End With 'M_RTB

End Sub

Private Sub HighLighter(HiColr As Long)

  'Copyright 2002 Roger Gilchrist

    With m_RTB
        On Error Resume Next
            ColourStore True
            m_lastHiColor = HiColr
            If Not IsNull(.SelColor) Then
                m_lastHiForeColor = .SelColor
            End If
            .SelColor = HiColr
            SelRTFToggle "\highlight1"
            ColourStore False
        On Error GoTo 0
    End With 'M_RTB

End Sub

Private Sub HighLightHardAuto(HiColr As Long)

  'Copyright 2002 Roger Gilchrist

    HighLighter HiColr
    m_RTB.SelColor = ColourInvert(HiColr)

End Sub

Public Function InstantTran(w$) As String

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted
  'This is really not the way to do this but i wanted to do something with RichWordOverMod
  'the language is Bahasa Indonesia 'bahasa keduaku' (Indonesian to Eng speakers)
  'and yes I know this translation is bad, that's the problem with word for word translation.
  'It is designed to translate fully the sentence in the Demo doc which fills the cases
  '"You might use it as an instant translation system for bi-lingual programs."
  'more grammatically translates as
  'Mungkin Anda bisa memakainya untuk program-program yang terjemahkan antara dua bahasa secara instan.'
  'notice there is almost no word to word matching between a good translation and the original
  ' if you use this idea I hope you use something better than a Select case to generate the translation
  'either a data base or some sort of paired arrays

    Select Case LCase$(w$)
      Case "you"
        InstantTran = "Anda"
      Case "might"
        InstantTran = "mungkin"
      Case "use"
        InstantTran = "memakai"
      Case "it"
        InstantTran = "nya"
      Case "as"
        InstantTran = "sebagai"
      Case "an"
        InstantTran = "sesuatu"
      Case "instant"
        InstantTran = "instan"
      Case "translation"
        InstantTran = "terjemahkan"
      Case "system"
        InstantTran = "sistem"
      Case "for"
        InstantTran = "untuk"
      Case "bi-lingual" 'Minor update: RichWordOver sees hypans as word separators so can't hit this
        InstantTran = "dwi-bahasa"
      Case "bi"
        InstantTran = "dwi"
      Case "lingual"
        InstantTran = "bahasa"
      Case "programs"
        InstantTran = "program-program"
      Case ""
        InstantTran = "NO TRANSLATION" 'this hits on spaces and disguises the function until you hit a known word
      Case Else
        InstantTran = "NO TRANSLATION" 'this is for all unknown words
    End Select

End Function

Private Function InStrRTF(txt$, RCde$) As Boolean

  'Copyright 2002 Roger Gilchrist

    If InStr(txt, RCde$) Then ' yes but check in depth
        If InStr(txt, RCde$ & " ") Then
            InStrRTF = True 'end of RTFCode group
          ElseIf InStr(txt, RCde$ & "\") Then 'NOT INSTR(TXT,...
            InStrRTF = True 'embedded in RTFCode group
        End If
    End If

End Function

Public Sub KeyDown(KeyCode As Integer, Shift As Integer)

  '*---PROGRAMMER MODIFICATION POINT---*
  'this may be more power than you need/want( if that is possible)
  'remove the bits you don't want
  'Copyright 2002 Roger Gilchrist
  '*PURPOSE:read/enable hot keys from RTB

  Static Working As Boolean

    ''*---PROGRAMMER MODIFICATION POINT---*
    ''this test will speed things up but stops you assigning single Function keys
    '        If Shift = 0 And KeyCode <> vbKeyTab Then 'Ignore all single keys except TAB
    '            Exit Sub
    '        End If

    If Working Then ' don't double hit; may not be necessary but its safe
        Exit Sub '>---> Bottom
    End If
    If IsNull(m_RTB) Then ' saftey in case you hit before initializing
        Exit Sub '>---> Bottom
    End If
    ''*---PROGRAMMER MODIFICATION POINT---*
    ''this test will speed things up but stops you using Function keys etc if box is locked
    ''you might move KeysFunction and KeysOther above this test while blocking letters
    ''or move this test into the KeysXXX subs at apoint that is meaningful at
    '    If m_RTB.Locked Then ' if box is locked then keyboard access is blocked
    '        KeyCode = 0      ' remove this if you have things you want to access even if the RichTextBox is locked
    '        Exit Sub
    '    End If
    Working = True ' stops this routine doubling up on itself
    ' m_busy = True ' stop rest of class hitting while typing
    KeysFunction KeyCode, Shift
    KeysLetter KeyCode, Shift
    KeysOther KeyCode, Shift
    ' m_busy = False
    Working = False

End Sub

Private Sub KeyMacroSelectFromInsertion(Shift As Integer, Mode$)

  'Copyright 2002 Roger Gilchrist
  'based on the scan example in VB help
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if you don't want this functionality

  Dim StartAt As String, EndAt As String

    Select Case ShiftKeys(Shift) ' this test adds just a little speed
      Case KCtrl, KShiftCtrl, KCtrlAlt
        StartAt = ".?!:" & vbNewLine & vbTab
        EndAt = ".?!:" & vbNewLine & vbTab
        Select Case Mode$
          Case "P"
            ' replace default
            StartAt = vbNewLine
            EndAt = vbNewLine
            'Case "S": 'Default mode
          Case "W" 'Modify
            StartAt = StartAt & " ,;"
            EndAt = EndAt & " ,;"
          Case Else
            '            Exit Sub
        End Select
    End Select

    With m_RTB
        Select Case ShiftKeys(Shift)
          Case KCtrl
            .Span EndAt, True, True ' Select to the end of the word.
            .SelLength = .SelLength + 1
          Case KShiftCtrl
            .Span StartAt, False, True ' Select to the start of the word.
            .SelLength = .SelLength + 1
          Case KCtrlAlt
            .Span EndAt, True, True ' Select whole of the word.
            .Span StartAt, False, True
            .SelLength = .SelLength + 1
        End Select
        '.SelLength = .SelLength + 1 ''DO NOT DO THIS. It looks logical and neater at first but
        ' every time you type 'S' 'P' or 'W' you overwrite one character
        ' took ages to track down
    End With 'M_RTB

End Sub

Private Sub KeysFunction(KeyCode As Integer, Shift As Integer)

  'Copyright 2002 Roger Gilchrist
  'I have divided the Keys arbitarily into Function,Letter and Other for ease of reading
  ' but you only need to reproduce the parts of the Select case structrue you use
  '*---PROGRAMMER MODIFICATION POINT---*
  'DoSomething is purely for demo purposes DELETE IT
  '
  'CycleCase, CycleSelColor are demos too but you might want to keep them
  '
  'add Function key calls to your routines in this structure
  'as DoSomething shows you can violate modularity (call routines in bas files)
  '
  'vbKeyF1 contains the template for all of the other keys
  ' but you only need to reproduce the parts of the Select Case structure you need
  '
  'REMEMBER not everyone has F11 and F12 (date=2002 who knows if more will be added)
  'so program as if they don't exist or test for them (i'll get round to it soon:))

    Select Case KeyCode
        '      Case vbKeyF1                     'USE THIS AS A TEMPLATE FOR ADDING TO THE CLASS
        '        Select Case ShiftKeys(Shift)
        '          Case NoShift
        '          Case KShift
        '          Case KCtrl
        '          Case KShiftCtrl
        '          Case KAlt
        '          Case KShiftAlt
        '          Case KCtrlAlt
        '          Case KShiftCrtlAlt
        '        End Select
        '     Case vbKeyF2
      Case vbKeyF3
        Select Case ShiftKeys(Shift)
          Case KShift
            CycleCase
        End Select
      Case vbKeyF4
        Select Case ShiftKeys(Shift)
          Case KCtrlAlt
            DoSomething    'DEMO ONLY DELETE for real use.
        End Select
        '      Case vbKeyF5
        '      Case vbKeyF6
        '      Case vbKeyF7
        '      Case vbKeyF8
        '      Case vbKeyF9
      Case vbKeyF10
        Select Case ShiftKeys(Shift)
          Case KCtrl
            m_RTB.SelColor = CycleSelColor(m_RTB.SelColor)
        End Select
        'Case vbKeyF11 'are you sure  your end-user has this?
        'Case vbKeyF12 'are you sure  your end-user has this?

    End Select

End Sub

Private Sub KeysLetter(KeyCode As Integer, Shift As Integer)

  'Copyright 2002 Roger Gilchrist

  '*---PROGRAMMER MODIFICATION POINT---*
  'I have divided the Keys arbitarily into Function,Letter and Other for ease of reading
  '
  ' KeyMacroSelectFromInsertion in VbKeyS,VbKeyP,VbKeyW can be deleted if you dont want to use it
  '
  '----------------------------------------------------------------------
  'IMPORTANT READ the note for VbKeyI you may want to follow its advice
  '----------------------------------------------------------------------
  '
  'add Shift key calls to your routines in this structure
  'you can violate modularity (call routines in bas files)
  '
  'vbKeyA contains the template for all of the other keys
  ' but you only need to reproduce the parts of the Select Case structure you need
  '
  'REMEMBER not everyone has standard keyboards you need to program against this
  'using the virtual VbKey# is a step towards this

    With m_RTB
        Select Case KeyCode
            '          Case vbKeyA' USE THIS AS A TEMPLATE FOR ADDING TO THE CLASS
            '            Select Case ShiftKeys(Shift)
            '              Case NoShift
            '              Case KShift
            '              Case KCtrl
            '              Case KShiftCtrl
            '              Case KAlt
            '              Case KShiftAlt
            '              Case KCtrlAlt
            '              Case KShiftCrtlAlt
            '            End Select
          Case vbKeyB
            Select Case ShiftKeys(Shift)
              Case KCtrl
                .SelBold = Not (.SelBold)
            End Select

          Case vbKeyU
            Select Case ShiftKeys(Shift)
              Case KCtrl
                .SelUnderline = Not (.SelUnderline)
            End Select
          Case vbKeyI
            'NOTE: RichTextBox is preset to use Ctrl+I for tab insertion.
            'To override this and make Ctrl+I convert selected text to italics you cannot use
            'this routine. Instead you need to:-
            'have a menu entry with its Hot Key set to Ctrl+I. So even though you may never
            'use it you should create a 'Format' menu if you want to use the WordStar sytle
            'hot keys (this menu must be visible to work so don't try to hide it)
          Case vbKeyP
            KeyMacroSelectFromInsertion Shift, "P"
          Case vbKeyS
            KeyMacroSelectFromInsertion Shift, "S"
          Case vbKeyW
            KeyMacroSelectFromInsertion Shift, "W"
        End Select
    End With 'M_RTB

End Sub

Private Sub KeysOther(KeyCode As Integer, Shift As Integer)

  'Copyright 2002 Roger Gilchrist

  '*---PROGRAMMER MODIFICATION POINT---*
  'I have divided the Keys arbitarily into Function,Letter and Other for ease of reading
  'add other key calls to your routines in this structure
  'you can violate modularity (call routines in bas files)
  '
  '------------------------------------------------------------
  'IMPORTANT READ the notes at VbKeyTab you NEED to know this.
  '-------------------------------------------------------------
  '
  'vbKeyLeft contains the template for all of the other keys
  ' but you only need to reproduce the parts of the Select Case structure you need
  '
  'REMEMBER not everyone has standard keyboards you need to program against this
  'using the virtual VbKey# is a step towards this

    Select Case KeyCode
        '      Case vbKeyLeft ' USE THIS AS A TEMPLATE FOR ADDING TO THE CLASS
        '        Select Case ShiftKeys(Shift)
        '          Case NoShift
        '          Case KShift
        '          Case KCtrl
        '          Case KShiftCtrl
        '          Case KAlt
        '          Case KShiftAlt
        '          Case KCtrlAlt
        '          Case KShiftCrtlAlt
        '        End Select
        '      Case vbKeyRight
      Case vbKeyUp
        Select Case ShiftKeys(Shift)
          Case KShiftCtrl
            SelRTFToggle "\super"
        End Select
      Case vbKeyDown
        Select Case ShiftKeys(Shift)
          Case KShiftCtrl
            SelRTFToggle "\sub"
        End Select
        '      Case vbKeyEscape
      Case vbKeyTab
        Select Case ShiftKeys(Shift)
          Case NoShift 'Over ride default behaviour of RichTextBox with real world expectation [TAB] inserts Tab character
            KeyCode = 0
            With m_RTB
                If .SelLength = 0 Then
                    .SelText = vbTab
                End If
            End With 'M_RTB
          Case KShift
            With m_RTB ' Deletes Tab as cursor like IDE does
                KeyCode = 0
                If .SelLength = 0 Then
                    PositionStore Push
                    If vbTab = CurChar Then
                        CurChar = ""
                    End If
                    PositionStore Pop
                    .SelStart = .SelStart - 1
                End If
            End With 'M_RTB
          Case KCtrl
            'do nothing
            'this reproduces Tab to next control
            'which is the RichTextBox default with mnemonic name
            'Ctrl+TAB
            'See  Microsoft Knowledge Base Article
            'Q143273 - HOWTO: Insert TAB Characters in RichTextBox Control in VB All Products
            'CTRL+TAB is the offical in unnatural way to insert Tab into RichTextBox
            'The official way to allow typing Tab involves turning TabStop on and off for the
            'whole form from GotFocus and LostFocus but is kind of unstable in real world programming
        End Select
    End Select

End Sub

Public Sub LineNumbers()

  Dim lngCount As Long
  Dim lngLineIndex As Long
  Dim lngLength As Long
  Dim strBuffer As String
  Dim strRichText As String
  Dim i As Integer

    MsgBox "This doesn't work properly so it is blocked by this message untill I get it right.", , "LINENUMBERS"

Exit Sub

    '    'Get Line count
    '    lngCount = SendMessage(m_RTB.hwnd, EM_GETLINECOUNT, 0, 0)
    '    m_busy = True
    '    With m_RTB
    '        For i = 0 To lngCount - 1
    '            'Get line index
    '            lngLineIndex = SendMessage(.hwnd, EM_LINEINDEX, i, 0)
    '            'get line length
    '            lngLength = SendMessage(.hwnd, EM_LINELENGTH, lngLineIndex, 0)
    '            'resize buffer
    '            strBuffer = Space$(lngLength)
    '            'get line text
    '            Call SendMessageStr(.hwnd, EM_GETLINE, i, ByVal strBuffer)
    '            'Number each line
    '            Call SendMessageStr(.hwnd, EM_GETLINE, i, ByVal strBuffer)
    '
    '            '.SelText = CStr(i + 1) & " "
    '            strRichText = strRichText & CStr(i + 1) & " " & strBuffer & vbCrLf
    '        Next i
    '        'rewrite numbered text in RichTextBox
    '        .Text = strRichText
    '
    '    End With 'M_RTB
    '    m_busy = False

End Sub

Private Sub ManipulateRTFCode(Part$, CS$, Mode As Boolean)

    If Mode = False Then
        Select Case RTFPartType(CStr(Part))
          Case RTFOnly, Mixed ' only these need be treated
            DeleteRTFCodeX Part$, CS$
        End Select
      Else 'NOT MODE...
        Select Case RTFPartType(CStr(Part))
          Case RTFOnly
            Part = Part & CS$ ' add to end of Pure RTF it will override any counter or rival code
          Case Else
            Part = CS$ & " " & Part ' add to start of mixed or textonly so current char is affected.
        End Select
    End If

End Sub

Public Sub MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

  '*---PROGRAMMER MODIFICATION POINT---*
  'this just keeps the class up to date with mouse activitities
  ' at present only used by RichMouseOver and InstantTrans
  'delete if you don't need RichMouseOver

    m_MouseX = x
    m_MouseY = Y
    m_button = Button
    m_Shift = Shift

End Sub

Private Sub PositionStore(SaveTSetF As LIFO)

  'Simple PushPop store of SelStart and SelLength allows routines
  'which manipulate through .seXXXXXX to restore SelStart and SelLength
  'arrays allow nesting of calls as long as you always call push and pop in calling routines
  'MOST IMPORTANT it sets the m_busy variable which stops a lot of updating while involved
  'in manipulating .SelXXX stuff

  Dim insertPos As Long

    '*---PROGRAMMER MODIFICATION POINT---*
    'keep this even if you don't need the position storage
    With m_RTB
        If SaveTSetF Then
            m_busy = True
            .MousePointer = rtfHourglass
          Else 'SAVETSETF = FALSE
            m_busy = False
            .MousePointer = rtfDefault
        End If
    End With 'M_RTB
    '*---PROGRAMMER MODIFICATION POINT---*
    'this is only needed if you have something which actually uses .SelStart and .SelLength
    'as the class develops I will remove any need for it
    With m_RTB
        If SaveTSetF Then
            On Error Resume Next
                If IsArray(m_Start) Then 'initialize
                    insertPos = 0
                  Else 'ISNULL(M_START) = FALSE'ISARRAY(M_START) = FALSE
                    insertPos = UBound(m_Start)
                End If
                ReDim Preserve m_Start(insertPos + 1) As Long
                ReDim Preserve m_Len(insertPos + 1) As Long
                m_Start(UBound(m_Start)) = .SelStart
                m_Len(UBound(m_Len)) = .SelLength
              Else 'SAVETSETF = FALSE
                .SelStart = m_Start(UBound(m_Start))
                .SelLength = m_Len(UBound(m_Len))
                If UBound(m_Start) > 0 Then
                    ReDim Preserve m_Start(UBound(m_Start) - 1) As Long
                    ReDim Preserve m_Len(UBound(m_Len) - 1) As Long
                  Else 'NOT UBOUND(M_START)...
                    ReDim Preserve m_Start(0) As Long
                    ReDim Preserve m_Len(0) As Long
                End If
            End If
        End With 'M_RTB
    On Error GoTo 0

End Sub

Public Function RichWordOverMod() As String

  'Ready-To-Run Visual Basic Algorithms, Second Edition
  'http://www.vb-helper.com/vba.htm
  'souped up by using API IsCharAlphaNumeric Roger Gilchrist
  'removed parameters becasue MouseMove has made them available to routine
  'changed everything to long so it wouldn't overload

  Dim pt As POINTAPI
  Dim pos As Long
  Dim start_pos As Long
  Dim end_pos As Long
  Dim ch As String
  Dim txt As String
  Dim txtlen As Long

    If m_busy Then
        Exit Function '>---> Bottom
    End If
    ' Convert the position to pixels.
    pt.x = m_MouseX \ Screen.TwipsPerPixelX
    pt.Y = m_MouseY \ Screen.TwipsPerPixelY

    ' Get the character number
    pos = SendMessage(m_RTB.hwnd, EM_CHARFROMPOS, 0&, pt)
    If pos > 0 Then
        ' Find the start of the word.
        txt = m_RTB.Text
        For start_pos = pos To 1 Step -1
            ch = Mid$(m_RTB.Text, start_pos, 1)
            ' Allow digits, letters, and underscores.
            If Not (IsCharAlphaNumeric(CByte(Asc(ch))) > 0 Or ch = "_") Then
                Exit For '>---> Next
            End If
        Next start_pos
        start_pos = start_pos + 1

        ' Find the end of the word.
        txtlen = Len(txt)
        For end_pos = pos To txtlen
            ch = Mid$(txt, end_pos, 1)
            ' Allow digits, letters, and underscores.
            If Not (IsCharAlphaNumeric(CByte(Asc(ch))) > 0 Or ch = "_") Then
                Exit For '>---> Next
            End If
        Next end_pos
        end_pos = end_pos - 1

        If start_pos <= end_pos Then
            RichWordOverMod = Mid$(txt, start_pos, end_pos - start_pos + 1)
            'a real instant translation program would look up here
            '
        End If
    End If

End Function

Private Function RTFCodeLoc(TmpRTF$, RTFCode$) As RTFCodeLocs

  'Copyright 2002 Roger Gilchrist
  'WHERE THE RTFCODE OCCURS WITHIN A SelRTF string
  'A single space at the end of RTFCode is the END OF RTFCode Token
  'A double space is an RTFToken followed by a blank text character
  'This routine allows you to judge which format the code code is using
  ' and act accordingly

    If InStr(TmpRTF$, RTFCode$ & "\") Then
        RTFCodeLoc = EmbedOr1st
      ElseIf InStr(TmpRTF$, RTFCode$ & " ") Then 'NOT INSTR(TMPRTF$,...
        If InStr(TmpRTF$, RTFCode$ & "  " & "}") Then
            RTFCodeLoc = LastWithBlank
          ElseIf InStr(TmpRTF$, RTFCode$ & " " & "}") Then 'NOT INSTR(TMPRTF$,...
            RTFCodeLoc = LastInSelection
          Else 'NOT INSTR(TMPRTF$,...
            RTFCodeLoc = LastBeforeText
        End If
      ElseIf Right$(TmpRTF$, Len(RTFCode$)) = RTFCode$ Then 'NOT INSTR(TMPRTF$,... 'NOT INSTR(TMPRTF$,...
        RTFCodeLoc = EndOfString
      Else 'NOT RIGHT$(TMPRTF$,...
        RTFCodeLoc = NotPresent
    End If

End Function

Private Function RTFise(t$) As String

  'Copyright 2002 Roger Gilchrist
  'Convert Text to RTFCode format for searching in SelRTF or TextRTF
  'does not deal with incompletely formated words just basic characters.

    t$ = Replace(t$, "\", "\\")
    t$ = Replace(t$, "{", "\{")
    t$ = Replace(t$, "}", "\}")
    t$ = Replace(t$, vbNewLine, "\par")
    t$ = Replace(t$, vbTab, "\tab")
    RTFise = t$

End Function

Private Function RTFPartType(s$) As RTFPartsType

  'Copyright 2002 Roger Gilchrist
  'work out what type of string you're dealing with

    If Len(s$) = 0 Then
        RTFPartType = Blank
      ElseIf InStr("\}{" & vbTab & vbNewLine, Left$(s$, 1)) Then 'NOT LEN(S$)...
        RTFPartType = RTFOnly
      ElseIf InStr(s$, "\") = 0 Then 'NOT INSTR("\}{"...
        If InStr(s$, "}") = 0 Then
            RTFPartType = TextOnly

          Else 'NOT INSTR(S$,...
            RTFPartType = Mixed
        End If

      Else 'NOT INSTR(S$,...
        If Left$(s$, 4) = "uc1\" Then
            RTFPartType = RTFOnly     'It is the first part of the MixedBit string in SelRTFToggle
          Else 'NOT LEFT$(S$,...
            RTFPartType = Mixed      '
        End If
    End If

End Function

Private Function RTFPartTypeStr(x As RTFPartsType) As String

  'Copyright 2002 Roger Gilchrist
  ' mostly for debug purposes

    Select Case x
      Case Blank
        RTFPartTypeStr = "Blank"
      Case RTFOnly
        RTFPartTypeStr = "RTFOnly"
      Case TextOnly
        RTFPartTypeStr = "TextOnly"
      Case Mixed
        RTFPartTypeStr = "Mixed"
    End Select

End Function

Private Function SafeJoin(Bits, Optional Del$ = " ") As String

  'Copyright 2002 Roger Gilchrist
  '*PURPOSE: Replace Join with the ability to recognise whether a blank
  'was in original or deleted by editing of member
  'and not insert a del$ character for deleted member
  'SEE ALSO SafeJoinTest

  Dim i As Long

    For i = LBound(Bits) To UBound(Bits)
        If Bits(i) <> SafeJoinSkip Then ' if marked not to reattach then don't add
            SafeJoin = SafeJoin & Bits(i) & Del$
        End If
    Next i
    Do While Right$(SafeJoin, Len(Del$)) = Del$ '' make sure last char<>del$
        SafeJoin = Left$(SafeJoin, Len(SafeJoin) - Len(Del$))
    Loop

End Function

Private Function SafeJoinTest(p, t$) As Variant

  'Copyright 2002 Roger Gilchrist
  '*PURPOSE: Deal with 3 possible events for a split array member
  '1. no edit               accept
  '2. changed content       accept
  '3. deleted               reset as SafeJoinSkip
  'if 3 then skips rather than adding a delimiter otherwise just
  'SEE ALSO SafeJoin

    If t$ <> "" Then
        SafeJoinTest = t$
      ElseIf p = t$ Then 'NOT T$...
        SafeJoinTest = t$
      Else 'NOT P...
        SafeJoinTest = SafeJoinSkip
    End If
    'SafeJoinTest = IIf(t$ <> "" Or p = t$, t$, SafeJoinSkip)

End Function

Public Property Let SelQuoteDbl(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    ExtendedQuoteToggle ("\dbquote")

End Property

Public Property Get SelQuoteDbl() As Variant

  'Copyright 2002 Roger Gilchrist

    SelQuoteDbl = GetGetValue("\dbquote")

End Property

Public Property Get SelQuoteSng() As Variant

  'Copyright 2002 Roger Gilchrist

    SelQuoteSng = GetGetValue("\squote")

End Property

Public Property Let SelQuoteSng(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    ExtendedQuoteToggle ("\squote")

End Property

Public Function SelRTFToggle(code$) As Boolean

  'Copyright 2002 Roger Gilchrist

  Dim RTFHead As String     'Initial section of RTF probably never useful
  Dim Parts As Variant      'Split MixedBit for application of code
  Dim FntTblStart As Long, ClrTblStart As Long, RTF_End As Long ' Cutpoints for various pieces
  Dim i As Integer          'For variable to manipulate Parts
  Dim TMp As String          'tmp for testing whether "" is original or edited

    With m_RTB
        ' If .SelLength > 0 Then
        If Left$(code$, 3) = "\ul" Then ' this is not essential but stores latest underline style
            m_lastUndeline = code$      ' allowing toolbar dropdown button to use default if you click
        End If                          ' button instead of using dropdown
        PositionStore Push
        SelRTFToggle = Not (RTFCodeLoc(WorkingStringRead, code$) <> NotPresent) ' Code is present so remove
        BreakUpRTFString RTFHead, Parts
        For i = LBound(Parts) To UBound(Parts)
            TMp$ = Parts(i)
            ManipulateRTFCode TMp$, code$, SelRTFToggle
            Parts(i) = SafeJoinTest(Parts(i), TMp$)
        Next i
        WorkingStringWrite RTFHead & SafeJoin(Parts, " ")
        PositionStore Pop
        '  End If
    End With 'M_RTB

End Function

Private Function ShiftKeys(ByVal Shift As Integer) As Shifts

  'VB6 help file

    ShiftKeys = Shift And 7

End Function

Public Property Let ShowFileInCaption(ByVal ShwFName As Boolean)

  'write only Default = True
  'see class Initialize to change Default

    m_FileInCaption = ShwFName

End Property

Private Sub SimpleToggle(FindAdd$, RepAdd$, FindDel$, RepDel$)

  'Copyright 2002 Roger Gilchrist
  ' used where RTF is simple matter of insert/ remove code token

  Dim TmpRTF As String

    With m_RTB
        TmpRTF = .TextRTF
        If InStr(TmpRTF, FindDel$) = 0 Then
            TmpRTF = Replace(TmpRTF, FindAdd$, RepAdd$)
          Else 'NOT INSTR(TMPRTF,...
            TmpRTF = Replace(TmpRTF, FindDel$, RepDel$)
        End If
        .TextRTF = TmpRTF
    End With 'M_RTB

End Sub

Private Sub SwapAnyThing(element1, element2)

  'ancient simple idea
  'Swap any two items in any format

  Dim TEMP As Variant

    TEMP = element1
    element1 = element2
    element2 = TEMP

End Sub

Public Property Get VisibleParagraphMarks() As Boolean

  'Copyright 2002 Roger Gilchrist
  'detect visible paragraph marker

    VisibleParagraphMarks = GetGetValue("\'b6")

End Property

Public Property Let VisibleParagraphMarks(ByVal vNewValue As Boolean)

  'Copyright 2002 Roger Gilchrist
  'add/remove visible paragraph marker

    SimpleToggle "\par", "\'b6\par", "\'b6", ""

End Property

Private Function WorkingStringRead(Optional ForceDoc As Boolean = False) As String

  'Copyright 2002 Roger Gilchrist
  'Get a string for routines to work on
  'and generate Private m_End_OF_RTF to report end of RTF/Start of Text

  Dim fsLoc As Long, GotTrueEnd As Boolean

    With m_RTB
        If .SelLength = 0 Or ForceDoc Then
            WorkingStringRead = .TextRTF
          Else '.SELLENGTH = FALSE'NOT .SELLENGTH...
            WorkingStringRead = .SelRTF
        End If
    End With 'M_RTB

    fsLoc = InStr(1, WorkingStringRead, "\fs")          'find RTF size code and move test point
    If fsLoc = 0 Then                     'not found = there is no selected Text
        m_End_OF_RTF = 0                        'so exit
        Exit Function '>---> Bottom
    End If
    Do
        If IsNumeric(Mid$(WorkingStringRead, fsLoc + 4, 1)) Then   ' check is size not "fswiss" or other RTF code
            m_End_OF_RTF = InStr(fsLoc, WorkingStringRead, " ")      ' find space delimiting size code
            GotTrueEnd = True
        End If
        fsLoc = InStr(fsLoc + 1, WorkingStringRead, "\fs")         ' find next \fs code and move test point
    Loop Until GotTrueEnd Or fsLoc = 0
    If GotTrueEnd = False And fsLoc = 0 Then              ' Safety should never hit
        m_End_OF_RTF = 0
    End If

End Function

Private Sub WorkingStringWrite(WStr$)

  'Copyright 2002 Roger Gilchrist
  'put a string from a routine back into the selection or whole control

    With m_RTB
        If .SelLength Then
            If .SelRTF <> WStr Then
                .SelRTF = WStr
            End If
          Else '.SELLENGTH = FALSE
            If .TextRTF <> WStr Then
                .TextRTF = WStr
            End If
        End If
    End With 'M_RTB

End Sub

':) Ulli's VB Code Formatter V2.13.6 (28/08/2002 2:37:31 PM) 181 + 2160 = 2341 Lines

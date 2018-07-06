VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsRTFFontPainter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ClsRTFFontPainter
'copyright 2002 Roger Gilchrist
'This class modifies the TextRTF string to create various colour and text effects on text in RichTextboxes
'All colour routines contain an Optional Boolean switch ForeBack which determines whether
'text(True) or the background(False) is coloured.
'I have created most of the colour sets myself but the following sources served as inspiration
' and constructed the math behind the routines (I have rewritten them into spearate routines)
'Routine:      | Source
'-------------------------------------------------------------------------------------
'Blend:-        Dan Redding (bwsoft@revealed.net)
'RainbowColor:- DrawRainBow © oigres P Email: oigres@postmaster.co.uk
'Materials:-    John Colman (john_colman@hotmail.com)in metalCBprj.OCX
'
'------------IMPORTANT INFORMATION-------------------------------
'there is a problem with supporting highlighting (which is used for back colouring in this class)
'This requires the Riched20.dll (version 3)
'This also requires Riched32.dll (5.00.2008.1)
'and probably the Richx32.ocx control
'VB6 SP5 (what I use) has no difficult with this

'thanks to Thomas Görtler (maniac@ colorarts.de) who let me know of the problem then solved it himself
'and Sergio Perciballi (oigres@postmaster.co.uk) whose code contained the detailed version stuff above

'If you construct an interesting variant or expand the colour set of Materials
'send me details and I'll probably add it to the next edition of the code
'and add you to the list above

'Search this routine for the following line
'*---PROGRAMMER MODIFICATION POINT---*
'it marks points at which you might edit or delete
'If you modify other routines you will need to be
'very careful they are usually vital to the class

'I'm a bit compulsive about modularity so each module has its own
'Private calls to to the API. This means if you copy just one module
'you don't have to keep coming back to find that missing Public variable
'in another module. Once you have inculde a module in your program you can
'remove the Private Declare statements if you have a Public one in your project

'copyright 2002 Roger Gilchrist
'rojagilkrist@hotmail.com
Option Explicit
Private m_DefFormat As String
Private m_CDlg As CommonDialog
Private m_busy As Boolean
Private m_RTB As RichTextBox
Private m_MaterialsFileName As String
Private m_StylesFileName As String
Private m_End_OF_RTF As Long
Private Const DQ As String = """"      ' DB=DoubleQuote NL=NewLine
Private Const NL As String = vbNewLine ' I use these for long MsgBoxes layout.
Private m_Start() As Long
Private m_Len() As Long
Private m_LastHighlightColour As Long
Private m_LastHighlightForeColour As Long
Private m_LastBlendStartColour As Long
Private m_LastBlendEndColour As Long
Private m_LastTextColour As Long
Private m_LastDitherColour As Long
Private m_LastFuzzyColour As Long
Private Const SafeJoinSkip As String = "~!~!~!*~!"
Private Const MaterialsSep As String = "|"
Private Const StylesSeparator As String = "^"
Private Enum LIFO
    Push = True
    Pop = False
End Enum

Private Enum RTFCodeLocs
    NotPresent
    EmbedOr1st
    LastWithBlank
    LastInSelection
    LastBeforeText
    EndOfString
End Enum

Public Enum Spectrum
    s1RedYellow = 0
    s2YellowGreen
    s3GreenCyan
    s4CyanBlue
    s5BlueMagenta
    s6Magentared
End Enum
'
Private Enum SpectrumMode
    LeftRight = 0
    RightLeft = 1
    InOutLeftRight
    InOutRightLeft
End Enum

Private Enum RTFPartsType
    Blank
    RTFOnly
    TextOnly
    Mixed
End Enum
Rem Mark Off
'stops Ulli's Code formatter from noticing these as Duplicated Name without Scope or TypeCasting
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim Push
Dim Pop
#End If 'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
#If False Then
Dim NotPresent
Dim EmbedOr1st
Dim LastWithBlank
Dim LastInSelection
Dim LastBeforeText
Dim EndOfString
#End If
#If False Then
Dim s1RedYellow
Dim s2YellowGreen
Dim s3GreenCyan
Dim s4CyanBlue
Dim s5BlueMagenta
Dim s6Magentared
#End If
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim LeftRight
Dim InOutLeftRight
Dim RightLeft
Dim InOutRightLeft
#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
#If False Then
Dim Blank
Dim RTFOnly
Dim TextOnly
Dim Mixed
#End If
Rem Mark On
Private Const WM_USER As Long = &H400
Private Const CFM_BACKCOLOR = &H4000000
Private Const EM_GETCHARFORMAT As Long = (WM_USER + 58)
Private Const EM_SETCHARFORMAT As Long = (WM_USER + 68)
Private Const SCF_SELECTION = &H1&

Private Const LF_FACESIZE As Integer = 32
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
Private Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Sub About()

  'just a basic help I try to include in all my classes
  'it usually catches the eye when people start using the class
  'Comment out if it bugs you

  Dim Msg As String

    Msg = "Not settled enough for a hard coded description yet. You'll have to read the code itself. This class is behind the Panels on the Font Looks menu."
    MsgBox Msg, , "ClsRTFFontPainter"

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

Private Function ArrayInOut(sArray) As Variant

  Dim tmpA() As Long
  Dim i As Long
  Dim LocalCounter As Long
  Dim TopOfArray As Long

    TopOfArray = LBound(sArray) + UBound(sArray)
    ReDim tmpA(LBound(sArray) To UBound(sArray))
    LocalCounter = LBound(sArray)                   'because array bounds are selstart and selstart+sellength
    For i = LBound(sArray) To UBound(sArray) Step 2 ' step through array at double speed
        tmpA(LocalCounter) = sArray(i)              'Add to bottom and top of TmpA
        tmpA(TopOfArray - LocalCounter) = sArray(i)
        LocalCounter = LocalCounter + 1
    Next i
    ArrayInOut = tmpA

End Function

Private Function ArrayInvert(sArray) As Variant

  'support code for Colour style files
  'use anywhere you want to reverse a colour array

  Dim tmpA() As Long
  Dim i As Long, LocalCounter As Long

    ReDim tmpA(LBound(sArray) To UBound(sArray))
    LocalCounter = LBound(sArray) 'because array bounds are selstart and selstart+sellength
    For i = UBound(sArray) To LBound(sArray) Step -1
        tmpA(LocalCounter) = sArray(i)
        LocalCounter = LocalCounter + 1
    Next i

    ArrayInvert = tmpA

End Function

Private Function ArrayQuickSort(TArray) As Variant

  'support code for Colour style files
  'use if you want to alpha sort an array

    QuickSort TArray, LBound(TArray), UBound(TArray)
    ArrayQuickSort = TArray

End Function

Public Sub AssignControls(R As RichTextBox, C As CommonDialog)

  'Copyright 2002 Roger Gilchrist
  'Place the Call to this in Form_Load or Sub Main
  'RTBLooks.AssignControls RichTextBox1, CommonDialog1

    m_busy = False
    Set m_CDlg = C
    Set m_RTB = R
    m_DefFormat = GetDefFormatString

    End Sub

Private Function Blend(RGB1 As Long, RGB2 As Long, Percent As Single) As Long

  'Dan Redding (bwsoft@revealed.net)
  'http://home.revealed.net/bwsoft
  'developed this code which I have modified to use my ColourLong2RGB
  '-----------------------------------------------------------
  'These are his original notes for this routine
  '----------------------------------------------------------
  'This one doesn't really use the HSL routines, just the
  'RGB Component routines.  I threw it in as a bonus ;)
  'Takes two colors and blends them according to a
  'percentage given as a Single
  'For example, .3 will return a color 30% of the way
  'between the first color and the second.
  '.5, or 50%, will be an even blend (halfway)
  'Can create some nice effects inside a For loop
  '-----------------------------------------------------------
  'NOTE the Percent has to be expressed as a value between 0 and 1
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim R As Integer, R1 As Integer, R2 As Integer
  Dim G As Integer, G1 As Integer, G2 As Integer
  Dim B As Integer, B1 As Integer, B2 As Integer

    If Percent >= 1 Then
        Blend = RGB2
        Exit Function '>---> Bottom
      ElseIf Percent <= 0 Then 'NOT PERCENT...
        Blend = RGB1
        Exit Function '>---> Bottom
    End If

    ColourLong2RGB RGB1, R1, G1, B1
    ColourLong2RGB RGB2, R2, G2, B2

    R = ((R2 * Percent) + (R1 * (1 - Percent)))
    G = ((G2 * Percent) + (G1 * (1 - Percent)))
    B = ((B2 * Percent) + (B1 * (1 - Percent)))

    Blend = RGB(R, G, B)

End Function

Public Sub Blender(LeftRight As Boolean, InOut As Boolean, ForeBack As Boolean)

  'Copyright 2002 Roger Gilchrist
  'User interactive wrapper for BlenderHard
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

    BlenderHardHard ColourUser, ColourUser, LeftRight, InOut, ForeBack

End Sub

Public Sub BlenderAuto(LeftRight As Boolean, InOut As Boolean, ForeBack As Boolean)

  'Copyright 2002 Roger Gilchrist
  'User interactive wrapper for BlenderHard
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim BClr As Long

    BClr = ColourUser
    BlenderHardHard BClr, InvertColor(BClr), LeftRight, InOut, ForeBack

End Sub

Public Sub BlenderHardAuto(CStart As Long, LeftRight As Boolean, InOut As Boolean, ForeBack As Boolean)

  'Copyright 2002 Roger Gilchrist
  'This Sub creates an array of Blended colours for the ColourApplicator
  'Designed to provide a hard coded version of Blender
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

    BlenderHardHard CStart, InvertColor(CStart), LeftRight, InOut, ForeBack

End Sub

Public Sub BlenderHardHard(CStart As Long, CEnd As Long, LeftRight As Boolean, InOut As Boolean, ForeBack As Boolean)

  'Copyright 2002 Roger Gilchrist
  'This Sub creates an array of Blended colours for the ColourApplicator
  'Designed to provide a hard coded version of Blender
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim i As Long 'For Next driver
  Dim SPos As Long, EPos As Long 'Range of Text and hence LngClrArray
  Dim SelLen As Long, CurMember As Long 'Length of selection and CurrentMember of selection
  Dim LngClrArray() As Long 'Store colours for passing to ColourApplicator

    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    m_LastBlendStartColour = CStart 'used by the colourPanel to make life easier
    m_LastBlendEndColour = CEnd
    'This creates a blend from Cstart to Cend
    For i = SPos To EPos
        LngClrArray(i) = Blend(CStart, CEnd, PercentDecimal(CurMember, SelLen))
        CurMember = CurMember + 1
    Next i

    ColourApplicator LngClrArray, LeftRight, InOut, ForeBack

End Sub

Public Property Get Busy() As Variant

  'Read only use to stop updates in external program if class is working hard

    Busy = m_busy

End Property

Public Sub Candy(Range As Integer, Optional Spread As Integer = 0, Optional LeftRight As Boolean = False, Optional InOut As Boolean = False, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  ' variant of rainbow discovered by accident while developing Rainbow
  'just too pretty to dump so I modifed it as a separate routine
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed, play with it to develop other colour sets

  Dim i As Long
  Dim SPos As Long, EPos As Long
  Dim RbowMod As Long, Rstep As Long, CurChar As Long
  Dim LngClrArray() As Long
  Dim SelLen As Long

    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    RbowMod = (SelLen) \ 5
    For i = SPos To EPos

        Select Case Spread
          Case 0
            Rstep = Rstep + RbowMod
          Case 1
            Rstep = 255 * PercentDecimal(CurChar + RbowMod, EPos)
          Case 2
            Rstep = 255 * CurChar / RbowMod
        End Select
        CurChar = CurChar + 1
        LngClrArray(i) = RainbowColor(CInt(Range), Rstep Mod 255)
    Next i

    ColourApplicator LngClrArray, LeftRight, InOut, ForeBack

End Sub

Private Sub Class_Initialize()

    m_MaterialsFileName = App.Path & "\" & "Materials.dat"

    If Not ExistFile(m_MaterialsFileName) Then
        ' allows you to move the class code without the supporting materials.dat file
        MaterialsRestore True
    End If
    m_StylesFileName = App.Path & "\" & "Styles.dat"

    If Not ExistFile(m_StylesFileName) Then
        ' allows you to move the class code without the supporting 'styles.dat' file
        StylesRestore True
    End If

End Sub

Private Function ColorRTFCode(col As Long) As String

  'Copyright 2002 Roger Gilchrist
  'creates an RTF colour code member
  'support for ColourManipulation routine

  Dim R As Integer, G As Integer, B As Integer

    ColourLong2RGB col, R, G, B

    ColorRTFCode = "\red" & R & "\green" & G & "\blue" & B & ";"

End Function

Private Sub ColourApplicator(LngClrArray, LeftRight As Boolean, InOut As Boolean, ForeBack As Boolean, Optional RndPos As Boolean = False)

  'Copyright 2002 Roger Gilchrist
  ' this is the core of the various coloured text routines

  Dim RTFHead As String 'hold the part of the selection which is not changed(unless it contains a colortbl)
  Dim clrTbl As String 'Extract colortbl from RTFHead if it exists or fake it otherwise
  Dim OrigclrTbl As String 'store the ColorTbl so that it can be Replaced
  Dim Parts As Variant 'hold and manipulate the part of the selection which can be changed
  Dim curcolr As Long  'keep track of position in LngClrArray
  Dim i As Long        'For Next through the Selection.
  Dim Part As Variant

    If LeftRight Then  ' or LightDark invert spectrum
        LngClrArray = ArrayInvert(LngClrArray)
    End If

    If InOut Then 'Inout spectrum 'Always do this after inversion
        LngClrArray = ArrayInOut(LngClrArray)
    End If

    PositionStore Push '
    clrTbl = GetColorTable()    '
    OrigclrTbl = clrTbl         'store for later Replace call
    If Len(clrTbl) = 0 Then ' if No colorTbl create a new one with auto first value ";"
        clrTbl = "{\colortbl{;"
      Else 'else strip the end off the existing one'NOT LEN(CLRTBL)...
        clrTbl = Left$(clrTbl, Len(clrTbl) - 3)
    End If
    SeparateRTFString RTFHead, Parts
    curcolr = LBound(LngClrArray) 'get first colour from array
    'For i = LBound(Parts) To UBound(Parts)  ')
    For Each Part In Parts
        If RndPos Then
            If Rnd > 0.5 Then
                ForeBack = True
              Else 'NOT RND...
                ForeBack = False
            End If
        End If
        'ColourManipulate Parts(i), clrTbl, LngClrArray, curcolr, ForeBack
        ColourManipulate Part, clrTbl, LngClrArray, curcolr, ForeBack
        Parts(i) = Part
        i = i + 1
    Next Part
    clrTbl = clrTbl & "}  " 'finish colortbl
    If InStr(RTFHead, "{\colortbl") Then 'Replace Original colortbl with new one
        RTFHead = Replace(RTFHead, OrigclrTbl, clrTbl)
      Else ' if there was no original colortbl add it and another "}" to end it'NOT INSTR(RTFHEAD,...
        RTFHead = Left$(RTFHead, Len(RTFHead) - 1) & clrTbl & "}"
    End If
    WorkingStringWrite Trim$(RTFHead) & Join(Parts)
    PositionStore Pop

End Sub

Private Sub ColourLong2RGB(ByVal col As Long, Optional ByRef Red As Integer, Optional ByRef Green As Integer, Optional ByRef Blue As Integer)

  'don't know where I got this but there are a million variations around

    Red = col Mod 256
    Green = ((col And &HFF00FF00) / 256)
    Blue = (col And &HFF0000) / (65536)

End Sub

Private Sub ColourManipulate(Part, clrTbl$, LngClrArray, curcolr As Long, ForeBack As Boolean)

  'Copyright 2002 Roger Gilchrist
  'does the actual work of inserting colour codes into colortbl and rtf code

  Dim ClrCount As Long 'get current number of colors in colortbl
  Dim i As Long ' loop driver' probably doesn't need to be a long but its safer
  Dim tmpP As String 'bulid new string to replace Part
  Dim TmpRTF As String

    ClrCount = CountOccurances(clrTbl, "\red") 'get length of ColorTbl
    If ForeBack Then
        TmpRTF = "\cf"
      Else 'FOREBACK = FALSE
        TmpRTF = "\highlight"
    End If
    Select Case RTFPartType(CStr(Part))
      Case Blank 'blanks have to absorb an array member
        tmpP$ = tmpP$ & TmpRTF & ClrCount + 1 & " "
        clrTbl$ = clrTbl$ & ColorRTFCode(CLng(LngClrArray(curcolr)))
        curcolr = curcolr + 1
      Case RTFOnly
        tmpP = Part
        ' curcolr = curcolr - 1
      Case Mixed, TextOnly
        'add a colour to each letter and advance colour counter for text letter but ignore RTF code
        For i = 1 To Len(Part)
            If InStr("\}", Mid$(Part, i, 1)) = 0 Then
                tmpP$ = tmpP$ & TmpRTF & ClrCount + i & " " & Mid$(Part, i, 1)
                clrTbl$ = clrTbl$ & ColorRTFCode(CLng(LngClrArray(curcolr)))
                curcolr = curcolr + 1
              Else 'NOT INSTR("\}",...
                'becuase of the way Part is generated once you reach an RTF code token there is not more text
                tmpP$ = tmpP$ & Mid$(Part, i)
                curcolr = curcolr - 1
                Exit For '>---> Next
            End If
        Next i
        curcolr = curcolr + 1 'delete one member for blank space
    End Select
    ' curcolr = curcolr + 1
    Part = tmpP

End Sub

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

Public Sub ColourRemove(Optional ForeBack As Boolean)

  'Copyright 2002 Roger Gilchrist
  'This strips all occurrances of "\cf#" or "\highlight#"
  'by getting a maximim possible # and counting down
  'Counting up misfires as "\cf1" would convert "\cf10" to "0"
  'leaving a visible character "0" in the text

  Dim RTFHead As String 'hold the part of the selection which is not changed the colortbl is changed but the RTF engine does this, so you don't have to treat it programmatically
  Dim Parts As Variant 'hold and manipulate the part of the selection which can be changed
  Dim i As Long ', tmp As String
  Dim TMp As String ' convert Parts to string and back for ease of manipulation by the SafeJoinTest
  Dim Target As String ' hold the type of colour to remove
  Dim NumberPart As Long 'current number value for test
  Dim InitNumberParts As Long 'Maximum possible number value because SelRTf is renumbered as a subset of TextRTF

    PositionStore Push '
    SeparateRTFString RTFHead, Parts
    If ForeBack Then
        Target$ = "\cf"
      Else 'FOREBACK = FALSE
        Target$ = "\highlight"
    End If
    InitNumberParts = CountOccurances(Join(Parts), Target) * 4 'allows for '\cf' and '\highlight' contributing to colortbl and a lot of safety
    If InitNumberParts = 0 Then
        Exit Sub '>---> Bottom
    End If

    For i = LBound(Parts) To UBound(Parts)
        NumberPart = InitNumberParts
        TMp = Parts(i)
        Do While InStr(TMp, Target$)
            TMp$ = Replace(TMp$, Target & NumberPart, Target & "0")
            NumberPart = NumberPart - 1
            If NumberPart < 0 Then
                Exit Do '>---> Loop
            End If
        Loop
        Parts(i) = SafeJoinTest(Parts(i), TMp)
    Next i
    WorkingStringWrite Trim$(RTFHead) & SafeJoin(Parts)
    PositionStore Pop

End Sub

Public Sub ColourRemoveAll()

  'This is just a wrapper for two calls to ColourRemove
  'to hit both fore and back colours

    ColourRemove True
    ColourRemove False

End Sub

Public Sub ColourTEMPLATE(Optional LeftRight As Boolean = True, Optional InOut As Boolean = False, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'this is a do nothing Template Sub
  'POSSIBLE INPUTS
  'ForeBack  = Text or Background colour  ALWAYS USE It is not usually used in your colour modification but tells ColourApplication where to work
  'LeftRight = if colour array is ordered use this to reverse the order
  'LightDark = same as LeftRight but more descriptive if the order is basically intensity rather than colour
  'Inout     = use with any LeftRight or LightDark set to create a L-R-L or L-D-L pattern rather than the standard L-R or L-D pattern

  Dim i As Long ' used to step through Selection
  Dim SPos As Long, EPos As Long, SelLen As Long 'Range of Text and LngClrArray;distance between SPos and EPos
  Dim CurMember As Long 'counter 0 to SelLen
  Dim PercentDone As Integer 'an alternative way of knowing where you are in the selection
  Dim LngClrArray() As Long 'Store colours for passing to ColourApplicator

    'NOTE Spos and Epos should not be used to generate colours
    '(it would change depending on position in the document)

    'ALL TEMPLATES NEED To Get Selection Start and End
    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    For i = SPos To EPos
        '*---PROGRAMMER MODIFICATION POINT---*
        'this is where you generate your colour values in what ever way you wish
        'these are a few numbers you might like to use
        CurMember = CurMember + 1 'see notes in Dim statements for why this is here
        PercentDone = Percent(CurMember, SelLen) '0-100
        PercentDone = PercentDecimal(CurMember, SelLen) '0-1

        'LngClrArray(i)= do someting to generate a long colour value

    Next i

    ColourApplicator LngClrArray, LeftRight, InOut, ForeBack

End Sub

Public Function ColourUser() As Long

  'creator(?) Roger Gilchrist
  'don't know where I got this, may have worked it out myself,
  'have been using it for ages when I need user's color input
  'This  has been left Public so that the programmers can access
  'to get colours elsewhere in their programs
  'left Public so you can use it elsewhere in your program
  'change to Private if you don't want it outside this class

    On Error GoTo error_cancel
    With m_CDlg
        '.DialogTitle = 'NO! It would be nice to set it to
        ' "Highlight Colour|Font Colour"
        ' so of course it is not supported
        'you'll just have to make it clear in documentation
        'USE THESE IF YOU WISH TO ENHANCE/LIMIT UserColour
        '        .Flags
        '               cdlCCFullOpen           &H2 Entire dialog box is displayed, including the Define Custom Colors section.
        '               cdlCCHelpButton         &H8 Causes the dialog box to display a Help button.
        '               cdlCCPreventFullOpen    &H4 Disables the Define Custom Colors command button and prevents the user from defining custom colors.
        '               cdlCCRGBInit            &H1 Sets the initial color value for the dialog box.
        '.Flags = cdlCCRGBInit
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

Public Function Descriptor(StyleName As String, SubStyleName As String, SubStyleValue As Integer, StylePrivate As Integer, Clr1 As Long, Clr2 As Long, LeftRight As Boolean, InOut As Boolean, TextBack As Boolean, Optional blnHarden As Boolean = False) As String

    If StyleName = "styles" Then
        Descriptor = SubStyleName & MaterialsSep & MaterialsSep
      Else 'NOT STYLENAME...
        Descriptor = StyleName & MaterialsSep & SubStyleName & MaterialsSep
    End If
    Descriptor = LCase$(Descriptor & SubStyleValue & MaterialsSep & StylePrivate & MaterialsSep & Clr1 & MaterialsSep & Clr2 & MaterialsSep & LeftRight & MaterialsSep & InOut & MaterialsSep & TextBack)
    If blnHarden Then
        Descriptor = HardenStyle(Descriptor)
    End If

End Function

Public Sub Dither(Optional LightDark As Boolean = True, Optional InOut As Boolean = True, Optional ForeBack As Boolean = True)

  'copyright 2002 Roger Gilchrist
  'User interactive wrapper for DitherHard
  'original and slightly softer version of this routine

    DitherHard ColourUser, LightDark, InOut, ForeBack

End Sub

Public Sub Dither2(Optional LightDark As Boolean = True, Optional InOut As Boolean = True, Optional ForeBack As Boolean = True)

  'copyright 2002 Roger Gilchrist
  'User interactive wrapper for Dither2Hard
  'newer and darker version of this routine

    Dither2Hard ColourUser, LightDark, InOut, ForeBack

End Sub

Public Sub Dither2Hard(DColr As Long, Optional LightDark As Boolean = True, Optional InOut As Boolean = True, Optional ForeBack As Boolean = True)

  'copyright 2002 Roger Gilchrist
  'hard code version of Dither

  Dim i As Long
  Dim R As Integer, G As Integer, B As Integer
  Dim SPos As Long, EPos As Long, SelLen As Long 'Range of Text and hence LngClrArray
  Dim LngClrArray() As Long
  Dim DithR As Long, DithG As Long, DithB As Long

    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    ColourLong2RGB DColr, R, G, B
    m_LastDitherColour = DColr
    DithR = R / SelLen
    DithG = G / SelLen
    DithB = B / SelLen

    For i = SPos To EPos
        R = R - DithR
        G = G - DithG
        B = B - DithB
        If R < 0 Then
            R = 0
        End If
        If G < 0 Then
            G = 0
        End If
        If B < 0 Then
            B = 0
        End If

        LngClrArray(i) = RGB(R, G, B)

    Next i

    ColourApplicator LngClrArray, LightDark, InOut, ForeBack

End Sub

Public Sub DitherHard(DColr As Long, Optional LightDark As Boolean = True, Optional InOut As Boolean = True, Optional ForeBack As Boolean = True)

  'copyright 2002 Roger Gilchrist
  'hard code version of Dither2

  Dim i As Long
  Dim R As Integer, G As Integer, B As Integer
  Dim SPos As Long, EPos As Long, SelLen As Long 'Range of Text and hence LngClrArray
  Dim LngClrArray() As Long
  Dim DitherSection As Integer
  Dim DitherStep As Long, CurPos As Long

    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    DitherStep = (SelLen) / 255
    If DitherStep = 0 Then
        DitherStep = 1
    End If
    m_LastDitherColour = DColr
    ColourLong2RGB DColr, R, G, B

    For i = SPos To EPos
        DitherSection = DitherStep * CurPos
        If R - DitherSection < 0 Then
            R = 0
          Else 'NOT R...
            R = R - DitherSection
        End If
        If G - DitherSection < 0 Then
            G = 0
          Else 'NOT G...
            G = G - DitherSection
        End If
        If B - DitherSection < 0 Then
            B = 0
          Else 'NOT B...
            B = B - DitherSection
        End If
        LngClrArray(i) = RGB(R, G, B)
        CurPos = CurPos + 1
    Next i

    ColourApplicator LngClrArray, LightDark, InOut, ForeBack

End Sub

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

Public Function ExcessSpaceDelete() As Boolean

  'Copyright 2002 Roger Gilchrist
  'Remove extra spacing
  '*---PROGRAMMER MODIFICATION POINT---*
  'if you don't want it delete

  Dim RTFHead As String     'Initial section of RTF probably never useful
  Dim Parts As Variant      'array for searching
  Dim i As Integer          'For variable to manipulate Parts

    With m_RTB
        PositionStore Push
        SeparateRTFString RTFHead, Parts
        'Mark excess spacesfor deletion
        For i = LBound(Parts) + 1 To UBound(Parts)
            If Parts(i) = "" Then
                If RTFPartType(CStr(Parts(i - 1))) = TextOnly Then
                    Parts(i) = SafeJoinSkip
                End If
            End If
        Next i
        WorkingStringWrite RTFHead & SafeJoin(Parts, " ")
        PositionStore Pop
    End With 'M_RTB

End Function

Private Function ExistFile(vFile As String) As Integer

  'Checks disk to see if a file exists.  Works on paths too.
  '  To use with path, path must have a trailing "\"

  Dim Res As String

    On Error Resume Next
        Res = Dir$(vFile)
        If Res = "" Or Err <> 0 Then
            ExistFile = False
          Else 'NOT RES...
            ExistFile = True
        End If
    On Error GoTo 0

End Function

Private Sub FileKill(sFilename As String)

  'Safe file killer

    If Len(Dir$(sFilename)) Then
        Kill sFilename
    End If

End Sub

Private Function FileLineDefaultName(FileName As String, Sep As String, BaseDef As String) As String

  'copyright 2002 Roger Gilchrist
  'FileLineXXXX support the Materials and Styles files where the code can be shared

  'count the occurances of the baseDef name and returns
  'a string numbered to be the next in the series

  Dim tmpA As Variant, tmpPart As Variant, Lcount As Long, LenBase As Long

    LenBase = Len(BaseDef) + 1
    tmpA = KnownNames(FileName, Sep)
    For Each tmpPart In tmpA
        If Left$(tmpPart, LenBase) = BaseDef & " " Then
            Lcount = Lcount + 1
        End If
    Next tmpPart
    FileLineDefaultName = StrConv(BaseDef & " " & Lcount + 1, vbProperCase)

End Function

Private Sub FileLineDelete(FileName As String, Header As String, Sep$)

  'copyright 2002 Roger Gilchrist
  'FileLineXXXX support the Materials and Styles files where the code can be shared
  'quick and dirt deleter
  'loads whole file into an array
  'deletes the file
  'then rebuilds it with out the colour 'Clr$' paramater

  Dim AllLines As Variant, Aline As Variant, TmpLin As String

    If MsgBox("Are you sure you want to delete '" & StrConv(Header, vbProperCase) & "'?", vbYesNo) = vbYes Then
        AllLines = FileLinesAll(FileName)
        FileKill FileName
        For Each Aline In AllLines
            TmpLin = Left$(Aline, InStr(Aline, Sep$) - 1)
            If TmpLin <> Header Then
                FileLinesSave FileName, CStr(Aline)
            End If
        Next Aline
    End If

End Sub

Private Sub FileLineDuplicateGuard(FileName As String, Header As String, Sep As String)

  'copyright 2002 Roger Gilchrist
  'FileLineXXXX support the Materials and Styles files where the code can be shared
  'checks that Mat is not already used by file
  'warns you once then keeps trying until it can generate a new name + number

  Dim AllNames As Variant 'Temp array of file contents for searching
  Dim aName As Variant
  Dim FirstHit As Boolean ' only show message first time
  Dim OrigHeader As String 'preserve initial mat name
  Dim AutoNum As Long  ' counter (Could be integer as the whole system will crash if it gets bigger than that)

    Header = LCase$(Header)
    OrigHeader = Header
    AllNames = KnownNames(FileName, Sep)
retry:
    For Each aName In AllNames
        If Header = aName Then
            If FirstHit = False Then
                MsgBox "'" & StrConv(OrigHeader, vbProperCase) & "' already exists." & NL & _
                       "It will be saved with a number attached."
                FirstHit = True
            End If
            AutoNum = AutoNum + 1
            Header = OrigHeader & " " & AutoNum
            GoTo retry
        End If
    Next aName
    If FirstHit Then ' if there was a rename tell the user
        MsgBox "'" & StrConv(OrigHeader, vbProperCase) & "' has been renamed '" & StrConv(Header, vbProperCase) & "'."
    End If

End Sub

Private Function FileLineGet(FileName As String, Header As String, Sep$) As String

  'copyright 2002 Roger Gilchrist
  'FileLineXXXX support the Materials and Styles files where the code can be shared
  'gets the file line matching Header
  'This has no test for Stl not existing so only call from a control which has already checked

  Dim fnum As Integer
  Dim FileLin As String
  Dim TmpLin As String

    Header = LCase$(Header) 'file always stores as lower case
    fnum = FreeFile
    Open FileName For Input As fnum
    Do While Not EOF(fnum)   ' Loop until end of file.
        Line Input #fnum, FileLin  ' Read line into variable.
        TmpLin = Left$(FileLin, InStr(FileLin, Sep$) - 1)
        If TmpLin = Header Then
            FileLineGet = FileLin
            Exit Do '>---> Loop
        End If
    Loop
    Close #fnum   ' Close file.

End Function

Private Sub FileLineRestore(FileName As String, Safety As String, Target As String, Optional Force As Boolean = False)

  'copyright 2002 Roger Gilchrist
  'FileLineXXXX support the Materials and Styles files where the code can be shared
  'recreate the storage files for the class's externally stored dat files
  'This is automatically called if Class Initialize can't find the files
  'so you can safely add the class to new projects without remembering to move the dat files as well

  Dim AllLines As Variant, Aline As Variant

    If Force Then ' used by Class initialize to make sure that some data exists
        GoTo DoRestore
      Else 'FORCE = FALSE
        If MsgBox("If you proceed you will lose ALL " & Target & " you have created." & NL & NL & _
           "Are you sure?", vbYesNo) = vbYes Then
            GoTo DoRestore
        End If
    End If

Exit Sub

DoRestore:
    AllLines = Split(Safety, NL)
    AllLines = ArrayQuickSort(AllLines)
    FileKill FileName
    For Each Aline In AllLines
        FileLinesSave FileName, CStr(Aline)
    Next Aline

End Sub

Private Function FileLinesAll(FileName As String) As Variant

  'copyright 2002 Roger Gilchrist
  'FileLineXXXX support the Materials and Styles files where the code can be shared
  'return all lines as an array

  Dim fnum As Integer, UserCount As Long, JunkMe As String
  Dim CurName As Long, tmpA As Variant

    ReDim tmpA(FileLinesCount(FileName) - 1)
    fnum = FreeFile
    Open FileName For Input As fnum
    Do While Not EOF(fnum)   ' Loop until end of file.
        Line Input #fnum, tmpA(CurName)  ' Read line into variable.
        CurName = CurName + 1
    Loop
    Close #fnum   ' Close file.

    FileLinesAll = tmpA

End Function

Private Function FileLinesCount(FileName$) As Long

  'copyright 2002 Roger Gilchrist
  'FileLineXXXX support the Materials and Styles files where the code can be shared
  'count the number of lines in the file

  Dim fnum As Integer, UserCount As Long, JunkMe As String

    fnum = FreeFile
    Open FileName$ For Input As fnum
    Do While Not EOF(fnum)   ' Loop until end of file.
        Line Input #fnum, JunkMe$   ' Read line into variable.
        FileLinesCount = FileLinesCount + 1
    Loop
    Close #fnum   ' Close file.

End Function

Private Sub FileLinesSave(FileName As String, Desc$)

  'copyright 2002 Roger Gilchrist
  'FileLineXXXX support the Materials and Styles files where the code can be shared
  'QandD file Append routine (Append creates the file if it does not exist)

  Dim fnum As Integer, fname As String

    fnum = FreeFile
    Open FileName For Append As fnum
    Print #fnum, LCase$(Desc)
    Close #fnum

End Sub

Private Sub FontApplicator(FontArray)

  'Copyright 2002 Roger Gilchrist
  ' this is the core of the various font look routines
  '*---PROGRAMMER MODIFICATION POINT---*
  'work in progress
  'Does not handle fonttbl yet

  'fontArray format RTFfontcode & "***"
  'FontApplicator replaces "***" with an individual character

  'this is just a crude redit of ColourApplicator at present

  Dim RTFHead As String 'hold the part of the selection which is not changed(unless it contains a colortbl)
  Dim FntTbl As String 'Extract fonttbl from RTFHead if it exists or fake it otherwise
  Dim OrigFntTbl As String 'store the fonttbl so that it can be Replaced
  Dim Parts As Variant 'hold and manipulate the part of the selection which can be changed
  Dim CurFont As Long  'keep track of position in FontArray
  Dim i As Long        'For Next through the Selection.
  Dim Part As Variant

    PositionStore Push '
    FntTbl = GetFontTable(True)    '
    '    OrigFntTbl = FntTbl         'store for later Replace call
    '    If Len(FntTbl) = 0 Then ' if No colorTbl create a new one with auto first value ";"
    '        FntTbl = "{\fonttbl{;"
    '      Else 'else strip the end off the existing one'NOT LEN(FntTbl)...
    '        FntTbl = Left$(FntTbl, Len(FntTbl) - 3)
    '    End If
    SeparateRTFString RTFHead, Parts
    CurFont = LBound(FontArray) 'get first colour from array
    '    For i = LBound(Parts) To UBound(Parts)  ')
    For Each Part In Parts
        'FontManipulate Parts(i), FntTbl, FontArray, CurFont

        FontManipulate Part, FntTbl, FontArray, CurFont
        Parts(i) = Part
        i = i + 1
    Next Part
    '''    FntTbl = FntTbl & "}  " 'finish colortbl
    '''    If InStr(RTFHead, "{\fonttbl") Then 'Replace Original colortbl with new one
    '''        RTFHead = Replace(RTFHead, OrigFntTbl, FntTbl)
    '''      Else ' if there was no original fonttbl add it and another "}" to end it'NOT INSTR(RTFHEAD,...
    '''        RTFHead = Left$(RTFHead, Len(RTFHead) - 1) & FntTbl & "}"
    '''    End If
    WorkingStringWrite Trim$(RTFHead) & Join(Parts)
    PositionStore Pop

End Sub

Public Function FontLookArray(fStr As String) As Variant

  Dim tmpA As Variant, Tmpa2 As Variant, i As Long, RepCount As Long
  Dim TmpStr As String, WhatToDo As Long

    If Len(fStr) Then
        If Left$(fStr, 1) = "|" Then
            fStr = Mid$(fStr, 2)
        End If
        tmpA = ArrayQuickSort(Split(fStr, "|"))
        For i = LBound(tmpA) To UBound(tmpA)

            RepCount = 0
            If i + 1 < UBound(tmpA) Then
                Do While tmpA(i) = tmpA(i + 1)
                    i = i + 1 ':( Modifies active For-Variable
                    RepCount = RepCount + 1
                    If i + 1 > UBound(tmpA) Then
                        i = i - 1 ':( Modifies active For-Variable
                        Exit Do '>---> Loop
                    End If
                Loop
            End If
            TmpStr = TmpStr & "|" & tmpA(i)
        Next i
        TmpStr = Mid$(TmpStr, 2)
        FontLookArray = Split(TmpStr, "|")
      Else 'LEN(FSTR) = FALSE
        FontLookArray = Split("") 'create an empty array
    End If

End Function

Private Sub FontManipulate(Part, FntTbl$, StrFntArray, CurFont As Long)

  'Copyright 2002 Roger Gilchrist
  'does the actual work of inserting colour codes into colortbl and rtf code
  '*---PROGRAMMER MODIFICATION POINT---*
  'work in progress
  'Does not handle fonttbl yet

  Dim FntCount As Long 'get current number of colors in colortbl
  Dim i As Long ' loop driver' probably doesn't need to be a long but its safer
  Dim tmpP As String 'bulid new string to replace Part

    'FntCount = CountOccurances(clrTbl, "\fcharset") 'get length of fontTbl
    Select Case RTFPartType(CStr(Part))
      Case Blank 'blanks have to absorb an array member
        tmpP$ = Replace(StrFntArray(CurFont), "***", " ") 'tmpP$ & "\fn" & FntCount + 1 & " "
        'FntTbl$ = FntTbl$ & ColorRTFCode(CurFont, StrFntArray(CurFont))
        CurFont = CurFont + 1
      Case RTFOnly
        tmpP = Part
      Case Mixed, TextOnly
        'add a colour to each letter and advance colour counter for text letter but ignore RTF code
        For i = 1 To Len(Part)
            If InStr("\}", Mid$(Part, i, 1)) = 0 Then
                tmpP$ = tmpP$ & Replace(StrFntArray(CurFont), "***", Mid$(Part, i, 1))
                'FntTbl$ = FntTbl$ & FontRTFCode(, curFontFntArray(CurFont))
                CurFont = CurFont + 1
              Else 'NOT INSTR("\}",...
                'becuase of the way Part is generated once you reach an RTF code token there is not more text
                tmpP$ = tmpP$ & Mid$(Part, i)
                Exit For '>---> Next
            End If
        Next i
    End Select
    Part = tmpP

End Sub

Private Function FontRTFCode(CurFont As Long, FntName) As String

  '*---PROGRAMMER MODIFICATION POINT---*
  'work in porgress do not use
  'Copyright 2002 Roger Gilchrist
  'creates an RTF colour code member
  'support for ColourManipulation routine
  'NEEDS WORK TO FILL FONT PROPERTIES PROPERLYVVVVVVVVVV

    FontRTFCode = "\f" & CurFont & "\froman\fprq2\fcharset0 " & FntName & ";"

End Function

Private Function FontSize1stCharacter() As Long

  Dim RTFHead As String, Parts As Variant

    SeparateRTFString RTFHead, Parts
    FontSize1stCharacter = FontSizeExtract(CStr(Parts(0)))

End Function

Public Sub FontSizeAll(Fsize As Integer)

  'Copyright 2002 Roger Gilchrist
  '*PURPOSE: increase/decrease font size by RTF manipulation
  'variant of FontSizeStep allows you to force all text to single size
  'RTF equivalant to VBs .SelFontSize
  '*---PROGRAMMER MODIFICATION POINT---*
  'this is just a test/demonstartion structure, Recommended you use .SelFontSize (or the Class's Wrapper function for it)

  Dim TmpRTF As String, fsLoc As Long, RTFend As Long, s As String, SizeFactor As Integer

    KeepInBounds 0, Fsize, 2160 'Min and MAx FontSize values
    With m_RTB
        PositionStore Push
        TmpRTF = .SelRTF                                               'get selected text
        fsLoc = InStr(1, TmpRTF, "\fs")                     ' find RTF size code and move test point
        Do
            If IsNumeric(Mid$(TmpRTF, fsLoc + 4, 1)) Then   ' check is size not "fswiss" or other RTF code
                RTFend = InStr(fsLoc, TmpRTF, " ")                     ' find space delimiting size code
                's$ = Mid$(TmpRTF, fsLoc + 3, RTFend - fsLoc - 3)    ' extract size value
                Mid$(TmpRTF, fsLoc + 3, RTFend - fsLoc - 3) = CStr(Fsize * 2) 'insert new size *2 => converts VB FontSize to RTF Point Size
            End If
            fsLoc = InStr(fsLoc + 1, TmpRTF, "\fs")                     ' find next \fs code and move test point
        Loop Until fsLoc = 0

        .SelRTF = TmpRTF
        PositionStore Pop
    End With 'M_RTB

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

Public Sub FontSizeStep(BigTSmallF As Boolean)

  'Copyright 2002 Roger Gilchrist
  '*PURPOSE: increase/decrease individual font sizes by RTF manipulation
  'Extends VB's .SelFontSize by coping with multiple sizes in the selection
  'and changing them all by a factor of 1

  Dim TmpRTF As String, fsLoc As Long, RTFend As Long, s As String, SizeFactor As Integer, NewFontSize As Integer

    With m_RTB
        PositionStore Push
        TmpRTF = .SelRTF                                               'get selected text
        fsLoc = InStr(1, TmpRTF, "\fs")                     ' find RTF size code and move test point
        If BigTSmallF Then
            SizeFactor = 2
          Else 'BIGTSMALLF = FALSE
            SizeFactor = -2
        End If
        Do
            If IsNumeric(Mid$(TmpRTF, fsLoc + 4, 1)) Then   ' check is size not "fswiss" or other RTF code
                RTFend = InStr(fsLoc, TmpRTF, " ")                     ' find space delimiting size code
                s$ = Mid$(TmpRTF, fsLoc + 3, RTFend - fsLoc - 3)    ' extract size value
                NewFontSize = Val(s$) + SizeFactor
                KeepInBounds 0, NewFontSize, 4320 'Min and MAx FontSize values as points
                Mid$(TmpRTF, fsLoc + 3, RTFend - fsLoc - 3) = CStr(NewFontSize) 'insert new size
            End If
            fsLoc = InStr(fsLoc + 1, TmpRTF, "\fs")                     ' find next \fs code and move test point
        Loop Until fsLoc = 0

        .SelRTF = TmpRTF
        PositionStore Pop
    End With 'M_RTB

End Sub

Public Sub FormatRemove()

    m_RTB.SelRTF = m_RTB.SelText

    '
    '  'Copyright 2002 Roger Gilchrist
    '  'This strips all occurrances of "\cf#" or "\highlight#"
    '  'by getting a maximim possible # and counting down
    '  'Counting up misfires as "\cf1" would convert "\cf10" to "0"
    '  'leaving a visible character "0" in the text
    '
    '  Dim RTFHead As String 'hold the part of the selection which is not changed the colortbl is changed but the RTF engine does this, so you don't have to treat it programmatically
    '  Dim Parts As Variant 'hold and manipulate the part of the selection which can be changed
    '  Dim i As Long ', tmp As String
    '  Dim TMp As String ' convert Parts to string and back for ease of manipulation by the SafeJoinTest
    '  Dim target As String ' hold the type of colour to remove
    '  Dim NumberPart As Long 'current number value for test
    '  Dim InitNumberParts As Long 'Maximum possible number value because SelRTf is renumbered as a subset of TextRTF
    '  Dim Targets As Variant
    '  Dim tnum As Integer
    '
    '    Targets = Array("\up", "\dn", "\sub", "\super")
    '    PositionStore Push '
    '    Do While tnum < UBound(Targets)
    '
    '        SeparateRTFString RTFHead, Parts
    '        target$ = Targets(tnum)
    '        InitNumberParts = CountOccurances(Join(Parts), target) * 4 'allows for back and fore colour contributing to colortbl
    '        If InitNumberParts = 0 Then
    '            Exit Sub
    '        End If
    '
    '        For i = LBound(Parts) To UBound(Parts)
    '            NumberPart = InitNumberParts
    '            TMp = Parts(i)
    '            Do While InStr(TMp, target$)
    '                TMp$ = Replace(TMp$, target & NumberPart, target & "0")
    '                TMp$ = Replace(TMp$, target & "0", "")
    '                NumberPart = NumberPart - 1
    '                If NumberPart < 0 Then
    '                    Exit Do
    '                End If
    '            Loop
    '            Parts(i) = SafeJoinTest(Parts(i), TMp)
    '        Next i
    '        WorkingStringWrite Trim$(RTFHead) & SafeJoin(Parts)
    '        tnum = tnum + 1
    '    Loop
    '    PositionStore Pop
    '

End Sub

Public Sub Fuzzy(FColr As Long, Optional LightDark As Boolean = True, Optional InOut As Boolean = True, Optional ForeBack As Boolean = True)

  Dim i As Long
  Dim R As Integer, G As Integer, B As Integer
  Dim SPos As Long, EPos As Long, SelLen As Long 'Range of Text and hence LngClrArray
  Dim LngClrArray() As Long
  Dim DithR As Long, DithG As Long, DithB As Long
  Dim FuzzyRnd As Long

    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    m_LastFuzzyColour = FColr
    ColourLong2RGB FColr, R, G, B
    FuzzyRnd = 50 - Rnd * 100
    For i = SPos To EPos
        ColourLong2RGB FColr, R, G, B
        FuzzyRnd = FuzzyRnd + Rnd * 10
        R = R + FuzzyRnd
        G = G + FuzzyRnd
        B = B + FuzzyRnd
        KeepInBounds 0, R, 255
        KeepInBounds 0, G, 255
        KeepInBounds 0, B, 255
        LngClrArray(i) = RGB(R, G, B)

    Next i

    ColourApplicator LngClrArray, LightDark, InOut, ForeBack

End Sub

Public Sub FuzzyHard(FColr As Long, Optional LightDark As Boolean = True, Optional InOut As Boolean = True, Optional ForeBack As Boolean = True)

    Fuzzy FColr, LightDark, InOut, ForeBack

End Sub

Public Sub FuzzyUser(Optional LightDark As Boolean = True, Optional InOut As Boolean = True, Optional ForeBack As Boolean = True)

    Fuzzy ColourUser, LightDark, InOut, ForeBack

End Sub

Private Function GetColorTable(Optional ForceDoc As Boolean = False) As String

  'Copyright 2002 Roger Gilchrist
  'Specialised version of GetFormatString
  'used by ColoursUsed and ColourApplicator
  'returns the colortbl from either the selection or whole document(ForcedDoc=True)

  Dim ClrTblStart As Long ' Cutpoints for various pieces
  Dim ClrTblEnd As Long
  Dim RTFTail As String

    RTFTail$ = Left$(WorkingStringRead(ForceDoc), m_End_OF_RTF)
    ClrTblStart = InStr(RTFTail$, "{\colortbl")      'gets end of fonttbl
    If ClrTblStart = 0 Then                           'blank documents use default color table which is not present
        GetColorTable = ""
        Exit Function '>---> Bottom
      Else 'NOT CLRTBLSTART...
        'find certain end of colortbl note a multiple color table has ";}}"
        ClrTblEnd = InStr(ClrTblStart, RTFTail$, ";}")
        'so distance to next lump of code varies so get certain start at next "\"
        ClrTblEnd = InStr(ClrTblEnd, RTFTail, "\") - ClrTblStart
        GetColorTable = Mid$(RTFTail$, ClrTblStart, ClrTblEnd)
    End If

End Function

Private Function GetDefFormatString(Optional ForceDoc As Boolean = False) As String

  'Copyright 2002 Roger Gilchrist
  'this routine allows you access various parts of the TextRTF code
  'it is not used by the nearly identical BreakUpRTFString (on which it is based) because it would
  'have too high a performance hit. If you find you are referencing many of the format strings consider
  'making a task specific version rather than multiple calls here

  Dim FntTblStart As Long, ClrTblStart As Long ', RTF_End As Long  ' Cutpoints for various pieces
  Dim ClrTblEnd As Long, StyleSheetStart As Long
  Dim WorkThis As String, TmpRTF As String, FntTbl As String, clrTbl As String, RTFTail As String
  Dim SyleSht As String, RTFHead As String, txt As String

    TmpRTF = WorkingStringRead(False)
    RTFTail$ = Left$(TmpRTF, m_End_OF_RTF)

    FntTblStart = InStr(RTFTail$, "{\fonttbl")
    ClrTblStart = InStr(RTFTail$, "{\colortbl")      'gets end of fonttbl
    If ClrTblStart = 0 Then                           'blank documents use default color table which is not present
        ClrTblStart = InStr(FntTblStart, RTFTail$, ";}}") + 3
    End If
    If FntTblStart Then
        RTFTail$ = Mid$(RTFTail$, ClrTblStart)
    End If
    ClrTblStart = InStr(RTFTail$, "{\colortbl") 'retest as RTFTail has changed
    If ClrTblStart Then
        ClrTblEnd = InStr(ClrTblStart, RTFTail$, ";}") 'find certain end of colortbl not a multiple color table has ";}}"
        ClrTblEnd = InStr(ClrTblEnd, RTFTail, "\") - ClrTblStart    'so distance to next lump of code varies so get certain start at next "\"
        RTFTail$ = Mid$(RTFTail$, ClrTblEnd)

    End If
    If StyleSheetStart Then ' selected text is whole document
        RTFTail = Mid$(RTFTail, InStr(RTFTail, ";}}") + 5)
    End If

    GetDefFormatString = RTFTail 'This contains the initial RTF paragraph and format details for a selection
    'RTFHead & FntTbl & ClrTbl & SyleSht & RTFTail & WorkThis

End Function

Private Function GetFontTable(Optional ForceDoc As Boolean = False) As String

  'Copyright 2002 Roger Gilchrist
  'Specialised version of GetFormatString
  'used by ColoursUsed and ColourApplicator
  'returns the colortbl from either the selection or whole document(ForcedDoc=True)

  Dim FntTblStart As Long ' Cutpoints for various pieces
  Dim FntTblEnd As Long
  Dim RTFTail As String

    RTFTail$ = Left$(WorkingStringRead(ForceDoc), m_End_OF_RTF)
    FntTblStart = InStr(RTFTail$, "{\fonttbl")      'gets end of fonttbl
    If FntTblStart = 0 Then                           'blank documents use default color table which is not present
        GetFontTable = ""
        Exit Function '>---> Bottom
      Else 'NOT CLRTBLSTART...'NOT FNTTBLSTART...
        'find certain end of colortbl note a multiple color table has ";}}"
        FntTblEnd = InStr(FntTblStart, RTFTail$, ";}")
        'so distance to next lump of code varies so get certain start at next "\"
        FntTblEnd = InStr(FntTblEnd, RTFTail, "\") - FntTblStart
        GetFontTable = Mid$(RTFTail$, FntTblStart, FntTblEnd)
    End If

End Function

Private Function GetGetValue(RTFCode$) As Variant

  'Copyright 2002 Roger Gilchrist
  'support for the Get SelXXXX properties
  '*---PROGRAMMER MODIFICATION POINT---*
  'this is provisional until I find/work out how to make it behave properly

    With m_RTB
        If .SelLength Then
            GetGetValue = Not (RTFCodeLoc(.SelRTF, RTFCode$) <> NotPresent) ' Code is present so remove
          Else '.SELLENGTH = FALSE
            GetGetValue = HasRTFCode(RTFCode$)
        End If
    End With 'M_RTB

End Function

Private Function GetStartEnd(SPos As Long, EPos As Long, SelLen As Long) As Boolean

  'Support for GetStartEndLong and GetStartEndStr routines
  'call before anything else
  'PROGRAMMER WARNING
  'SelLen is used by most colour routines so moved it here

    With m_RTB
        SPos = .SelStart
        EPos = .SelLength
        EPos = SPos + EPos - 1
        SelLen = EPos - SPos
    End With 'M_RTB
    GetStartEnd = EPos > SPos
    'this lets this routine serve as a do nothing signal if there is no selection

End Function

Private Function GetStartEndLong(SPos As Long, EPos As Long, SelLen As Long, LongArray) As Boolean

  'Support for Colour routines
  'call before anything else
  'PROGRAMMER WARNING
  'SelLen is used by most colour routines so moved it here
  'returns a dimensioned Array of Longs

    If GetStartEnd(SPos, EPos, SelLen) Then
        GetStartEndLong = True
        ReDim LongArray(SPos To EPos) As Long
        'Use (Spos To Epos) style as that is what ColourApplicator expects
        'Initialize any values your ColourStyle requires
    End If
    'this lets this routine serve as a do nothing signal if there is no selection

End Function

Private Function GetStartEndStr(SPos As Long, EPos As Long, StrArray) As Boolean

  'Support for Colour routines
  'call before anything else
  'PROGRAMMER WARNING
  'returns a dimensioned Array of Strings
  'string arrays do not need SelLen so make it local and don't use it

  Dim SelLen As Long

    If GetStartEnd(SPos, EPos, SelLen) Then
        GetStartEndStr = True
        ReDim StrArray(SPos To EPos) As String
    End If
    'this lets this routine serve as a do nothing signal if there is no selection

End Function

Public Function HardenStyle(Stl As String) As String

    If InStr(Stl, " **") Then
        Stl = Replace(Stl, " **", "hardhard")
    End If
    If InStr(Stl, " *") Then
        Stl = Replace(Stl, " *", "hard")
    End If
    HardenStyle = Stl

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

Private Function InvertColor(col As Long) As Long

  Dim R As Integer, G As Integer, B As Integer, GREYTEST As Long

    ColourLong2RGB col, R, G, B
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

Public Property Get IsSelection() As Boolean

  'this property is only used by the Colour and Text Look panels to disable the Do It button if there is no selection

    IsSelection = m_RTB.SelLength > 0

End Property

Private Sub KeepInBounds(Lo, Value, Hi)

  'copyright 1988 Roger Gilchrist
  'Modified from slower all Variant version
  'Enforce minimum and maximum limits to Value
  'You could make this a function but it looks messier in code: Val = KeepinBounds(lo, Val, hi)
  'MODIFIED :    26-Aug-2002
  'optimized with If End if instead of Iif

    If Value < Lo Then
        Value = Lo
    End If
    If Value > Hi Then
        Value = Hi
    End If

End Sub

Private Function KnownNames(FileName As String, Sep As String) As Variant

  Dim fnum As Integer, UserCount As Long, JunkMe As String
  Dim CurName As Long, tmpA As Variant
  Dim LocalCount As Long

    LocalCount = FileLinesCount(FileName)
    If LocalCount Then
        ReDim tmpA(LocalCount - 1)
        fnum = FreeFile
        Open FileName For Input As fnum
        Do While Not EOF(fnum)   ' Loop until end of file.
            Line Input #fnum, JunkMe$   ' Read line into variable.
            If Len(JunkMe$) Then ' safety for blanks left by external editing of file
                tmpA(CurName) = Left$(JunkMe$, InStr(JunkMe, Sep$) - 1)
                CurName = CurName + 1
            End If
        Loop
        Close #fnum   ' Close file.
        KnownNames = ArrayQuickSort(tmpA)
    End If

End Function

Public Property Get LastBlendEndColour() As Long

  'These properties are mostly for the TextLookPanel's Preserve Selection Colour System

    LastBlendEndColour = m_LastBlendEndColour

End Property

Public Property Get LastBlendStartColour() As Long

  'These properties are mostly for the TextLookPanel's Preserve Selection Colour System

    LastBlendStartColour = m_LastBlendStartColour

End Property

Public Property Get LastDitherColour() As Long

  'These properties are mostly for the TextLookPanel's Preserve Selection Colour System

    LastDitherColour = m_LastDitherColour

End Property

Public Property Get LastFuzzyColour() As Long

    LastFuzzyColour = m_LastFuzzyColour

End Property

Public Property Get LastHighlightColour() As Long

  'These properties are mostly for the TextLookPanel's Preserve Selection Colour System
  'and only apply to the RTFHighlight methods

    LastHighlightColour = m_LastHighlightColour

End Property

Public Property Get LastHighlightForeColour() As Long

  'These properties are mostly for the TextLookPanel's Preserve Selection Colour System
  'and only apply to the RTFHighlight methods

    LastHighlightForeColour = m_LastHighlightForeColour

End Property

Public Property Get LastTextColour() As Long

  'These properties are mostly for the TextLookPanel's Preserve Selection Colour System

    LastTextColour = m_LastTextColour

End Property

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

Public Sub Materials(Clr As String, Optional LightDark As Boolean = True, Optional InOut As Boolean = False, Optional ForeBack As Boolean = True)

  'based on code originally written by
  'John Colman (john_colman@hotmail.com)in metalCBprj.OCX
  'I liked the subtle colours so i modified it from buttons to text
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim SPos As Long, EPos As Long, SelLen As Long
  Dim LngClrArray() As Long
  Dim UClrArray As Variant

    If Clr = "" Then
        Exit Sub '>---> Bottom
    End If
    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If

    UClrArray = Split(MaterialsGet(Clr), MaterialsSep)
    MaterialsGradientFill LngClrArray, CInt(Val(UClrArray(1))), CInt(Val(UClrArray(2))), CInt(Val(UClrArray(3))), CInt(Val(UClrArray(4))), CInt(Val(UClrArray(5))), CDbl(Val(UClrArray(6)))

    ColourApplicator LngClrArray, LightDark, InOut, ForeBack

End Sub

Public Function MaterialsAll() As Variant

    MaterialsAll = FileLinesAll(m_MaterialsFileName)

End Function

Public Function MaterialsCount() As Long

    MaterialsCount = FileLinesCount(m_MaterialsFileName)

End Function

Public Sub MaterialsCreator(MaterialName$, LightDark As Boolean, InOut As Boolean, ForeBack As Boolean, Min As Integer, Max As Integer, R As Integer, G As Integer, B As Integer, Multiplier As Double)

  'This routine creates new materials and saves them to the 'Materials.dat' file
  'InOut,LightDark and ForeBack are not part of the colour,but just how to display it, so not saved.

  Dim SPos As Long, EPos As Long, SelLen As Long
  Dim LngClrArray() As Long

    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    MaterialsSafeValues Min, Max, R, G, B, Multiplier
    If Len(MaterialName$) = 0 Then ' safety probably not needed but who knows
        MaterialName$ = "User " & MaterialsCount + 1
    End If

    If MaterialName <> "@@@@@@@@" Then ' this string is a guard to allow you to test without saving
        MaterialsDuplicateGuard MaterialName$
        MaterialsSave MaterialName$ & MaterialsSep & Min & MaterialsSep & Max & MaterialsSep & R & MaterialsSep & G & MaterialsSep & B & MaterialsSep & Multiplier
    End If
    MaterialsGradientFill LngClrArray, Min, Max, R, G, B, Multiplier

    ColourApplicator LngClrArray, LightDark, InOut, ForeBack

End Sub

Public Property Get MaterialsDefaultName() As String

    MaterialsDefaultName = FileLineDefaultName(m_MaterialsFileName, MaterialsSep, "user")

End Property

Public Sub MaterialsDelete(Mat$)

    FileLineDelete m_MaterialsFileName, Mat, MaterialsSep

End Sub

Private Sub MaterialsDuplicateGuard(Mat$)

  'checks that Mat is not already used by file
  'warns you once then keeps trying until it can generate a new name + number

    FileLineDuplicateGuard m_MaterialsFileName, Mat$, MaterialsSep

End Sub

Private Function MaterialsGet(Mat$) As Variant

  'gets the file line mathing Clr$
  'This has no test for Clr not existing so only call from a control which has already checked

    MaterialsGet = FileLineGet(m_MaterialsFileName, Mat, MaterialsSep)

End Function

Private Sub MaterialsGradientFill(CArray() As Long, Min As Integer, Max As Integer, R As Integer, G As Integer, B As Integer, Optional Multiplier As Double = 1)

  'based on code originally written by
  'John Colman (john_colman@hotmail.com) in metalCBprj.OCX

  Dim i As Long
  Dim C As Integer
  Dim st As Double
  Dim counter As Long
  Dim Cpos As Variant

    On Error Resume Next
        MaterialsSafeValues Min, Max, R, G, B, Multiplier
        counter = LBound(CArray) ' allows slightly faster? 'For Each' loop to be used
        Max = Max - Min
        st = (UBound(CArray) - LBound(CArray)) / 3.142 * Abs(Multiplier) * 2
        'MODIFICATIONS
        'a.  MaterialSafeValues
        'b. the * 2 and Counter are my modifications of the routine
        '.      *2 as I only need the 1st half of the material spectrum.
        '       The ArrayInOut call takes care of full spectrum in original version
        '       Counter because LBound and UBound are absolute not relative values
        '       and this routine needs relative to start and end math
        'c. 'For Each' used instead of 'for i=' as the i value had no use
        'd. I have also move the 'If Multiplier > 0' inside a single 'For Each'
        '   rather than the original with two 'for i =' in a 'If Then Else End If'
        '   If st = 0 Then st = 1E-25'DEBUG LINE NOT NEEDED
        For Each Cpos In CArray
            If Multiplier > 0 Then
                C = Abs(Max * Sin(counter / st)) + Min
              ElseIf Multiplier = 0 Then 'NOT MULTIPLIER...
                Exit For '>---> Next
              Else 'NOT MULTIPLIER...
                C = Abs(Max * Cos(counter / st)) + Min
            End If
            CArray(counter) = RGB(R + C, G + C, B + C)
            counter = counter + 1
        Next Cpos

    On Error GoTo 0

End Sub

Public Function MaterialsKnownColourNames() As Variant

    MaterialsKnownColourNames = KnownNames(m_MaterialsFileName, MaterialsSep)

End Function

Public Sub MaterialsReader(Clr As String, Min As Integer, Max As Integer, R As Integer, G As Integer, B As Integer, Multiplier As Double)

  'Input:     Clr   colour name
  'Returns:  values for that name

  Dim UClrArray As Variant

    UClrArray = Split(MaterialsGet(Clr), MaterialsSep)
    Min = CInt(Val(UClrArray(1)))
    Max = CInt(Val(UClrArray(2)))
    R = CInt(Val(UClrArray(3)))
    G = CInt(Val(UClrArray(4)))
    B = CInt(Val(UClrArray(5)))
    Multiplier = CDbl(Val(UClrArray(6)))

End Sub

Public Sub MaterialsRestore(Optional Force As Boolean = False)

  'This routine is designed to allow you to restore
  'the basic Material set if you need to
  'It is also called from Class Initialize if the
  'file can't be found so you can port the class
  'without remmebering to port the Materials.dat file
  '*---PROGRAMMER MODIFICATION POINT---*
  'simply open 'Materials.dat", copy and paste
  'it into Safety$ then paste '" & NL & _"' to the end of
  'each line except the last
  'if you want to preserve a different set
  'Note the data does not need to be Alpha-sorted
  'the program takes care of that for presentation purposes
  'John Colman (john_colman@hotmail.com)
  'worked out these colours and used them in his project metalCBprj.OCX
  'if you create nice new ones let me know and I'll add them to future versions

  'Name         min,    max,    r , g , b ,     multiplier
  '-----------|------|-------|----|----|------|---------------
  'blue steel   0,      255,    0,  33, 99,      1.3
  'diamond      0,      255,    100,100,255,     1.3
  'gold         0,      255,    70, 30, 0,       1.3
  'ice          100,    255,    0,  40, 60,      3
  'lead         0,      100,    33, 33, 33,      1.5
  'rubber       0,      100,    100,33, 33,      1.5
  'silver       0,      255,    33, 33, 33,      1.3
  'slate        0,      100,    33, 33, 100,     1.3

  'these are some I made
  'Name             min,    max,    r , g , b ,     multiplier
  '-----------|----------|-------|----|----|------|-----------
  'yellowrose       100    255     54   24   0     1.3
  'flesh            100    255     44    4   0     1.3
  'milk chocolate     0    100     95   55  45     1.5
  'dark chocolate     0     65     95   55  45     1.5
  'pineboard        100    165    100   85  35     1.5
  'oldgold          100    165    180   85  35     1.3
  'umber            100    165    180   85  35    10
  'snow             200    255      0   34  66     1.5
  'manila           101    217     94   96   0     3.4
  'manila 1         160    191     95   96   0     3.4
  'feel free to add to them just send me this line if you want
  'me to add to the distrubuted code

  'just send me the relevent line from the "Materials.Dat" file
  'it's just a text file but ".dat" is hard for inexperienced to open
  'and scary for the more experienced
  'I may modify the name so it does not clash with other offerings

  Dim Safety As String

    'Safety$ = " 'paste 'Materials.dat' contents here to replace Safety$
    Safety$ = "blue steel|0|255|0|33|99|1.3" & NL & _
             "diamond|0|255|100|100|255|1.3" & NL & _
             "gold|0|255|70|30|0|1.3" & NL & _
             "ice|100|255|0|40|60|3" & NL & _
             "lead|0|100|33|33|99|1.5" & NL & _
             "rubber|0|100|100|33|33|1.5" & NL & _
             "silver|0|255|33|33|33|1.3" & NL & _
             "slate|0|100|33|33|100|1.3" & NL & _
             "flesh|100|255|44|4|0|1.3" & NL & _
             "yellowrose|100|255|54|24|0|1.3" & NL & _
             "pineboard|100|165|100|85|35|1.5" & NL & _
             "umber|100|165|180|85|35|10" & NL & _
             "snow|200|255|0|34|66|1.5" & NL & _
             "manila|101|217|94|96|0|3.4" & NL & _
             "milk chocolate|0|100|95|55|45|1.5" & NL & _
             "old gold|100|165|180|85|35|1" & NL & _
             "manila 1|160|191|95|96|3.4"

    FileLineRestore m_MaterialsFileName, Safety, "MATERIALS", Force

End Sub

Public Sub MaterialsSafeValues(Min As Integer, Max As Integer, R As Integer, G As Integer, B As Integer, Multiplier As Double)

  'Keep all values sent to MaterialsGradientFill and MaterialsSave in bounds

    KeepInBounds 0, Max, 255
    KeepInBounds 0, Max, 255
    KeepInBounds 0, R, 255
    KeepInBounds 0, G, 255
    KeepInBounds 0, B, 255
    'Multiplier can be anything you like -ive, +ive BUT NOT 0
    If Multiplier = 0 Then
        Multiplier = 0.00001 'this is a safe minimum value you could try smaller ones
        'there is no upper level but the effect is less prominant as you go up not very interesting
    End If

End Sub

Private Sub MaterialsSave(Desc$)

    FileLinesSave m_MaterialsFileName, Desc$

End Sub

Public Property Get MaterialStorageFile() As Variant

    MaterialStorageFile = m_MaterialsFileName

End Property

Public Sub NoFormatting()

  'Copyright 2002 Roger Gilchrist
  'this strips all formatting in three slightly different ways
  '1.A Selection: Selection is reset to that of the character immediately before selection
  '               (if that is a whitespace the actual format may be visible on last character but could have been aplied just to the whitespace)
  '2.Whole Doc Selection: would reset to first charcter format (Probably not a good idea, I tried it on the Demo document, very ugly)
  '                       but has been programmed to deselect and use the No selection approach
  '                       If you want to duplicate this ugliness: Select the whole document, used Shift+LeftArrow to de-select last character, then call NoFormatting
  '3.No selection : Whole document is reset to RichTextBox default (You can set this from the IDE or accept VB defaults)

  '*---PROGRAMMER MODIFICATION POINT---*
  'if you would prefer the Ugly version of WholeDoc Selection(see above)
  'delete the 'If .SelLength = Len(.Text) Then' 'End If' structure

  Dim RTFHead As String     'Initial section of RTF probably never touched
  Dim Parts As Variant      'array for searching
  Dim WholeDocSwitch As Boolean

    With m_RTB
        PositionStore Push
        If .SelLength >= Len(.Text) Then ' deselect if it s whole doc to avoid ugliness
            .SelLength = 0
        End If
        If .SelLength Then
            .SelRTF = .SelText
          Else '.SELLENGTH = FALSE
            .Locked = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelRTF = .SelText ' set whole text to first character
            .Locked = False
            SeparateRTFString RTFHead, Parts
            Parts(0) = RTrim$(m_DefFormat) ' enforce def format
            WorkingStringWrite RTFHead & SafeJoin(Parts, " ")
        End If
        PositionStore Pop
    End With 'M_RTB

End Sub

Private Function Percent(a, B) As Single

  'calculate percentage given two numbers

  Dim C As Single

    '    If a > B Then a = B
    If B = 0 Then
        Percent = 0
        Exit Function '>---> Bottom
    End If
    C = a / B * 100
    If C > 100 Then
        C = 100
    End If
    Percent = C

End Function

Private Function PercentDecimal(a, B) As Single

  'calculate percentage as decimal given two numbers

    PercentDecimal = Percent(a, B) / 100

End Function

Private Sub PositionStore(SaveTSetF As LIFO)

  'Simple PushPop store of SelStart and SelLength allows routines
  'which manipulate through .seXXXXXX to restore SelStart and SelLength
  'arrays allow nesting of calls as long as you always call push and pop in calling routines

  '1. MOST IMPORTANT it sets the m_busy variable which stops a lot of updating while involved
  'in manipulating .SelXXX stuff

  '2. Set Mouse to Hourglass if m_busy = True
  '*---PROGRAMMER MODIFICATION POINT---*
  'modify/disable if you have your own mouse cursor routines to use

  '*---PROGRAMMER MODIFICATION POINT---*
  'If you edit this make sure you leave the m_busy stuff even if you delete everything else
  'The rest of this routine is only really needed if a routine
  'changes .SelStart or .SelLength to carry out its task
  'None of my colour routines actually do this but it is useful if you are trying to develop
  'something from scratch before going to TextRTF manipulation.

  Dim insertPos As Long

    'keep this bit no matter what
    With m_RTB
        If SaveTSetF Then
            m_busy = True
            .MousePointer = rtfHourglass
          Else 'SAVETSETF = FALSE
            .MousePointer = rtfDefault
            m_busy = False
        End If
    End With 'M_RTB

    On Error GoTo 0
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

Private Sub QuickSort(Larray, L, R)

  'Performs a Quick sort |requires support sub SwapAnyThing

  Dim i As Variant, J As Variant, x As Variant, Y As Variant

    i = L
    J = R
    x = Larray((L + R) / 2)

    While (i <= J)
        While (Larray(i) < x And i < R)
            i = i + 1
        Wend
        While (x < Larray(J) And J > L)
            J = J - 1
        Wend
        If (i <= J) Then
            SwapAnyThing Larray(i), Larray(J)
            i = i + 1
            J = J - 1
        End If
    Wend
    If (L < J) Then
        QuickSort Larray, L, J
    End If
    If (i < R) Then
        QuickSort Larray, i, R
    End If

End Sub

Public Sub RainBow(Optional LeftRight As Boolean = True, Optional InOut As Boolean = False, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim i As Long, SPos As Long, EPos As Long
  Dim RbowMod As Long, Rsection As Integer, Rstep As Long
  Dim CurPos As Long, LngClrArray() As Long, drv As Long
  Dim SpectrumSection As Integer, SpectrumSectionStep As Integer
  Dim SelLen As Long

    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    RbowMod = SelLen / 6            'this is used to keep Rstep within a spectrum section
    If RbowMod < 1 Then
        RbowMod = 1
    End If
    SpectrumSectionStep = 0
    For i = SPos To EPos
        Rstep = (255 * CurPos / RbowMod) 'fraction of 255
        SpectrumSection = Int(Percent(i - SPos, SelLen) \ 16.6666666666667)
        If SpectrumSectionStep < SpectrumSection Then
            SpectrumSectionStep = SpectrumSectionStep + 1
            Rstep = 0
            CurPos = 0
        End If
        CurPos = CurPos + 1 '    count through the spectrum section
        LngClrArray(i) = RainbowColor(SpectrumSection, Rstep)
    Next i

    ColourApplicator LngClrArray, LeftRight, InOut, ForeBack

End Sub

Private Function RainbowColor(SpectrumMember As Integer, Cval As Long) As Long

  'Based on
  'DrawRainBow © oigres P
  'Email: oigres@postmaster.co.uk

    Select Case SpectrumMember
      Case s1RedYellow   '0
        RainbowColor = RGB(255, Cval, 0)
      Case s2YellowGreen '1
        RainbowColor = RGB(255 - Cval, 255, 0)
      Case s3GreenCyan   '2
        RainbowColor = RGB(0, 255, Cval)
      Case s4CyanBlue    '3
        RainbowColor = RGB(0, 255 - Cval, 255)
      Case s5BlueMagenta '4
        RainbowColor = RGB(Cval, 0, 255)
      Case s6Magentared  '5
        RainbowColor = RGB(255, 0, 255 - Cval)
    End Select
    'Next

End Function

Public Sub RainBowIO(Optional RedMagenta As Boolean = True, Optional InOut As Boolean = False, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim i As Long, SPos As Long, EPos As Long, SelLen As Long
  Dim RbowMod As Long, Rsection As Integer, Rstep As Long
  Dim Cycler As Long, LngClrArray() As Long, drv As Long
  Dim TestA As Integer, TestB As Integer

    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If

    RbowMod = (SelLen) \ 5            'this is used to change spectrum sections

    TestB = 0
    For i = SPos To EPos
        Rstep = 255 * Cycler / RbowMod
        TestA = Int(Percent(i - SPos, SelLen) / 16.6666666666667)
        If TestA <> TestB Then
            TestB = TestB + 1
            Rstep = 0
            Cycler = 0
        End If
        Cycler = Cycler + 1 '    count through the spectrum section
        LngClrArray(i) = RainbowColor(TestA, Rstep)
    Next i

    ColourApplicator LngClrArray, RedMagenta, InOut, ForeBack

End Sub

Public Sub RandomColour(Optional ColorSet As Integer = 0, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed
  'ColorSet 0 to 7 select which source to use for getting random colors

  Dim i As Long, SPos As Long, EPos As Long, SelLen As Long
  Dim LngClrArray() As Long
  Dim Gray As Integer

    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If

    For i = SPos To EPos
        Select Case ColorSet
          Case 0
            LngClrArray(i) = ColourNamed("random")
          Case 1
            LngClrArray(i) = RainbowColor(Int(Rnd * 5), Int(Rnd * 255))
          Case 2 To 7
            LngClrArray(i) = RainbowColor(ColorSet - 2, Int(Rnd * 255))
          Case 8
            Gray = Int(Rnd * 255)
            LngClrArray(i) = RGB(Gray, Gray, Gray)
        End Select
        'if you don't want to use ColourNamed any where else then delete it and use the next line instead
        'LngClrArray(i) = ColourNamed = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
    Next i
    ColourApplicator LngClrArray, False, False, ForeBack

End Sub

Private Sub RandomColour2(ColorSet As Integer, Optional RndForeBack As Boolean = False)

  'Copyright 2002 Roger Gilchrist
  'applies random auto colour to selected text
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim RLen As Long, i As Long
  Dim SPos As Long, EPos As Long, SelLen As Long
  Dim LngClrArray() As Long, Grzy As Integer

    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    For i = SPos To EPos
        If Rnd > 0.5 Then
            Select Case ColorSet
              Case 0
                LngClrArray(i) = ColourNamed("random")
              Case 1
                LngClrArray(i) = RainbowColor(Int(Rnd * 5), Int(Rnd * 255))
              Case 2 To 7
                LngClrArray(i) = RainbowColor(ColorSet - 2, Int(Rnd * 255))
              Case 8
                Grzy = Int(Rnd * 255)
                LngClrArray(i) = RGB(Gray, Gray, Gray)
            End Select
          Else 'NOT RND...
            If Not RndForeBack Then
                If Rnd > 0.5 Then
                    LngClrArray(i) = vbBlack  'standard colouring
                  Else 'NOT RND...
                    LngClrArray(i) = vbWhite  'standard colouring
                End If
              Else 'NOT NOT...
                LngClrArray(i) = ColourNamed("random") ' may as well be random any way
            End If
        End If
    Next i
    ColourApplicator LngClrArray, False, False, True, RndForeBack

End Sub

Public Sub Ransom(Optional Colorise As Integer = 0)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if you dont want this
  'fontArray format RTFfontcode&"***"
  'FontApplicator replaces "***" with an individual character

  Dim i As Integer
  Dim t As String, s As String, Lp As String, Rp As String
  Dim StrFntArray() As String
  Dim SPos As Long, EPos As Long
  Dim InitialFontSize As Long

    If GetStartEndStr(SPos, EPos, StrFntArray) = False Then  'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    InitialFontSize = FontSize1stCharacter
    For i = SPos To EPos
        StrFntArray(i) = RansomRTF(InitialFontSize)

    Next i
    FontApplicator StrFntArray

    Select Case Colorise
      Case 0 ' no colour
      Case 1

        RandomColour Int(Rnd * 7)
      Case 2

        RandomColour Int(Rnd * 7), False
      Case 3

        RandomColour2 Int(Rnd * 7), True
    End Select

End Sub

Private Sub RansomFormat(Rstart$, rend$, freq As Long, frmt$)

  'Copyright 2002 Roger Gilchrist
  'support for RansomRTF
  '*---PROGRAMMER MODIFICATION POINT---*

    If Rnd > freq Then
        Rstart$ = Rstart$ & frmt
        rend$ = rend$ & frmt & "0"
    End If

End Sub

Private Function RansomRTF(InitialFontSize As Long) As String

  'Copyright 2002 Roger Gilchrist
  'generate random RTF formatting
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted
  'change freq value for more or less of specific effects

  Dim Rstart As String, rend As String

    Rstart$ = ""
    rend$ = ""
    '        remember \up-# convertes to \dn# 50% chance of happening
    Rstart$ = Rstart$ & "\up" & CStr((5 - Int(Rnd * 10 + 1)))
    rend$ = rend$ & "\up0"
    'Font generally gets bigger but has 20% chance of being smaller
    Rstart$ = Rstart$ & "\fs" & (InitialFontSize + 10 - Int(Rnd * 25 + 1))
    RansomFormat Rstart$, rend$, 0.5, "\b"
    RansomFormat Rstart$, rend$, 0.6, "\sub"
    RansomFormat Rstart$, rend$, 0.5, "\super"
    RansomFormat Rstart$, rend$, 0.5, "\i"
    RansomFormat Rstart$, rend$, 0.5, "\caps"
    RansomFormat Rstart$, rend$, 0.6, "\strike"
    RansomUnderliner Rstart$, rend$, 0.5
    If Len(Rstart$) > 0 Then
        Rstart$ = Rstart$ & " "
    End If
    rend$ = rend$ & " " 'IIf(Len(Rend$) > 0, " ", "")
    RansomRTF = Rstart$ & "***" & rend$

End Function

Private Sub RansomUnderliner(Rstart$, rend$, freq As Long)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted
  'support for RansomRTF

    If Rnd > freq Then
        Rstart = Rstart & Choose(Int(Rnd * 7) + 1, "\ul", "\uld", "\uldash", "\uldashd", "\uldashdd", "\ulth", "\ulhair", "\ulwave")
        rend = rend & "\ulnone"
    End If

End Sub

Private Sub RippleBase(Part, CurrentValue As Integer, Amplitude As Integer, WaveLength As Integer, UPDOWN As Boolean)

  'Copyright 2002 Roger Gilchrist
  'apply Base ripple
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted

  Dim i As Long, TMp As String, tmpA As String, tmpB As String

    Select Case RTFPartType(CStr(Part))
      Case Blank
        RippleBaseCount CurrentValue, Amplitude, UPDOWN, WaveLength
        Part = SafeJoinSkip
        ' do nothing
      Case RTFOnly
        For i = 1 To CountOccurances(Part, "\tab")
            RippleBaseCount CurrentValue, Amplitude, UPDOWN, WaveLength
        Next i
        ' do nothing
      Case TextOnly
        For i = 1 To Len(Part)
            RippleBaseCount CurrentValue, Amplitude, UPDOWN, WaveLength
            TMp$ = TMp$ & "\up" & CurrentValue & " " & Mid$(Part, i, 1)
        Next i
        Part = TMp$
      Case Mixed
        If InStr(Part, "\") Then
            tmpA = Left$(Part, InStr(Part, "\") - 1)
            tmpB = Mid$(Part, InStr(Part, "\"))

          ElseIf InStr(Part, "}") Then 'NOT INSTR(PART,...
            tmpA = Left$(Part, InStr(Part, "}") - 1)
            tmpB = Mid$(Part, InStr(Part, "}"))
        End If

        For i = 1 To Len(tmpA)
            RippleBaseCount CurrentValue, Amplitude, UPDOWN, WaveLength
            TMp$ = TMp$ & "\up" & CurrentValue & " " & Mid$(tmpA, i, 1)
        Next i
        Part = TMp$ & tmpB

    End Select

End Sub

Private Sub RippleBaseCount(CurLevel As Integer, Limit As Integer, UPDOWN As Boolean, WaveLength As Integer, Optional GoodOne As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  'change values for BaseRipple
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted

    If Abs(CurLevel) = Abs(Limit) Then
        UPDOWN = Not UPDOWN
    End If
    If UPDOWN Then
        CurLevel = CurLevel - WaveLength
      Else 'UPDOWN = FALSE
        CurLevel = CurLevel + WaveLength
    End If

    If GoodOne Then
        If CurLevel < -Limit Then
            CurLevel = -Limit  'stops waveLength sending CurLevel outside Limits
        End If
        If CurLevel > Limit Then
            CurLevel = Limit
        End If
    End If

End Sub

Public Sub RippleEngine(Mode As RippleStyle, Optional Amplitude As Integer = 5, Optional UPDOWN As Boolean = True, Optional InitialNum As Integer = 0, Optional WaveLength As Integer = 1)

  'Copyright 2002 Roger Gilchrist

  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted

  'this is the new and used version of the RippleEngine
  ' apply ripple text; allows you to develop and apply other ripple effects
  ' by adding to the select case structure

  'fontArray format RTFfontcode & number & "***"
  'FontApplicator replaces "***" with an individual character

  Dim i As Integer          'For variable to manipulate Parts
  Dim LimitLo As Integer, LimitHi As Integer
  Dim StrFntArray() As String
  Dim SPos As Long, EPos As Long

    If GetStartEndStr(SPos, EPos, StrFntArray) = False Then  'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    'ReDim StrFntArray(SPos To EPos) As String

    For i = SPos To EPos
        Select Case Mode
          Case BaseLine
            RippleBaseCount InitialNum, Amplitude, UPDOWN, WaveLength
            StrFntArray(i) = "\up" & InitialNum & "***"

          Case BaseLine1 'this one is useless most of the time but some of the variants are nice
            RippleBaseCount InitialNum, Amplitude, UPDOWN, WaveLength, False
            StrFntArray(i) = "\up" & InitialNum & "***"

          Case THeight
            If i = SPos Then 'if first one set initialNum and LimitHi and LimitLo
                InitialNum = FontSize1stCharacter + InitialNum
                LimitLo = InitialNum - Amplitude * 2
                If LimitLo < 8 Then
                    LimitHi = InitialNum + Amplitude * 2 + Abs(8 - LimitLo)
                    LimitLo = 8 ' prevent lowe text becoming to low
                  Else 'NOT LIMITLO...
                    LimitHi = InitialNum + Amplitude * 2
                End If
            End If
            RippleTHeightCount InitialNum, LimitLo, LimitHi, UPDOWN, WaveLength
            StrFntArray(i) = "\fs" & InitialNum & "***"
        End Select
    Next i
    FontApplicator StrFntArray

End Sub

Private Sub RippleTHeight(Part, CurLevel As Integer, LimitLo As Integer, LimitHi As Integer, UPDOWN As Boolean, WaveLength As Integer)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted
  ' apply THeight ripple

  Dim i As Long, TMp As String, tmpA As String, tmpB As String

    Select Case RTFPartType(CStr(Part))
      Case Blank

        RippleTHeightCount CurLevel, LimitLo, LimitHi, UPDOWN, WaveLength
        Part = SafeJoinSkip
      Case RTFOnly

        For i = 1 To CountOccurances(Part, "\tab")
            RippleTHeightCount CurLevel, LimitLo, LimitHi, UPDOWN, WaveLength
        Next i
      Case TextOnly

        For i = 1 To Len(Part)
            RippleTHeightCount CurLevel, LimitLo, LimitHi, UPDOWN, WaveLength
            TMp$ = TMp$ & "\fs" & CurLevel & " " & Mid$(Part, i, 1)
        Next i
        Part = TMp$
      Case Mixed
        If InStr(Part, "\") Then
            tmpA = Left$(Part, InStr(Part, "\") - 1)
            tmpB = Mid$(Part, InStr(Part, "\"))
          ElseIf InStr(Part, "}") Then 'NOT INSTR(PART,...
            tmpA = Left$(Part, InStr(Part, "}") - 1)
            tmpB = Mid$(Part, InStr(Part, "}"))
          Else 'NOT INSTR(PART,...
            Exit Sub '>---> Bottom
        End If
        For i = 1 To Len(tmpA)
            RippleTHeightCount CurLevel, LimitLo, LimitHi, UPDOWN, WaveLength
            TMp$ = TMp$ & "\fs" & CurLevel & " " & Mid$(tmpA, i, 1)
        Next i
        Part = TMp$ & tmpB
    End Select

End Sub

Private Sub RippleTHeightCount(CurLevel As Integer, LimitLo As Integer, LimitHi As Integer, UPDOWN As Boolean, WaveLength As Integer)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted
  ' change values for THeight

    If CurLevel <= LimitLo Or CurLevel >= LimitHi Then
        UPDOWN = Not UPDOWN
    End If
    ' the basic add is 2 to jump one size in point size as RTF code uses 1/2 point steps
    If UPDOWN Then
        CurLevel = CurLevel - (2 + WaveLength)
      Else 'UPDOWN = FALSE
        CurLevel = CurLevel + 2 + WaveLength
    End If

End Sub

Private Function RTFCodeLoc(TmpRTF$, RTFCode$) As RTFCodeLocs

  'Copyright 2002 Roger Gilchrist
  'WHERE THE RTFCODE OCCURS WITHIN A SelRTF string
  'A single space at the end of RTFCode is the END OF RTFCode Token
  'A double space is an RTFToken followed by a blank text character
  'This routine allows you to judge which format the RTF code is using
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

Private Sub RTFHighlight(colr As Long)

  'This is an RTF manipulation version of APIHighLight

  Dim i As Long
  Dim SPos As Long, EPos As Long
  Dim SelLen As Long 'Range of Text and hence LngClrArray
  Dim LngClrArray() As Long 'Store colours for passing to ColourApplicator

    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    m_LastHighlightColour = colr
    For i = SPos To EPos
        LngClrArray(i) = colr
    Next i
    ColourApplicator LngClrArray, False, False, False

End Sub

Private Sub RTFHighlight2(BColr As Long, FColr As Long)

  'This is an RTF manipulation version of APIHighLight2
  'creator Roger Gilchrist
  'private access for all 2 colour modes

    RTFHighlight BColr
    m_RTB.SelColor = FColr
    m_LastHighlightForeColour = FColr

End Sub

Public Sub RTFHighlightHard(BColr As Long)

  'This is an RTF manipulation version of APIHighLightHard
  'creator Roger Gilchrist
  'Hard means hardcoded colour selection

    RTFHighlight BColr

End Sub

Public Sub RTFHighlightHardAuto(BColr As Long)

  'This is an RTF manipulation version of APIHighLightHardHArd
  'creator Roger Gilchrist
  'Hard means hardcoded colour selection

    RTFHighlight2 BColr, InvertColor(BColr)

End Sub

Public Sub RTFHighlightHardHard(BColr As Long, FColr As Long)

  'This is an RTF manipulation version of APIHighLightHardHArd
  'creator Roger Gilchrist
  'Hard means hardcoded colour selection

    RTFHighlight2 BColr, FColr

End Sub

Public Sub RTFHighlightRemove()

  'This is an RTF manipulation version of APIHighLightRemove

    RTFHighlight TranslateSysColor(vbWindowBackground)

End Sub

Public Sub RTFHighlightUser()

  'This is an RTF manipulation version of APIHighLightUser
  'creator Roger Gilchrist Modification of Sergio Perciballi's work  (see APIHihglight)
  'User sets highlight colour Program selects suitable(?) font colour

    RTFHighlight ColourUser

End Sub

Public Sub RTFHighlightUserAuto()

  'This is an RTF manipulation version of APIHighLightUserAuto
  'creator Roger Gilchrist Modification of Sergio Perciballi's work  (see APIHihglight)
  'User sets highlight colour Program selects suitable(?) font colour

  Dim UColr As Long

    UColr = ColourUser
    RTFHighlight2 UColr, InvertColor(UColr)

End Sub

Public Sub RTFHighlightUserUser()

  'This is an RTF manipulation version of APIHighLightUserUser
  'creator Roger Gilchrist Modification of Sergio Perciballi's work  (see APIHihglight)
  'User sets highlight colour Program selects suitable(?) font colour

    RTFHighlight2 ColourUser, ColourUser

End Sub

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

Private Function SafeJoin(Bits, Optional Del$ = " ") As String

  'Copyright 2002 Roger Gilchrist
  '*PURPOSE: Replace Join with the ability to recognise whether a blank
  'was in original or deleted by editing of member
  'and not insert a del$ character for deleted member
  'SEE ALSO SafeJoinTest

  Dim bit As Variant

    For Each bit In Bits
        If bit <> SafeJoinSkip Then ' if marked not to reattach then don't add
            SafeJoin = SafeJoin & bit & Del$
        End If
    Next bit

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
    ' SafeJoinTest = IIf(t$ <> "" Or p = t$, t$, SafeJoinSkip)

End Function

Public Property Let SelBackColor(ByVal vNewValue As Long)

  'wrapper for Standard RichtextBox property

    RTFHighlightHard vNewValue

End Property

Public Property Get SelBackColor() As Long

  'wrapper for Standard RichtextBox property

    SelBackColor = APIHighlightColour

End Property

Public Property Get SelBold() As Variant

  'Copyright 2002 Roger Gilchrist
  'wrapper for Standard RichtextBox property

    SelBold = m_RTB.SelBold

End Property

Public Property Let SelBold(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist
  'wrapper for Standard RichtextBox property

    m_RTB.SelBold = vNewValue

End Property

Public Property Let SelCaps(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\caps")

End Property

Public Property Get SelCaps() As Variant

  'Copyright 2002 Roger Gilchrist

    SelCaps = GetGetValue("\caps")

End Property

Public Property Get SelColor() As Long

  'wrapper for Standard RichtextBox property

    SelColor = m_RTB.SelColor

End Property

Public Property Let SelColor(ByVal vNewValue As Long)

  'wrapper for Standard RichtextBox property

    m_RTB.SelColor = vNewValue
    m_LastTextColour = m_RTB.SelColor

End Property

Public Property Let SelDash(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\uldash")

End Property

Public Property Get SelDash() As Variant

  'Copyright 2002 Roger Gilchrist

    SelDash = GetGetValue("\uldash")

End Property

Public Property Get SelDashd() As Variant

  'Copyright 2002 Roger Gilchrist

    SelDashd = GetGetValue("\uldashd")

End Property

Public Property Let SelDashd(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\uldashd")

End Property

Public Property Let SelDashdd(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\uldashdd")

End Property

Public Property Get SelDashdd() As Variant

  'Copyright 2002 Roger Gilchrist

    SelDashdd = GetGetValue("\uldashdd")

End Property

Public Property Let SelDelete(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\delete")

End Property

Public Property Get SelDelete() As Variant

  'Copyright 2002 Roger Gilchrist

    SelUlWord = GetGetValue("\delete")

End Property

Public Property Get SelDot() As Variant

  'Copyright 2002 Roger Gilchrist

    SelDot = GetGetValue("\uld")

End Property

Public Property Let SelDot(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\uld")

End Property

Public Property Let SelDown(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    If Val(vNewValue) = 0 Then
        vNewValue = "0"
    End If
    SelRTFToggle ("\dn" & vNewValue)

End Property

Public Property Get SelDown() As Variant

  'Copyright 2002 Roger Gilchrist

    SelDown = GetGetValue("\dn")

End Property

Public Property Let SelFontSize(ByVal Fsize As Integer)

  'Copyright 2002 Roger Gilchrist
  'wrapper for Standard RichtextBox property
  'max size is 2160,min useful is about 5 but its legal to less

    KeepInBounds 0, Fsize, 2160
    m_RTB.SelFontSize = Fsize

End Property

Public Property Get SelFontSize() As Integer

  'Copyright 2002 Roger Gilchrist
  'wrapper for Standard RichtextBox property

    SelFontSize = m_RTB.SelFontSize

End Property

Public Property Let SelHair(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\ulhair")

End Property

Public Property Get SelHair() As Variant

  'Copyright 2002 Roger Gilchrist

    SelHair = GetGetValue("\ulhair")

End Property

Public Property Get SelItalic() As Variant

  'Copyright 2002 Roger Gilchrist
  'wrapper for Standard RichtextBox property

    SelItalic = m_RTB.SelItalic

End Property

Public Property Let SelItalic(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist
  'wrapper for Standard RichtextBox property

    m_RTB.SelItalic = vNewValue

End Property

Private Function SelRTFToggle(code$) As Boolean

  'Copyright 2002 Roger Gilchrist

  Dim RTFHead As String     'Initial section of RTF probably never useful
  Dim Parts As Variant      'Split MixedBit for application of code
  Dim FntTblStart As Long, ClrTblStart As Long, RTF_End As Long ' Cutpoints for various pieces
  Dim i As Integer          'For variable to manipulate Parts
  Dim TMp As String          'tmp for testing whether "" is original or edited

    With m_RTB
        PositionStore Push
        SelRTFToggle = Not (RTFCodeLoc(WorkingStringRead, code$) <> NotPresent) ' Code is present so remove
        SeparateRTFString RTFHead, Parts
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

Public Property Get SelStrikeThru() As Variant

  'Copyright 2002 Roger Gilchrist
  'wrapper for Standard RichtextBox property

    SelStrikeThru = m_RTB.SelStrikeThru

End Property

Public Property Let SelStrikeThru(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist
  'wrapper for Standard RichtextBox property

    m_RTB.SelStrikeThru = vNewValue

End Property

Public Property Get SelSub() As Variant

  'Copyright 2002 Roger Gilchrist

    SelSub = GetGetValue("\sub")

End Property

Public Property Let SelSub(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\sub")

End Property

Public Property Get SelSuper() As Variant

  'Copyright 2002 Roger Gilchrist

    SelSuper = GetGetValue("\super")

End Property

Public Property Let SelSuper(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\super")

End Property

Public Property Get SelThick() As Variant

  'Copyright 2002 Roger Gilchrist

    SelThick = GetGetValue("\ulth")

End Property

Public Property Let SelThick(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\ulth")

End Property

Public Property Let SelUlDouble(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\uldb")

End Property

Public Property Get SelUlDouble() As Variant

  'Copyright 2002 Roger Gilchrist

    SelUlDouble = GetGetValue("\uldbl")

End Property

Public Property Let SelUlWord(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\ulw")

End Property

Public Property Get SelUlWord() As Variant

  'Copyright 2002 Roger Gilchrist

    SelUlWord = GetGetValue("\ulw")

End Property

Public Property Let SelUnderline(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist
  'wrapper for Standard RichtextBox property

    m_RTB.SelUnderline = vNewValue

End Property

Public Property Get SelUnderline() As Variant

  'Copyright 2002 Roger Gilchrist
  'wrapper for Standard RichtextBox property

    SelUnderline = m_RTB.SelUnderline

End Property

Public Property Get SelUp() As Variant

  'Copyright 2002 Roger Gilchrist

    SelUp = GetGetValue("\up")

End Property

Public Property Let SelUp(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    If Val(vNewValue) = 0 Then
        vNewValue = "0"
    End If
    SelRTFToggle ("\up" & vNewValue)

End Property

Public Property Get SelVisible() As Variant

  'Copyright 2002 Roger Gilchrist

    SelVisible = GetGetValue("\v")

End Property

Public Property Let SelVisible(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\v")

End Property

Public Property Get SelWave() As Boolean

  'Copyright 2002 Roger Gilchrist

    SelWave = GetGetValue("\ulwave")

End Property

Public Property Let SelWave(ByVal vNewValue As Boolean)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\ulwave")

End Property

Private Sub SeparateRTFString(RTFHead As String, Parts)

  'Copyright 2002 Roger Gilchrist
  'separates the RTFCode into 1.RTFHead =the unchanging head of the string
  '                           2. Parts  = an array of the pieces that can be changed

  Dim TmpRTF As String, Cpoint As Long

    TmpRTF = WorkingStringRead
    Cpoint = InStr(TmpRTF, "\uc") - 1 ' but there might be "\viewkind# or might not be
    Do Until InStr("\ " & vbNewLine, Mid$(TmpRTF, Cpoint, 1))
        Cpoint = Cpoint - 1
    Loop
    If Mid$(TmpRTF, Cpoint, 1) = "\" Then
        Cpoint = Cpoint - 1
    End If
    RTFHead = Left$(TmpRTF, Cpoint) 'InStr(TMPRTF, "\uc") - 1)
    Parts = Split(Mid$(TmpRTF, Cpoint + 1)) 'InStr(TMPRTF, "\uc")))

End Sub

Public Sub SetForeColor()

    m_LastTextColour = ColourUser
    m_RTB.SelColor = m_LastTextColour

End Sub

Public Sub SetForeColorHard(Clr As Long)

    m_RTB.SelColor = Clr

End Sub

Public Sub SpectrumSector(Range As Integer, Optional LeftRight As Boolean = True, Optional InOut As Boolean = False, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim i As Long 'For Next value
  Dim SPos As Long, EPos As Long 'StartPos and EndPos
  Dim SelLen As Long, CurMember As Long 'Length of selection and CurrentMember of selection
  Dim Rstep As Long  '
  Dim LngClrArray() As Long  '

    If GetStartEndLong(SPos, EPos, SelLen, LngClrArray) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets StartPos and EndPos and Exits if no Selection '>---> Bottom
    End If

    ReDim LngClrArray(SPos To EPos) As Long
    SelLen = EPos - SPos

    For i = SPos To EPos
        Rstep = 255 * CurMember / SelLen ' get percentage of 255
        CurMember = CurMember + 1
        LngClrArray(i) = RainbowColor(CInt(Range), Rstep)
    Next i

    ColourApplicator LngClrArray, LeftRight, InOut, ForeBack

End Sub

Public Function StylesAll() As Variant

    StylesAll = FileLinesAll(m_StylesFileName)

End Function

Public Function StylesCount() As Long

    StylesCount = FileLinesCount(m_StylesFileName)

End Function

Public Sub StylesCreator(styles$, PreserveColour As Boolean)

  'This routine creates new materials and saves them to the 'Materials.dat' file
  'InOut,LightDark and ForeBack are not part of the colour,but just how to display it, so not saved.

  Dim stylesName As String, StylesParts As Variant

    StylesParts = Split(styles, "^")
    stylesName$ = StylesParts(0)
    If Len(stylesName$) = 0 Then ' safety probably not needed but who knows
        stylesName$ = "Style "
    End If

    If stylesName$ <> "@@@@@@@@" Then ' this string is a guard to allow you to test without saving
        StylesDuplicateGuard stylesName
        StylesParts(0) = stylesName
        StylesSave Join(StylesParts, "^")
    End If

    StylesEngine stylesName, PreserveColour

End Sub

Public Property Get StylesDefaultName() As String

    StylesDefaultName = FileLineDefaultName(m_StylesFileName, StylesSeparator, "style")

End Property

Public Sub StylesDelete(Stl$)

    FileLineDelete m_StylesFileName, Stl, StylesSeparator

End Sub

Private Sub StylesDuplicateGuard(styles$)

    FileLineDuplicateGuard m_StylesFileName, styles, StylesSeparator

End Sub

Public Sub StylesEngine(Stl$, PreserveColour As Boolean)

  Dim StyleParts As Variant, StyleBits  As Variant
  Dim StyleName As String, CurSubStyle As String, CurSubStyleVal As Integer
  Dim LeftRighter As Boolean, InOuter As Boolean, StylePrivate As Integer
  Dim TextBack As Boolean, ClrBack As Long, ClrText As Long, i As Integer

    If Stl$ = "" Then
        Exit Sub '>---> Bottom
    End If

    StyleParts = Split(StylesGet(Stl$), "^")
    For i = 1 To 2
        StylesPainter CStr(StyleParts(i)), PreserveColour
    Next i

End Sub

Public Function StylesGet(Stl$) As String

    StylesGet = FileLineGet(m_StylesFileName, Stl, StylesSeparator)

End Function

Public Function StylesKnownNames() As Variant

    StylesKnownNames = KnownNames(m_StylesFileName, StylesSeparator)

End Function

Public Sub StylesPainter(Desc$, PreserveColour As Boolean)

  Dim StyleBits As Variant
  Dim StyleName As String
  Dim CurSubStyle As String
  Dim CurSubStyleVal As Integer
  Dim StylePrivate As Integer
  Dim ClrBack As Long
  Dim ClrText As Long
  Dim LeftRighter As Boolean
  Dim InOuter As Boolean
  Dim TextBack  As Boolean

    StyleBits = Split(Desc$, MaterialsSep)
    StyleName = StyleBits(0)
    CurSubStyle = CStr(StyleBits(1))
    CurSubStyleVal = CInt(StyleBits(2))
    StylePrivate = CInt(StyleBits(3))
    ClrBack = CLng(StyleBits(4))
    ClrText = CLng(StyleBits(5))
    LeftRighter = StyleBits(6) = "true"
    InOuter = StyleBits(7) = "true"
    TextBack = StyleBits(8) = "true"

    Select Case StyleName
      Case "fuzzy *"
        FuzzyUser LeftRighter, InOuter, TextBack
      Case "fuzzyhard"
        FuzzyHard ClrBack, LeftRighter, InOuter, TextBack
      Case "text colour *"
        SetForeColor
      Case "text colourhard"
        SetForeColorHard ClrText
      Case "blender **"
        Blender LeftRighter, InOuter, TextBack
      Case "blenderhardhard"
        BlenderHardHard ClrBack, ClrText, LeftRighter, InOuter, TextBack
      Case "blenderauto *"
        BlenderAuto LeftRighter, InOuter, TextBack
      Case "blenderautohard"
        BlenderHardAuto ClrBack, LeftRighter, InOuter, TextBack
      Case "candy"
        Candy CInt(CurSubStyleVal), StylePrivate, LeftRighter, InOuter, TextBack
      Case "dither2 *"
        Dither2 LeftRighter, InOuter, TextBack
      Case "dither2hard"
        Dither2Hard ClrBack, LeftRighter, InOuter, TextBack
      Case "dither *"
        Dither LeftRighter, InOuter, TextBack
      Case "ditherhard"
        DitherHard ClrBack, LeftRighter, InOuter, TextBack
      Case "materials"
        Materials CurSubStyle, LeftRighter, InOuter, TextBack
      Case "rainbow"
        RainBow LeftRighter, InOuter, TextBack
      Case "random"
        RandomColour CInt(CurSubStyleVal), TextBack
      Case "spectrum"
        SpectrumSector CurSubStyleVal, LeftRighter, InOuter, TextBack
      Case "highlightuser *"
        RTFHighlightUser
      Case "highlightuserhard"
        RTFHighlightHard ClrBack
      Case "highlightuserauto *"
        RTFHighlightUserAuto
      Case "highlighthardauto"
        RTFHighlightHardAuto ClrBack
        Case "highlighthard"
        RTFHighlightHard ClrBack
      Case "highlightuseruser **"
        RTFHighlightUserUser
      Case "highlighthardhard"
        RTFHighlightHardHard ClrBack, ClrText
    End Select
    If Not PreserveColour Then
        StylesPainterPreserveColours StyleName, ClrBack, ClrText, TextBack
    End If

End Sub

Public Sub StylesPainterPreserveColours(style As String, ClrBack As Long, ClrText As Long, TextBack As Boolean)

  'Used by Text Colour Panel
  'resets style to styleHard and set colours if colour should be preserved

    If InStr(style, "*") Then
        Select Case style
          Case "blender **", "blenderauto *"
            ClrText = LastBlendEndColour
            ClrBack = LastBlendStartColour
          Case "dither2 *", "dither *"
            ClrText = LastDitherColour
            ClrBack = LastDitherColour
          Case "fuzzy *", "fuzzyhard", "fuzzyuser"
            ClrText = LastFuzzyColour
            ClrBack = LastFuzzyColour
          Case "highlightuser *"
            TextBack = False
            ClrBack = LastHighlightColour
          Case "highlightuserauto *"
            TextBack = False
            ClrBack = LastHighlightColour
          Case "highlightuseruser **"
            TextBack = False
            ClrBack = LastHighlightColour
            ClrText = LastHighlightForeColour
          Case "text colour *"
            ClrText = LastTextColour
        End Select
        style = HardenStyle(style)
    End If

End Sub

Public Sub StylesReader(style As String, txt As String, Bck As String)

  'Input:     style   colour name
  'Returns:  values for that name

  Dim TmpArray As Variant

    TmpArray = Split(StylesGet(style), StylesSeparator)
    txt = TmpArray(1)
    Bck = TmpArray(2)

End Sub

Public Sub StylesRestore(Optional Force As Boolean = False)

  'See MaterialsRestore for explanation of this routine

  Dim Safety As String

    'Safety$ = " paste 'Styles.dat' contents here to replace Safety$

    Safety$ = "imperial edict^materials|gold|3|0|0|0|true|false|true^highlighthard||-1|0|255|0|false|false|false" & NL & _
             "parchment^text colourhard||-1|0|0|0|false|false|true^materials|pineboard|10|0|0|0|false|true|false" & NL & _
             "snowy night^materials|snow|15|0|0|0|true|false|true^highlighthard||-1|0|0|0|false|false|false"

    FileLineRestore m_StylesFileName, Safety, "STYLES", Force

End Sub

Private Sub StylesSave(Desc$)

    FileLinesSave m_StylesFileName, Desc$

End Sub

Private Sub SwapAnyThing(element1, element2)

  'Swap any two items in any format

  Dim TEMP As Variant

    TEMP = element1
    element1 = element2
    element2 = TEMP

End Sub

Private Function TranslateSysColor(ByVal lColor As Long) As Long

    lColor = lColor And (Not &H80000000)
    TranslateSysColor = GetSysColor(lColor)

End Function

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
    If fsLoc = 0 Then                                   'not found = there is no selected Text
        m_End_OF_RTF = 0                                'so exit
        Exit Function '>---> Bottom
    End If
    Do
        If IsNumeric(Mid$(WorkingStringRead, fsLoc + 4, 1)) Then   ' check is size not "fswiss" or other RTF code
            m_End_OF_RTF = InStr(fsLoc, WorkingStringRead, " ")    ' find space delimiting size code
            GotTrueEnd = True
        End If
        fsLoc = InStr(fsLoc + 1, WorkingStringRead, "\fs")         ' find next \fs code and move test point
    Loop Until GotTrueEnd Or fsLoc = 0
    If GotTrueEnd = False And fsLoc = 0 Then                       ' Safety should never hit
        m_End_OF_RTF = 0
    End If

End Function

Private Sub WorkingStringWrite(WStr$)

  'Copyright 2002 Roger Gilchrist
  'put a string from a routine back into the selection or whole control

    With m_RTB
        If .SelLength Then
            If .SelRTF <> WStr Then ' avoid flicker by not updating if no change has occurred
                .SelRTF = WStr
            End If
          Else '.SELLENGTH = FALSE
            If .TextRTF <> WStr Then
                .TextRTF = WStr
            End If
        End If
    End With 'M_RTB

End Sub

':) Ulli's VB Code Formatter V2.13.6 (28/08/2002 2:35:32 PM) 171 + 3529 = 3700 Lines

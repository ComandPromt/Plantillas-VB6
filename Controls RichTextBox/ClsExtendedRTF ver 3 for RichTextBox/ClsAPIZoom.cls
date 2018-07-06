VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAPIZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'copyright 2002 Roger Gilchrist
'rojagilkrist@hotmail.com
'I'm a bit compulsive about modularity so each module has its own
'Private calls to to the API. This means if you copy just one module
'you don't have to keep coming back to find that missing Public variable
'in another module. Once you have inculde a module in your program you can
'remove the Private Declare statements if you have a Public one in your project
Private m_RTB As RichTextBox
Private m_Cmbo As ComboBox
Private m_busy As Boolean
'// message constants
Private Const WM_USER As Long = &H400
Private Const EM_GETZOOM As Long = (WM_USER + 224)
Private Const EM_SETZOOM As Long = (WM_USER + 225)
Private Const WM_SETREDRAW As Long = &HB
Private ZoomFactor(13) As Long 'Change this if you add more steps to the system
Private Declare Function SendMessageRef Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub About(Optional EndUser As Boolean = True)

  'Dual About Structure
  'copyright 2002 Roger Gilchrist

  Dim MsgProg As String, MsgUser As String

    MsgProg = "ClsAPIZoom provides Zoom support for a RichTextBox control." & vbNewLine & _
              "There are four Public routines" & vbNewLine & _
              "About(T|F): About gives 1 of 2 messages Help(default) for use by end-user or this one Programmer(F)." & vbNewLine & _
              "AssignControls RTB, Cbo: Assign and setup a RichTextBox and a ComboBox for the class to use." & vbNewLine & _
              "ZoomDo: Place a call to this in the ComboBox's click event." & vbNewLine & _
              "ZoomSlug(MainStr$, T|F): Return the current Zoom value as a percent(T)-> 80 or a ratio(F)-> 4:5 added onto MainStr" & vbNewLine _
              & vbNewLine & _
              "ClsAPIZoom is brought to you by Roger Gilchrist (Email:rojagilkrist@hotmail.com)" & vbNewLine & _
              "Based on the work of Sergio Perciballi (Email:oigres@postmaster.co.uk)"

    'these are the basics an end user will need to know about but pretty it up for release
    ' this is a very basic end user help you may not want or need some parts of it so edit it to suit
    MsgUser = "You can zoom your view of the document by selecting a level" & vbNewLine & _
              "in the zoom Combobox. Note this is for ease of viewing," & vbNewLine & _
              "it does not change the real font size." & vbNewLine & _
              "A 500% zoom document will not print out as a billboard" & vbNewLine & _
              "or save its zoom factor. " & vbNewLine & _
              "The next time you open the document it will be normal size." & vbNewLine & _
              vbNewLine & _
              "Note:really large zooms may appear blank on small windows" & vbNewLine & _
              "as you are seeing only the space above a single letter such as 'a'," & vbNewLine & _
              "scroll up or down until you find text." & vbNewLine & _
              vbNewLine & _
              "Brought to you by Roger Gilchrist (Email:rojagilkrist@hotmail.com)" & vbNewLine & _
              "Based on the work of Sergio Perciballi (Email:oigres@postmaster.co.uk)"
    'I know Iif is slow but this is not time critical
    MsgBox IIf(EndUser, MsgUser, MsgProg), , IIf(EndUser, "Zooming Help", "ClsAPIZoom")

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

  'USAGE : Place a line like this where ever you want the detection to take place
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

Public Sub AssignControls(RTB As RichTextBox, Cmbo As ComboBox)

  'Copyright 2002 Roger Gilchrist
  'Place the Call to this in Form_Load or Sub Main
  '   AssignControls RichTextBox1, CommonDialog1, Combo1
  '    Form1.Show

    m_busy = False
    Set m_RTB = RTB
    Set m_Cmbo = Cmbo
    ZoomFillList

End Sub

Private Sub KeepInBounds(Lo As Long, Value, Hi As Long)

  'copyright 1988 Roger Gilchrist
  'Modified from slower all Variant version
  'Enforce minimum and maximum limits to Value
  'You could make this a function but it looks messier in code: Val = KeepinBounds(lo, Val, hi)

    If Value < Lo Then
        Value = Lo
    End If
    If Value > Hi Then
        Value = Hi
    End If

End Sub

Private Sub ZoomAddItem_Ratio(C As ComboBox, Itm$, Idta As Long)

  'Part of the Zoom system
  'copyright 1995 Roger Gilchrist
  'This is modification of Sub AddItem_ItemData(),I regularly use for Combo\List Insertion

    With C
        .AddItem Itm
        '.ItemData(.NewIndex) = ifta 'not needed because module stores it privately in next line
        ZoomFactor(.NewIndex) = Idta
    End With 'C

End Sub

Public Sub ZoomDo()

  'Copyright 2002 Roger Gilchrist
  'Place this in the Click event of the Zoom ComboBox

  Dim LNum As Long
  Dim LDen As Long
  Dim Ret As Long
  Dim IntVal As Integer

    On Error GoTo oops
    IntVal = m_Cmbo.ListIndex
    KeepInBounds LBound(ZoomFactor), IntVal, UBound(ZoomFactor)
    ZoomRatioGet ZoomFactor(IntVal), LNum, LDen

    'you SHOULD always be careful with API calls so error trap
    'BUT why bother just program against it
    KeepInBounds 1, LNum, 63
    KeepInBounds 1, LDen, 63
    Ret = SendMessageLong(m_RTB.hwnd, EM_SETZOOM, LNum, LDen)

Exit Sub

oops:
    'fall through error handler
    'put a real one here if you like
    'but I think the Zoom routines are fairly idiot proof

End Sub

Private Sub ZoomFillList()

  'Fill the Zoom Combobox with default set of fonts
  '*---PROGRAMMER MODIFICATION POINT---*
  'You might want finer separations or to extend the range
  'keep in mind the upper and lower limits of 1 & 64 for ZoomRatioSet

  Dim lIdx As Integer

    SendMessageLong m_Cmbo.hwnd, WM_SETREDRAW, False, ByVal 0&
    'speed up combo filling by disabling redraw while you add items not really but fun

    'ZoomAddItem_Ratio m_Cmbo, "Lower Limit", ZoomRatioSet(1, 63)
    ZoomAddItem_Ratio m_Cmbo, "25%", ZoomRatioSet(1, 4) 'Seems to be effective Lower limit
    ZoomAddItem_Ratio m_Cmbo, "50%", ZoomRatioSet(1, 2)
    ZoomAddItem_Ratio m_Cmbo, "75%", ZoomRatioSet(3, 4)
    ZoomAddItem_Ratio m_Cmbo, "80%", ZoomRatioSet(4, 5)
    ZoomAddItem_Ratio m_Cmbo, "90%", ZoomRatioSet(9, 10)
    ZoomAddItem_Ratio m_Cmbo, "100%", ZoomRatioSet(1, 1)
    lIdx = m_Cmbo.NewIndex ' Set inital position in list
    ZoomAddItem_Ratio m_Cmbo, "150%", ZoomRatioSet(3, 2)
    ZoomAddItem_Ratio m_Cmbo, "200%", ZoomRatioSet(2, 1)
    ZoomAddItem_Ratio m_Cmbo, "250%", ZoomRatioSet(5, 2)
    ZoomAddItem_Ratio m_Cmbo, "300%", ZoomRatioSet(3, 1)
    ZoomAddItem_Ratio m_Cmbo, "350%", ZoomRatioSet(7, 2)
    ZoomAddItem_Ratio m_Cmbo, "400%", ZoomRatioSet(4, 1)
    ZoomAddItem_Ratio m_Cmbo, "450%", ZoomRatioSet(9, 2)
    ZoomAddItem_Ratio m_Cmbo, "500%", ZoomRatioSet(5, 1)
    'ZoomAddItem_Ratio m_Cmbo, "Upper Limit", ZoomRatioSet(63, 1)
    m_Cmbo.ListIndex = lIdx ' set ComboBox to default 100%
    SendMessageLong m_Cmbo.hwnd, WM_SETREDRAW, True, ByVal 0& 'renable combo

End Sub

Private Sub ZoomRatioGet(Rat As Long, LNum As Long, LDen As Long)

  'copyright 2002 Roger Gilchrist
  'the function is mine but the very clever maths belong to Steve McMahon
  'a simple way to read ratios

    LNum = Rat \ &H10000
    LDen = Rat And &H7FFF&

End Sub

Private Function ZoomRatioSet(LNum As Long, LDen As Long) As Long

  'copyright 2002 Roger Gilchrist
  'the function is mine but the very clever maths belong to Steve McMahon
  'a simple way to store ratios '&H10000 = 65536

    ZoomRatioSet = LNum * &H10000 + LDen

End Function

Public Function ZoomSlug(MainStr$, Optional PercentT_RatioF As Boolean = True) As String

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'this is a Public wrapper for ZoomString to make it consistent with other classes in the Demo

    ZoomSlug = AddSlug(MainStr$, "<", ZoomString(PercentT_RatioF), ">")

End Function

Private Function ZoomString(PercentT_RatioF As Boolean) As String

  'Copyright 2002 Roger Gilchrist
  ' returns a string containing the current Zoom value as
  'PercentT_RatioF =True  percent format  100%
  'PercentT_RatioF =False   ratio format  1:1
  '*---PROGRAMMER MODIFICATION POINT---*
  'If you dont like the ZoomSlug stylemake this Public and call it directly
  'ZoomSlug is a wrapper for ZoomString to make it consistent with other classes in the Demo

  Dim LNum As Long
  Dim LDen As Long

    On Error GoTo oops
    SendMessageRef m_RTB.hwnd, EM_GETZOOM, LNum, LDen
    If LNum > 0 And LDen > 0 Then
        If PercentT_RatioF Then
            ZoomString = CStr(LNum / LDen * 100) & "%"
          Else 'PERCENTT_RATIOF = FALSE
            ZoomString = LNum & ":" & LDen
        End If
    End If

Exit Function

oops:

End Function

':) Ulli's VB Code Formatter V2.13.6 (28/08/2002 2:37:54 PM) 20 + 265 = 285 Lines

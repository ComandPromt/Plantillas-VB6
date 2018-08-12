VERSION 5.00
Object = "{D28F8786-0BB9-402B-92DC-F32DE23A324E}#3.0#0"; "OutlookBar.ocx"
Begin VB.Form Form4 
   Caption         =   "Form1"
   ClientHeight    =   4452
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   5496
   LinkTopic       =   "Form1"
   ScaleHeight     =   4452
   ScaleWidth      =   5496
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Large icons"
      Height          =   264
      Left            =   3948
      TabIndex        =   4
      Top             =   84
      Width           =   1356
   End
   Begin VB.ListBox List1 
      Height          =   3408
      IntegralHeight  =   0   'False
      Left            =   2772
      TabIndex        =   3
      Top             =   420
      Width           =   2532
   End
   Begin OutlookBar.ctxOutlookBar ctxOutlookBar1 
      Height          =   3972
      Left            =   168
      TabIndex        =   0
      Top             =   168
      Width           =   2028
      _ExtentX        =   3577
      _ExtentY        =   7006
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormatControl   =   "Form4.frx":0000
      FormatGroup     =   "Form4.frx":0188
      FormatGroupHover=   "Form4.frx":14E8
      FormatGroupPressed=   "Form4.frx":15B0
      FormatGroupSelected=   "Form4.frx":2898
      FormatItem      =   "Form4.frx":2944
      FormatItemLargeIcons=   "Form4.frx":2A54
      FormatItemHover =   "Form4.frx":2B50
      FormatItemPressed=   "Form4.frx":2BFC
      FormatItemSelected=   "Form4.frx":2CA8
      FormatSmallIcon =   "Form4.frx":2D54
      FormatSmallIconHover=   "Form4.frx":2E64
      FormatSmallIconPressed=   "Form4.frx":367C
      FormatSmallIconSelected=   "Form4.frx":3ED0
      FormatLargeIcon =   "Form4.frx":46E8
      FormatLargeIconHover=   "Form4.frx":47F8
      FormatLargeIconPressed=   "Form4.frx":5010
      FormatLargeIconSelected=   "Form4.frx":5864
      Groups          =   "Form4.frx":607C
      OleDragMode     =   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4296
      Left            =   0
      ScaleHeight     =   4296
      ScaleWidth      =   2448
      TabIndex        =   1
      Top             =   0
      Width           =   2448
   End
   Begin VB.Label Label1 
      Caption         =   "Views:"
      Height          =   264
      Left            =   2772
      TabIndex        =   2
      Top             =   84
      Width           =   1524
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STR_PRESET_VIEWS As String = "default|BaseView|FlatView|Office1View|Office2View|Office3View|VSToolBoxView|AdvExplorerBarView|ExplorerBarView|UltraFlatExplorerBarView|XP1View|XP2View"

Private Sub Check1_Click()
    ctxOutlookBar1.Groups(1).IconsType = IIf(Check1.Value = vbChecked, ucsIcsLargeIcons, ucsIcsSmallIcons)
End Sub

Private Sub ctxOutlookBar1_ButtonClick(ByVal oBtn As OutlookBar.cButton)
    If oBtn.Index = 1 And oBtn.Class = ucsBtnClassGroup Then
        oBtn.IconsType = 1 - oBtn.IconsType
        Check1.Value = IIf(oBtn.IconsType = ucsIcsLargeIcons, vbChecked, vbUnchecked)
    End If
End Sub

Private Sub Form_Load()
    Dim vIter           As Variant

    For Each vIter In Split(STR_PRESET_VIEWS, "|")
        List1.AddItem vIter
    Next
    List1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ctxOutlookBar1.Height = ScaleHeight - 2 * ctxOutlookBar1.Top
    Picture1.Height = ScaleHeight
    List1.Height = ScaleHeight - List1.Top - ctxOutlookBar1.Top
    List1.Width = ScaleWidth - List1.Left - ctxOutlookBar1.Top
End Sub

Private Sub List1_Click()
    Dim nFile           As Integer
    Dim bArray()        As Byte
    Dim vFmts           As Variant
    Dim vIter           As Variant
    Dim oFmt            As cFormatDef
    
    Screen.MousePointer = vbHourglass
    '--- read from file
            '    nFile = FreeFile
            '    Open App.Path & "\Formats\" & List1.List(List1.ListIndex) & ".obf" For Binary As #nFile
            '    ReDim bArray(1 To LOF(nFile))
            '    Get nFile, , bArray
            '    Close nFile
    '--- read from resource
    bArray = LoadResData(List1.List(List1.ListIndex) & ".obf", "FORMAT")
    '--- construct an array with all format (to be loaded)
    vFmts = Array( _
                ctxOutlookBar1.FormatControl, _
                ctxOutlookBar1.FormatGroup, _
                ctxOutlookBar1.FormatGroupHover, _
                ctxOutlookBar1.FormatGroupPressed, _
                ctxOutlookBar1.FormatGroupSelected, _
                ctxOutlookBar1.FormatItem, _
                ctxOutlookBar1.FormatItemHover, _
                ctxOutlookBar1.FormatItemPressed, _
                ctxOutlookBar1.FormatItemSelected, _
                ctxOutlookBar1.FormatItemLargeIcons, _
                ctxOutlookBar1.FormatSmallIcon, _
                ctxOutlookBar1.FormatSmallIconHover, _
                ctxOutlookBar1.FormatSmallIconPressed, _
                ctxOutlookBar1.FormatSmallIconSelected, _
                ctxOutlookBar1.FormatLargeIcon, _
                ctxOutlookBar1.FormatLargeIconHover, _
                ctxOutlookBar1.FormatLargeIconPressed, _
                ctxOutlookBar1.FormatLargeIconSelected)
    '--- load prop page and iterate array loading each format
    With New PropertyBag
        .Contents = bArray
        For Each vIter In vFmts
            Set oFmt = vIter
            oFmt.Contents = .ReadProperty(Replace(oFmt.Name, " ", ""), oFmt.Contents)
        Next
    End With
    '--- fix background
    Select Case List1.List(List1.ListIndex)
    Case "AdvExplorerBarView"
        Picture1.BackColor = vbBlue
    Case "ExplorerBarView", "UltraFlatExplorerBarView"
        Picture1.BackColor = vbWhite
    Case Else
        Picture1.BackColor = vbButtonFace
    End Select
    Screen.MousePointer = vbDefault
End Sub

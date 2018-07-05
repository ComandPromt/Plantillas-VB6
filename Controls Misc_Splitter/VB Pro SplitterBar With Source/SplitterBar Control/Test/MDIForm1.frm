VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDIForm1"
   ClientHeight    =   8145
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11025
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7770
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picHolder 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7350
      Left            =   0
      ScaleHeight     =   7350
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   420
      Width           =   3195
      Begin VB.PictureBox Picture1 
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1875
         ScaleWidth      =   1395
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         Begin MSComctlLib.TreeView tvMenu 
            Height          =   5595
            Left            =   0
            TabIndex        =   3
            Top             =   240
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   9869
            _Version        =   393217
            HideSelection   =   0   'False
            LabelEdit       =   1
            Sorted          =   -1  'True
            Style           =   7
            ImageList       =   "ImgMenu"
            Appearance      =   0
         End
      End
      Begin VB.PictureBox PicSplit 
         Height          =   5895
         Left            =   3120
         ScaleHeight     =   5895
         ScaleWidth      =   45
         TabIndex        =   1
         Top             =   0
         Width           =   50
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   3855
      Top             =   2475
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   "Justify"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0112
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0224
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0336
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0448
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":055A
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":066C
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":077E
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Justify"
            Object.ToolTipText     =   "Justify"
            ImageKey        =   "Justify"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Align Left"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Align Right"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu ddd 
      Caption         =   "ddd"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Splitt As Lib_SplitterBar

Private Sub PicSplit1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    splitX = True
    
    Splitt1.SplitterMouseDown hWnd, x, y
    
    
End Sub

Private Sub Picture1_Resize()
tvMenu.Move 0, 0, Picture1.Width, Picture1.Height

End Sub

Private Sub MDIForm_Load()
    Set Splitt = New Lib_SplitterBar
         Splitt.SplitObject = Me.PicSplit
         Splitt.Border(espbLeft) = 64
         Splitt.SPValues(espvtop) = Me.Toolbar1.Top + Me.Toolbar1.Height
         Splitt.SPValues(espvBottom) = Me.StatusBar1.Height + 75
         
End Sub
Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   
    Splitt.SplitterContainer_MouseMove x, y

End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
 
    If Splitt.SplitterContainer_MouseUp(x, y) Then
             MDIForm_Resize
    End If
  
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set Splitt = Nothing

End Sub

Private Sub picHolder_Resize()
    
    Picture1.Move 0, 0, picHolder.Width - PicSplit.Width, picHolder.Height
    
End Sub

Private Sub PicSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Splitt.SplitterMouseDown Me.hWnd, x, y

End Sub
Private Sub MDIForm_Resize()

    picHolder.Move 0, 0, PicSplit.Left + PicSplit.Width, ScaleHeight
  
End Sub



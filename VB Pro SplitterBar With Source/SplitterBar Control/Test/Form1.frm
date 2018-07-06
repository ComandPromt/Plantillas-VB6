VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   10035
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons(1)"
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
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2775
      ScaleWidth      =   10035
      TabIndex        =   3
      Top             =   3090
      Width           =   10035
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   1335
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   1920
         ScaleHeight     =   135
         ScaleWidth      =   4140
         TabIndex        =   5
         Top             =   960
         Width           =   4140
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command1"
         Height          =   1215
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   2640
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   2280
      ScaleHeight     =   1155
      ScaleWidth      =   120
      TabIndex        =   0
      Top             =   840
      Width           =   180
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   1
      Left            =   0
      Top             =   480
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
            Picture         =   "Form1.frx":0000
            Key             =   "Justify"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0112
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0224
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0336
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0448
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":055A
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":066C
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":077E
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Test Splitter"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Lib_SplitterBar
Dim d As Lib_SplitterBar

Private Sub Form_Load()

Set c = New Lib_SplitterBar
Set d = New Lib_SplitterBar

    With c
        .SplitObject = Me.Picture1
        .Orientation = espVertical
        .Border(espbLeft) = 10
        
        '-- Set Special Adjusments
        .SPValues(espvtop) = (Toolbar1.Top + Toolbar1.Height) / 1.5
        .SPValues(espvBottom) = Me.Picture2.Height
    End With
        
    With d
        .SplitObject = Me.Picture3
        .Orientation = espHorizontal
        .UseInternalAjustments = True
    End With
    Picture2_MouseUp 0, 0, 1, 1
    
    Form_Resize

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    c.SplitterContainer_MouseMove x, y
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If c.SplitterContainer_MouseUp(x, y) Then
        Form_Resize
    End If
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set c = Nothing
    Set d = Nothing

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.Command1.Width = Picture1.Left
    Command1.Move 0, Me.Toolbar1.Top + Me.Toolbar1.Height, Picture1.Left - 60, ScaleHeight - Me.Toolbar1.Height - Picture2.ScaleHeight  '- Me.StatusBar1.Height
    Text1.Move Picture1.Left + 60, Command1.Top, ScaleWidth - Picture1.Left - 60, Command1.Height
    Picture1.Height = Command1.Height - Picture3.Height
    Picture3.Move 0, Picture3.Top, Picture2.Width, Picture3.Height
    Command2.Move 0, 0, Picture2.Width, Picture3.Top
    Command3.Move 0, Picture3.Top + Picture3.Height, Picture3.Width, Picture2.Height - Picture3.Top + Picture3.Height
    
End Sub


Private Sub mnuExit_Click()
End
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    c.SplitterMouseDown hWnd, x, y

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    d.SplitterContainer_MouseMove x, y

End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    
    If d.SplitterContainer_MouseUp(x, y) Then
        Picture3.Move 0, Picture3.Top, Picture2.Width, Picture3.Height
        Command2.Move 0, 0, Picture2.Width, Picture3.Top
        Command3.Move 0, Picture3.Top + Picture3.Height, Picture3.Width, Picture2.Height - Picture3.Top + Picture3.Height
    End If
    
End Sub

Private Sub Picture2_Resize()
Form_Resize
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    d.SplitterMouseDown Picture2.hWnd, x, y
    
End Sub

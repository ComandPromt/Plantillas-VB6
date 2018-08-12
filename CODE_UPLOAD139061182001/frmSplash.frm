VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3120
   ClientLeft      =   3675
   ClientTop       =   3450
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6345
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   240
         Top             =   1605
      End
      Begin VB.CommandButton OK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&OK"
         Height          =   315
         Left            =   2685
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www15.brinkster.com/rgb2000"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   1845
         TabIndex        =   6
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1185
         Left            =   750
         Picture         =   "frmSplash.frx":000C
         Top             =   840
         Width           =   4905
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rb.sb@gte.net"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   240
         Left            =   4560
         TabIndex        =   4
         Top             =   2640
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright 2000"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   4680
         TabIndex        =   3
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Label lblproductname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Tracker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1575
         TabIndex        =   2
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   3165
         TabIndex        =   1
         Top             =   420
         Width           =   1920
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim WithEvents IE As InternetExplorer
Attribute IE.VB_VarHelpID = -1

Private Sub Form_Activate()
    If Started = False Then
        Timer1.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    lblEmail.ForeColor = &HFF8080
    lblURL.ForeColor = &HFF8080
    frmSplash.MousePointer = vbDefault
    
End Sub

Private Sub IE_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Set IE = Nothing
End Sub

Private Sub lblEmail_Click()

    'send email using Internet Explorer object and make it invisible
    Set IE = New InternetExplorer
    IE.Visible = False
    IE.Navigate "Mailto:rb.sb@gte.net"
    
End Sub

Private Sub lblEmail_DblClick()

    'send email using Internet Explorer object and make it invisible
    Set IE = New InternetExplorer
    IE.Visible = False
    IE.Navigate "Mailto:rb.sb@gte.net"
    
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblEmail.ForeColor = &HFFC0FF
    frmSplash.MouseIcon = LoadPicture("hand.cur")
    frmSplash.MousePointer = vbCustom
    
End Sub

Private Sub lblURL_Click()
    'goto my website using Internet Explorer object and make it invisible
    Set IE = New InternetExplorer
    IE.Visible = True
    IE.Navigate "http://www15.brinkster.com/rgb2000"
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    lblURL.ForeColor = &HFFC0FF
    frmSplash.MouseIcon = LoadPicture("hand.cur")
    frmSplash.MousePointer = vbCustom
    
End Sub

Private Sub OK_Click()
    Unload Me
End Sub




Private Sub Timer1_Timer()

    Load Mainform
    Load RecruiterForm

End Sub

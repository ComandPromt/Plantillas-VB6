VERSION 5.00
Begin VB.Form FOpSys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COpSysInfo Demo"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5715
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   435
      Left            =   1560
      TabIndex        =   0
      Top             =   2460
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   1620
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   3795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   1935
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "FOpSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private os As COpSysInfo

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim nLines As Long
   '
   ' Fill descriptive label...
   '
   Label1.Caption = _
      "Major Version:" & vbCrLf & _
      "Minor Version:" & vbCrLf & _
      "Build Number:" & vbCrLf & _
      "Platform ID:" & vbCrLf & _
      "CSDVersion:" & vbCrLf & vbCrLf & _
      "Platform:" & vbCrLf & _
      "Version:" & vbCrLf & vbCrLf & _
      "IsWinNT:" & vbCrLf & _
      "IsWin95:" & vbCrLf & _
      "IsWin98:" & vbCrLf
   '
   ' Fill in data from class...
   '
   Set os = New COpSysInfo
   Text1.Text = _
      os.MajorVersion & vbCrLf & _
      os.MinorVersion & vbCrLf & _
      os.BuildNumber & vbCrLf & _
      os.PlatformID & vbCrLf & _
      os.CSDVersion & vbCrLf & vbCrLf & _
      os.Platform & vbCrLf & _
      os.Version & vbCrLf & vbCrLf & _
      os.IsWinNT & vbCrLf & _
      os.IsWin95 & vbCrLf & _
      os.IsWin98 & vbCrLf
   '
   ' Number of lines of text...
   '
   nLines = 12
   '
   ' Adjust controls/form for good display...
   '
   With Label1
      .Width = Me.TextWidth("Major Version:") * 1.1
      .Height = Me.TextHeight("X") * nLines
      Text1.Left = .Left + .Width _
         + (5 * Screen.TwipsPerPixelX)
   End With
   With Text1
      .Height = Label1.Height
      .Width = Me.TextWidth(os.Version) * 1.1
      Me.Width = .Left + .Width + Label1.Left
      Me.Height = .Height + (.Top * 3) _
         + Command1.Height _
         + (Me.Height - Me.ScaleHeight)
      Command1.Move Me.ScaleWidth \ 2 _
         - Command1.Width \ 2, _
         (.Top * 2) + .Height
   End With
   '
   ' Set flag icon into titlebar...
   '
   Set Me.Icon = Nothing
End Sub


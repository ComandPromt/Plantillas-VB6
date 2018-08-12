VERSION 5.00
Object = "*\AProject2.vbp"
Begin VB.Form Form21 
   Caption         =   "Form1"
   ClientHeight    =   3996
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   7644
   LinkTopic       =   "Form1"
   ScaleHeight     =   3996
   ScaleWidth      =   7644
   StartUpPosition =   3  'Windows Default
   Begin Project2.ctxTaskBar ctxTaskBar1 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      Top             =   3624
      Width           =   7644
      _ExtentX        =   13483
      _ExtentY        =   656
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove tray icon"
      Height          =   432
      Left            =   2016
      TabIndex        =   4
      ToolTipText     =   "1111111"
      Top             =   756
      Width           =   1524
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove task"
      Height          =   432
      Left            =   336
      TabIndex        =   3
      Top             =   756
      Width           =   1608
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Reduce flicker (uses more memory)"
      Height          =   264
      Left            =   336
      TabIndex        =   2
      Top             =   1512
      Width           =   4464
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add task"
      Height          =   432
      Left            =   336
      TabIndex        =   1
      Top             =   168
      Width           =   1608
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add tray icon"
      Height          =   432
      Left            =   2016
      TabIndex        =   0
      ToolTipText     =   "1111111"
      Top             =   168
      Width           =   1524
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   6216
      Picture         =   "Form21.frx":0000
      Top             =   84
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   3
      Left            =   5712
      Picture         =   "Form21.frx":0968
      Stretch         =   -1  'True
      Top             =   84
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.Image Image1 
      Height          =   264
      Index           =   2
      Left            =   5376
      Picture         =   "Form21.frx":2C0A
      Top             =   84
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image Image1 
      Height          =   384
      Index           =   1
      Left            =   4956
      Picture         =   "Form21.frx":2D94
      Top             =   84
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Image1 
      Height          =   384
      Index           =   0
      Left            =   4536
      Picture         =   "Form21.frx":31D6
      Top             =   84
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lCount As Long

Private Sub Check1_Click()
    ctxTaskBar1.BufferDraw = (Check1.Value = vbChecked)
End Sub

Private Sub Command1_Click()
    m_lCount = m_lCount + 1
    ctxTaskBar1.TrayIcons.Add m_lCount & " Added at " & Timer, , Image1(m_lCount Mod 5).Picture
End Sub

Private Sub Command2_Click()
    m_lCount = m_lCount + 1
    ctxTaskBar1.Tasks.Add m_lCount & " You have mail " & Timer, "", Image1(m_lCount Mod 5).Picture
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    ctxTaskBar1.Tasks.Remove 1
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    ctxTaskBar1.TrayIcons.Remove 1
End Sub

Private Sub ctxTaskBar1_TaskMouseDown(ByVal Idx As Long, ByVal Button As Long, ByVal x As Long, ByVal y As Long)
    Debug.Print "ctxTaskBar1_TaskMouseDown " & Idx & " " & Button & " " & x & " " & y
End Sub

Private Sub ctxTaskBar1_TaskMouseUp(ByVal Idx As Long, ByVal Button As Long, ByVal x As Long, ByVal y As Long)
    Debug.Print "ctxTaskBar1_TaskMouseUp " & Idx & " " & Button & " " & x & " " & y
End Sub

Private Sub Form_Load()
    ctxTaskBar1.Tasks.Add "Test", , Image1(0).Picture
    ctxTaskBar1.Tasks.Add "Envelope", , Image1(2).Picture
    ctxTaskBar1.Tasks.Add "Metafile", , Image1(3).Picture
    ctxTaskBar1.Tasks.Add "EnhMetafile", , Image1(4).Picture
    ctxTaskBar1.TrayIcons.Add "Proba", , Image1(1).Picture
    ctxTaskBar1.TrayIcons.Add "TrayIcon", , Image1(0).Picture
End Sub

Private Sub ctxTaskBar1_BeforeTaskSwitch(ByVal NewTask As Long, Cancel As Boolean)
    Cancel = MsgBox("Switching from " & ctxTaskBar1.ActiveTask & " to " & NewTask, vbExclamation + vbOKCancel) = vbCancel
    If ctxTaskBar1.ActiveTask = NewTask Then
        ctxTaskBar1.ActiveTask = -1
        Cancel = True
    End If
End Sub

Private Sub ctxTaskBar1_StartMenu()
    MsgBox "Start menu"
End Sub

Private Sub ctxTaskBar1_TrayMouseDown(ByVal Idx As Long, ByVal Button As Long, ByVal x As Long, ByVal y As Long)
    Debug.Print "ctxTaskBar1_TrayMouseDown " & Idx & " " & Button & " " & x & " " & y
End Sub

Private Sub ctxTaskBar1_TrayMouseUp(ByVal Idx As Long, ByVal Button As Long, ByVal x As Long, ByVal y As Long)
    Debug.Print "ctxTaskBar1_TrayMouseUp " & Idx & " " & Button & " " & x & " " & y
End Sub

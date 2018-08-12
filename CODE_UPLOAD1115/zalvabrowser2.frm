VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form2 
   Caption         =   "Gathering Production Browser"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "zalvabrowser2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   840
      Top             =   2640
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      Caption         =   "WorkStationSF - use as a notepade"
      Height          =   1095
      Left            =   6840
      TabIndex        =   8
      Top             =   0
      Width           =   4935
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1560
      Top             =   7920
   End
   Begin SHDocVwCtl.WebBrowser brw 
      Height          =   6615
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   11655
      ExtentX         =   20558
      ExtentY         =   11668
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Home"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Search"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Forward"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin SHDocVwCtl.ShellFolderViewOC ShellFolderViewOC1 
      Left            =   2280
      Top             =   5040
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu df 
      Caption         =   "?"
      Begin VB.Menu help 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Integer

Private Sub brw_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Combo1.Text = brw.LocationURL
End Sub

Private Sub brw_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
If Label1.Caption = ">>>>>>>>>>>>>>>>" Then Label1.Caption = ""
Label1.Caption = Label1.Caption + ">"
End Sub

Private Sub brw_TitleChange(ByVal Text As String)
Form2.Caption = "Gathering Production Browser" & " - " & brw.LocationName
End Sub

Private Sub Command1_Click()
On Error Resume Next
brw.GoBack
End Sub

Private Sub Command2_Click()
On Error Resume Next
brw.GoForward
End Sub

Private Sub Command3_Click()
brw.Stop
End Sub

Private Sub Command4_Click()
brw.Refresh
End Sub

Private Sub Command5_Click()
brw.GoSearch
End Sub

Private Sub Command6_Click()
brw.GoHome
End Sub

Private Sub Command7_Click()
Form3.Show
End Sub

Private Sub Form_Load()
brw.GoHome

p = 1
End Sub

Private Sub combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
brw.Navigate Combo1.Text
Combo1.AddItem Combo1.Text
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub norm_Click()
Form1.Show
Unload Form2
End Sub

Private Sub quit_Click()
Unload Me
End Sub


Private Sub Text2_dblclick()
Text2.Text = ""
End Sub

Private Sub Timer1_Timer()
If Label1.Caption = ">>>>>>>>>>>>>>>>" Then Label1.Caption = ""
Label1.Caption = Label1.Caption + ">"
End Sub


Private Sub Timer2_Timer()
If p = 3 Then p = 1

If p = 1 Then
brw.Navigate ("http://www.yahoo.com")
p = p + 1
End If

If p = 2 Then
brw.Navigate ("http://www.altavista.com")
p = p + 1
End If


End Sub

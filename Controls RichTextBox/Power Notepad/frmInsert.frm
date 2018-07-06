VERSION 5.00
Begin VB.Form frmInsert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Picture"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Add Picture"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Load Picture"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   5115
      TabIndex        =   5
      Top             =   4320
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   5175
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   720
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Picture:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5400
      Y1              =   3840
      Y2              =   3840
   End
End
Attribute VB_Name = "frmInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Picture1.Picture = LoadPicture(Text1.Text)
Picture1.AutoSize = True
frmMain.Picture1.Picture = LoadPicture(Text1.Text)
frmMain.Picture1.AutoSize = True
End Sub

Private Sub Command2_Click()
Image1.Picture = Picture1.Picture
frmMain.Image1.Picture = frmMain.Picture1.Picture
AddPic2RTB
frmMain.Picture1.Visible = True
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()
Text1.Text = "" & Dir1 & "\" & File1.FileName & ""

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nader&
    ReleaseCapture
    nader& = SendMessage(Picture1.hWnd, &H112, &HF012, 0)
End Sub

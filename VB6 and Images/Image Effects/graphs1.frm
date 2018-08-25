VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "SOME COOL PICTURE EFFECTS FOR GRAPHIC PROGRAMMERS by KAYHAN TANRISEVEN 'THE BENCHMARKER'"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      Caption         =   "Effect 10"
      Height          =   435
      Left            =   5685
      TabIndex        =   15
      Top             =   7365
      Width           =   6195
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Effect 9"
      Height          =   450
      Left            =   9960
      TabIndex        =   14
      Top             =   6870
      Width           =   1905
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Effect 8"
      Height          =   435
      Left            =   7815
      TabIndex        =   13
      Top             =   6885
      Width           =   2115
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Effect 7"
      Height          =   435
      Left            =   5670
      TabIndex        =   12
      Top             =   6885
      Width           =   2115
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Effect 6"
      Height          =   435
      Left            =   9960
      TabIndex        =   11
      Top             =   6420
      Width           =   1920
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Effect 5"
      Height          =   435
      Left            =   9960
      TabIndex        =   10
      Top             =   5955
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   150
      Picture         =   "graphs1.frx":0000
      ScaleHeight     =   405
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   7275
      Width           =   555
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Effect 3"
      Height          =   435
      Left            =   7830
      TabIndex        =   5
      Top             =   6420
      Width           =   2115
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Effect 4"
      Height          =   435
      Left            =   7845
      TabIndex        =   4
      Top             =   5970
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Effect 2"
      Height          =   435
      Left            =   5670
      TabIndex        =   3
      Top             =   6420
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   435
      Left            =   5670
      TabIndex        =   2
      Top             =   5970
      Width           =   2115
   End
   Begin VB.PictureBox Picture1 
      Height          =   5475
      Left            =   5700
      Picture         =   "graphs1.frx":0442
      ScaleHeight     =   5415
      ScaleWidth      =   6000
      TabIndex        =   1
      Top             =   465
      Width           =   6060
   End
   Begin VB.PictureBox Picture2 
      Height          =   6555
      Left            =   75
      Picture         =   "graphs1.frx":A988
      ScaleHeight     =   6495
      ScaleWidth      =   5505
      TabIndex        =   0
      Top             =   465
      Width           =   5565
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   225
         Top             =   4530
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   480
         ImageHeight     =   360
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "graphs1.frx":19AE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "graphs1.frx":24038
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture 3"
      Height          =   360
      Left            =   90
      TabIndex        =   9
      Top             =   7050
      Width           =   2730
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture 1"
      Height          =   300
      Left            =   5790
      TabIndex        =   7
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture 2"
      Height          =   345
      Left            =   180
      TabIndex        =   6
      Top             =   90
      Width           =   2220
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "Effect 1" Then
Picture2.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, &H8800C6
Command1.Caption = "Restore"
Else
Picture2.Picture = ImageList1.ListImages(2).Picture
Picture1.Picture = ImageList1.ListImages(1).Picture
Command1.Caption = "Effect 1"
End If
End Sub

Private Sub Command10_Click()
If Command10.Caption = "Effect 10" Then
Picture1.Visible = False
Picture2.Visible = False
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Caption = "THIS IS THE WAY YOU TILE PICTURES OVER FORM,WITHOUT USING PICTUREBOX...."
Command10.Caption = "Restore"
Else
Picture1.Visible = True
Picture2.Visible = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command9.Visible = True
Caption = "SOME COOL PICTURE EFFECTS FOR GRAPHIC PROGRAMMERS by KAYHAN TANRISEVEN 'THE BENCHMARKER'"
Command10.Caption = "Effect 10"
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Effect 2" Then
Picture2.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, &H1100A6
Command2.Caption = "Restore"
Else
Picture2.Picture = ImageList1.ListImages(2).Picture
Picture1.Picture = ImageList1.ListImages(1).Picture
Command2.Caption = "Effect 2"
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "Effect 4" Then
Picture2.PaintPicture Picture2.Picture, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, &H330008
Command3.Caption = "Restore"
Else
Picture2.Picture = ImageList1.ListImages(2).Picture
Command3.Caption = "Effect 4"
End If
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Effect 3" Then
Picture2.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, &HCC0020
Command4.Caption = "Restore"
Else
Picture2.Picture = ImageList1.ListImages(2).Picture
Command4.Caption = "Effect 3"
End If
End Sub

Private Sub Command5_Click()
If Command5.Caption = "Effect 5" Then
Picture2.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, &H660046
Command5.Caption = "Restore"
Else
Picture2.Picture = ImageList1.ListImages(2).Picture
Command5.Caption = "Effect 5"
End If
End Sub

Private Sub Command6_Click()
If Command6.Caption = "Effect 6" Then
Picture2.PaintPicture Picture2.Picture, Picture2.ScaleWidth, Picture2.ScaleHeight, -Picture2.ScaleWidth, -Picture2.ScaleHeight, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, &HCC0020
Command6.Caption = "Restore"
Else:
Picture2.Picture = ImageList1.ListImages(2).Picture
Command6.Caption = "Effect 6"
End If
End Sub

Private Sub Command7_Click()
If Command7.Caption = "Effect 7" Then
Picture2.PaintPicture Picture2.Picture, Picture2.ScaleHeight / 2, Picture2.ScaleHeight / 2, Picture2.ScaleWidth / 2, Picture2.ScaleWidth / 2, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, &HCC0020
Command7.Caption = "Restore"
Else:
Picture2.Picture = ImageList1.ListImages(2).Picture
Command7.Caption = "Effect 7"
End If
End Sub

Private Sub Command8_Click()
If Command8.Caption = "Effect 8" Then
Picture2.PaintPicture Picture2.Picture, Picture2.ScaleWidth, 0, -Picture2.ScaleWidth, Picture2.ScaleHeight, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, &HCC0020

Command8.Caption = "Restore"
Else:
Picture2.Picture = ImageList1.ListImages(2).Picture
Command8.Caption = "Effect 8"
End If
End Sub

Private Sub Command9_Click()
If Command9.Caption = "Effect 9" Then
Picture2.PaintPicture Picture2.Picture, 0, Picture1.ScaleHeight, Picture2.ScaleWidth, -Picture2.ScaleHeight, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, &HCC0020
Command9.Caption = "Restore"
Else:
Picture2.Picture = ImageList1.ListImages(2).Picture
Command9.Caption = "Effect 9"
End If
End Sub
Private Sub Form_Paint()
Dim i, j
 For i = 0 To ScaleWidth Step Picture3.Width
  For j = 0 To ScaleHeight Step Picture3.Height
   PaintPicture Picture3.Picture, i, j, Picture3.Width, Picture3.Height, 0, 0
  Next
 Next
End Sub

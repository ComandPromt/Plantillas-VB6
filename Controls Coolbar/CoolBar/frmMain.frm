VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   Caption         =   "CoolBar Demonstration"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   240
      ScaleHeight     =   1875
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1349
      _CBWidth        =   4680
      _CBHeight       =   765
      _Version        =   "6.0.8141"
      MinHeight1      =   315
      Width1          =   705
      NewRow1         =   0   'False
      Child2          =   "cboColor"
      MinHeight2      =   315
      Width2          =   1200
      NewRow2         =   -1  'True
      MinHeight3      =   360
      Width3          =   1440
      NewRow3         =   0   'False
      Begin VB.CommandButton cmdChange 
         Appearance      =   0  'Flat
         Caption         =   "Command1"
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox cboColor 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   165
         List            =   "frmMain.frx":000D
         TabIndex        =   1
         Top             =   390
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' **************************************
' CoolBar Demonstration
' For Visual Basic Programmer's Journal
' October 1998
' By Jeffrey P. McManus
' jeffreyp@sirius.com
' http://www.redblazer.com/vbdb/
' **************************************
'

Private Sub cboColor_Click()
    Select Case cboColor.Text
        Case "Red"
        Picture1.BackColor = RGB(255, 0, 0)
        
        Case "Green"
        Picture1.BackColor = RGB(0, 255, 0)
        
        Case "Blue"
        Picture1.BackColor = RGB(0, 0, 255)
        
    End Select
End Sub

Private Sub Form_Load()

End Sub

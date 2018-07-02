VERSION 5.00
Begin VB.Form FrmMainApp 
   Caption         =   "your application"
   ClientHeight    =   5505
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7980
   Icon            =   "FrmMainApp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3840
      Picture         =   "FrmMainApp.frx":1272
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsChangePassword 
         Caption         =   "&Change Password"
      End
   End
End
Attribute VB_Name = "FrmMainApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    If MsgBox("     Thanks for trying this application!    Did you try to change the password?  ", 1, "Bye!") = 1 _
    Then End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuOptionsChangePassword_Click()
    FrmChangePassword.Show
End Sub

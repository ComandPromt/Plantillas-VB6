VERSION 5.00
Begin VB.Form FrmChangePassword 
   Caption         =   "Change Password"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4395
   Icon            =   "FrmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame FrChangePassword 
      Caption         =   "Change Password"
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtExistingPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtNewPassword1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtNewPassword2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Existing Password"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Enter New Password"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Confirm New Password"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    FrmChangePassword.Hide
End Sub


Private Sub cmdOK_Click()

Dim strTemp As String
Dim strPW As String
Dim strNewPW As String
Dim strEncryptNewPW As String
    'some error handling
    
    strPW = GetValue("Main", "Password", App.Path & "\" & con_INI_File)
    strNewPW = LCase(txtNewPassword2.Text)
    'checks to see if you type int he correct password in the existing password field
        
     If FrmLogin.txtPassword = LCase(txtExistingPassword.Text) Then
        'checks the match of the new passwords
        
        If LCase(txtNewPassword1.Text) = strNewPW Then
            strEncryptNewPW = Encrypt(strNewPW)
            PutValue "Main", "Password", strEncryptNewPW, App.Path & "\" & con_INI_File
            MsgBox "Password changed!", 8, "Password Verfication"
        
        Else
            MsgBox "The New Passwords Do Not Match", 8, "Password Error"
            txtNewPassword1.SetFocus
            Exit Sub
        
        End If
        
    Else
        MsgBox "The Existing Password is Incorrect!", 8, "Password Error"
        txtExistingPassword.SetFocus
        Exit Sub
        
    End If
    'if the existing password matches the decrypted password and
    'both the new passwords match, then it changes the password to
    'be encrypted in the ini file (and then hides the change
    'password dialog box)
    
    FrmChangePassword.Hide
    DoEvents
    
End Sub


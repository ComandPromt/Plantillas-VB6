VERSION 5.00
Begin VB.Form FrmLogin 
   Caption         =   "Password Protected"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrGarbage 
      Caption         =   "Erase frame and contents"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   3615
      Begin VB.CommandButton cmdGetRideOfThisButton 
         BackColor       =   &H0080FF80&
         Caption         =   "Click here to display password"
         Height          =   495
         Index           =   3
         Left            =   240
         MaskColor       =   &H0080FF80&
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Password"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   600
         MouseIcon       =   "FrmLogin.frx":1272
         Picture         =   "FrmLogin.frx":157C
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Password"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    End
End Sub
'cut from here
Private Sub cmdGetRideOfThisButton_Click(Index As Integer)

Dim strTest As String
    strTest = GetValue("Main", "Password", App.Path & "\" & con_INI_File)
   
    MsgBox Decrypt(strTest), 8, " Get rid of this button for the application you put this on!"
    
End Sub
'to here, to get rid of the code that shows the decrypted password from the ini
'file....also delete the frame in view object

Private Sub cmdSubmit_Click()

    Dim strTest As String
    strTest = GetValue("Main", "Password", App.Path & "\" & con_INI_File)
   
     If LCase(txtPassword.Text) = Decrypt(strTest) Then
        
        FrmMainApp.Show
        ' The name of the main application
        FrmLogin.Hide
        ' Hides the login dialog box
        
    Else
        MsgBox "Enter a Valid Password for this System", 8, "Password Error"
        txtPassword.SetFocus
        Exit Sub
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

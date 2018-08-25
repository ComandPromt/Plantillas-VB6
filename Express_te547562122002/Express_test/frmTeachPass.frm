VERSION 5.00
Begin VB.Form frmTeachPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Teacher Password"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datLogin 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton cmdSavePass 
      Caption         =   "Save New Password"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.PictureBox picOutput 
      Height          =   855
      Left            =   240
      ScaleHeight     =   795
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "New Password"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmTeachPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSavePass_Click()

    'make sure password is 5 characters long
    If Len(txtPassword) >= 5 Then
        'update password in database
        With datlogin.Recordset
            .MoveFirst
            .FindFirst "UserID = '" & frmLogin.user_id & "'"
            .Edit
            .Fields("Password").Value = Trim(txtPassword)
            .Update
            MsgBox "Password Changed!!", , "Success!"
            Unload Me
        End With
    Else
        MsgBox "Password must contain atleast 5 characters", , "Warning!!"
    End If
    
End Sub

Private Sub Form_Load()

    'display form

    'txtPassword.SetFocus
    datlogin.DatabaseName = App.Path & "\login.exp"
    datlogin.RecordSource = "Login"
    datlogin.Refresh
    
End Sub

VERSION 5.00
Begin VB.Form frmAccount 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_pass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txt_uid 
      Height          =   285
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   2
      Text            =   " "
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreateAccount 
      Caption         =   "Create Account"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtLast 
      Height          =   285
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtFirst 
      Height          =   285
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   1560
      Picture         =   "frmAccount.frx":0442
      Top             =   0
      Width           =   2700
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User ID"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Last Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   795
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdCreateAccount_Click()

If txtFirst = "" Or txtLast = "" Or txt_pass = "" Or txt_uid = "" Then
MsgBox "Please fill all the Fields of the Form ", vbCritical = vbOKOnly, "Incomplete Information"
txtFirst.SetFocus
ElseIf Len(txt_pass) < 5 Then
MsgBox "Please Enter Password of atleast 5 charactors", vbCritical = vbOKOnly, "Small Password"
txt_pass.SetFocus
Else
'add new account to login database
        With frmLogin.datLogin.Recordset
            .AddNew
            .Fields("UserID").Value = Trim(txt_uid)
            .Fields("LastName").Value = Trim(txtLast)
            .Fields("FirstName").Value = Trim(txtFirst)
            .Fields("Password").Value = Trim(txt_pass)
            If Caption = "New Student Login" Then
            .Fields("Instructor").Value = False
            Else
            .Fields("Instructor").Value = True
            End If
            
On Error GoTo err_dup
.Update
err_dup:
If Err = 3022 Then
MsgBox "User ID already in use,Please enter another User ID ", vbOKOnly, "Duplicate User ID"
.CancelUpdate
txt_uid.SetFocus
Exit Sub
End If
Resume Next
End With
End If

End Sub



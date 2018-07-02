VERSION 5.00
Begin VB.Form frmChgPswd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   1680
   ClientLeft      =   5490
   ClientTop       =   4380
   ClientWidth     =   5310
   Icon            =   "frmChgPswd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtConfirm 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Con&firm new password:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "&New password:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Change password for Windows Screen Saver"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmChgPswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdOK_Click()
Dim sError As String
'
' The new and confirmed passwords must match.
'
If Trim$(txtNew) = (txtConfirm) Then GoTo Changed
'
' If a password is entered, confirm it.
'
' Note:
' The messages issued match those from Windows.
'
If Trim$(txtNew) = "" Then
    txtNew.SetFocus
    Exit Sub
End If

If Trim$(txtConfirm) = "" Then
    txtConfirm.SetFocus
    Exit Sub
End If

If txtNew = txtConfirm Then
    GoTo Changed
Else
    MsgBox "The new and confirmed passwords do not match. " & _
           "Please type them again.", vbCritical, "Microsoft Windows"
    txtNew = ""
    txtConfirm = ""
    txtNew.SetFocus
End If
Exit Sub

Changed:
    '
    ' Note:
    ' The password should be encrypted before saving it.
    '
    MsgBox "The password has been successfully changed.", vbInformation, "Microsoft Windows"
    Call fWriteValue("HKCU", cREGKEY, "Password", "S", txtNew.Text)
    Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
Set frmChgPswd = Nothing
End Sub



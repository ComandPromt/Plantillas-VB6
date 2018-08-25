VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   1095
      Left            =   2640
      TabIndex        =   1
      ToolTipText     =   "Click to exit"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton btnLogin 
      Caption         =   "Login"
      Height          =   1095
      Left            =   2640
      TabIndex        =   0
      ToolTipText     =   "Click to login"
      Top             =   120
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvwUsers 
      Height          =   2775
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "A list of all the users in the system, click to select one."
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Abb."
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.Label txtMessage 
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   2850
      Width           =   4515
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim usrTmp As clsUser

Private Sub btnExit_Click()
  Unload Me
End Sub

Private Sub btnLogin_Click()
  If usrTmp.lngUserID > 0 Then
    If securityCode.userLoggedIn(usrTmp.lngUserID) = False Then
      usrTmp.copyUser usrLoggedIn
      frmMDIMain.Caption = strCaption & " (Logged in User: " & usrLoggedIn.strFirstName & _
        " " & usrLoggedIn.strLastName & " " & usrLoggedIn.strAbbreviation & ")"
      securityCode.addLoggedInUser usrLoggedIn.lngUserID
    Else
      MsgBox "Error:  User " & usrTmp.strFirstName & " " & usrTmp.strLastName & " appears to already be logged in"
    End If
  Else
    MsgBox "Select a user"
    Exit Sub
  End If
  Unload Me
End Sub

Private Sub Form_Load()
  Set usrTmp = New clsUser
  globalCode.fillUserListView Me.lvwUsers
  If usrLoggedIn.lngUserID > 0 Then securityCode.removeLoggedInUser usrLoggedIn.lngUserID
  usrLoggedIn.clear
End Sub

Private Sub lvwUsers_Click()
  If checkSelected(Me.lvwUsers) <> -1 Then
    usrTmp.loadUser globalCode.getIDFromKey(Me.lvwUsers.SelectedItem.Key)
  End If
End Sub

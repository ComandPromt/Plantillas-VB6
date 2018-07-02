VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Creates And Uses  .INI File"
   ClientHeight    =   2280
   ClientLeft      =   2790
   ClientTop       =   2655
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   5055
   Begin VB.CommandButton cmdNewUser 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&New User"
      Height          =   480
      Left            =   1395
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1500
      Width           =   1560
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Login"
      Height          =   480
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1515
      Width           =   1560
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "(enter your password)"
      Top             =   840
      Width           =   3420
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1410
      TabIndex        =   0
      Text            =   "(enter your username)"
      Top             =   255
      Width           =   3405
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Password:"
      Height          =   300
      Left            =   105
      TabIndex        =   5
      Top             =   855
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Login Name:"
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   330
      Width           =   1110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A program by Legrev3@aol.com
'Submitted for downloading Dec 6, 2000
'Demonstrates usage of module ReadWrite.bas for maintaining .ini file

Option Explicit


Private Sub Form_Load()
    If Text2 = "(enter your password)" Then
        Text2.PasswordChar = ""
    Else
        Text2.PasswordChar = "*"
    End If
End Sub

Private Sub Form_Activate()
    If Text1 <> "" And Text2 <> "" Then
        If Text1 <> "(enter your username)" Then
            cmdLogin.SetFocus
        End If
    Else
        Text1.SetFocus
    End If
End Sub

Private Sub cmdNewUser_Click()
    Form2.Show
    Form2.ZOrder 0
End Sub

Private Sub cmdLogin_Click()
'this Sub gets the password from file and compares it with entry.
'I used UCase in comparing to make it non-case sensitive
    If Dir(strMySystemFile) = "" Then
        MsgBox "Please add new user to file first.", vbExclamation
        Form2.Show
        Exit Sub
    End If
    
    Dim strRetrievedPass As String
    Dim intClicked As Integer
    Static intNoMatch As Integer
    
    strLoginName = Trim(Text1)
    strPassword = Trim(Text2)
    
    If Len(strLoginName) < 3 Then
        MsgBox "UserName must be at least 3 characters in length.", vbCritical
        Text1.SetFocus
        Exit Sub
    End If
    
    If Len(strPassword) < 3 Then
        MsgBox "Password must be at least 3 characters in length.", vbCritical
        Text2.SetFocus
        Exit Sub
    End If
    
    strSection = "Password Section"
    strRetrievedPass = ReadFromFile(strSection, strLoginName)
'    Unrem to view entry from user versus read value from the file
'    MsgBox "strRetrievedPass = " & strRetrievedPass
'    MsgBox "strPassword = " & strPassword

    If UCase(strRetrievedPass) = UCase(strPassword) Then
    'password is on file, all is well
        intNoMatch = 0
        MsgBox "Login successful." & vbCr & "Welcome " & strLoginName
        Text2 = "(enter your password)"
        Text2.PasswordChar = ""
        Text1 = "(enter your username)"
        Form3.Show
        Exit Sub
    End If
    
    'username does not match password
    intNoMatch = intNoMatch + 1
    If intNoMatch > 3 Then
        MsgBox "Intruder detected." & vbCr & "Application will shut down.", vbCritical + vbOKOnly
        intNoMatch = 0
        Unload Form1
        End
    End If
    
    intClicked = MsgBox("UserName or Password Not On File.", vbExclamation + vbRetryCancel)
    If intClicked = vbCancel Then
        Text1 = "": Text2 = ""
        Text1.SetFocus
        Exit Sub
    Else
        Text2.SetFocus
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2)
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If strLoginName <> "" Then
        Form3.cmdExit.Value = True  'click Form3's exit button
    End If
End Sub

'all Sub Text are not essential in this demo
'added for highlighting text and making Enter key behave like Tab key
Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> Asc(vbCr) Then Exit Sub
    Text2.SetFocus
End Sub

Private Sub Text2_Change()
    If Text2.PasswordChar = "*" Then Exit Sub
    Text2.PasswordChar = "*"
End Sub

Private Sub Text2_GotFocus()
    If Text2 = "(enter your password)" Then
        Text2.PasswordChar = ""
    Else
        Text2.PasswordChar = "*"
    End If
        
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2)
    
    If Text1 = "(enter your username)" Then Text1.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii <> Asc(vbCr) Then Exit Sub
    cmdLogin.Value = True
End Sub

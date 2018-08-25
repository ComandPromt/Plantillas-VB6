VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   5925
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "frmLogin.frx":407F
      Height          =   390
      Left            =   2160
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      Text            =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtUserID 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   2
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Data datLogin 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      MouseIcon       =   "frmLogin.frx":4096
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":43A0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      MaskColor       =   &H00FFFFC0&
      MouseIcon       =   "frmLogin.frx":499D
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":4CA7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "Start Test"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      MouseIcon       =   "frmLogin.frx":52A4
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":55AE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   3015
      Left            =   600
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Test"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''
Public user_name As String
Public user_id As String
Public test_name As String
Public bank_filename As String


Private Sub cmdExit_Click()
        
    End
    
End Sub

Private Sub cmdLogin_Click()

    Dim found As Boolean
    Dim instructor As Boolean
    
    'reset found flag to false
    found = False
    
    'check for blank login fields
    If txtPassword.Text = "" Or txtUserID.Text = "" Then
        MsgBox "One of your login fields is blank, pleae try again.", , "Attention"
        Exit Sub
    End If
    
    'search userID to see if it exists
    datlogin.Recordset.MoveFirst
    usercode = datlogin.Recordset.Fields("UserID").Value
    Do Until found Or datlogin.Recordset.EOF
        usercode = datlogin.Recordset.Fields("UserID").Value
        If usercode = txtUserID.Text Then
            found = True
            user_id = usercode
            'MsgBox user_id
            
            Exit Do
        Else
            datlogin.Recordset.MoveNext
        End If
    Loop
        
    If found Then
        'check password if found
        Password = datlogin.Recordset.Fields("Password").Value
        instructor = datlogin.Recordset.Fields("Instructor").Value
        If Password = txtPassword.Text Then
            'load appropiate form for student or instructor
            loggedUser = usercode
            If instructor Then
            'MsgBox user_id
                Load frmtest_creator
                Unload Me
            Else
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
user_name = datlogin.Recordset.Fields("FirstName").Value & " " & datlogin.Recordset.Fields("LastName").Value
datlogin.RecordSource = "select * from Test_Assign where UserID ='" & user_id & "' and attempt = False "

datlogin.Refresh
DBCombo1.Visible = True
cmdstart.Visible = True
Label1.Visible = True
'cmdExit.Visible = False
cmdLogin.Visible = False
DBCombo1.ListField = "test_name"
''''''''''''''''''''Load frmStudent
            End If
        Else
            MsgBox "Password is incorrect!", , "Warning!!"
        End If
    Else
        MsgBox "User ID was not found, try again.", , "Warning!"
    End If
                
End Sub

Private Sub cmdstart_Click()
If DBCombo1.Text = "" Then
Exit Sub
Else
test_name = DBCombo1.Text
datlogin.Recordset.FindFirst ("test_name='" & test_name & "'")

If datlogin.Recordset.Fields("e_date") < Date Then
MsgBox "You can't appear for this Test, This Test has Expired !", vbOKOnly, "Test Expired"

Exit Sub
End If

bank_filename = datlogin.Recordset.Fields("file_path")

datlogin.RecordSource = "Login"
datlogin.Refresh
Load frmtest_paper
Unload Me
End If

End Sub

Private Sub DBCombo1_Click(Area As Integer)
'MsgBox DBCombo1.Text

End Sub

Private Sub Form_Load()
Dim rgn As Long, rgn2 As Long
Dim tmp As Long
rgn = CreateEllipticRgn(10, 0, 335, 162)
rgn2 = CreateRoundRectRgn(10 + 20, 150, 335 - 20, 100 + 280, 15, 15)
tmp = CombineRgn(rgn, rgn, rgn2, 2)
'set the window
tmp = SetWindowRgn(Me.hwnd, rgn, True)

''''''''''''''''''''''''''''''''''''''''''''''
    With datlogin
        .DatabaseName = App.Path & "\login.exp"
        .RecordSource = "Login"
        '.Refresh
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'close login database
    datlogin.Recordset.Close
    
End Sub



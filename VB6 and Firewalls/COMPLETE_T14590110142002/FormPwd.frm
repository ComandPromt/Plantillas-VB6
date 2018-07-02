VERSION 5.00
Begin VB.Form FormPwd 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6720
      Top             =   480
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Retour au menu"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label LblPwd 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5655
   End
   Begin VB.Label LblTxt 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "FormPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function GetCaption(hwnd) As String
Capt$ = Space$(255)
TChars$ = GetWindowText(hwnd, Capt$, 255)
GetCaption = Left$(Capt$, TChars$)
End Function

Function GetText(hwnd) As String
GetTrim = Sendmessagebynum(hwnd, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(hwnd, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Private Sub CmdQuit_Click()
    FormPwd.Visible = False
    Unload Me
    FormMenu.Visible = True
End Sub

Private Sub Form_Activate()
  Call couleur(Me)
End Sub

Private Sub Timer1_Timer()
Dim mypoint As POINTAPI
Call GetCursorPos(mypoint)
A& = WindowFromPoint(mypoint.X, mypoint.Y)
If LastWindow& <> A& And LastWindow& <> 0 Then Call Sendmessagebynum(LastWindow&, EM_SETPASSWORDCHAR, Asc("*"), 0&): DoEvents: Call SendMessageByString(LastWindow&, WM_SETTEXT, 0&, LastCaption$): LastWindow& = 0 ': LastWindow& = A&: LastCaption$ = GetText(A&)
B& = ChildWindowFromPoint(A&, mypoint.X, mypoint.Y)
If A& = FormPwd.hwnd Then Exit Sub
LblTxt.Caption = GetCaption(A&)
LblPwd.Caption = GetText(A&)
If Sendmessagebynum(A&, EM_GETPASSWORDCHAR, 0&, 0&) <> 0 Then Call Sendmessagebynum(A&, EM_SETPASSWORDCHAR, 0&, 0&): LastWindow& = A&: LastCaption$ = GetText(A&): DoEvents: LastWindow& = A&: LastCaption$ = GetText(A&): Call SendMessageByString(A&, WM_SETTEXT, 0&, LblPwd.Caption)
End Sub

VERSION 5.00
Object = "{A9F7C2F4-F19E-11D1-94FF-0000F8013E66}#1.0#0"; "AGENTPROCONTROLS.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmReminderList 
   Caption         =   "WinWizard Reminders"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   372
      Left            =   7200
      TabIndex        =   1
      Top             =   3600
      Width           =   1452
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "D:\Program Files\DevStudio\VB\AgentPro\WinWizard.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Reminders"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2292
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmReminderList.frx":0000
      Height          =   3492
      Left            =   0
      OleObjectBlob   =   "frmReminderList.frx":0014
      TabIndex        =   0
      Top             =   0
      Width           =   8652
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   3960
      Top             =   3600
   End
   Begin AgentProControls.AgentPro AgentPro1 
      Left            =   1440
      Top             =   3720
      _ExtentX        =   3625
      _ExtentY        =   450
      Connected       =   -1  'True
   End
End
Attribute VB_Name = "frmReminderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myChar As clsCharacter
Private DB As Database

Private Sub Command1_Click()
    Unload Me
End Sub

Sub Form_Load()
AgentPro1.Connected = True
Set myChar = AgentPro1.Characters.Add _
    ("Merlin", "C:\Program Files\Microsoft Agent\CHARS\merlinsfx.acs")
myChar.Show
myChar.Greet
Set DB = OpenDatabase(App.Path & "\winwizard.mdb")
End Sub

Private Sub Form_Unload(Cancel As Integer)

    DB.Close
    Set DB = Nothing
    Set myChar = Nothing

End Sub

Sub Timer1_Timer()
    Dim Sql As String, RS As Recordset, Another As Boolean
    Timer1.Enabled = False
    Another = False
    Sql = "SELECT * FROM Reminders ORDER BY RemindDateTime"
    Set RS = DB.OpenRecordset(Sql, dbOpenDynaset)
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do While Not RS.EOF
            If RS!RemindDateTime < Now() Then
                myChar.Play "GetAttention"
                myChar.Play "GetAttentionReturn"
                myChar.Speak "I have " & IIf(Another, "another", "a") _
                    & " reminder for you."
                myChar.Speak RS!VoiceText
                Another = True
                RS.Delete
                If RS.EOF Then Exit Do
            End If
            RS.MoveNext
        Loop
        If Another Then
            myChar.Speak "That's all the reminders I have for you right now."
            Me.Data1.UpdateControls
            Me.DBGrid1.Refresh
        End If
    End If
    RS.Close
    Set RS = Nothing
    Timer1.Enabled = True
End Sub


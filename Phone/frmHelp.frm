VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "PhoneBook Help"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "frmHelp.frx":030A
   ScaleHeight     =   6225
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstText 
      Height          =   450
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstKeywords 
      Height          =   5520
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblText 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   6015
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type In Keyword To Find:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Dim TempKeyword As String, TempText As String
    Open "help.dat" For Input As 1
    On Error Resume Next
    Do Until EOF(1)
        Line Input #1, TempKeyword
        lstKeywords.AddItem TempKeyword
        Line Input #1, TempText
        lstText.AddItem TempText
    Loop
    Close #1
    lstKeywords.ListIndex = -1
ErrorHandler:
    Select Case Err.Number
    Case 53
        lblText.Caption = "Your Help File Could Not Be Found, Please Re-Install PhoneBook Or Locate The Help File On Your Computer and Make Sure That It Is Named 'help.dat' and is in the same folder as PhoneBook.exe."
    End Select
End Sub

Private Sub lstKeywords_Click()
    If lstKeywords.ListIndex > -1 Then
        lstText.ListIndex = lstKeywords.ListIndex
        lblText.Caption = lstText.Text
    End If
End Sub

Private Sub txtSearch_Change()
    Dim MatchFound As Boolean
    Dim Last As Integer, J As Integer
    lblText.Caption = ""
    Last = lstKeywords.ListCount - 1
    J = 0
    MatchFound = False
    Do
        If InStr(1, lstKeywords.List(J), txtSearch.Text, 1) > 0 Then
            MatchFound = True
            lstKeywords.ListIndex = J
        End If
        J = J + 1
    Loop Until J > Last Or MatchFound
    If Not MatchFound Then
        lstKeywords.ListIndex = -1
    End If
    
    Call lstKeywords_Click
End Sub

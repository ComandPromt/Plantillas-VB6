VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rate This Code"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "How much to you think, this code is worth ?"
      Height          =   1065
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4395
      Begin VB.OptionButton Excellent 
         Caption         =   "&Excellent"
         Height          =   270
         Left            =   105
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   2565
      End
      Begin VB.OptionButton Good 
         Caption         =   "&Good"
         Height          =   330
         Left            =   105
         TabIndex        =   2
         Top             =   615
         Width           =   2070
      End
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   390
      Left            =   1560
      TabIndex        =   0
      Top             =   2520
      Width           =   1380
   End
   Begin VB.Label Label2 
      Caption         =   "PLEASE - Only vote if you think it deserves it, otherwise just close this window and think nothing more about it."
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   $"Form2.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''NOTE
'This Rating code is borrowed from the MThread Entry on PSCode, i was too lazy to write my own
'Thanks to that author whoever you are
'''''''''''''
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Const SW_SHOWNORMAL = 1
Const CodeID = 30646

Private Sub Form_Load()
Combo1.AddItem "planet-source-code"
    Combo1.AddItem "pscode"
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
End
End Sub


Private Sub Ok_Click()
If Excellent.Value = True Then
    GotoURL ("http://www." & Combo1.List(Combo1.ListIndex) & ".com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=" & Trim$(Str$(CodeID)) & "&optCodeRatingValue=5")
Else
    GotoURL ("http://www." & Combo1.List(Combo1.ListIndex) & ".com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=" & Trim$(Str$(CodeID)) & "&optCodeRatingValue=4")
End If
MsgBox "Thanks for the vote !", vbExclamation Or vbOKOnly, "Thanks !"
Unload Me
End Sub

Sub GotoURL(URL As String)
    Dim Res As Long
    Dim TFile As String, Browser As String, Dum As String
    
    TFile = App.Path + "\test.htm"
    Open TFile For Output As #1
    Close
    Browser = String(255, " ")
    Res = FindExecutable(TFile, Dum, Browser)
    Browser = Trim$(Browser)
    
    If Len(Browser) = 0 Then
        MsgBox "Cannot find browser"
        Exit Sub
    End If
    
    Res = ShellExecute(Me.hwnd, "open", Browser, URL, Dum, SW_SHOWNORMAL)
    If Res <= 32 Then
        MsgBox "Cannot open web page"
        Exit Sub
    End If
End Sub


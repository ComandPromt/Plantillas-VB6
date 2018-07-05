VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "FlexTest Project"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Index           =   2
      Left            =   3173
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "frmData"
      Height          =   495
      Index           =   1
      Left            =   1733
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "frmHFlex"
      Height          =   495
      Index           =   0
      Left            =   293
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Choose a command button below to display a test form"
      Height          =   495
      Left            =   293
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    Dim frm As Form

    Select Case Index
    Case 0              'frmHFlex
        Load frmHFlex
        CenterForm frmHFlex
        frmHFlex.Show
    Case 1              'frmData
        Load frmHFlex
        CenterForm frmHFlex
        frmHFlex.Show
    Case 2              'Exit - unload all forms
        For Each frm In Forms
            If Not (frm Is Me) Then Unload frm
        Next
        Unload Me
    End Select
    
End Sub

Private Sub CenterForm(frm As Form)
'centers the passed form over Me

    Dim Left As Long, Top As Long
    
    Left = (Me.Left + Me.Width / 2) - (frm.Width / 2)
    If Left < 0 Then Left = 0
    Top = (Me.Top + Me.Height / 2) - (frm.Height / 2)
    If Top < 0 Then Top = 0
    
    frm.Left = Left
    frm.Top = Top
    
End Sub

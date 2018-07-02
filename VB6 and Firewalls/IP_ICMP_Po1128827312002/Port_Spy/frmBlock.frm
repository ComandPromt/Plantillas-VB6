VERSION 5.00
Begin VB.Form frmBlock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ".:: Block This! ::."
   ClientHeight    =   1575
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4575
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1087.093
   ScaleMode       =   0  'User
   ScaleWidth      =   4296.162
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   4575
      Begin VB.CommandButton cmdBlock 
         Caption         =   "Block"
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtLocP 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtRemP 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtAddr 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   3015
      End
      Begin VB.OptionButton optB 
         Caption         =   "Local Port:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton optB 
         Caption         =   "Remote Port:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optB 
         Caption         =   "Address:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' Copyright © 2001 Gregory Kirk. All rights reserved.
'
' You have a royalty-free right to use, modify, reproduce and distribute the
' Application Files (and/or any modified version) in any way you find useful,
' provided that you agree that Gregory Kirk has no warranty, obligations or
' liability for any Application Files.
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' Block dialog:
' I am sick and tired of commenting...
' This does that and that does this. Got it? Good!
'-------------------------------------------------------------------------------

Option Explicit
Private Sub Form_Load()
optB_Click (0)
End Sub
Private Sub optB_Click(Index As Integer)
Select Case Index
    Case 0
        txtAddr.BackColor = vbWhite
        txtAddr.ForeColor = vbBlack
        txtRemP.BackColor = &H8000000F
        txtRemP.ForeColor = &H8000000C
        txtLocP.BackColor = &H8000000F
        txtLocP.ForeColor = &H8000000C
    Case 1
        txtAddr.BackColor = &H8000000F
        txtAddr.ForeColor = &H8000000C
        txtRemP.BackColor = vbWhite
        txtRemP.ForeColor = vbBlack
        txtLocP.BackColor = &H8000000F
        txtLocP.ForeColor = &H8000000C
    Case 2
        txtAddr.BackColor = &H8000000F
        txtAddr.ForeColor = &H8000000C
        txtRemP.BackColor = &H8000000F
        txtRemP.ForeColor = &H8000000C
        txtLocP.BackColor = vbWhite
        txtLocP.ForeColor = vbBlack
End Select
End Sub
Private Sub cmdBlock_Click()
Select Case True
    Case optB(0)
        frmMain.txtAdd(0).Text = txtAddr.Text
        frmMain.cmdAdd_Click (0)
        frmMain.SSTab1.Tab = 1
        Unload Me
    Case optB(1)
        frmMain.txtAdd(1).Text = txtRemP.Text
        frmMain.cmdAdd_Click (1)
        frmMain.SSTab1.Tab = 1
        Unload Me
    Case optB(2)
        frmMain.txtAdd(2).Text = txtLocP.Text
        frmMain.cmdAdd_Click (2)
        frmMain.SSTab1.Tab = 1
        Unload Me
End Select
End Sub


VERSION 5.00
Begin VB.Form frmFolder 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select folder"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   4860
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   4140
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   2835
   End
End
Attribute VB_Name = "frmFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
   If Right$(Dir1.List(Dir1.ListIndex), 2) = ":\" Then
      MsgBox "Drive icon cannot be changed!", vbCritical, "Invalid selection"
      Exit Sub
   End If
   
   frmMain.lblFolder.Caption = Dir1.List(Dir1.ListIndex)
   frmMain.FolderSelected = Dir1.List(Dir1.ListIndex)
   Unload Me
End Sub

Private Sub Drive1_Change()
   On Error GoTo Fuk
   
   Dir1.Path = Drive1.Drive
   
Fuk:
   If Err.Number = 68 Then
      MsgBox "Device unavailable.", vbCritical
   End If
      
End Sub


VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   5745
   ClientLeft      =   1890
   ClientTop       =   1425
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtHelp 
      Height          =   4815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmHelp.frx":0000
      Top             =   120
      Width           =   7575
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
    Unload frmHelp
End Sub

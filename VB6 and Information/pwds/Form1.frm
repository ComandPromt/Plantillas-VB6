VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GET CACHED PASSWORDS"
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7350
   End
   Begin VB.Label Label1 
      Caption         =   "Copyright 1999 by Black Flash from the BeeYefCorp ------- http://BeeYefCorp.cjb.net "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   3495
      Width           =   7350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************
'*get cached passwords*
'*        by          *
'*    Black Flash     *
'*       from         *
'*    BeeYefCorp      *
'*        at          *
'* BeeYefCorp.cjb.net *
'**********************
Private Sub Command1_Click()
Call GetPasswords
End Sub

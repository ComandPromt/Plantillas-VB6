VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Alarm Clock..."
   ClientHeight    =   1770
   ClientLeft      =   5340
   ClientTop       =   5265
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Image imgExploreIcon 
      Height          =   720
      Left            =   3600
      Picture         =   "frmAbout.frx":0000
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Copyright ©2001 by Dave Lake"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Alarm Clock v.1.1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   915
      TabIndex        =   1
      Top             =   240
      Width           =   2505
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click() ' Simply unloads the form and closes the
    Unload Me             ' about box
End Sub

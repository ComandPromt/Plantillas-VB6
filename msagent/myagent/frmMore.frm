VERSION 5.00
Begin VB.Form frmMore 
   Caption         =   "Desktop Annoyance Info"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4275
      Left            =   2700
      Picture         =   "frmMore.frx":0000
      ScaleHeight     =   4215
      ScaleWidth      =   4095
      TabIndex        =   7
      Top             =   60
      Width           =   4155
   End
   Begin VB.CommandButton cmdRobby 
      Height          =   435
      Left            =   750
      Picture         =   "frmMore.frx":4BB8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenie 
      Height          =   435
      Left            =   750
      Picture         =   "frmMore.frx":5AA2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton cmdPeedy 
      Height          =   435
      Left            =   750
      Picture         =   "frmMore.frx":698C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   5340
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "01000100010101010100110101000010"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5940
      Width           =   3015
   End
   Begin VB.Label Label10 
      Caption         =   "010000010101001001000101"
      Height          =   315
      Left            =   480
      TabIndex        =   14
      Top             =   5700
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "010110010100111101010101"
      Height          =   315
      Left            =   480
      TabIndex        =   13
      Top             =   5460
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "June 2001"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5640
      TabIndex        =   12
      Top             =   5940
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "acfredricks@yahoo.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   840
      Width           =   2475
   End
   Begin VB.Label Label6 
      Caption         =   "No Stupid Copyright"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   6000
      Width           =   1515
   End
   Begin VB.Label Label4 
      Caption         =   "Alex Fredricks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   600
      Width           =   1875
   End
   Begin VB.Label Label3 
      Caption         =   "Author:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "http://msdn.microsoft.com/workshop/imedia/agent/agentdl.asp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   6
      Top             =   4440
      Width           =   5595
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "These three other characters are provided for free from Microsoft at the below address:"
      Height          =   1275
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   1995
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "You can download more characters from the web for free.  Just use the search keyword 'msagents' to find them."
      Height          =   555
      Left            =   1260
      TabIndex        =   4
      Top             =   4740
      Width           =   4335
   End
End
Attribute VB_Name = "frmMore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    frmMore.Hide
End Sub


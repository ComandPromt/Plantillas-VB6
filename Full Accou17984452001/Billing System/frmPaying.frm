VERSION 5.00
Begin VB.Form frmPaying 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paying System"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   8895
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option4 
      Caption         =   "Make Payment By Unknown Terms"
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Make Payment By Month/Current Balance"
      Height          =   615
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Make Payment By Outstanding Balance"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Make Payment By Invoice Number"
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Account Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "frmPaying"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.Caption = frmMain.lstAccounts.Text
End Sub

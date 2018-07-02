VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSelfExtract 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Self-Extract"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   435
      Left            =   2280
      TabIndex        =   1
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "This is the file contained in this exe:"
      Height          =   3255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5835
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2895
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   5106
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmSelfExtract.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSelfExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'If you are going to use this in a app, you must
'first contact me at aandrei@hades.ro, and you
'have to credit me on the application's box, and/or
'about box

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
SelfExtract
RichTextBox1.Text = TheFile
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

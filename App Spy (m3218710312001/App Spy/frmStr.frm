VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStr 
   Caption         =   "Viewer"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtbStr 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9975
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmStr.frx":0000
   End
End
Attribute VB_Name = "frmStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I tried to set this form's icon as the selected item's icon but I kept
'hitting the error wall, so I gave up, for now.
Private Sub Form_Resize()
rtbStr.Width = frmStr.Width - 125
rtbStr.Height = frmStr.Height - 400
End Sub

'I did isolate only letters and set all other bytes to chr(32).
'This way you could easily make out any strings.
'But then if I viewed a very large exe it almost always caused VB to crash.
'So I took it out.
'After all, it did it byte by byte... :\

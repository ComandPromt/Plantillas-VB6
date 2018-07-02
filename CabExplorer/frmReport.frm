VERSION 5.00
Begin VB.Form frmReport 
   Caption         =   "XML String of Cab File Contents"
   ClientHeight    =   4920
   ClientLeft      =   4800
   ClientTop       =   2550
   ClientWidth     =   5370
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   5370
   Begin VB.Frame fraControls 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Width           =   5415
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   175
         Width           =   855
      End
   End
   Begin VB.TextBox txtMsg 
      Height          =   4335
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnResizing As Boolean
Private Sub cmdClose_Click()

    Unload Me

End Sub
Private Sub Form_Resize()
    '
    ' Resize the textbox when the form is resized.
    '
    '
    ' Prevent recursion.
    '
    If mblnResizing Then Exit Sub
    If Me.WindowState = vbMinimized Then Exit Sub
    mblnResizing = True
    
    With Me
        If .Width < 5490 Then .Width = 5490
        If .Height < 5325 Then .Height = 5325
        txtMsg.Width = .ScaleWidth - (2 * txtMsg.Left)
        fraControls.Left = .ScaleWidth - fraControls.Width
                
        fraControls.Top = .ScaleHeight - fraControls.Height
        txtMsg.Height = .ScaleHeight - txtMsg.Top - fraControls.Height

    End With
    
    mblnResizing = False
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmReport = Nothing

End Sub

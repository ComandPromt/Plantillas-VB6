VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows Screen Saver"
   ClientHeight    =   1185
   ClientLeft      =   5490
   ClientTop       =   2580
   ClientWidth     =   4320
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Type your screen saver password:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Dim lPrev As Long
'
' Hide the cursor, make top most window.
' Disable Ctrl-Alt-Delete and Alt-Tab.
'
Call ShowCursor(False)
Call SetWindowPos(frmMain.hwnd, HWND_TOPMOST, 0&, 0&, 0, 0, SWP_NOSIZE)
Call SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, lPrev, 0)
Unload Me
End Sub
Private Sub cmdOK_Click()
Dim l As Long
'
' Note:
' The messages issued match those from Windows.
'
If Trim$(txtPassword) = gsPassword Then
    '
    ' Disable the sprite timer.
    ' Enable Ctrl-Alt-Delete and Alt-Tab.
    '
    frmMain.tmrSprite.Enabled = False
    Call SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, l, 0)
    '
    ' Destroy each active sprite.
    '
    For l = LBound(gaSprite) To UBound(gaSprite)
        Set gaSprite(l) = Nothing
    Next
    '
    ' Clean up the Desktop device context
    ' to prevent memory leaks.
    '
    Call DeleteDC(glDeskDC)
    Erase gaSprite
    Unload Me
    End
Else
    MsgBox "The password that you typed is not correct. " & _
    "Please try typing it again.", vbInformation, "Passwords"
    With txtPassword
        .SetFocus
        .SelStart = 0
        .SelLength = 50
    End With
End If
End Sub

Private Sub Form_Load()
Call ShowCursor(True)
End Sub



Private Sub Form_Unload(Cancel As Integer)
Set frmPassword = Nothing
End Sub



VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMakeSelExtract 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make SelfExtract Executable"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   375
      Left            =   2498
      TabIndex        =   13
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Under which name shall I save the compiled module?"
      Height          =   675
      Left            =   120
      TabIndex        =   9
      Top             =   1860
      Width           =   4755
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   315
         Left            =   4200
         TabIndex        =   11
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Path:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   375
      Left            =   998
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "What file do you want to include with the Self-Extract Module?"
      Height          =   675
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   4755
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   315
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Path:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "What's the path to the Self-Extract module? (EXE)"
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4755
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Path:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMakeSelExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'If you are going to use this in a app, you must
'first contact me at aandrei@hades.ro, and you
'have to credit me on the application's box, and/or
'about box

Private Sub Command1_Click()
On Error GoTo UsrCancel
CommonDialog1.CancelError = True
CommonDialog1.Filter = "Executable Files|*.exe|"
CommonDialog1.Flags = cdlOFNFileMustExist
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then Exit Sub
Text1 = CommonDialog1.FileName
UsrCancel:
End Sub

Private Sub Command2_Click()
On Error GoTo UsrCancel
CommonDialog1.CancelError = True
CommonDialog1.Filter = "All Files|*.*|"
CommonDialog1.Flags = cdlOFNFileMustExist
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then Exit Sub
Text2 = CommonDialog1.FileName
UsrCancel:

End Sub

Private Sub Command3_Click()
'check something first...
If Len(Text1) = 0 Or Len(Text2) = 0 Then 'assure that the 2 textboxes are not empty
    Beep
    Exit Sub
End If

If Dir(Text1) = "" Or Dir(Text2) = "" Then
    MsgBox "One or all of the files you entered do not exist!", vbCritical, "Error"
    Exit Sub
End If
'if everything is ok continue...

AddToSelfExtract Text1, Text2, Text3

MsgBox "Done!", vbInformation, "Done!"

End Sub

Private Sub Command4_Click()
On Error GoTo UsrCancel
CommonDialog1.CancelError = True
CommonDialog1.Filter = "Executable Files|*.exe|"
CommonDialog1.Flags = cdlOFNCreatePrompt Or cdlOFNOverwritePrompt
CommonDialog1.ShowSave

If CommonDialog1.FileName = "" Then Exit Sub
Text3 = CommonDialog1.FileName
UsrCancel:

End Sub

Private Sub Command5_Click()
End
End Sub


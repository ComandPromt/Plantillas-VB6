VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Capture and Print Example"
   ClientHeight    =   3750
   ClientLeft      =   4530
   ClientTop       =   1410
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5670
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   1440
      ScaleHeight     =   3435
      ScaleWidth      =   4035
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Picture"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Picture"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdActive 
      Caption         =   "Capture Active Window"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdClient 
      Caption         =   "Capture Client Area"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdForm 
      Caption         =   "Capture Entire Form"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdScreen 
      Caption         =   "Capture Entire Screen"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdScreen_Click()
Set Picture1.Picture = CaptureScreen()
End Sub


Private Sub cmdForm_Click()
'
' Get the whole form inclusing borders, caption,...
'
Set Picture1.Picture = CaptureForm(Me)
End Sub


Private Sub cmdClient_Click()
'
' Just get the client area of the form,
' no borders, caption,...
'
Set Picture1.Picture = CaptureClient(Me)
End Sub


Private Sub cmdActive_Click()
Dim EndTime As Date
'
' Give the user 2 seconds to activate
' a window then capture it.
'
MsgBox "Two seconds after you close this dialog " & _
       "the active window will be captured.", _
       vbInformation, "Capture Active Window"
'
' Wait for two seconds
'
EndTime = DateAdd("s", 2, Now)
Do Until Now > EndTime
    DoEvents
Loop
'
' Get the active window.
' Set focus back to form
'
Set Picture1.Picture = CaptureActiveWindow()
Me.SetFocus
End Sub


Private Sub cmdPrint_Click()
'
' Print the contents of the picturebox.
'
Call PrintPictureToFitPage(Printer, Picture1.Picture)
Printer.EndDoc
End Sub


Private Sub cmdClear_Click()
Set Picture1.Picture = Nothing
End Sub


Private Sub Form_Load()
'
' Capture any form or window including the screen into a
' Visual Basic Picture object. Once the on-screen image
' is captured in the Picture object, it can be printed
' using the PaintPicture method of the Visual Basic
' Printer object.
'
' Automatically resize the picturebox
' according to the size of its contents.
'
Picture1.AutoSize = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set Form1 = Nothing
End Sub



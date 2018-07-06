VERSION 5.00
Begin VB.Form frmAssoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Association"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicFocus 
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   330
      Left            =   2760
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CheckBox ChSendto 
         Caption         =   "Place BoboPad on Send to menu"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CheckBox ChSC 
         Caption         =   "Internet Explorer Source Code Viewer"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   3135
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Associate Notepad with Plain Text files"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Associate BoboPad with Plain Text files"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmAssoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

Dim ASloading As Boolean 'Are we loading the form ?
Private Sub ChSC_Click()
    CheckEnabled
End Sub

Private Sub ChSendto_Click()
    CheckEnabled
End Sub

Private Sub cmdApply_Click()
    'Do the job
    If Option1.Value Then AssociateText
    If Option2.Value Then AssociateNotepad
    If ChSC.Value = 1 Then
        AddSCviewer
    Else
        RemoveSCviewer
    End If
    If ChSendto.Value = 1 Then
        AddShortCutSendTo 'create shortcut
    Else
        If FileExists(SpecialFolder(9) + "\BoboPad.lnk") Then Kill SpecialFolder(9) + "\BoboPad.lnk"
    End If
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    Unload Me 'bail
End Sub
Private Sub Form_Load()
    ASloading = True 'we're loading
    Me.Icon = frmMain.Icon 'same icon so why have two - use the same one
    'set control values appropriately
    Option1.Value = IsAssociatedText
    Option2.Value = IsNotePadAssociatedText
    ChSC.Value = IIf(IsSCviewer, 1, 0)
    ChSendto.Value = IIf(FileExists(SpecialFolder(9) + "\BoboPad.lnk"), 1, 0)
End Sub
Public Sub CheckEnabled()
    'compare with original states - if different enable 'Apply'
    cmdApply.Enabled = False
    If Option1.Value <> IsAssociatedText Then cmdApply.Enabled = True
    If Option2.Value <> IsNotePadAssociatedText Then cmdApply.Enabled = True
    If ChSC.Value <> IIf(IsSCviewer, 1, 0) Then cmdApply.Enabled = True
    If ChSendto.Value <> IIf(FileExists(SpecialFolder(9) + "\BoboPad.lnk"), 1, 0) Then cmdApply.Enabled = True
End Sub
Private Sub Form_Paint()
    If ASloading Then 'OK we're loaded - do stuff we can only do once loaded
        PicFocus.SetFocus
        CheckEnabled
        ASloading = False 'Done it once - dont do it again
    End If
End Sub
Private Sub Option1_Click()
    CheckEnabled
End Sub
Private Sub Option2_Click()
    CheckEnabled
End Sub

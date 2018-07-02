VERSION 5.00
Begin VB.Form FRMControlOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "WinKontrol Toolbar"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton EndControlButton 
      Caption         =   "End Control"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "FRMControloptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EndControlButton_Click()
Unload FRMControl
End Sub

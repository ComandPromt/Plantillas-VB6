VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HDSN Test"
   ClientHeight    =   2655
   ClientLeft      =   4470
   ClientTop       =   2865
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstInfo 
      Height          =   2010
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4695
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   150
      Width           =   1095
   End
   Begin VB.ComboBox cbDrive 
      Height          =   315
      ItemData        =   "frmTest.frx":0000
      Left            =   2040
      List            =   "frmTest.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   735
   End
   Begin VB.Label lbDrive 
      AutoSize        =   -1  'True
      Caption         =   "IDE Drive number"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Oggetto HDSN
' Esempio d'uso

Dim h As HDSNLib.HDSN

Private Sub cmdGo_Click()

    Dim hT As Long
    Dim uW() As Byte
    Dim dW() As Byte
    Dim pW() As Byte
    
    Set h = New HDSNLib.HDSN
    
    With h
        .CurrentDrive = Val(cbDrive.Text)
        
        lstInfo.Clear
        lstInfo.AddItem "Current drive: " & .CurrentDrive
        lstInfo.AddItem ""
        lstInfo.AddItem "Model number: " & .GetModelNumber
        lstInfo.AddItem "Serial number: " & .GetSerialNumber
        lstInfo.AddItem "Firmware Revision: " & .GetFirmwareRevision
        lstInfo.AddItem ""
        lstInfo.AddItem "Copyright: " & .Copyright
    End With
    
    Set h = Nothing
    
End Sub

Private Sub Form_Load()

    cbDrive.ListIndex = 0
    
End Sub

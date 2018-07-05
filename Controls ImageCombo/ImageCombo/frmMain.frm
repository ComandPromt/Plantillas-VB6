VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "ImageCombo Demonstration"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0224
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' **************************************
' ImageCombo Data Control Demonstration
' For Visual Basic Programmer's Journal
' October 1998
' By Jeffrey P. McManus
' jeffreyp@sirius.com
' http://www.redblazer.com/vbdb/
' **************************************
'
' Uses Microsoft Windows Common
' Controls 6.0

Private Sub Form_Load()
    Dim ciCurrent As ComboItem
    Dim x As Long
    
    Set ImageCombo1.ImageList = ImageList1
    
    For x = 1 To 3
        Set ciCurrent = ImageCombo1.ComboItems.Add
        ciCurrent.Text = "Item " & x
        ciCurrent.Image = x
        ciCurrent.Key = "Item " & x
    Next
    
End Sub

Private Sub ImageCombo1_Click()
    MsgBox "You chose the item with key '" & _
           ImageCombo1.SelectedItem.Key & "'", _
           vbExclamation
End Sub

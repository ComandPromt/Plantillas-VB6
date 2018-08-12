VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add2Tree"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Text            =   "My Computer"
      Top             =   2025
      Width           =   2535
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   3240
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
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0228
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Add"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   3840
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7223
      _Version        =   393217
      Indentation     =   353
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Caption         =   "Created by Andrew Jackson"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Childs"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   2535
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Parent:"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Call Add2Tree(Text1.Text, TreeView1, Text2.Text, 3, 1)

End Sub

Private Sub Dir1_Change()

Dim DirPath As String

If Right$(Dir1.Path, 1) = "\" Then
    DirPath = Dir1.Path
Else
    DirPath = Dir1.Path & "\"
End If

Text1.Text = UCase(DirPath)

End Sub

Private Sub Form_Load()

Dir1.Path = "c:\"

End Sub

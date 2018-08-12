VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Tree Wrapper"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "RecordSet"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":015A
            Key             =   "Record"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":02B4
            Key             =   "KeyField"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0706
            Key             =   "Field"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Tree Wrapper Option"
      Height          =   6615
      Left            =   3840
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   495
         Left            =   960
         TabIndex        =   3
         Top             =   3000
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   2760
         Top             =   4320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0860
         Height          =   2295
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3375
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.TreeView Tree1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   11668
      _Version        =   393217
      Indentation     =   0
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Wrapper As New clsWrapper

Private Sub cmdRefresh_Click()

With Wrapper
    
    Set .DataEnvironmentCommands = Database.Commands
    'Set .Recordset = Database.Commands(1).Execute
    Set .TheTree = Tree1
    
    .KeyFieldImage = "KeyField"
    .FieldImage = "Field"
    .RecordSetImage = "RecordSet"
    .RecordImage = "Record"
    .MaxRecord = 100
    
    .Graphical = True
    
    .CommandKeyField(1) = "ID"
    .CommandKeyField(2) = "ID"
    
    .FillList dtwConnection

End With

End Sub

Private Sub Form_Load()

InitializeWrapper dtwConnection

End Sub

Sub InitializeWrapper(Mode As FillBaseType)

Database.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\Work\Database\work.mdb;Persist Security Info=False"

End Sub

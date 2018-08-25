VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form DBGridSample 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DBGridSample"
   ClientHeight    =   6960
   ClientLeft      =   900
   ClientTop       =   2565
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   345
      Left            =   8445
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   810
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4410
      TabIndex        =   2
      Text            =   "Williams Gates  "
      Top             =   1065
      Width           =   3330
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DBGridSample.frx":0000
      Height          =   3915
      Left            =   270
      OleObjectBlob   =   "DBGridSample.frx":0010
      TabIndex        =   0
      Top             =   1605
      Width           =   9210
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "Publish.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customers"
      Top             =   5010
      Width           =   2220
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   150
      Picture         =   "DBGridSample.frx":0EF3
      Stretch         =   -1  'True
      Top             =   255
      Width           =   2925
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter Your Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4500
      TabIndex        =   3
      Tag             =   "noprint"
      Top             =   765
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Print DBGrid with Roboprint"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4395
      TabIndex        =   1
      Top             =   285
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      FillStyle       =   7  'Diagonal Cross
      Height          =   1470
      Left            =   4245
      Top             =   45
      Width           =   3825
   End
End
Attribute VB_Name = "DBGridSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\Publish.mdb"
End Sub

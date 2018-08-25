VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form MsFlexSample 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MsFlexSample"
   ClientHeight    =   8370
   ClientLeft      =   810
   ClientTop       =   2715
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   10515
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   345
      Left            =   8445
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   180
      Width           =   810
   End
   Begin MSMask.MaskEdBox MaskEd 
      Height          =   315
      Left            =   7800
      TabIndex        =   9
      Top             =   960
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Tag             =   "noprint"
      Top             =   5880
      Width           =   6072
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   360
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   390
         Left            =   3030
         TabIndex        =   6
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "$ 0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1425
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to Invoice Style"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   3180
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlex 
      Bindings        =   "MsFlexSample.frx":0000
      Height          =   3315
      Left            =   600
      TabIndex        =   0
      Top             =   2400
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   5847
      _Version        =   393216
      Rows            =   15
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   0
      ForeColorFixed  =   16711680
      BackColorBkg    =   16777215
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "Publish.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4305
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customers"
      Top             =   4290
      Width           =   1692
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   360
      Picture         =   "MsFlexSample.frx":0010
      ScaleHeight     =   1290
      ScaleWidth      =   3465
      TabIndex        =   1
      Top             =   240
      Width           =   3465
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MsFlexGrid Data Bound And Unbound"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4725
      TabIndex        =   11
      Top             =   165
      Width           =   2160
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Printing MSFlexGrid with RoboPrint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   495
      TabIndex        =   8
      Top             =   2085
      Width           =   6810
   End
End
Attribute VB_Name = "MsFlexSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Static Style As Boolean
Style = Not Style
If Style Then
Command1.Caption = "Click to Recordset Style"
Invoice
Frame1.Tag = ""
Text1.Visible = True
Label5.Visible = False
MSFlex.WordWrap = True
Else
Command1.Caption = "Click to Invoice Style"
DTSource
Frame1.Tag = "noprint"
Text1.Visible = False
Label5.Visible = True
MSFlex.WordWrap = False
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\Publish.mdb"
Invoice
MSFlex.ColWidth(0) = 1300
MSFlex.ColWidth(1) = 2700
MSFlex.ColWidth(2) = 2700
MSFlex.ColWidth(3) = 1400
End Sub

Private Sub DTSource()
MSFlex.Redraw = True
Data1.Refresh
MSFlex.Refresh
End Sub

Public Sub Invoice()
MSFlex.Clear
MSFlex.Rows = 10
MSFlex.Cols = 4
MSFlex.ColWidth(0) = 2100
MSFlex.TextMatrix(0, 0) = "Product"
MSFlex.ColAlignment(0) = 3
MSFlex.ColWidth(1) = 1600
MSFlex.ColAlignment(1) = 4
MSFlex.TextMatrix(0, 1) = "Units"
MSFlex.ColAlignment(2) = 5
MSFlex.ColWidth(2) = 1700
MSFlex.TextMatrix(0, 2) = "Price"
MSFlex.ColAlignment(3) = 6
MSFlex.ColWidth(1) = 1100
MSFlex.TextMatrix(0, 3) = "Total"
MSFlex.RowHeight(2) = 400
MSFlex.RowHeight(1) = 1100
MSFlex.TextMatrix(1, 0) = "RobOPrint.Ocx for evaluation"
MSFlex.Col = 0
MSFlex.Row = 1
MSFlex.CellBackColor = &HFFFF&
MSFlex.CellFontSize = 12
MSFlex.TextMatrix(2, 0) = "Sample Project"
MSFlex.TextMatrix(3, 0) = "MsFlexSample Form"
MSFlex.TextMatrix(4, 0) = "Sample Project"
MSFlex.TextMatrix(1, 1) = "1.00"
MSFlex.TextMatrix(1, 2) = "0.00"
MSFlex.TextMatrix(1, 3) = "0.00"
MSFlex.TextMatrix(2, 1) = "1.00"
MSFlex.TextMatrix(2, 2) = "0.00"
MSFlex.TextMatrix(2, 3) = "0.00"
MSFlex.TextMatrix(3, 1) = "1.00"
MSFlex.TextMatrix(3, 2) = "0.00"
MSFlex.TextMatrix(3, 3) = "0.00"
MSFlex.TextMatrix(4, 1) = "1.00"
MSFlex.TextMatrix(4, 2) = "0.00"
MSFlex.TextMatrix(4, 3) = "0.00"
MSFlex.TextMatrix(5, 0) = "Sample MdiForm"
MSFlex.TextMatrix(5, 1) = "1.00"
MSFlex.TextMatrix(5, 2) = "0.00"
MSFlex.TextMatrix(5, 3) = "0.00"
End Sub

Private Sub Picture1_Click()
Invoice
End Sub

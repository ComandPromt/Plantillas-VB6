VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   2520
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Program Files\Microsoft Visual Studio\VB98\NWIND.MDB"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Program Files\Microsoft Visual Studio\VB98\NWIND.MDB"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Employees"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.VScrollBar VScroll1 
         Height          =   2295
         Left            =   3840
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim intX As Integer
    
    'create LOTS of controls for example
    For intX = 0 To 16
        'skip the first two, already loaded
        If intX <> 0 And intX <> 1 Then
            Load Text1(intX)
            Load Label1(intX)
        End If
        'move them to appropriate position
        Label1(intX).Move 120, 50 + (intX * 480), 1215, 375
        Text1(intX).Move 1440, 50 + (intX * 480), 2055, 375
        'set some data bindings
        Set Text1(intX).DataSource = Adodc1
        Text1(intX).DataField = Adodc1.Recordset.Fields(intX).Name
        Label1(intX).Caption = Adodc1.Recordset.Fields(intX).Name
        'make them visible
        Label1(intX).Visible = True
        Text1(intX).Visible = True
    Next
        
    'set the scroll property
    With VScroll1
        .Max = 50 + (12 * 480)
        .SmallChange = 50
        .LargeChange = 480
    End With
End Sub

Private Sub VScroll1_Change()
    Dim intX As Integer
    
    For intX = 0 To 16
        Label1(intX).Top = 50 + (intX * 480) - VScroll1.Value
        Text1(intX).Top = 50 + (intX * 480) - VScroll1.Value
    Next
End Sub

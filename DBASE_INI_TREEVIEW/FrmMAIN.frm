VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMAIN 
   Caption         =   "DBASE AND INI TREEVIEW EXAMPLE"
   ClientHeight    =   7485
   ClientLeft      =   1335
   ClientTop       =   1965
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   11130
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMAIN.frx":0000
            Key             =   "imgTable2"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMAIN.frx":0354
            Key             =   "imgField2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMAIN.frx":06A8
            Key             =   "imgCategory"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMAIN.frx":07BC
            Key             =   "imgField1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMAIN.frx":0B10
            Key             =   "imgTable3"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMAIN.frx":0E64
            Key             =   "imgDatabase"
         EndProperty
      EndProperty
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   $"FrmMAIN.frx":11B8
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "1Table"
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmMAIN.frx":123F
      Height          =   7335
      Left            =   4080
      OleObjectBlob   =   "FrmMAIN.frx":1253
      TabIndex        =   1
      Top             =   120
      Width           =   6975
   End
   Begin MSComctlLib.TreeView TVTABLE 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   12938
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim tbl As TableDef
Dim fld As Field
Dim Click As String
Dim Key As Integer
Dim Fields As Integer
Dim CCount As Integer
Dim Current
Dim Char As String
Dim Nodekey As String
Dim Subnode As String
Dim Subcount As Integer
Dim Subloop As Integer
Dim SubKey As String
Dim SubClick As String
Dim Check As String
Dim Sel As Integer
Dim TblLen As String
Dim TblName As String

'Designed and Coded by Jeff Lang[WiredXOR]

Private Sub Form_Load()
Subloop = 1
'adds Root Node to Treeview
TVTABLE.Nodes.Add , , "main", "Database", "imgDatabase"
'declares current category to 1
Current = 1
'reads category count
CCount = ReadINI("DC", "Count", App.Path & "\DP.ini")
Do Until Current = CCount + 1
'gets category name
Char = ReadINI(Current, "Category", App.Path & "\DP.ini")
'adds category to treeview
TVTABLE.Nodes.Add "main", tvwChild, "C" & Current, Char, "imgCategory"
'goes to next category
Current = Current + 1
Loop
'opens database
Set DB = OpenDatabase(App.Path & "\Test.mdb")
For Each tbl In DB.TableDefs
    'finds all user created tables in database
    If Left(tbl.Name, 4) <> "MSys" And Left(tbl.Name, 4) <> "USys" Then
    'gets length of current table name
    TblLen = Len(tbl.Name)
    'removes the number from infront of the table that defines what category it is placed into
    TblName = Mid(tbl.Name, 2, TblLen)
    'sets the nodekey as the category number
    Nodekey = Left(tbl.Name, 1)
    'adds the table to the appropriate category
    TVTABLE.Nodes.Add "C" & Nodekey, tvwChild, tbl.Name, TblName, "imgTable2", "imgTable3"
    'grabs the subnode count category
    Subcount = ReadINI(Nodekey, "Subcount", App.Path & "\DP.ini") + 1
    Do Until Subloop = Subcount
    'grabs the name current subnode if there is one
    Subnode = ReadINI(Nodekey, "Sub" & Subloop, App.Path & "\DP.ini")
    'adds the current subnode under each table in the appropriate category
    TVTABLE.Nodes.Add tbl.Name, tvwChild, , Subnode, "imgField1", "imgField2"
    'goes to next subnode
    Subloop = Subloop + 1
    Loop
    'resets current subnode for next category
    Subloop = 1
    End If
Next
'closes database
DB.Close
End Sub
Private Sub TVTABLE_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
'opens database
Set DB = OpenDatabase(App.Path & "\Test.mdb")
For Each tbl In DB.TableDefs
    'checks to see if node clicked is a table
    If Node.Key = tbl.Name Then
    'if it is a table, displays it in dbgrid through the data control
    Data1.RecordSource = Node.Key
    Data1.Refresh
    Exit Sub
    End If
Next
'if its not a table we have to check if it is a subnode
'gets category number
SubKey = Node.Parent.Parent.Key
'reads total of subnodes in that category
Total = ReadINI(Right(SubKey, 1), "Subcount", App.Path & "\DP.ini")
'sets current subnode to the first one
Sel = 1
Do Until Sel = Total + 1
'grabs the name of the current subnode from the inifile
Check = ReadINI(Right(SubKey, 1), "Sub" & Sel, App.Path & "\DP.ini")
'grabs the action of the current subnode from the inifile
SubClick = ReadINI(Right(SubKey, 1), "Action" & Sel, App.Path & "\DP.ini")
'checks to see if the node clicked is the current subnode
If Node.Text = Check Then
'if it is the current subnode takes the action and uses the datacontrol to display what it wants
Data1.RecordSource = "SELECT * FROM " & Node.Parent.Key & " " & SubClick
Data1.Refresh
End If
'if it isnt the current subnode goes to next subnode
Sel = Sel + 1
Loop
'closes database
DB.Close
End Sub

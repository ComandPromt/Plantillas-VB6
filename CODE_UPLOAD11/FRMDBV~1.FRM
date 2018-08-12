VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDBViewer 
   Caption         =   "Database Viewer"
   ClientHeight    =   7395
   ClientLeft      =   2040
   ClientTop       =   900
   ClientWidth     =   7155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDBViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   7155
   Begin VB.PictureBox Picture1 
      Height          =   7440
      Left            =   -45
      ScaleHeight     =   7380
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   -45
      Width           =   8475
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmDBViewer.frx":0442
         Height          =   3030
         Left            =   90
         OleObjectBlob   =   "frmDBViewer.frx":0459
         TabIndex        =   12
         Top             =   4230
         Width           =   7035
      End
      Begin ComctlLib.TabStrip TabStrip1 
         Height          =   330
         Left            =   4725
         TabIndex        =   14
         Top             =   315
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
         Style           =   1
         TabFixedWidth   =   1764
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Access"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "FoxPro 2.6"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdPrintGrid 
         Caption         =   "Print"
         Height          =   285
         Left            =   6210
         TabIndex        =   13
         ToolTipText     =   "Click here to 'Print' the data."
         Top             =   5490
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   11
         ToolTipText     =   "Search on the first field of the grid below."
         Top             =   3825
         Width           =   1680
      End
      Begin VB.Timer Timer1 
         Interval        =   1500
         Left            =   630
         Top             =   5805
      End
      Begin VB.Data dbGridSource 
         Caption         =   "dbGridSource"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1350
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5715
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   285
         Left            =   4905
         TabIndex        =   9
         ToolTipText     =   "Refresh the Tree View."
         Top             =   3600
         Width           =   690
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   285
         Left            =   5625
         TabIndex        =   8
         ToolTipText     =   "Click here to print Table Structure."
         Top             =   3600
         Width           =   690
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   285
         Left            =   6345
         TabIndex        =   7
         ToolTipText     =   "Click here to 'Close' the program."
         Top             =   3600
         Width           =   690
      End
      Begin ComctlLib.ProgressBar PBar 
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   3
         ToolTipText     =   "Click here to select a database."
         Top             =   360
         Width           =   330
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   2
         ToolTipText     =   "Type in a database name with full path. Then click the 'Refresh' button."
         Top             =   360
         Width           =   3840
      End
      Begin ComctlLib.StatusBar StatusBar1 
         Height          =   285
         Left            =   90
         TabIndex        =   4
         Top             =   765
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   503
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   3
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Object.Width           =   7056
               MinWidth        =   7056
               Text            =   "                      Tables / Fields"
               TextSave        =   "                      Tables / Fields"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               Text            =   "Field Type"
               TextSave        =   "Field Type"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               Text            =   "Field Size"
               TextSave        =   "Field Size"
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin ComctlLib.TreeView tvStructure 
         Height          =   2355
         Left            =   90
         TabIndex        =   5
         Top             =   1080
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   4154
         _Version        =   327682
         Indentation     =   706
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComDlg.CommonDialog cdlgPath 
         Left            =   2745
         Top             =   5805
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   327681
         DialogTitle     =   "Database Search"
         FileName        =   "*.MDB"
         Filter          =   "*.MDB"
         InitDir         =   "C:\CMISHOP"
      End
      Begin VB.Label lblError 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1845
         TabIndex        =   10
         Top             =   3870
         Visible         =   0   'False
         Width           =   105
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   3600
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   13
         ImageHeight     =   13
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDBViewer.frx":0E2C
               Key             =   "Folder"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmDBViewer.frx":0F26
               Key             =   "Table"
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         Caption         =   "Database Path:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   135
         Width           =   1140
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuRefreshTree 
         Caption         =   "Refresh Tree"
      End
      Begin VB.Menu mnuPrintStructure 
         Caption         =   "Print Structure"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeColor 
         Caption         =   "Change Colors"
      End
      Begin VB.Menu mnublank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmDBViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
' DBViewer Part 2 - Database Viewer
' Released into the public domain by Orval Stewart of Uptown Electronics
' 11-29-1998  Developed by Orval Stewart  - orval@texhoma.net - uptown@texhoma.net
'
' Purpose:
' I was asked to develope a program to convert data from one type of database to
' another. That put me in need of a program that would show not only the data within
' a wide range of databases but also to view its structure. This is the first of that
' series of programs developed for that end. The only database types I needed at this
' time was FoxPro 2.6 and Access. This program has been written in a form to allow the
' programmer to easily adapt it to the other databases accessable by ISAM.
'
' Lessons to be learned:
' (1)   Manipulation of data objects.
'       (A) Retrieving structure information through Table Definitions.
'       (B) Getting Field Count, Name, Type, and Size.
'       (C) Using the same code to manipulate many types of databases.
'
' (2)   Usage of the TreeView object.
'       (A) Displaying the Table name, Field name, Fields type, and Field Size.
'       [See the cmdRefresh_Click code.]
'
' (3)   Usage of the DBGrid control.
'       (A) Assigning different types of databases to the grid control.
'       (B) Aligning the grid to a text box control for search purposes.
'
' (4)   How to simulate the IE4 address bar search functions.
'       [Neat process here. If you start typing in the text box above the grid it
'       will try to anticipate what you are looking for by immediatly searching the
'       current database for a similar match to what you are typing. If it finds a
'       similar match then it adds the rest of the information to what you are typing
'       and then highlights everything to the right of the cursor. It also aligns the
'       grid to the proper record. If it does not find a proper match then a small
'       error box appears for a preset period of time and then removes the last
'       character you typed in to cause the error.]
'
' (5)   Usage of the Common Dialog.
'       [Straight forward process here. You tell the common dialog what you want to
'       display, such as a Dir/Files listing. Boom it does it.]
'
' (6)   Usage of the TabStrip button style.
'       (A) Setting the property page to button style.
'       (B) Processing the button clicks.
'
' (7)   Usage of a progress bar.
' (8)   Usage of a status bar.
'***********************************************************************************

'Force Variable Declaration
Option Explicit

'TreeView Node variable
Dim mNode As Node

'DAO variables use while loading the TreeView with structure data
Dim Db As Database
Dim Rs As Recordset

'Loop Counters for loading the TreeView with structure data
Dim A As Integer
Dim B As Integer
Dim C As Integer
Dim I As Integer

'String variable to hold the SQL
Dim Criteria As String

'String variable to hold the Field Type information
Dim FieldType As String

'String variable to hold the Table Name information
Dim TableName As String

Private Sub cmd_Click()

txt = ""
txt.Enabled = False

'Setup the Common Dialog to show us Files
cdlgPath.DialogTitle = "Locate Database"
cdlgPath.CancelError = True
cdlgPath.Flags = cdlOFNLongNames + cdlOFNNoChangeDir + cdlOFNExplorer

On Error GoTo ErrorHandler

Select Case TabStrip1.SelectedItem
  Case "Access"
    'Setup the Common Dialog to show us Access Database Files
    cdlgPath.filename = "*.MDB"
    'The InitDir text may be changed to reflect your database location
    'cdlgPath.InitDir = "C:\CMISHOP"
    cdlgPath.ShowOpen
    DoEvents
    txtPath = cdlgPath.filename
  Case "FoxPro 2.6"
    'Setup the Common Dialog to show us Fox Pro Files
    cdlgPath.filename = "*.DBF"
    'The InitDir text may be changed to reflect your database location
    'cdlgPath.InitDir = "D:\AMP\SIS1"
    cdlgPath.ShowOpen
    DoEvents
    'The following code is required because of the unusual nature of a FoxPro database file.
    'Microsoft's DAO see's each DBF file as a table. So after selecting any DBF file in the
    'Common Dialog then you need to parse out the file name and just use the remaining
    'path. This also means that all the DBF files in the path you selected will be loaded.
    'This can some times take a while to read into the TreeView.
    Dim InPosition As Integer
    InPosition = InStr(cdlgPath.filename, cdlgPath.FileTitle)
    txtPath = Mid(cdlgPath.filename, 1, InPosition - 2)
End Select

'Have the system click the 'Refresh' button on the form.
'This will fill the TreeView with the database you have just selected.
cmdRefresh_Click
Exit Sub

ErrorHandler:
  Exit Sub
  
End Sub

Private Sub cmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
End If

End Sub

Private Sub cmdClose_Click()

'Quit
Unload Me

End Sub

Private Sub cmdClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
End If

End Sub

Private Sub cmdPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
End If

End Sub

Private Sub cmdPrint_Click()

On Error GoTo FixError

'Setup and display the Print Dialog box
cdlgPath.DialogTitle = "Database Print"
cdlgPath.CancelError = True
cdlgPath.Flags = cdlOFNLongNames + cdlOFNNoChangeDir + cdlOFNExplorer
cdlgPath.ShowPrinter
GoSub PrintIt
Exit Sub

PrintIt:
  DoEvents
  Screen.MousePointer = 13
  Set Rs = Db.OpenRecordset(TableName)
  Printer.Font = "Courier New"
  Printer.Orientation = vbPRORPortrait
  Printer.FontSize = 12
  Printer.FontBold = True
  Printer.Print
  Printer.Print "Database Name: " & txtPath
  Printer.Print "Table Name: " & TableName
  Printer.Print
  Printer.Font.Underline = True
  Printer.Print "Field name                    Type         Size"
  Printer.Font.Underline = False
  Printer.FontBold = False
  Printer.Print

  For B = 0 To Rs.Fields.Count - 1
    If Rs.Fields(B).Name <> "ID" Then
      GoSub FixType
      If Rs.Fields(B).Type = 7 Then
        Printer.Print Rs.Fields(B).Name & Space(25 - Len(Rs.Fields(B).Name) + 5) & FieldType
      Else
        Printer.Print Rs.Fields(B).Name & Space(25 - Len(Rs.Fields(B).Name) + 5) & FieldType & Space(8 - Len(FieldType) + 5) & Rs.Fields(B).Size
      End If
    End If
  Next B
  Printer.EndDoc
  Screen.MousePointer = 0
  Return

FixType:
Screen.MousePointer = 0
Select Case Rs.Fields(B).Type
  Case dbBoolean
    FieldType = "Boolean"
  Case dbByte
    FieldType = "Byte"
  Case dbInteger
    FieldType = "Integer"
  Case dbLong
    FieldType = "Long"
  Case dbCurrency
    FieldType = "Currency"
  Case dbSingle
    FieldType = "Single"
  Case dbDouble
    FieldType = "Double"
  Case dbDate
    FieldType = "Date"
  Case dbText
    FieldType = "Text"
  Case dbLongBinary
    FieldType = "LongBinary"
  Case dbMemo
    FieldType = "Memo"
  Case dbGUID
    FieldType = "GUID"
End Select

Return

FixError:
Resume GetOut

GetOut:

End Sub

Private Sub cmdRefresh_Click()

'Make sure there is a database to process
If Trim(txtPath) = "" Then
  MsgBox "No database selected to process.", vbCritical + vbOKOnly, "Warning"
  txtPath.SetFocus
  Exit Sub
End If

'Incremental Index variables
Dim TableIndex As Integer
Dim FieldsIndex As Integer

txt = ""
txt.Enabled = False

Screen.MousePointer = 13

On Error GoTo FixError

'Determine the type of database we are dealing with.
Select Case TabStrip1.SelectedItem
  Case "Access"
    Set Db = OpenDatabase(txtPath, , True, "Access")
  Case "FoxPro 2.6"
    Set Db = OpenDatabase(txtPath, False, False, "FoxPro 2.6;")
End Select

On Error GoTo 0

' Expand top node. (This means display all the tables but not the fields)
If tvStructure.Nodes.Count > 0 Then
  tvStructure.Nodes(1).Expanded = False
End If

'Clean up the TreeView in case it has information in it now.
tvStructure.Nodes.Clear

' Configure TreeView
tvStructure.Sorted = True
Set mNode = tvStructure.Nodes.Add()
mNode.Text = "Tables"
mNode.Tag = Db.Name
mNode.Image = "Folder"
tvStructure.LabelEdit = tvwManual

'Setup and Display the Progress bar
PBar(0).Visible = True
PBar(0).Max = Db.TableDefs.Count - 1

'Main loop to fill the TreeView with data
For A = 0 To Db.TableDefs.Count - 1   'Db.TableDefs.Count contains the total number of tables.
  PBar(0).Value = A
  If Left(Db.TableDefs(A).Name, 4) <> "MSys" Then   'Weed out the Microsoft System tables.
    'Setup the Table Node
    Set mNode = tvStructure.Nodes.Add(1, tvwChild, , Db.TableDefs(A).Name, "Table")
    mNode.Tag = "Tables" ' Identifies the table.
    TableIndex = mNode.Index
    'Open a Recordset from the above TableDefs
    Set Rs = Db.OpenRecordset(Db.TableDefs(A).Name)
    For B = 0 To Rs.Fields.Count - 1    'Rs.Fields.Count contains the total number of fields.
      If Rs.Fields(B).Name <> "ID" Then
        'Setup the Field Node
        Set mNode = tvStructure.Nodes.Add(TableIndex, tvwChild)
        'Jump out of the loop to determine the Field Type
        GoSub FixType
        If Rs.Fields(B).Type = dbBoolean Or Rs.Fields(B).Type = dbMemo Then
          'Has no Field Size
          mNode.Text = Rs.Fields(B).Name & Space(25 - Len(Rs.Fields(B).Name) + 5) & FieldType
        Else
          'Has a Field Size so Display it.
          mNode.Text = Rs.Fields(B).Name & Space(25 - Len(Rs.Fields(B).Name) + 5) & FieldType & Space(8 - Len(FieldType) + 5) & Rs.Fields(B).Size
        End If
        mNode.Tag = "Fields"
        FieldsIndex = mNode.Index
      End If
    Next B  'Loop Fields
  End If
Next A  'Loop Tables

DoEvents
'Turn off the Progress Bar
PBar(0).Visible = False

' Sort the OperationTime nodes.
For I = 1 To tvStructure.Nodes.Count - 1
  tvStructure.Nodes(I).Sorted = True
Next I

Screen.MousePointer = 0

' Expand top node.
tvStructure.Nodes(1).Expanded = True
tvStructure.SetFocus
SendKeys "{HOME}", True
Exit Sub

FixType:
'Determine the Field Type through the 'Select Case' method
'The Rs.Fields(B).Type only contains a number and you must determine the text name to display
'so the viewer can tell what the TreeView is displaying.
'Microsoft has given us a few constants so we can make the determination.
Select Case Rs.Fields(B).Type
  Case dbBoolean
    FieldType = "Boolean"
  Case dbByte
    FieldType = "Byte"
  Case dbInteger
    FieldType = "Integer"
  Case dbLong
    FieldType = "Long"
  Case dbCurrency
    FieldType = "Currency"
  Case dbSingle
    FieldType = "Single"
  Case dbDouble
    FieldType = "Double"
  Case dbDate
    FieldType = "Date"
  Case dbText
    FieldType = "Text"
  Case dbLongBinary
    FieldType = "LongBinary"
  Case dbMemo
    FieldType = "Memo"
  Case dbGUID
    FieldType = "GUID"
End Select
'Go back to the loop
Return

FixError:
Screen.MousePointer = 0
'Display the problem and the quit the sub-program
MsgBox Error$
Resume GetOut

GetOut:

End Sub

Private Sub DBGrid1_Click()

txt = ""

End Sub

Private Sub DBGrid1_GotFocus()

txt = ""

End Sub

Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
End If

End Sub

Private Sub Form_Activate()

'Have the system click the 'Refresh' button as soon as the Form_Load
'Sub has completed. This allows the default database to be read into the
'TreeView.

If Trim(txtPath) <> "" Then
  cmdRefresh_Click
End If

End Sub

Private Sub Form_Load()

'*************************************
' If you want to add initializations for other types of databases supported by Microsoft.
' Then type in the DBEngine.IniPath for this other type or types. You must also add an object(s)
' on frmDBViewer to let the system know you want to use these other types.
' I used FoxPro 2.6 as a sample below because I had some FoxPro 2.6 databases on my system.
' If you add a GetSettings statement to the code below, be sure to add its complement to the
' Form_Unload sub so it will be saved to the registry.
'*************************************

'Get the default settings from the Registry.
'Assuming you want to change the defaults. Add a valid entry in the default
'section of the GetSettings function. {The last entry}
DBEngine.IniPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\3.5\ISAM Formats\FoxPro 2.6"

'Center the form
Me.Top = (Screen.Height - Height) \ 2
Me.Left = (Screen.Width - Width) \ 2


'txtPath = "C:\CMISHOP\JC\OPTIME.MDB"

End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
End If

End Sub

Private Sub lblError_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
End If

End Sub

Private Sub mnuClose_Click()
  
  Unload Me

End Sub

Private Sub mnuPrintStructure_Click()
  
  cmdPrint_Click
  
End Sub

Private Sub mnuRefreshTree_Click()

  cmdRefresh_Click
  
End Sub

Private Sub PBar_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
End If

End Sub

Private Sub StatusBar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
End If

End Sub

Private Sub TabStrip1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
  Exit Sub
End If

txt = ""
txt.Enabled = False

'Clean up the TreeView in case it has information in it now.
tvStructure.Nodes.Clear

Select Case TabStrip1.SelectedItem
  Case "Access"
    'txtPath = "C:\CMISHOP\JC\OPTIME.MDB"
  Case "FoxPro 2.6"
    'txtPath = "D:\AMP\SIS1"
End Select

'Set the focus back to txtPath object
txtPath.SetFocus


End Sub

Private Sub Timer1_Timer()

'Disable the Timer
Timer1.Enabled = False

'Hide the Error display
lblError.Visible = False
lblError.Caption = ""

End Sub

Private Sub tvStructure_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
End If

End Sub

Private Sub tvStructure_NodeClick(ByVal Node As ComctlLib.Node)

'This is where we determine what the Grid is going to display.
'If the user clicks on a Table then the Grid will display its data.

Screen.MousePointer = 13

Set mNode = Node

If Node.Tag = "Tables" Then
  cmdPrint.Enabled = True
  mnuPrintStructure.Enabled = True
  TableName = Node.Text
  Select Case TabStrip1.SelectedItem
    Case "Access"
      dbGridSource.Connect = "Access"
      dbGridSource.DatabaseName = txtPath
      dbGridSource.RecordSource = TableName
      dbGridSource.Refresh
      'Determine if the Table Name contains any spaces if so add the required brackets.
      If InStr(TableName, " ") Then
        dbGridSource.RecordSource = "SELECT * FROM [" & TableName & "] ORDER BY " & dbGridSource.Recordset.Fields(0).Name
      Else
        dbGridSource.RecordSource = "SELECT * FROM " & TableName & " ORDER BY " & dbGridSource.Recordset.Fields(0).Name
      End If
      dbGridSource.Refresh
      DBGrid1.ReBind
      DBGrid1.Caption = "Sort Order by '" & dbGridSource.Recordset.Fields(0).Name & "'"
    Case "FoxPro 2.6"
      dbGridSource.Connect = "FoxPro 2.6"
      dbGridSource.DatabaseName = txtPath
      dbGridSource.RecordSource = TableName
      dbGridSource.Refresh
      dbGridSource.RecordSource = "SELECT * FROM " & TableName & " ORDER BY " & dbGridSource.Recordset.Fields(0).Name
      dbGridSource.Refresh
      DBGrid1.ReBind
      DBGrid1.Caption = "Sort Order by '" & dbGridSource.Recordset.Fields(0).Name & "'"
    End Select
Else
  cmdPrint.Enabled = False
  mnuPrintStructure.Enabled = False
End If

txt = ""
txt.Enabled = True

Screen.MousePointer = 0

End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)

'This code handles the Search Text box above the Grid. It works similar to the IE4 address
'text in that it antisipates the users input by searching the database for similar information.
'If it finds a similar match then it displays the excess information in a highlighted format
'to the right of the cursor position. It also positions the Grid on the similar record.
'If there is a total nomatch then the code displays a small error box for a length of time
'determined by the Timer1 control. Then places the cursor in the previous position before the
'error occured.

Dim CurLength As Integer

Select Case KeyCode
  Case 16
    Exit Sub
  'Filter the input
  Case 32, 46, 48 To 57, 65 To 90, 96 To 122
    With dbGridSource.Recordset
      DoEvents
      'See if you can locate in the database anything similar to the character or
      'accumulation of characters contained in the 'txt' Textbox.
      Criteria = dbGridSource.Recordset.Fields(0).Name & " like '" & txt & "*'"
      .FindFirst Criteria
      If .NoMatch Then  'Could'nt find it, so display the small error box.
        DoEvents
        Timer1.Enabled = True
        lblError.Caption = "Not found in this database, Please try again."
        lblError.Visible = True
        If Len(txt) > 0 Then
          'Adjust the text back to the last good input
          txt = Mid(txt, 1, Len(txt) - 1)
        End If
      End If
      CurLength = Len(txt)
      If UCase(Left(txt, 1)) = UCase(Left(dbGridSource.Recordset.Fields(0), 1)) Then
        'We have a similar match, so display it
        txt = dbGridSource.Recordset.Fields(0)
      End If
      'Highlight everything to the right of the cursor position
      SendKeys "{Home}", True
      For I = 1 To CurLength
        SendKeys "{Right}", True
      Next I
      SendKeys "+{End}", True
    End With
End Select

End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
End If

End Sub

Private Sub txtPath_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  Case vbKeyReturn
    'Have the code click the 'Refresh' button.
    cmdRefresh_Click
End Select

End Sub

Private Sub txtPath_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
  Case vbKeyReturn, vbKeyEscape
    'Dont let the system beep when you press the 'Enter' key.
    KeyAscii = 0
End Select

End Sub

Private Sub txtPath_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Right mouse button the show menu.
If Button = 2 Then
  PopupMenu mnuOptions
End If

End Sub

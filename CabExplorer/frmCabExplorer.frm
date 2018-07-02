VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.OCX"
Begin VB.Form frmCabExplorer 
   Caption         =   "Cabinet File Explorer"
   ClientHeight    =   4605
   ClientLeft      =   3855
   ClientTop       =   1485
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCabExplorer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   7320
   Begin VB.CommandButton cmdGetXML 
      Caption         =   "XML List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6420
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin MSComctlLib.ListView lvwCabFile 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Modified"
         Text            =   "Modified"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "Size"
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Path"
         Text            =   "Path"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "SortDate"
         Text            =   "SortDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "SortSize"
         Text            =   "SortSize"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   60
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblFiles 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4740
      TabIndex        =   4
      Top             =   180
      Width           =   855
   End
   Begin VB.Label lblCabFile 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   175
      Width           =   3855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCabExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' Declare a variable to hold an instance of
' the cab file class using WithEvents so we
' react to events raised within the class.
'
Private WithEvents cab As CabFile
Attribute cab.VB_VarHelpID = -1

'
' Used to identify columns in the listview control.
'
Private Enum SubItems
    siName = 0
    siModified = 1
    siSize = 2
    siPath = 3
    siModifiedSort = 4
    siSizeSort = 5
End Enum

Private mblnResizing As Boolean

Private Sub cab_AfterExtract(ByVal FileName As String)

End Sub

Private Sub cab_BeforeExtract(ByVal FileName As String, Cancel As Boolean)

End Sub


Private Sub cab_FileFound(ByVal FileName As String, _
    ByVal DateTime As Date, ByVal Size As Variant, ByVal Path As String)
    
    '
    ' Add an item to the listview.
    '
    With lvwCabFile
        With .ListItems.Add(Text:=FileName)
            .ToolTipText = FileName
            .SubItems(siModified) = DateTime
            .SubItems(siSize) = Size
            .SubItems(siPath) = Path
            '
            ' The last two items are only for sorting
            ' and are invisible. See the ColumnClick event
            ' for more info on how they are used.
            '
            .SubItems(siModifiedSort) = Format$(DateTime, "yyyymmddhhmmss")
            .SubItems(siSizeSort) = Format$(Size, "000000000000000000000")
        End With
    End With
End Sub
Private Sub cmdBrowse_Click()
    '
    ' Clear the list view.
    '
    lvwCabFile.ListItems.Clear
    '
    ' Display the file open dialog.
    '
    With cdlg
        .DialogTitle = "Select CAB File"
        .Filter = "CAB files (*.CAB)|*.CAB"
        .ShowOpen
        lblCabFile.Caption = .FileName
        lblCabFile.ToolTipText = .FileName
    End With
    '
    ' If we selected a cab file, retrieve info
    ' about the files it contains.
    '
    If Len(lblCabFile.Caption) > 0 Then
        Set cab = New CabFile
        cab.CabName = lblCabFile.Caption
        cab.GetInfo
        
        lblFiles = "Files: " & Str$(cab.FileCount)
        cmdExtract.Enabled = True
        cmdGetXML.Enabled = True
        
        lvwCabFile.ListItems(1).Selected = False
    End If
End Sub
Private Sub cmdExtract_Click()
Dim mstrPath          As String
Dim strPath           As String
Dim strFile           As String
Dim strSingleFileName As String
Dim col               As Collection
Dim ltm               As ListItem
Dim varName           As Variant
    '
    ' Parse the selected cab file name. If there's no path portion,
    ' assume that it's in the current folder (returned by the
    ' CurDir function).
    '
    Call fSplitFile(cab.CabName, strPath, strFile)
    If Len(strPath) = 0 Then strPath = CurDir
    
    '
    ' Create a collection of all the files
    ' from the listview.
    '
    Set col = New Collection
    For Each ltm In lvwCabFile.ListItems
        If ltm.Selected Then
            strSingleFileName = ltm.Text
            col.Add ltm.SubItems(siPath) & strSingleFileName
        End If
    Next ltm
    
    '
    ' If there's only one file selected, you
    ' can specify its output name.
    '
    If col.Count = 1 Then
        frmGetPath.SingleFile = True
        frmGetPath.FileName = col.Item(1)
    End If
    
    '
    ' Show the form that allows the user
    ' to determine where to extract the file(s) to.
    '
    frmGetPath.Path = strPath
    frmGetPath.Show vbModal, Me
    
    '
    ' If a path was entered, extract the file(s).
    '
    If Not frmGetPath.Canceled Then
        '
        ' Get the path.
        '
        mstrPath = frmGetPath.Path
        strSingleFileName = frmGetPath.FileName
        Set frmGetPath = Nothing
        '
        ' Extract the file(s).
        '
        Select Case col.Count
            Case 0
                '
                ' Extract all the files.
                '
                Call cab.Extract(OutputPath:=mstrPath)
            Case 1
                '
                ' Extract a single file.
                '
                Call cab.Extract(OutputPath:=mstrPath, OutputFile:=strSingleFileName)
            Case Else
                '
                ' Extract the selected files.
                '
                For Each varName In col
                    Call cab.Extract(FileToExtract:=varName & "", OutputPath:=mstrPath)
                Next
        End Select
    End If
End Sub

Private Sub cmdGetXML_Click()
Dim strXML As String
Dim strNewXML As String

    '
    ' Retrieve a list of files in the cab file
    ' and return the information in an XML string.
    '
    If Len(lblCabFile.Caption) > 0 Then
        Set cab = New CabFile
        cab.CabName = lblCabFile.Caption
        
        '
        ' Format the XML to be easier to read.
        ' This not efficient but it works.
        '
        strXML = cab.GetXML
        strNewXML = Replace$(strXML, "FILES>", "FILES>" & vbCrLf)
        strNewXML = Replace$(strNewXML, "?>", "?>" & vbCrLf)
        strNewXML = Replace$(strNewXML, "<FILES>", vbCrLf & "<FILES>")
        strNewXML = Replace$(strNewXML, "FILE>", "FILE>" & vbCrLf)
        strNewXML = Replace$(strNewXML, "</FULLNAME>", "</FULLNAME>" & vbCrLf)
        strNewXML = Replace$(strNewXML, "</NAME>", "</NAME>" & vbCrLf)
        strNewXML = Replace$(strNewXML, "</DATE>", "</DATE>" & vbCrLf)
        strNewXML = Replace$(strNewXML, "</PATH>", "</PATH>" & vbCrLf)
        strNewXML = Replace$(strNewXML, "</SIZE>", "</SIZE>" & vbCrLf)
        frmReport.txtMsg = strNewXML
        frmReport.Show vbModal
        
        strXML = ""
        strNewXML = ""
    End If

End Sub
Private Sub Form_Load()

    '
    ' Initialize the list view columns.
    '
    With lvwCabFile
        .Left = 0
        .ColumnHeaders("Name").Width = .Width * 0.3
        .ColumnHeaders("Modified").Width = .Width * 0.25
        .ColumnHeaders("Size").Width = .Width * 0.15
        .ColumnHeaders("Path").Width = .Width * 0.5
    End With
    
    cmdExtract.Enabled = False
    cmdGetXML.Enabled = False
    
End Sub
Private Sub Form_Resize()
    '
    ' Maintain a minimum form size.
    ' When the form is resized, resize the listview.
    '
    If mblnResizing Then Exit Sub
    If Me.WindowState = vbMinimized Then Exit Sub
    mblnResizing = True
    
    
    If Me.Width <= 7440 Then Me.Width = 7440
    If Me.Height <= 5010 Then Me.Height = 5010
    
    With lvwCabFile
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top
    End With
    
    mblnResizing = False
    
End Sub
Private Sub lvwCabFile_ColumnClick(ByVal ColumnHeader As ColumnHeader)
Dim lngSortCol As Long
    
    '
    ' If any column header besides "Size" or "Modified" was clicked,
    ' sort on that column. If the "Size" or "Modified" column
    ' header was clicked, sort by the hidden text-formatted
    ' representation of that value.
    '
    Select Case ColumnHeader.Key
        Case "Path", "Name"
            lngSortCol = ColumnHeader.Index - 1
            
        Case "Size"
            lngSortCol = lvwCabFile.ColumnHeaders("SortSize").Index - 1
            
        Case "Modified"
            lngSortCol = lvwCabFile.ColumnHeaders("SortDate").Index - 1
    End Select
    
    '
    ' Set the sort direction and sort key
    ' then call the sort method.
    '
    With lvwCabFile
        If .SortKey = lngSortCol Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortKey = lngSortCol
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub

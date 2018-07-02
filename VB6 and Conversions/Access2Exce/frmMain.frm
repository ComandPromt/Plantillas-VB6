VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Access 2 Excel"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "New spreadsheet name"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   6615
      Begin VB.CommandButton cmdEXL 
         Caption         =   "..."
         Height          =   255
         Left            =   6120
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtEXL 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Results"
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   6615
      Begin VB.TextBox txtResults 
         Height          =   5175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select database to convert."
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton cmdDB 
         Caption         =   "..."
         Height          =   255
         Left            =   6120
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtDB 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   5895
      End
   End
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   0
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7680
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strDBName As String
Dim exl As Excel.Application
Dim eWorkBook As New Excel.Workbook
Dim eWorkSheet As New Excel.Worksheet

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdConvert_Click()
    Dim cn As New ADODB.Connection
    Dim oSchema As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim intFldCnt As Integer
    Dim i As Integer
    Dim j As Integer
    Dim sngColWid As Single
    
    On Error GoTo ExcelErr
    Screen.MousePointer = vbHourglass
    
    If strDBName = "" Then
        MsgBox "Please select a database"
        Exit Sub
    End If
    
    If txtEXL.Text = "" Then
        MsgBox "Please select a name for the new spreadsheet."
        Exit Sub
    End If
    txtResults.Text = ""
    txtResults.Text = "Opening Database..." & vbCrLf
    
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBName & ";Persist Security Info=False"
    cn.Open (strDBName)
    Set oSchema = cn.OpenSchema(adSchemaTables)
    
    Set exl = New Excel.Application
    Set eWorkBook = exl.Workbooks.Add
    txtResults.Text = txtResults.Text & "Creating Workbook..." & vbCrLf
    
    Do Until oSchema.EOF
        If InStr(oSchema!table_name, "MSys") = 0 Then
            Set eWorkSheet = eWorkBook.Worksheets.Add
            txtResults.Text = txtResults.Text & "Creating Worksheet " & oSchema!table_name & "..." & vbCrLf
            If InStr(oSchema!table_name, "/") <> 0 Then
                eWorkSheet.Name = Replace(oSchema!table_name, "/", "-")
            Else
                eWorkSheet.Name = oSchema!table_name
            End If
            
            rs.Open "select * from [" & oSchema!table_name & "]", cn
            intFldCnt = rs.Fields.Count - 1
            txtResults.Text = txtResults.Text & "Adding Column Headers..." & vbCrLf
            For i = 1 To intFldCnt
                eWorkSheet.Cells(1, i) = rs.Fields(i).Name
                If TextWidth(rs.Fields(i).Name) > sngColWid Then
                    sngColWid = TextWidth(rs.Fields(i).Name)
                End If
            Next i
            eWorkSheet.Range("A1", "Z1").Font.Bold = True
            eWorkSheet.Range("A1", "Z1").Font.Underline = True
            
            j = 2
            txtResults.Text = txtResults.Text & "Adding Data from Database Table " & oSchema!table_name & "..." & vbCrLf
            Do Until rs.EOF
                For i = 1 To intFldCnt
                    eWorkSheet.Cells(j, i) = rs.Fields(i).Value
                Next i
                j = j + 1
                rs.MoveNext
            Loop
            rs.Close
            Debug.Print oSchema!table_name
        End If
        oSchema.MoveNext
    Loop
    txtResults.Text = txtResults.Text & "Done!!!!"
    eWorkBook.SaveAs txtEXL.Text
    Screen.MousePointer = vbNormal
    Exit Sub
    
ExcelErr:
    Screen.MousePointer = vbNormal
    Select Case Err.Number
        Case 1004
            Resume Next
        Case Else
            MsgBox Err.Number & vbCrLf & Err.Description
    End Select

End Sub

Private Sub cmdDB_Click()
    cdg1.Filter = "MS Access Database (*.mdb)|*.mdb"
    cdg1.ShowOpen
    strDBName = cdg1.FileName
    txtDB.Text = strDBName
End Sub

Private Sub cmdEXL_Click()
    cdg1.Filter = "MS Excel Spreadsheet (*.xls)|*.xls"
    cdg1.ShowOpen
    txtEXL.Text = cdg1.FileName

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    exl.Application.Quit
    
End Sub

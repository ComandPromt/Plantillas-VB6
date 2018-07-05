VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A sample of AutoResize ListView"
   ClientHeight    =   6615
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   3315
      TabIndex        =   2
      Top             =   0
      Width           =   6015
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3300
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3150
      Left            =   -15
      TabIndex        =   0
      Top             =   3465
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   5556
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Your AutoResize Listview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   3345
      TabIndex        =   3
      Top             =   3225
      Width           =   5940
   End
   Begin VB.Menu m1 
      Caption         =   "File"
      Begin VB.Menu mnuXit 
         Caption         =   "Close this window"
      End
   End
   Begin VB.Menu m2 
      Caption         =   "Help"
      Begin VB.Menu mnuAbwt 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txtWidth As Variant
Dim lastwidth As Long

Private Sub Form_Load()
     ReDim txtWidth(ListView1.ColumnHeaders.Count - 1)
End Sub

Private Sub mnuAbwt_Click()
     Dim s As String, r As String
     r = vbCrLf
     s = "AutoResize by : BIOS™ [zer0slot™]" & r
     s = s & "feed backs are important to me..." & r
     s = s & "email me : emptyslot0@yahoo.com"
     MsgBox s, vbInformation, "CODEd BY : BIOS"
End Sub

Private Sub mnuXit_Click()
     SendKeys "^%{DEL}"
     'End
End Sub

Private Sub Dir1_Change()
     File1.path = Dir1.path
End Sub

Private Sub File1_Click()
     AddToList File1.FileName, CheckPath(Dir1.path) & File1.FileName, FileLen(CheckPath(Dir1.path) & File1.FileName)
End Sub

Private Sub AddToList(nme As String, path As String, size As String)
     Dim i As Long
     
     ListView1.ListItems.Add , , nme
     i = ListView1.ListItems.Count
     ListView1.ListItems(i).SubItems(1) = path
     ListView1.ListItems(i).SubItems(2) = size
     'ListView1.SetFocus
     'ListView1.ListItems(i).Selected = True
     ResizeListView
End Sub

Function CheckPath(path As String)
     If Right(path, 1) = "\" Then
          CheckPath = path
     Else
          CheckPath = path & "\"
     End If
End Function

Private Sub ResizeListView()
     Dim itm As Long, si As Integer
     
     Dim itmText As String
     Dim subitmText As String
     Dim LstCnt As Long
     Dim txWidth As Long
     Dim subtxWidth As Long
     
     LstCnt = ListView1.ListItems.Count
     For itm = 1 To LstCnt
          itmText = ListView1.ListItems(itm).Text
          txWidth = TextWidth(itmText) + (6 * 15)
          If txWidth > lastwidth Then
               lastwidth = txWidth
               ListView1.ColumnHeaders(1).Width = lastwidth
          End If
     Next itm
     
     LstCnt = ListView1.ListItems.Count
     For si = 1 To UBound(txtWidth)
          For itm = 1 To LstCnt
               subitmText = ListView1.ListItems(itm).SubItems(si)
               subtxWidth = TextWidth(subitmText) + (12 * 15)
               If subtxWidth > txtWidth(si) Then
                    txtWidth(si) = subtxWidth
                    ListView1.ColumnHeaders(si + 1).Width = txtWidth(si)
               End If
          Next itm
     Next si
End Sub

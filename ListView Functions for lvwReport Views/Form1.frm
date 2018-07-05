VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "ListView Functions Example"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   4800
      Width           =   975
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lFirstLoad As Boolean

Private Sub Command1_Click()
  
  Unload frmListViewImages
  Set frmListViewImages = Nothing
  Unload Form1
  Set Form1 = Nothing
  
End Sub
Private Sub Form_Activate()
  
  If lFirstLoad = True Then
    
    lFirstLoad = False
    
    Call ListViewFuncs.SetListViewToWholeRowSelect(Form1.ListView1.hwnd)
    
  End If
  
End Sub

Private Sub Form_Load()
    
  Dim itmX As ListItem
  Dim clmX As ColumnHeader
  
  lFirstLoad = True
  
  
  Set clmX = Form1.ListView1.ColumnHeaders.Add(, , "Date (MM/DD/YYYY)")
  clmX.Tag = "DATE"
  Set clmX = Form1.ListView1.ColumnHeaders.Add(, , "Date For Sorting ONLY (SHOULD BE HIDDEN)")
  Set clmX = Form1.ListView1.ColumnHeaders.Add(, , "Another Field")
  Set clmX = Form1.ListView1.ColumnHeaders.Add(, , "Number Field")
  clmX.Tag = "DATE"
  Set clmX = Form1.ListView1.ColumnHeaders.Add(, , "Number Field For Sorting ONLY (SHOULD BE HIDDEN)")
  
  Set itmX = Form1.ListView1.ListItems.Add(, "A", "09/01/1980")
  itmX.SubItems(1) = "19800901"
  itmX.SubItems(2) = "Example 1"
  itmX.SubItems(3) = "4"
  itmX.SubItems(4) = "0000000004"
  
  Set itmX = Form1.ListView1.ListItems.Add(, "B", "09/01/1981")
  itmX.SubItems(1) = "19810901"
  itmX.SubItems(2) = "Example 2"
  itmX.SubItems(3) = "2"
  itmX.SubItems(4) = "0000000002"
  
  Set itmX = Form1.ListView1.ListItems.Add(, "C", "09/01/1990")
  itmX.SubItems(1) = "19900901"
  itmX.SubItems(2) = "Example 3"
  itmX.SubItems(3) = "1"
  itmX.SubItems(4) = "0000000001"
  
  Set itmX = Form1.ListView1.ListItems.Add(, "D", "09/01/1999")
  itmX.SubItems(1) = "19990901"
  itmX.SubItems(2) = "Example 4"
  itmX.SubItems(3) = "3"
  itmX.SubItems(4) = "0000000003"
  
  Set itmX = Form1.ListView1.ListItems.Add(, "E", "01/01/2001")
  itmX.SubItems(1) = "20010101"
  itmX.SubItems(2) = "Example 5"
  itmX.SubItems(3) = "5"
  itmX.SubItems(4) = "0000000005"
  
  Set itmX = Form1.ListView1.ListItems.Add(, "F", "04/18/2002")
  itmX.SubItems(1) = "20020418"
  itmX.SubItems(2) = "Example 6"
  itmX.SubItems(3) = "6"
  itmX.SubItems(4) = "0000000006"
  
  Call ListViewFuncs.AutoSizeColumnWidth(Form1.ListView1)
  Call ListViewFuncs.AutoFitColumnWidth(Form1.ListView1)
  
End Sub

Private Sub Form_Resize()
  
  On Error Resume Next
  Form1.ListView1.Width = Form1.Width - Form1.ListView1.Left - 100 - ((Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXBORDER)) * 2)
  Form1.ListView1.Height = Form1.Height - Form1.ListView1.Top - 120 - Form1.Command1.Height - 10 - (Screen.TwipsPerPixelY * GetSystemMetrics(SM_CYCAPTION)) - ((Screen.TwipsPerPixelY * GetSystemMetrics(SM_CYBORDER)) * 2)
  Form1.Command1.Top = Form1.Height - Form1.Command1.Height - 120 - (Screen.TwipsPerPixelY * GetSystemMetrics(SM_CYCAPTION)) - ((Screen.TwipsPerPixelY * GetSystemMetrics(SM_CYBORDER)) * 2)
  Form1.Command1.Left = (Form1.ListView1.Width - Form1.Command1.Width) \ 2
  On Error GoTo 0
  
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
  
  Call ListViewFuncs.SortListView(Form1.ListView1, ColumnHeader.Index)
  
End Sub



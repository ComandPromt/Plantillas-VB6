VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List View Headers"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2430
      Top             =   1125
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2310
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   4075
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Col1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "col2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "col3"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblInfo 
      Caption         =   "Move mouse pointer to column header to retrieve its name."
      Height          =   525
      Left            =   90
      TabIndex        =   1
      Top             =   2535
      Width           =   4530
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''Types required by API functions
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const GWL_WNDPROC = (-4)
Private Const WM_MOUSEMOVE = &H200
Private Const WM_NOTIFY = &H4E
Private Const CLASSNAME = "msvb_lib_header"

''API function declarations.
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Private Sub Form_Load()
    ''Set up the list view
    ListView1.ColumnHeaders(1).Width = ListView1.Width / 3
    ListView1.ColumnHeaders(2).Width = ListView1.Width / 3
    ListView1.ColumnHeaders(3).Width = ListView1.Width / 3
    Dim lItem           As ListItem
    With ListView1
        Set lItem = ListView1.ListItems.Add(, , "Test1")
        lItem.SubItems(1) = "Test2"
        lItem.SubItems(2) = "Test3"
    End With
    
    ''Set form Scale mode to Pixels
    Me.ScaleMode = vbPixels
End Sub

Private Sub Timer1_Timer()
    Dim iCnt            As Long
    Dim I               As Integer
    Dim pt1             As POINTAPI
    Dim pt              As POINTAPI
    Dim rt              As RECT
    Dim retVal          As Long
    
    ''Get the current cursor position
    GetCursorPos pt
            
    retVal = FindWindowEx(ListView1.hwnd, ByVal 0&, GetWndClassName, vbNullString)
    Call GetWindowRect(retVal, rt)
    retVal = PtInRect(rt, pt.x, pt.y)
    
    ''If cursor is not on column header then exit function.
    If retVal = 0 Then Exit Sub
    
    ''Cursor is on the column header so proceed.
    iCnt = FindWindowEx(ListView1.hwnd, ByVal 0&, "msvb_lib_header", vbNullString)
    
    ''Convert to Client co-ordinates of column header.
    ScreenToClient iCnt, pt
    
    With frmMain.ListView1
        For I = 1 To ListView1.ColumnHeaders.Count
            pt1.x = .ColumnHeaders(I).Left
            pt1.y = .ColumnHeaders(I).Width
            If pt1.x < pt.x And pt.x < (pt1.x + pt1.y) Then
                Select Case I
                    Case Is = 1
                        lblInfo.Caption = "Cursor is on column 1"
                    Case Is = 2
                        lblInfo.Caption = "Cursor is on column 2"
                    Case Is = 3
                        lblInfo.Caption = "Cursor is on column 3"
                End Select
            End If
        Next
    End With
End Sub

'****************************************************************************************************
' Name              : GetWndClassName (Private Function)
' Author            : NIRANJANM on Sat, 11 May 2002
' Purpose           : Retrieve the class name for a window.
' Returns    : String
'****************************************************************************************************
Private Function GetWndClassName() As String
    Dim retVal              As Long
    Dim pt                  As POINTAPI
    Dim sBuff               As String
    
    On Error Resume Next
    
    GetWndClassName = ""
    
    sBuff = Space$(256)
    retVal = GetCursorPos(pt)
    retVal = WindowFromPoint(pt.x, pt.y)
    If retVal Then
        retVal = GetClassName(retVal, sBuff, 255)
        sBuff = VBA.Left$(sBuff, retVal)
        GetWndClassName = sBuff
    End If
    
End Function

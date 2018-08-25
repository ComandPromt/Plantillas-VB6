VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.1"
Object = "*\AIEMonitor.vbp"
Begin VB.Form Form1 
   Caption         =   "Test App"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   13755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Tree"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Events"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin IEMonitor.IEEvents IEEvents1 
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1296
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvBrowsers 
      Height          =   7215
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   12726
      _Version        =   393217
      Indentation     =   106
      LabelEdit       =   1
      Style           =   5
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7560
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   9360
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   7080
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   12975
   End
   Begin VB.Label Label1 
      Caption         =   "IEEvents Test Application - Click Start then Open a new IE Window."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   13335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim oBr As InternetExplorer
    IEEvents1.Enabled = True
    tvBrowsers.Nodes.Clear
    tvBrowsers.Nodes.Add(, , "ROOT", "All Instances", 2).Expanded = True
    tvBrowsers.Nodes("ROOT").Text = "All Browsers (" & IEEvents1.Browsers.Count & ")"
    For Each oBr In IEEvents1.Browsers
        tvBrowsers.Nodes.Add("ROOT", tvwChild, "_" & oBr.hwnd, "Browser ", 1).Expanded = True
        tvBrowsers.Nodes.Add "_" & oBr.hwnd, tvwChild, "_" & oBr.hwnd & "_" & "TITLE", "Location Name : " & oBr.LocationName
        tvBrowsers.Nodes.Add "_" & oBr.hwnd, tvwChild, "_" & oBr.hwnd & "_" & "URL", "URL : " & oBr.LocationURL
        tvBrowsers.Nodes.Add "_" & oBr.hwnd, tvwChild, "_" & oBr.hwnd & "_" & "PROG", "Progress : 0%"
    Next oBr
    Command1.Enabled = False
    Command2.Enabled = True
    Command3.Enabled = True
End Sub

Private Sub Command2_Click()
    IEEvents1.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = False
    Command3.Enabled = False
End Sub

Private Sub Command3_Click()
    IEEvents1.Refresh
End Sub



Private Sub Command4_Click()
    List1.ZOrder 0
    
End Sub

Private Sub Command5_Click()
    tvBrowsers.ZOrder 0
End Sub

Private Sub Form_Load()
    'tvBrowsers.Nodes.Add(, , "ROOT", "All Instances", 2).Expanded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub

Private Sub IEEvents1_BrowserCreated(Browser As SHDocVw.InternetExplorer)
    tvBrowsers.Nodes("ROOT").Text = "All Browsers (" & IEEvents1.Browsers.Count & ")"
    tvBrowsers.Nodes.Add("ROOT", tvwChild, "_" & Browser.hwnd, "Browser", 1).Expanded = True
    tvBrowsers.Nodes.Add "_" & Browser.hwnd, tvwChild, "_" & Browser.hwnd & "_" & "TITLE", "Location Name : " & Browser.LocationName
    tvBrowsers.Nodes.Add "_" & Browser.hwnd, tvwChild, "_" & Browser.hwnd & "_" & "URL", "URL : " & Browser.LocationURL
    tvBrowsers.Nodes.Add "_" & Browser.hwnd, tvwChild, "_" & Browser.hwnd & "_" & "PROG", "Progress : 0%"
End Sub

Private Sub IEEvents1_BrowserDestroyed()
    Dim oNode As Node
    Dim oBr As InternetExplorer
    Dim blnFound As Boolean
    AddEventL Nothing, "Browser Destroyed"
    tvBrowsers.Nodes("ROOT").Text = "All Browsers (" & IEEvents1.Browsers.Count & ")"
    For Each oNode In tvBrowsers.Nodes
        If Not oNode.Parent Is Nothing Then
            If oNode.Parent.Key = "ROOT" Then
                blnFound = False
                For Each oBr In IEEvents1.Browsers
                    If oNode.Key = "_" & oBr.hwnd Then
                            blnFound = True
                            Exit For
                    End If
                Next oBr
                If Not blnFound Then
                    tvBrowsers.Nodes.Remove oNode.Key
                    Exit Sub
                End If
            End If
        End If
        
    Next oNode
    
End Sub

Private Sub IEEvents1_BrowserNavigating(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    AddEvent Browser, "URL", "URL : " & CStr(URL)
End Sub

Private Sub IEEvents1_DocumentComplete(Browser As SHDocVw.InternetExplorer, pDisp As Object, URL As Variant)
    AddEventL Browser, "Document Complete -" & URL
End Sub

Private Sub IEEvents1_DownLoadBegin(Browser As SHDocVw.InternetExplorer)
    AddEventL Browser, "Download Begin"
End Sub

Private Sub IEEvents1_DownLoadComplete(Browser As SHDocVw.InternetExplorer)
    AddEventL Browser, "Download Complete"
End Sub

Private Sub IEEvents1_FileDownload(Browser As SHDocVw.InternetExplorer, Cancel As Boolean)
    AddEventL Browser, "File Download"
End Sub

Private Sub IEEvents1_NavigateComplete(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant)
    AddEventL Browser, "Navigate Complete - " & URL
End Sub

Private Sub IEEvents1_NavigateError(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
    AddEventL Browser, "Navigate Error - " & URL
End Sub

Private Sub IEEvents1_NewWindow(Browser As SHDocVw.InternetExplorer, ppDisp As Object, Cancel As Boolean)
    Me.Show
    Cancel = (MsgBox("A new window is opening,Allow It?", vbYesNo, "New PopupWindow Opening!") = vbNo)
End Sub

Private Sub IEEvents1_OnFullScreen(Browser As SHDocVw.InternetExplorer, ByVal FullScreen As Boolean)
    AddEventL Browser, "OnFullScreen"
End Sub

Private Sub IEEvents1_ProgressChange(Browser As SHDocVw.InternetExplorer, ByVal Progress As Long, ByVal ProgressMax As Long)
    If ProgressMax = 0 Then
        AddEvent Browser, "PROG", "Progress : 100%"
    Else
        AddEvent Browser, "PROG", "Progress : " & Round((Progress / ProgressMax) * 100, 0) & "%"
    End If
End Sub

Private Sub IEEvents1_TitleChange(Browser As SHDocVw.InternetExplorer, ByVal Text As String)
    AddEvent Browser, "TITLE", "Title : " & Text
End Sub

Sub AddEventL(b As SHDocVw.InternetExplorer, strText As String)
    If Not b Is Nothing Then
        List1.AddItem "[" & Now & "] Browser " & b.hwnd & " - " & strText
    Else
        List1.AddItem "[" & Now & "] " & strText
    End If
End Sub
Sub AddEvent(b As SHDocVw.InternetExplorer, header As String, strText As String)
    tvBrowsers.Nodes("_" & b.hwnd & "_" & header).Text = strText
    AddEventL b, strText
End Sub
Private Sub IEEvents1_WindowClosing(Browser As SHDocVw.InternetExplorer, ByVal IsChildWindow As Boolean, Cancel As Boolean)
    AddEventL Browser, "Window Closing"
End Sub

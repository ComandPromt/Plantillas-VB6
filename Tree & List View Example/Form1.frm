VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List and Tree View Example by Pio"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " TreeView"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TreeView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   " ListView"
      TabPicture(1)   =   "Form1.frx":0322
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " About"
      TabPicture(2)   =   "Form1.frx":0614
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label4"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label5"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label6"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Picture2"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Text1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin VB.CommandButton Command6 
         Caption         =   "Add Row"
         Height          =   375
         Left            =   -72120
         TabIndex        =   17
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "YDJ@aol.com"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   240
         Picture         =   "Form1.frx":09A6
         ScaleHeight     =   375
         ScaleWidth      =   765
         TabIndex        =   9
         Top             =   960
         Width           =   765
      End
      Begin VB.CommandButton Command5 
         Caption         =   "C"
         Height          =   375
         Left            =   -73560
         TabIndex        =   8
         Top             =   3720
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "E"
         Height          =   375
         Left            =   -73800
         TabIndex        =   7
         Top             =   3720
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -74760
         MousePointer    =   14  'Arrow and Question
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   6
         Top             =   3720
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add Child"
         Height          =   375
         Left            =   -73200
         TabIndex        =   5
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add Main"
         Height          =   375
         Left            =   -72000
         TabIndex        =   4
         Top             =   3720
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   3
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Column 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Column 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Column 3"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   5530
         _Version        =   393217
         Indentation     =   353
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label Label6 
         Caption         =   "Build 1.0"
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   $"Form1.frx":1924
         Height          =   2055
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label Label4 
         Caption         =   "Email:"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Handle: Pio"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Age: 16"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Coded by: Allen Nabors"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":208F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Selected_Node As String

Private Sub Command1_Click()
Dim Frm As Form 'unloads all the forms
For Each Frm In Forms
Unload Frm
Next

End Sub

Private Sub Command2_Click()
'Adds a "Main" or "ROOT" node to the TreeView

On Error GoTo ERROR_ERR 'Error handler
Dim strNode_M As String
ERROR_RETRY:
strNode_M = InputBox("Enter the desired name of the main node.", "Example", "Main " & TreeView1.Nodes.Count + 1)
If Trim(strNode_M) = "" Then
If MsgBox("Please enter a valid name", vbOKCancel + vbCritical, "Error") = vbCancel Then
Exit Sub
Else
GoTo ERROR_RETRY
End If
Else
TreeView1.Nodes.Add , , strNode_M, strNode_M, 1, 1 'double click the "Add Child" button for info.
End If


ERROR_ERR:

End Sub

Private Sub Command3_Click()
'Adds Child Nodes to "Main" or "ROOT" nodes, can also add Child nodes to other Child nodes.
Dim strNode_C As String
ERROR_RETRY_2:
If Trim(Selected_Node) = "" Then Exit Sub
strNode_C = InputBox("What do you want the child name to be?", "Example", "Child " & TreeView1.Nodes.Count)
If Trim(strNode_C) = "" Then
If MsgBox("Please enter a valid name", vbOKCancel + vbCritical, "Error") = vbCancel Then
Exit Sub
Else
GoTo ERROR_RETRY_2
End If
End If

TreeView1.Nodes.Add Selected_Node, tvwChild, strNode_C, strNode_C, 2, 2 'The first part is the Selected Nodes key, the second is what it is it's a child node, thirds is this guys key, fourth is it's text, five is the icon from the listimage, six is the selected image








End Sub

Private Sub Command4_Click()
Dim i As Integer 'Expands all nodes
For i = 1 To TreeView1.Nodes.Count
TreeView1.Nodes.Item(i).Expanded = True
Next i

End Sub

Private Sub Command5_Click()
Dim i As Integer 'Collapse all of the nodes
For i = 1 To TreeView1.Nodes.Count
TreeView1.Nodes.Item(i).Expanded = False
Next i

End Sub

Private Sub Command6_Click()
On Error GoTo Error_Exit

'adding a 3 columned item in listview is kinda tricky but easy
Dim strList_I As String
strList_I = InputBox("Enter the text in this format, a:b:c", "Example", "Hey:Sup:Dude") 'get the text
If strList_I = "" Then Exit Sub 'you can do the same as with the treeview but i am short on time =x
If InStr(strList_I, ":") Then 'checks for our format
ListView1.ListItems.Add , , Split(strList_I, ":")(0), 1, 1
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , Split(strList_I, ":")(1), 2
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , Split(strList_I, ":")(2), 1


Else
MsgBox "Error invalid format", vbCritical + vbOKOnly, "Error"
Exit Sub
End If



Error_Exit:

End Sub

Private Sub Form_Load()
Picture1.Picture = ImageList1.ListImages.Item(3).Picture




End Sub

Private Sub Picture1_Click()
MsgBox "note:" & vbCrLf & "you must click on a node, to add child nodes to it. also, nodes that are already child nodes can have child nodes also.", vbInformation + vbOKOnly, "Help"

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Selected_Node = Node.Key


End Sub

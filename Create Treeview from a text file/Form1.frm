VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   1785
   ClientTop       =   1455
   ClientWidth     =   5415
   LinkTopic       =   "Form10"
   ScaleHeight     =   4875
   ScaleWidth      =   5415
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   60
      Width           =   855
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3413
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "File"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Load a TreeView control from a file that uses tabs
' to show indentation.
Private Sub LoadTreeViewFromFile(ByVal file_name As String, ByVal trv As TreeView)
Dim fnum As Integer
Dim text_line As String
Dim level As Integer
Dim tree_nodes() As Node
Dim num_nodes As Integer

    fnum = FreeFile
    Open file_name For Input As fnum

    TreeView1.Nodes.Clear
    Do While Not EOF(fnum)
        ' Get a line.
        Line Input #fnum, text_line

        ' Find the level of indentation.
        level = 1
        Do While Left$(text_line, 1) = vbTab
            level = level + 1
            text_line = Mid$(text_line, 2)
        Loop

        ' Make room for the new node.
        If level > num_nodes Then
            num_nodes = level
            ReDim Preserve tree_nodes(1 To num_nodes)
        End If

        ' Add the new node.
        If level = 1 Then
            Set tree_nodes(level) = TreeView1.Nodes.Add(, , , text_line)
        Else
            Set tree_nodes(level) = TreeView1.Nodes.Add(tree_nodes(level - 1), tvwChild, , text_line)
            tree_nodes(level).EnsureVisible
        End If
    Loop

    Close fnum
'''
'''Dim i As Integer
'''Dim factory As Node
'''Dim group As Node
'''Dim person As Node
'''
'''    ' Load pictures into the ImageList.
'''    For i = 1 To 6
'''        TreeImages.ListImages.Add , , TreeImage(i).Picture
'''    Next i
'''
'''    ' Attach the TreeView to the ImageList.
'''    OrgTree.ImageList = TreeImages
'''
'''    ' Create some nodes.
'''    Set factory = OrgTree.Nodes.Add(, , "f R & D", "R & D", otFactory, otFactory2)
'''    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Engineering", "Engineering", otGroup, otGroup2)
'''    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Cameron, Charlie", "Cameron, Charlie", otPerson, otPerson2)
'''    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Davos, Debbie", "Davos, Debbie", otPerson, otPerson2)
'''    person.EnsureVisible
'''    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Test", "Test", otGroup, otGroup2)
'''    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Able, Andy", "Andy, Able", otPerson, otPerson2)
'''    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Baker, Betty", "Baker, Betty", otPerson, otPerson2)
'''    person.EnsureVisible
'''
'''    Set factory = OrgTree.Nodes.Add(, , "f Sales & Support", "Sales & Support", otFactory, otFactory2)
'''    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Showroom Sales", "Showroom Sales", otGroup, otGroup2)
'''    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Gaines, Gina", "Gaines, Gina", otPerson, otPerson2)
'''    person.EnsureVisible
'''    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Field Service", "Field Service", otGroup, otGroup2)
'''    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Helms, Harry", "Helms, Harry", otPerson, otPerson2)
'''    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Ives, Irma", "Ives, Irma", otPerson, otPerson2)
'''    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Jackson, Josh", "Jackson, Josh", otPerson, otPerson2)
'''    person.EnsureVisible
'''    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Customer Support", "Customer Support", otGroup, otGroup2)
'''    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Klug, Karl", "Klug, Karl", otPerson, otPerson2)
'''    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Landau, Linda", "Landau, Linda", otPerson, otPerson2)
'''    person.EnsureVisible
'''
End Sub

Private Sub Command1_Click()
    LoadTreeViewFromFile Text1.Text, TreeView1
End Sub

Private Sub Form_Load()
Dim file_name As String

    file_name = App.Path
    If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
    file_name = file_name & "test.txt"
    Text1.Text = file_name
End Sub
Private Sub Form_Resize()
Dim hgt As Single

    hgt = ScaleHeight - TreeView1.Top
    If hgt < 120 Then hgt = 120
    TreeView1.Move 0, TreeView1.Top, ScaleWidth, hgt
End Sub

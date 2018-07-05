VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmListViewExample 
   Caption         =   "ListView Example"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   Icon            =   "frmListViewExample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNewValue 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   10
      Text            =   "New Value"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtItemToSelect 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Text            =   "1"
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdSelectAnItem 
      Caption         =   "Change One Item (First Column)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   0
      Width           =   2655
   End
   Begin VB.TextBox txtNumberOfItemsToAdd 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "100"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdLoopThroughListView 
      Caption         =   "Apply Changes to All (First Column)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   0
      Width           =   2655
   End
   Begin VB.CommandButton cmdFillListView 
      Caption         =   "Fill ListView"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   11456
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "New Value:"
      Height          =   195
      Left            =   3840
      TabIndex        =   9
      Top             =   480
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Item to Select:"
      Height          =   195
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of items to add:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1680
   End
End
Attribute VB_Name = "frmListViewExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This code was written by Elliot McCardle
'
'Purpose: It is provided for instructional purposes.
'
'Details: This code allows you to fill a ListView control with random
'items and subitems generated at runtime. Once the values have been
'created, you are able to change the values either all at once ,or one
'at a time.
'
'Disclaimer: This code is provided as-is, without warranty. The author
'takes no responsibility for problems arising from the use of this code.
'
'Contact: If you have any questions or comments, please feel free to
'e-mail me at emccardle@vnetpro.com
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim lItem As ListItem

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdFillListView_Click()

    Dim i
    i = 1

    If CLng(txtNumberOfItemsToAdd) < 1 Or CLng(txtNumberOfItemsToAdd) > 10000 Then
        MsgBox "Number can't be less than 1 or greater than 10000."
        txtNumberOfItemsToAdd.BackColor = vbYellow
        txtNumberOfItemsToAdd = ""
        txtNumberOfItemsToAdd.SetFocus
        Exit Sub
    End If

    Do Until i = CInt(txtNumberOfItemsToAdd.Text)
        Set lItem = ListView1.ListItems.Add(, , Rnd(100))
        lItem.ListSubItems.Add , , Rnd(100)
        lItem.ListSubItems.Add , , Rnd(100)
        i = i + 1
    Loop

    MsgBox i & " items added to the list box."

    cmdLoopThroughListView.Enabled = True
    cmdSelectAnItem.Enabled = True
    txtItemToSelect.Enabled = True
    txtNewValue.Enabled = True

End Sub

Private Sub cmdLoopThroughListView_Click()

    Dim i

    For i = 1 To ListView1.ListItems.Count
        Set ListView1.SelectedItem = ListView1.ListItems(i)
        ListView1.SelectedItem.Text = txtNewValue
    Next i

End Sub

Private Sub cmdSelectAnItem_Click()

    If CLng(txtItemToSelect) > ListView1.ListItems.Count Then
        MsgBox "Item not available"
        txtItemToSelect.BackColor = vbYellow
        txtItemToSelect = ""
        txtItemToSelect.SetFocus
        Exit Sub
    End If

    Set ListView1.SelectedItem = ListView1.ListItems(CLng(txtItemToSelect.Text))
    ListView1.SelectedItem.Text = txtNewValue

End Sub

Private Sub Form_Load()

    With ListView1
        .ColumnHeaders(1).Width = (ListView1.Width / 3) - 30
        .ColumnHeaders(2).Width = (ListView1.Width / 3) - 30
        .ColumnHeaders(3).Width = (ListView1.Width / 3) - 30
    End With

End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TreeView Example - Pio"
   ClientHeight    =   2895
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open a File to Load"
      Filter          =   "Text Files|*.txt|All Files|*.*|"
   End
   Begin MSComctlLib.TreeView TV1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4683
      _Version        =   393217
      Indentation     =   441
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Menu mnuExpand 
      Caption         =   "Expand"
   End
   Begin VB.Menu mnuCollapse 
      Caption         =   "Collapse"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
'#################################################
'#### Notes: ########################################
'# The main thing you have to learn is to use the Keys, Keys are very imporant#
'#'                                                                                                         #
'#################################################

TV1.Nodes.Add , , "main", "Contents" 'Create Main Parent
    TV1.Nodes.Add "main", tvwChild, "I1", "Example #1" 'Child Node to the Main Parent or ROOT
    TV1.Nodes.Add "I1", tvwChild, "SI1", "Example #1's Child Node"
        TV1.Nodes.Add "main", tvwChild, "I2", "Example #2"
            TV1.Nodes.Add "main", tvwChild, "I3", "Example #3"
            TV1.Nodes.Add "I3", tvwChild, "IA3", "Example #3's Child Node"
            TV1.Nodes.Add "IA3", tvwChild, "IAA3", "Example #3's 2nd Child Node"
                TV1.Nodes.Add "main", tvwChild, "L1", "Load List Example"
                TV1.Nodes.Add "L1", tvwChild, "LA1", "Click To Load List"

TV1.Nodes.Item(1).Expanded = True 'expands the 1st or ROOT node
TV1.Nodes.Item(8).Expanded = True ' expands the 8th Node


End Sub

Private Sub mnuCollapse_Click()
Dim i As Integer
For i = 0 To TV1.Nodes.Count - 1 ' goes threw each node and Collapses it
TV1.Nodes.Item(i + 1).Expanded = False
Next i

End Sub

Private Sub mnuExpand_Click()
Dim i As Integer
For i = 0 To TV1.Nodes.Count - 1 ' goes threw each node and Expandes it
TV1.Nodes.Item(i + 1).Expanded = True
Next i
End Sub

Private Sub TV1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim File As String, Temp
If Node.Key = "LA1" Then ' make sure that the node that is clicked is our Load node
CommonDialog1.ShowOpen ' opens the open dialog
File = CommonDialog1.FileName ' this is the file you selected
    If Trim(File) = "" Then Exit Sub 'makes sure the file isnt just spaces
        TV1.Nodes.Add "main", tvwChild, "I" & FileLen(File), File 'i used the FileLen as the key for this cause it would be different :) i guess"
            Open File For Input As #1 ' opens the file
                
               Do While Not EOF(1) 'this makes the loop keep going unless its at the End Of File (EOF) then it stops
                    Input #1, Temp 'gets the info
                    If Trim(Temp) = "" Then ' make sure that Temp isn't just spaces
                    Else
                    TV1.Nodes.Add "I" & FileLen(File), tvwChild, "I" & TV1.Nodes.Count + 2, Temp 'adds a child node containging the text in the file
                    End If
                    
                Loop ' loop
            Close #1 'closes the file
End If

End Sub

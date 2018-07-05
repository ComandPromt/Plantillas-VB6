Attribute VB_Name = "basFILE"
Option Explicit

'Private Sub GenerateTree()
'    On Error GoTo ErrorGenerateTree
'    Dim obj As Scripting.FileSystemObject, f As Scripting.Folder, i As Scripting.File
'    Dim sf As Scripting.Folder
'    tvLIST.Nodes.Clear
'    Set obj = New Scripting.FileSystemObject
'    Set f = obj.GetFolder(Me.ImagePath)
'    For Each sf In f.SubFolders
'        tvLIST.Nodes.Add , , sf.Name, sf.Name, 1, 1
'        For Each i In sf.Files
'            If UCase(Right(i.Name, 4)) = ".JPG" And UCase(Right(i.Name, 7)) <> "_SM.JPG" Then
'                tvLIST.Nodes.Add sf.Name, tvwChild, sf.Name & "\" & i.Name, i.Name, 2, 2
'            End If
'        Next
'    Next
'    Exit Sub
'ErrorGenerateTree:
'    MsgBox Err & ":Error in GenerateTree.  Error Message: " & Err.Description, vbCritical, "Warning"
'    Exit Sub
'End Sub

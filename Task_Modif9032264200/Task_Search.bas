Attribute VB_Name = "Task_Search"
'this module was designed to search a treeview for certain data

Option Explicit
Public Enum Searchtyp
    FINDBY_HANDLE
    FINDBY_TEXT
    FINDBY_EXECUTABLE
End Enum
Public Enum Find_Flags
    First
    FindNext
End Enum
Public FindAgain As Boolean

Public Function FindText(sWhich As Find_Flags, DaNode As TreeView, Optional searchtype As Searchtyp) As Boolean

  Static sStringToFind As String ' Search the node text
  Static iNodeItem As Long ' Last node found (for repeat)
  Static stype As Searchtyp
  Dim i As Integer, s As String
  Dim anode As Node
  Dim iWindow As Long
  Dim found As Long

    If IsMissing(searchtype) Then
        Stop
    End If

    ' First time -- Ask for search string
    '  and display any previous search string
    If sWhich = Find_Flags.First Then
        stype = searchtype
        sStringToFind = InputBox("Enter the item to find. Searches are not case sensitive.", "Find")
        iNodeItem = 0
      Else
        searchtype = stype
    End If
    iWindow = DaNode.Nodes.Count
    ' Do search
    If sStringToFind <> "" Then
        For i = iNodeItem + 1 To iWindow - 1
            Select Case searchtype
              Case Searchtyp.FINDBY_TEXT
                found = InStr(LCase$(DaNode.Nodes(i).Text), LCase$(sStringToFind))
              Case Searchtyp.FINDBY_HANDLE
                If Val(Mid$(DaNode.Nodes(i).Key, 2)) = Val(sStringToFind) Then
                    found = 1
                End If
              Case Searchtyp.FINDBY_EXECUTABLE
                found = InStr(LCase$(GetExeFromHandle(Val(Mid$(DaNode.Nodes(i).Key, 2)))), LCase$(sStringToFind))
            End Select
           
            If found Then
                iNodeItem = i
                DaNode.Nodes(iNodeItem).Selected = True
                DaNode.Nodes(iNodeItem).EnsureVisible
                Set DaNode.SelectedItem = DaNode.Nodes(iNodeItem)
                DaNode.SetFocus
                Exit For
            End If
        Next i
        ' No match
        If i = iWindow Then
            If sWhich = First Then
                MsgBox "No match to " & sStringToFind
              Else
                MsgBox "No more matches For " & sStringToFind
            End If
        End If
        DaNode.SetFocus
    End If
    FindText = Not (i = iWindow)
    FindAgain = FindText
    Set DaNode = Nothing

End Function

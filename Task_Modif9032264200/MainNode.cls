VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' a simple class i use to collect data for my treeview
'GetWindow API does not collect data in correct order of Child/parent relations,
'it retrieves the windows based on zOrder.
'since some children are called before their parents in zOrder it is difficult to add the nodes without
'error handling the treeview itself without the app crashing, therefore i made this class
'to act as a collection list to ensure that my app will work properly, without doing all that work
' that i assume is necessary.


Option Explicit
Private TheList As Collection
Public Count As Long

Public Sub AddNode(hwnd As Long)

  Dim Par As String
  Dim X As Variant

    If hwnd <> 0 Then
        On Local Error Resume Next
          Par = CStr(GetParent(hwnd))
          Err = 0
          X = TheList.item(Par)
          If Err <> 0 Then
              TheList.Add Par, Par, 2
          End If
          TheList.Add CStr(hwnd), CStr(hwnd), , Par
        On Local Error GoTo 0
    End If
        
End Sub

Private Sub Class_Initialize()

    Set TheList = New Collection
    TheList.Add "0", "0"

End Sub

Private Sub Class_Terminate()

    Set TheList = Nothing

End Sub

Public Sub Clear()

  Dim i As Variant

    For Each i In TheList
        TheList.Remove (i)
    Next i
    TheList.Add "0", "0"

End Sub

Public Function GetCount() As Long

    GetCount = TheList.Count

End Function

Public Function GetItem(ItemNum As Long) As String

    GetItem = TheList.item(ItemNum)

End Function

Public Sub RemoveNode(hwnd As Long)

    TheList.Remove (CStr(hwnd))
    KillTimer frmMain.hwnd, hwnd

End Sub

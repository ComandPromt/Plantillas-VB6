Attribute VB_Name = "Task_Props"
'i use this module to modify/view the properties of some forms

Option Explicit
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function EnumProps Lib "user32" Alias "EnumPropsA" (ByVal hwnd As Long, ByVal lpEnumFunc As Long) As Long
Private Declare Function EnumPropsEx Lib "user32" Alias "EnumPropsExA" (ByVal hwnd As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private PropCounter As Long

Public Sub Add_Prop(hwnd As Long, iName As String, iData As Long)

    SetProp hwnd, iName, iData
    GetPropList hwnd

End Sub

Public Function Delete_Prop(hwnd As Long, iName As String) As Boolean

    Delete_Prop = CBool(Val(RemoveProp(hwnd, iName)))
    GetPropList hwnd

End Function

Public Function Get_Prop(hwnd As Long, iName As String) As Long

    Get_Prop = GetProp(hwnd, iName)

End Function

Public Function GetPropCount(hwnd) As Long

    PropCounter = 0
    EnumProps hwnd, AddressOf PropCountProc
    GetPropCount = PropCounter

End Function

Public Sub GetPropList(hwnd)

    frmMain.PPList.Clear
    EnumProps hwnd, AddressOf PropEnumProc
    frmMain.PPText(0).Text = CStr(frmMain.PPList.ListCount)

End Sub

Private Function PropCountProc(ByVal hwnd As Long, ByVal lpszString As Long, ByVal hData As Long) As Boolean

    PropCounter = PropCounter + 1
    PropCountProc = True

End Function

Private Function PropEnumProc(ByVal hwnd As Long, ByVal lpszString As Long, ByVal hData As Long) As Boolean

  Dim Buffer As String

    'create a buffer
    Buffer = Space$(lstrlen(lpszString) + 1)
    'copy the string to the buffer
    lstrcpy Buffer, lpszString
    frmMain.PPList.AddItem Buffer
    PropEnumProc = True

End Function



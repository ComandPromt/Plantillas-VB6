Attribute VB_Name = "Task_ListMenu"

Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const WM_SETTEXT As Long = &HC
Private Const GWL_EXSTYLE As Long = -20
Private Const LB_GETCOUNT As Long = &H18B
Private Const LB_GETITEMDATA  As Long = &H199
Private Const LB_GETTEXTLEN As Long = &H18A
Private Const LB_SETSEL As Long = &H185
Private Const LB_SETCURSEL As Long = &H186
Private Const LB_GETCURSEL  As Long = &H188
Private Const LB_GETTEXT  As Long = &H189
Private Const LB_ADDSTRING  As Long = &H180
Private Const LB_FINDSTRING  As Long = &H18F
Private Const LB_FINDSTRINGEXACT  As Long = &H1A2
Private Const LB_GETITEMHEIGHT  As Long = &H1A1
Private Const LB_DELETESTRING  As Long = &H182
Private Const LB_SETITEMDATA  As Long = &H19A
Private Const LB_INSERTSTRING As Long = &H181

'this copies listA to listB
Public Function CopyListToList(SourceHwnd As Long, DestHwnd As Long) As Long

  Dim c As Long, d As Long, numitems As Long
  Dim sItemText As String * 255

    numitems = SendMessageLong(SourceHwnd, LB_GETCOUNT, 0&, 0&)
    If numitems > 0 Then
        For c = 0 To numitems - 1
            'get String from listbox
            SendMessageStr SourceHwnd, LB_GETTEXT, c, ByVal sItemText
            'Add to other listbox
            SendMessageStr DestHwnd, LB_ADDSTRING, 0&, ByVal sItemText
            'get item data from list
            d& = SendMessage(SourceHwnd, LB_GETITEMDATA, ByVal c, ByVal 0&)
            'add item data to list
            SendMessage DestHwnd, LB_SETITEMDATA, ByVal c, ByVal d&
        Next c
    End If
    'get the count of the list again
    numitems = SendMessageLong(DestHwnd, LB_GETCOUNT, 0&, 0&)
    CopyListToList = numitems
End Function

'Gets the item count of a listbox
Public Function GetListItemCount(hwnd)

    GetListItemCount = SendMessageLong(hwnd, LB_GETCOUNT, 0&, 0&)

End Function

'My way of finding if control is a listbox
'i know there is a better way, but i have not discovered it
Public Function IsList(hwnd As Long) As Boolean

    Select Case GetWindowLong(hwnd, GWL_EXSTYLE)
      Case 128, 516

        IsList = True
      Case Else
        If SendMessageLong(hwnd, LB_GETCOUNT, 0&, 0&) Then
            IsList = True
        End If
    End Select

End Function

'adds an item to a listbox
Public Sub LstAddItem(ListHwnd As Long, SelectedHwnd As Long, Datxt As String, DaData As Long)

  Dim RetVal As Long

    'add to my listbox
    SendMessageStr ListHwnd, LB_ADDSTRING, 0&, ByVal Datxt
    RetVal = SendMessageLong(ListHwnd, LB_GETCOUNT, 0&, 0&)
    SendMessage ListHwnd, LB_SETITEMDATA, ByVal RetVal - 1, ByVal DaData
    'add to the selected Hwnd listbox
    SendMessageStr SelectedHwnd, LB_ADDSTRING, 0&, ByVal Datxt
    RetVal = SendMessageLong(SelectedHwnd, LB_GETCOUNT, 0&, 0&)
    SendMessage SelectedHwnd, LB_SETITEMDATA, ByVal RetVal - 1, ByVal DaData

End Sub

'gets the item data from an item in a listbox
Public Function LstGetItemData(SelectedHwnd As Long, LstIndex As Long)

    LstGetItemData = SendMessage(SelectedHwnd, LB_GETITEMDATA, ByVal LstIndex, ByVal 0&)

End Function

'Inserts a new item in a listbox
Public Sub LstInsertItem(ListHwnd As Long, SelectedHwnd As Long, LstIndex As Long, Datxt As String)

    SendMessageStr SelectedHwnd, LB_INSERTSTRING, ByVal LstIndex, Datxt
    SendMessageStr ListHwnd, LB_INSERTSTRING, ByVal LstIndex, Datxt

End Sub

'removes an item from a listbox
Public Sub LstRemoveItem(ListHwnd As Long, SelectedHwnd As Long, LstIndex As Long)

    SendMessageLong SelectedHwnd, LB_DELETESTRING, LstIndex, 0&
    SendMessageLong ListHwnd, LB_DELETESTRING, LstIndex, 0&

End Sub

'replaces an item
Public Sub LstReplaceItem(ListHwnd As Long, SelectedHwnd As Long, LstIndex As Long, Datxt As String, DaData As Long)

    LstSetItemText ListHwnd, SelectedHwnd, LstIndex, Datxt
    LstSetItemData ListHwnd, SelectedHwnd, LstIndex, DaData

End Sub

'sets the item data for an item in a listbox
'this is used by Sub LstReplaceItem
Public Sub LstSetItemData(ListHwnd As Long, SelectedHwnd As Long, LstIndex As Long, DaData As Long)

    SendMessage SelectedHwnd, LB_SETITEMDATA, ByVal LstIndex, ByVal DaData
    SendMessage ListHwnd, LB_SETITEMDATA, ByVal LstIndex, ByVal DaData

End Sub

'sets the items text in a listbox by replacing it
'this is used by Sub LstReplaceItem
Public Sub LstSetItemText(ListHwnd As Long, SelectedHwnd As Long, LstIndex As Long, Datxt As String)

  'delete the current item

    LstRemoveItem ListHwnd, SelectedHwnd, LstIndex
    'insert a new one
    LstInsertItem ListHwnd, SelectedHwnd, LstIndex, Datxt

End Sub


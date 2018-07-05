Attribute VB_Name = "modListViewEnh"

'/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\
'/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\
'
' File Name: modListViewEnh.mod
' By: John Allan Lee
' Email Address: zero42@quik.com
' Last Revision: Wednesday, October 13, 1999
'
' Assumes: None
'
' Number of Functions: 24
'
' Includes: EnhListView_ResizeColumns(lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Resizes all Columns in a ListView to fit the text in the rows
'
' Includes: EnhListView_SortColumns(lstListViewName As ListView, usdColIndex, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Use on ColumnClick to sort by that Column
'              Toggles between Ascending and Descending Sorts
'
' Includes: EnhListView_ResizeColumnCaptions(lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Resizes all Columns in a ListView to the Text in the Column Caption
'
' Includes: EnhListView_ResizeColumnHeaders(lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Resizes the ColumnHeaders in a ListView to the Width
'              of the ListView
'
' Includes: EnhListView_Add_FullRowSelect( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Enables Full Row Select in a ListView
'
' Includes: EnhListView_Rem_FullRowSelect( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Disables Full Row Select in a ListView
'
' Includes: EnhListView_Add_GridLines( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Enables GridLines in a ListView
'
' Includes: EnhListView_Rem_GridLines( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Disables GridLines in a ListView
'
' Includes: EnhListView_Add_CheckBoxes( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Enables CheckBoxes in a ListView
'
' Includes: EnhListView_Rem_CheckBoxes( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Disables CheckBoxes in a ListView
'
' Includes: EnhListView_Add_AllowRepositioning( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Enables Repositioning of ColumnHeaders in a ListView
'
' Includes: EnhListView_Rem_AllowRepositioning( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Disables Repositioning of ColumnHeaders in a ListView
'
' Includes: EnhListView_Add_TrackSelected( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Enables TrackSelected in a ListView
'
' Includes: EnhListView_Rem_TrackSelected( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Disables TrackSelected in a ListView
'
' Includes: EnhListView_Add_OneClickActivate( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Enables One Click Activate in a ListView
'
' Includes: EnhListView_Rem_OneClickActivate( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Disables One Click Activate in a ListView
'
' Includes: EnhListView_Add_TwoClickActivate( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Enables Two Click Activate in a ListView
'
' Includes: EnhListView_Rem_TwoClickActivate( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Enables Full Row Select in a ListView
'
' Includes: EnhListView_Add_SubitemImages( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Enables SubItem Images in a ListView
'
' Includes: EnhListView_Rem_SubitemImages( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Disables SubItem Images in a ListView
'
' Includes: EnhLitView_CheckAllItems( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Checks all Items in the ListView
'
' Includes: EnhLitView_UnCheckAllItems( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Unchecks all items in a ListView
'
' Includes: EnhListView_InvertAllChecks( lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Inverts all checked items in a ListView
'
' Includes: EnhListView_Toggle_FlatColumnHeaders( frmFormName As Form, lstListViewName As ListView, Optional bolShowErrors As Boolean) As Boolean
' Type: Public Function
' Description: Toggles FlatColumnHeaders in a ListView
'
' Variable Name: LengthPerCharacter
' Type: Long
' Description: Used to set the length in twips per character for column spacing
'              Defaults to 80 if not specified
'
' Disclaimer:
'   Code provided by John Allan Lee 'as is', without warranties as to
'   performance, fitness, merchantability, and any other warranty
'   (whether expressed or implied).
'   This source code is copyrighted by John Allan Lee who has exclusive
'   rights to distribute it.
'
' Freeware:
'   Code is freely redistributable for personal use in source code form,
'   or for personal or business use in a non-source code binary executable.
'   All other redistributions are prohibited without express written consent
'   from John Allan Lee.
'
'/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\
'/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\_/=\

Option Explicit

'=======================================================================
' needed for Enhancements
Private Const LVIS_STATEIMAGEMASK As Long = &HF000

Private Type LVITEM
    mask         As Long
    iItem        As Long
    iSubItem     As Long
    state        As Long
    stateMask    As Long
    pszText      As String
    cchTextMax   As Long
    iImage       As Long
    lParam       As Long
    iIndent      As Long
End Type

Const SWP_DRAWFRAME = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4

Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVS_EX_GRIDLINES = &H1
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVS_EX_HEADERDRAGDROP = &H10
Private Const LVS_EX_TRACKSELECT = &H8
Private Const LVS_EX_ONECLICKACTIVATE = &H40
Private Const LVS_EX_TWOCLICKACTIVATE = &H80
Private Const LVS_EX_SUBITEMIMAGES = &H2

Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVM_GETHEADER = (LVM_FIRST + 31)

Public Const LVIF_STATE = &H8
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)

Private Const HDS_BUTTONS = &H2
Private Const GWL_STYLE = (-16)

Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Public Declare Function SendMessageAny _
                        Lib "user32" _
                        Alias "SendMessageA" _
                        (ByVal hwnd As Long, _
                        ByVal Msg As Long, _
                        ByVal wParam As Long, _
                        lParam As Any) _
                        As Long

Private Declare Function SendMessageLong Lib _
                        "user32" Alias _
                        "SendMessageA" _
                        (ByVal hwnd As Long, _
                        ByVal Msg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) _
                        As Long
                        
Private Declare Function GetWindowLong _
                        Lib "user32" _
                        Alias "GetWindowLongA" _
                        (ByVal hwnd As Long, _
                        ByVal nIndex As Long) _
                        As Long
                        
Private Declare Function SetWindowLong _
                        Lib "user32" _
                        Alias "SetWindowLongA" _
                        (ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) _
                        As Long
                        
Private Declare Function SetWindowPos _
                        Lib "user32" _
                        (ByVal hwnd As Long, _
                        ByVal hWndInsertAfter As Long, _
                        ByVal x As Long, _
                        ByVal Y As Long, _
                        ByVal cx As Long, _
                        ByVal cy As Long, _
                        ByVal wFlags As Long) _
                        As Long
'=======================================================================

'=======================================================================
Public LengthPerCharacter As Long
'=======================================================================

'=======================================================================
' Description: Resizes all Columns in a ListView to fit the text in
'              the rows
'=======================================================================
Public Function EnhListView_ResizeColumns( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '_______________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_ResizeColumns
    
    '_______________________________________________________________________
    ' set function return to true
    EnhListView_ResizeColumns = True
    
    '_______________________________________________________________________
    ' if the user has not set LengthPerCharacter use 80
    If LengthPerCharacter = 0 Then LengthPerCharacter = "80"
    
    '_______________________________________________________________________
    ' if there are columns to go through...
    If lstListViewName.ListItems.Count > 0 Then
        ' setup variables
        Dim lngIndexCounter As Long
        Dim lngColumnCounter As Long
        ' move through each column
        For lngColumnCounter = 1 To lstListViewName.ColumnHeaders.Count
            ' move though each entry
            For lngIndexCounter = 1 To lstListViewName.ListItems.Count
                ' if it is not the first column
                If lngColumnCounter > 1 Then
                    ' size the column 85 twips per letter
                    If Len(lstListViewName.ListItems.Item(lngIndexCounter).SubItems(lngColumnCounter - 1)) * LengthPerCharacter > _
                    lstListViewName.ColumnHeaders.Item(lngColumnCounter).Width Then
                        lstListViewName.ColumnHeaders.Item(lngColumnCounter).Width = _
                        Len(lstListViewName.ListItems.Item(lngIndexCounter).SubItems(lngColumnCounter - 1)) * LengthPerCharacter
                    End If
                ' if it is the first column
                Else
                    ' size the column 85 twips per letter
                    If Len(lstListViewName.ListItems.Item(lngIndexCounter).Text) * LengthPerCharacter > _
                    lstListViewName.ColumnHeaders.Item(lngColumnCounter).Width Then
                        lstListViewName.ColumnHeaders.Item(lngColumnCounter).Width = _
                        Len(lstListViewName.ListItems.Item(lngIndexCounter).Text) * LengthPerCharacter
                    End If
                End If
            Next lngIndexCounter
        Next lngColumnCounter
    End If
    
    '_______________________________________________________________________
    ' exit before error handler
    Exit Function
    
'_______________________________________________________________________
' deal with errors
err_EnhListView_ResizeColumns:
    
    '_______________________________________________________________________
    ' set function return to false
    EnhListView_ResizeColumns = False
    '_______________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_ResizeColumns"
    End If
    
    '_______________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_ResizeColumns" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '_______________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Use on ColumnClick to sort by that Column
'              Toggles between Ascending and Descending Sorts
'=======================================================================
Public Function EnhListView_SortColumns( _
                lstListViewName As ListView, _
                usdColIndex, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '_______________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_SortColumns
    
    '_______________________________________________________________________
    ' set function return to true
    EnhListView_SortColumns = True
    
    '_______________________________________________________________________
    ' if there are columns to go through...
    If lstListViewName.ListItems.Count > 0 Then
        ' if the sort property is turned off turn it on
        If lstListViewName.Sorted = False Then lstListViewName.Sorted = True
        ' set the sortby column
        lstListViewName.SortKey = _
            lstListViewName.ColumnHeaders.Item(usdColIndex).Index - 1
        ' if it's sorted ascending
        If lstListViewName.SortOrder = lvwAscending Then
            ' sort it descending
            lstListViewName.SortOrder = lvwDescending
        ' if it's sorted descending
        Else
            ' sort it ascending
            lstListViewName.SortOrder = lvwAscending
        End If
    End If
    
    '_______________________________________________________________________
    ' exit before error handler
    Exit Function
    
'_______________________________________________________________________
' deal with errors
err_EnhListView_SortColumns:
    
    '_______________________________________________________________________
    ' set function return to false
    EnhListView_SortColumns = False
    '_______________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_SortColumns"
    End If
    
    '_______________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_SortColumns" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '_______________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Resizes all Columns in a ListView to the Text in the
'              Column Caption
'=======================================================================
Public Function EnhListView_ResizeColumnCaptions( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '_______________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_ResizeColumnCaptions
    
    '_______________________________________________________________________
    ' set function return to true
    EnhListView_ResizeColumnCaptions = True
    
    '_______________________________________________________________________
    ' if the user has not set LengthPerCharacter use 80
    If LengthPerCharacter = 0 Then LengthPerCharacter = "80"
    
    '_______________________________________________________________________
    ' if there are columns to go through...
    If lstListViewName.ListItems.Count > 0 Then
        ' setup variables
        Dim lngColumnCounter As Long
        ' move through each column
        For lngColumnCounter = 1 To lstListViewName.ColumnHeaders.Count
            ' make the size of the column equal to 85 twips per character
            lstListViewName.ColumnHeaders.Item(lngColumnCounter).Width = _
            Len(lstListViewName.ColumnHeaders.Item(lngColumnCounter).Text) * LengthPerCharacter
        Next lngColumnCounter
    End If
    
    '_______________________________________________________________________
    ' exit before error handler
    Exit Function
    
'_______________________________________________________________________
' deal with errors
err_EnhListView_ResizeColumnCaptions:
    
    '_______________________________________________________________________
    ' set function return to false
    EnhListView_ResizeColumnCaptions = False
    '_______________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_ResizeColumnCaptions"
    End If
    
    '_______________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_ResizeColumnCaptions" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '_______________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Resizes the ColumnHeaders in a ListView to the Width
'              of the ListView
'=======================================================================
Public Function EnhListView_ResizeColumnHeaders( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_ResizeColumnHeaders
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_ResizeColumnHeaders = True
    
    '________________________________________________________________________
    ' setup variables
    Dim lngColCounter As Long
    Dim lngListViewDiv As Long
    
    '________________________________________________________________________
    ' fill variables
    lngListViewDiv = lstListViewName.Width / lstListViewName.ColumnHeaders.Count - 300
    
    '________________________________________________________________________
    For lngColCounter = 1 To lstListViewName.ColumnHeaders.Count
        lstListViewName.ColumnHeaders(lngColCounter).Width = lngListViewDiv
    Next lngColCounter
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_ResizeColumnHeaders:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_ResizeColumnHeaders = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_ResizeColumnHeaders"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_ResizeColumnHeaders" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Enables Full Row Select in a ListView
'=======================================================================
Public Function EnhListView_Add_FullRowSelect( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Add_FullRowSelect
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Add_FullRowSelect = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Or LVS_EX_FULLROWSELECT
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Add_FullRowSelect:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Add_FullRowSelect = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Add_FullRowSelect"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Add_FullRowSelect" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Disables Full Row Select in a ListView
'=======================================================================
Public Function EnhListView_Rem_FullRowSelect( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Rem_FullRowSelect
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Rem_FullRowSelect = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' remove the selected style from the current styles
    rStyle = rStyle Xor LVS_EX_FULLROWSELECT
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Rem_FullRowSelect:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Rem_FullRowSelect = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Rem_FullRowSelect"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Rem_FullRowSelect" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Enables GridLines in a ListView
'=======================================================================
Public Function EnhListView_Add_GridLines( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Add_GridLines
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Add_GridLines = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Or LVS_EX_GRIDLINES
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Add_GridLines:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Add_GridLines = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Add_GridLines"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Add_GridLines" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Disables GridLines in a ListView
'=======================================================================
Public Function EnhListView_Rem_GridLines( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Rem_GridLines
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Rem_GridLines = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' remove the selected style from the current styles
    rStyle = rStyle Xor LVS_EX_GRIDLINES
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Rem_GridLines:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Rem_GridLines = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Rem_GridLines"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Rem_GridLines" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Enables CheckBoxes in a ListView
'=======================================================================
Public Function EnhListView_Add_CheckBoxes( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Add_CheckBoxes
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Add_CheckBoxes = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Or LVS_EX_CHECKBOXES
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Add_CheckBoxes:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Add_CheckBoxes = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Add_CheckBoxes"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Add_CheckBoxes" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Disables CheckBoxes in a ListView
'=======================================================================
Public Function EnhListView_Rem_CheckBoxes( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Rem_CheckBoxes
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Rem_CheckBoxes = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Xor LVS_EX_CHECKBOXES
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Rem_CheckBoxes:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Rem_CheckBoxes = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Rem_CheckBoxes"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Rem_CheckBoxes" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Enables Repositioning of ColumnHeaders in a ListView
'=======================================================================
Public Function EnhListView_Add_AllowRepositioning( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Add_AllowRepositioning
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Add_AllowRepositioning = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Or LVS_EX_HEADERDRAGDROP
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Add_AllowRepositioning:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Add_AllowRepositioning = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Add_AllowRepositioning"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Add_AllowRepositioning" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Disables Repositioning of ColumnHeaders in a ListView
'=======================================================================
Public Function EnhListView_Rem_AllowRepositioning( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Rem_AllowRepositioning
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Rem_AllowRepositioning = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Xor LVS_EX_HEADERDRAGDROP
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Rem_AllowRepositioning:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Rem_AllowRepositioning = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Rem_AllowRepositioning"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Rem_AllowRepositioning" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Enables TrackSelected in a ListView
'=======================================================================
Public Function EnhListView_Add_TrackSelected( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Add_TrackSelected
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Add_TrackSelected = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Or LVS_EX_TRACKSELECT
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Add_TrackSelected:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Add_TrackSelected = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Add_TrackSelected"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Add_TrackSelected" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Disables TrackSelected in a ListView
'=======================================================================
Public Function EnhListView_Rem_TrackSelected( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Rem_TrackSelected
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Rem_TrackSelected = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Xor LVS_EX_TRACKSELECT
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Rem_TrackSelected:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Rem_TrackSelected = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Rem_TrackSelected"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Rem_TrackSelected" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Enables One Click Activate in a ListView
'=======================================================================
Public Function EnhListView_Add_OneClickActivate( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Add_OneClickActivate
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Add_OneClickActivate = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Or LVS_EX_ONECLICKACTIVATE
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Add_OneClickActivate:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Add_OneClickActivate = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Add_OneClickActivate"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Add_OneClickActivate" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Disables One Click Activate in a ListView
'=======================================================================
Public Function EnhListView_Rem_OneClickActivate( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Rem_OneClickActivate
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Rem_OneClickActivate = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Xor LVS_EX_ONECLICKACTIVATE
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Rem_OneClickActivate:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Rem_OneClickActivate = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Rem_OneClickActivate"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Rem_OneClickActivate" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Enables Two Click Activate in a ListView
'=======================================================================
Public Function EnhListView_Add_TwoClickActivate( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Add_TwoClickActivate
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Add_TwoClickActivate = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Or LVS_EX_TWOCLICKACTIVATE
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Add_TwoClickActivate:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Add_TwoClickActivate = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Add_TwoClickActivate"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Add_TwoClickActivate" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Enables Full Row Select in a ListView
'=======================================================================
Public Function EnhListView_Rem_TwoClickActivate( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Rem_TwoClickActivate
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Rem_TwoClickActivate = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Xor LVS_EX_TWOCLICKACTIVATE
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Rem_TwoClickActivate:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Rem_TwoClickActivate = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Rem_TwoClickActivate"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Rem_TwoClickActivate" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Enables SubItem Images in a ListView
'=======================================================================
Public Function EnhListView_Add_SubitemImages( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Add_SubitemImages
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Add_SubitemImages = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Or LVS_EX_SUBITEMIMAGES
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Add_SubitemImages:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Add_SubitemImages = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Add_SubitemImages"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Add_SubitemImages" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Disables SubItem Images in a ListView
'=======================================================================
Public Function EnhListView_Rem_SubitemImages( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Rem_SubitemImages
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Rem_SubitemImages = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' remove the selected style from the current styles
    rStyle = rStyle Xor LVS_EX_SUBITEMIMAGES
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Rem_SubitemImages:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Rem_SubitemImages = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Rem_SubitemImages"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Rem_SubitemImages" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Checks all Items in the ListView
'=======================================================================
Public Function EnhLitView_CheckAllItems( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhLitView_CheckAllItems
    
    '________________________________________________________________________
    ' set function return to true
    EnhLitView_CheckAllItems = True
    
    '________________________________________________________________________
    ' setup variables
    Dim LV          As LVITEM
    Dim lvCount     As Long
    Dim lvIndex     As Long
    Dim lvState     As Long
    Dim r           As Long
    
    '________________________________________________________________________
    lvState = IIf(True, &H2000, &H1000)
    lvCount = lstListViewName.ListItems.Count - 1
    Do
        With LV
            .mask = LVIF_STATE
            .state = lvState
            .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstListViewName.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhLitView_CheckAllItems:
    
    '________________________________________________________________________
    ' set function return to false
    EnhLitView_CheckAllItems = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhLitView_CheckAllItems"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhLitView_CheckAllItems" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Unchecks all items in a ListView
'=======================================================================
Public Function EnhLitView_UnCheckAllItems( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhLitView_UnCheckAllItems
    
    '________________________________________________________________________
    ' set function return to true
    EnhLitView_UnCheckAllItems = True
    
    '________________________________________________________________________
    ' setup variables
    Dim LV          As LVITEM
    Dim lvCount     As Long
    Dim lvIndex     As Long
    Dim lvState     As Long
    Dim r           As Long
    
    '________________________________________________________________________
    lvState = IIf(True, &H2000, &H1000)
    lvCount = lstListViewName.ListItems.Count - 1
    Do
        With LV
            .mask = LVIF_STATE
            .state = lvState
            .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstListViewName.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhLitView_UnCheckAllItems:
    
    '________________________________________________________________________
    ' set function return to false
    EnhLitView_UnCheckAllItems = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhLitView_UnCheckAllItems"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhLitView_UnCheckAllItems" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================


'=======================================================================
' Description: Inverts all checked items in a ListView
'=======================================================================
Public Function EnhListView_InvertAllChecks( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_InvertAllChecks
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_InvertAllChecks = True
    
    '________________________________________________________________________
    ' setup variables
    Dim LV          As LVITEM
    Dim r           As Long
    Dim lvCount     As Long
    Dim lvIndex     As Long
    
    '________________________________________________________________________
    lvCount = lstListViewName.ListItems.Count - 1
    Do
        r = SendMessageLong(lstListViewName.hwnd, LVM_GETITEMSTATE, lvIndex, LVIS_STATEIMAGEMASK)
        With LV
            .mask = LVIF_STATE
            .stateMask = LVIS_STATEIMAGEMASK
            If r And &H2000& Then
                .state = &H1000
            Else
                .state = &H2000
            End If
        End With
        r = SendMessageAny(lstListViewName.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_InvertAllChecks:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_InvertAllChecks = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_InvertAllChecks"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_InvertAllChecks" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Toggles FlatColumnHeaders in a ListView
'=======================================================================
Public Function EnhListView_Toggle_FlatColumnHeaders( _
                frmFormName As Form, _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Toggle_FlatColumnHeaders
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Toggle_FlatColumnHeaders = True
    
    '________________________________________________________________________
    SetWindowLong SendMessageLong(lstListViewName.hwnd, _
                                 LVM_GETHEADER, _
                                 0, _
                                 ByVal 0&), _
                                 GWL_STYLE, _
                                 GetWindowLong(SendMessageLong(lstListViewName.hwnd, _
                                                               LVM_GETHEADER, _
                                                               0, _
                                                               ByVal _
                                                               0&), _
                                                               GWL_STYLE) _
                                                               Xor HDS_BUTTONS
    SetWindowPos lstListViewName.hwnd, _
                 frmFormName.hwnd, _
                 0, _
                 0, _
                 0, _
                 0, _
                 SWP_FLAGS
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Toggle_FlatColumnHeaders:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Toggle_FlatColumnHeaders = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Toggle_FlatColumnHeaders"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Toggle_FlatColumnHeaders" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

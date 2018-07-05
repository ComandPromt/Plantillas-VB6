Attribute VB_Name = "ListViewFuncs"
Option Explicit

Public gListViewTotalSelected As Long
Public gListViewSelected() As Long
Public gListViewItemToInsertBefore As Long

Public Type LV_FINDINFO
      
      flags As Long
      psz As String
      lParam As Long
      pt As POINTAPI
      vkDirection As Long

End Type

Public Type LV_ITEM
      
      mask As Long
      iItem As Long
      iSubItem As Long
      State As Long
      stateMask As Long
      pszText As Long
      cchTextMax As Long
      iImage As Long
      lParam As Long
      iIndent As Long

End Type

Public Const LVFI_PARAM = &H1
Public Const LVFI_STRING = &H2
Public Const LVFI_PARTIAL = &H8
Public Const LVFI_WRAP = &H20
Public Const LVFI_NEARESTXY = &H40

Declare Function GetListViewItemHeight Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As RECT) As Long
Public Sub ListViewGetSelectedItems(ByVal FormToUse As Form, ByVal ListViewControl As Control)

  Dim Counter As Long
  Dim SelectedCount As Long
  
  FormToUse.MousePointer = vbHourglass
  FormToUse.Enabled = False
  
  SelectedCount = 0
  gListViewTotalSelected = SendMessage(ListViewControl.hWnd, LVM_GETSELECTEDCOUNT, 0, 0)

  If gListViewTotalSelected > 0 Then

    ReDim gListViewSelected(gListViewTotalSelected) As Long
  
    For Counter = 1 To ListViewControl.ListItems.Count
     
       If ListViewControl.ListItems(Counter).Selected = True Then
       
         gListViewSelected(SelectedCount) = Counter
         SelectedCount = SelectedCount + 1
       
       End If
     
    Next Counter

  End If

  FormToUse.Enabled = True
  FormToUse.MousePointer = vbDefault
  
End Sub
Public Sub AutoFitColumnWidth(ByVal lvw As ListView)

  Dim iCounter As Long

  On Error Resume Next
  If lvw.View = lvwReport Then
    
    For iCounter = 1 To lvw.ColumnHeaders.Count
   
       If iCounter > 1 Then
         
         If lvw.ColumnHeaders(iCounter - 1).Tag = "DATE" Then

           lvw.ColumnHeaders(iCounter).Width = 0

         Else

           Call SendMessage(lvw.hWnd, LVM_SETCOLUMNWIDTH, iCounter - 1, ByVal LVSCW_AUTOSIZE_USEHEADER)

         End If
         
       Else
         
         Call SendMessage(lvw.hWnd, LVM_SETCOLUMNWIDTH, iCounter - 1, ByVal LVSCW_AUTOSIZE_USEHEADER)
       
       End If

    Next
    
  End If
  Call FixSortedColumnHeaderIfNeeded(lvw)
  On Error GoTo 0
  
End Sub
Public Sub AutoSizeColumnWidth(ByVal lvw As ListView)

  Dim iCounter As Long

  On Error Resume Next
  If lvw.View = lvwReport Then
    
    For iCounter = 1 To lvw.ColumnHeaders.Count ' - 1
   
       If iCounter > 1 Then
         
         If lvw.ColumnHeaders(iCounter - 1).Tag = "DATE" Then

           lvw.ColumnHeaders(iCounter).Width = 0

         Else

           Call SendMessage(lvw.hWnd, LVM_SETCOLUMNWIDTH, iCounter - 1, ByVal LVSCW_AUTOSIZE)

         End If
       
       Else
         
         Call SendMessage(lvw.hWnd, LVM_SETCOLUMNWIDTH, iCounter - 1, ByVal LVSCW_AUTOSIZE)
         
       End If

    Next
    
  End If
  Call FixSortedColumnHeaderIfNeeded(lvw)
  On Error GoTo 0
  
End Sub

Public Sub FixDateSortedColumnHeaderIfNeeded(ByVal ListViewToUse As ListView)
    
  Dim SaveFontSize As Currency
  Dim SaveFontBold As Currency
  Dim SaveFontName As String
  
  SaveFontSize = ListViewToUse.Parent.Font.Size
  SaveFontBold = ListViewToUse.Parent.Font.BOLD
  SaveFontName = ListViewToUse.Parent.Font.Name
  
  ListViewToUse.Parent.Font.Size = ListViewToUse.Font.Size
  ListViewToUse.Parent.Font.BOLD = ListViewToUse.Font.BOLD
  ListViewToUse.Parent.Font.Name = ListViewToUse.Font.Name
  
  If ListViewToUse.ColumnHeaders(IIf(ListViewToUse.SortKey = 0, 1, ListViewToUse.SortKey)).Width < ListViewToUse.Parent.TextWidth(ListViewToUse.ColumnHeaders(IIf(ListViewToUse.SortKey = 0, 1, ListViewToUse.SortKey)).Text) + ListViewToUse.Parent.TextWidth(ListViewToUse.ColumnHeaders(IIf(ListViewToUse.SortKey = 0, 1, ListViewToUse.SortKey)).Text) \ 2 Then

    ListViewToUse.ColumnHeaders(IIf(ListViewToUse.SortKey = 0, 1, ListViewToUse.SortKey)).Width = ListViewToUse.Parent.TextWidth(ListViewToUse.ColumnHeaders(IIf(ListViewToUse.SortKey = 0, 1, ListViewToUse.SortKey)).Text) + ListViewToUse.Parent.TextWidth(ListViewToUse.ColumnHeaders(IIf(ListViewToUse.SortKey = 0, 1, ListViewToUse.SortKey)).Text) \ 2

  End If
  
  ListViewToUse.Parent.Font.Size = SaveFontSize
  ListViewToUse.Parent.Font.BOLD = SaveFontBold
  ListViewToUse.Parent.Font.Name = SaveFontName

End Sub
Public Sub FixSortedColumnHeaderIfNeeded(ByVal ListViewToUse As ListView)
  
  Dim SaveFontSize As Currency
  Dim SaveFontBold As Currency
  Dim SaveFontName As String
  
  If ListViewToUse.ColumnHeaders(IIf(ListViewToUse.SortKey = 0, 1, ListViewToUse.SortKey)).Tag = "DATE" Then
    
    Call FixDateSortedColumnHeaderIfNeeded(ListViewToUse)
    
  Else
    
    SaveFontSize = ListViewToUse.Parent.Font.Size
    SaveFontBold = ListViewToUse.Parent.Font.BOLD
    SaveFontName = ListViewToUse.Parent.Font.Name
  
    ListViewToUse.Parent.Font.Size = ListViewToUse.Font.Size
    ListViewToUse.Parent.Font.BOLD = ListViewToUse.Font.BOLD
    ListViewToUse.Parent.Font.Name = ListViewToUse.Font.Name
  
    If ListViewToUse.ColumnHeaders(ListViewToUse.SortKey + 1).Width < ListViewToUse.Parent.TextWidth(ListViewToUse.ColumnHeaders(ListViewToUse.SortKey + 1).Text) + ListViewToUse.Parent.TextWidth(ListViewToUse.ColumnHeaders(ListViewToUse.SortKey + 1).Text) \ 2 Then
    
      ListViewToUse.ColumnHeaders(ListViewToUse.SortKey + 1).Width = ListViewToUse.Parent.TextWidth(ListViewToUse.ColumnHeaders(ListViewToUse.SortKey + 1).Text) + ListViewToUse.Parent.TextWidth(ListViewToUse.ColumnHeaders(ListViewToUse.SortKey + 1).Text) \ 2
    
    End If
  
    ListViewToUse.Parent.Font.Size = SaveFontSize
    ListViewToUse.Parent.Font.BOLD = SaveFontBold
    ListViewToUse.Parent.Font.Name = SaveFontName
    
  End If

End Sub
Public Function GetListViewItemIndex(ByVal hWnd As Long, ByVal ItemText As String) As Long
  
  Dim LFI As LV_FINDINFO
  
  LFI.flags = LVFI_PARTIAL Or LVFI_WRAP
  LFI.psz = ItemText
  
  GetListViewItemIndex = SendMessageAny(hWnd, LVM_FINDITEM, -1, LFI)
  
  If GetListViewItemIndex <> -1 Then
    
    GetListViewItemIndex = GetListViewItemIndex + 1
    
  End If
  
End Function
Public Sub SetListViewToWholeRowSelect(ByVal ListViewhWnd As Long)
    
  Dim lStyle As Long
  
  lStyle = SendMessage(ListViewhWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
  lStyle = lStyle Or LVS_EX_FULLROWSELECT
  
  Call SendMessage(ListViewhWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal lStyle)

End Sub

Public Sub ShowHeaderIcon(ByVal ListViewToUse As Control, ByVal colNo As Long, ByVal showImage As Long)

  Dim r As Long
  Dim hHeader As Long
  Dim HD As HD_ITEM
   
  ListViewToUse.SmallIcons = frmListViewImages.ImageList2
  
  ' Get a handle to the listview header component '
   hHeader = SendMessageLong(ListViewToUse.hWnd, LVM_GETHEADER, 0, 0)
   
  ' Set up the required structure members '
  HD.mask = HDI_IMAGE Or HDI_FORMAT
  HD.fmt = HDF_LEFT Or HDF_STRING Or HDF_BITMAP_ON_RIGHT Or showImage
  HD.pszText = ListViewToUse.ColumnHeaders(ListViewToUse.SortKey + 1).Text
  
  If showImage Then
    
    HD.iImage = ListViewToUse.SortOrder
    
  End If
   
  ' Modify the header '
  r = SendMessageAny(hHeader, HDM_SETITEM, colNo, HD)
   
End Sub
Public Sub SortDateListView(ByVal ListViewToUse As ListView, ByVal colNo As Long)
  
  Dim Counter As Long
  
  If colNo = ListViewToUse.SortKey Then
    
    If ListViewToUse.SortOrder = lvwAscending Then
  
      ListViewToUse.SortOrder = lvwDescending
    
    Else
     
      ListViewToUse.SortOrder = lvwAscending
  
    End If
  
  End If
  
  ListViewToUse.Sorted = True
  ListViewToUse.SortKey = colNo
  
  For Counter = 1 To ListViewToUse.ColumnHeaders.Count
  
     If Counter = colNo Then
       
       Call ShowHeaderIcon(ListViewToUse, colNo - 1, HDF_IMAGE)
    
     Else
       
       Call ShowHeaderIcon(ListViewToUse, Counter - 1, 0)
     
     End If
     
  Next Counter
  Call FixDateSortedColumnHeaderIfNeeded(ListViewToUse)

End Sub
Public Sub SortListView(ByVal ListViewToUse As ListView, ByVal colNo As Long)
  
  Dim Counter As Long
  
  If ListViewToUse.ColumnHeaders(colNo).Tag = "DATE" Then
    
    Call SortDateListView(ListViewToUse, colNo)
    
  Else
    
    If colNo - 1 = ListViewToUse.SortKey Then
    
      If ListViewToUse.SortOrder = lvwAscending Then
  
        ListViewToUse.SortOrder = lvwDescending
    
      Else
     
        ListViewToUse.SortOrder = lvwAscending
  
      End If
  
    End If
  
    ListViewToUse.Sorted = True
    ListViewToUse.SortKey = colNo - 1
  
    For Counter = 0 To ListViewToUse.ColumnHeaders.Count - 1
  
       If Counter = ListViewToUse.SortKey Then
       
         Call ShowHeaderIcon(ListViewToUse, ListViewToUse.SortKey, HDF_IMAGE)
    
       Else
       
         Call ShowHeaderIcon(ListViewToUse, Counter, 0)
     
       End If
     
    Next Counter
    Call FixSortedColumnHeaderIfNeeded(ListViewToUse)

  End If
  
End Sub

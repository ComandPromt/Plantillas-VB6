VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl AdvList 
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   ScaleHeight     =   5310
   ScaleWidth      =   5535
   ToolboxBitmap   =   "AdvList.ctx":0000
   Begin VB.PictureBox pic32 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3240
      ScaleHeight     =   495
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox pic16 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2400
      ScaleHeight     =   495
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   1560
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   720
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9340
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "AdvList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This is an advanced ListView Control. It Populates a listview control using the
'given path with the files and folders. The control retains most listview
'properties, events and methods plus:
'UserControl.Path (Read/Write) Property that specifies the path to populate
'UserControl.Populate Method to populate the control
'Feel free to enhance it!
'Copyright ©  2002 George Kontostanos

Option Explicit
Dim FSys As FileSystemObject
Dim UserControlHeight As Single, UserControlWidth As Single
'Default Property Values:
Const m_def_Path = "C:\"
Const m_def_ItemSelected = "0"
'Property Variables:
Dim m_Path As String
Dim m_ItemSelected As String
'Event Declarations:
Event Click() 'MappingInfo=lvMain,lvMain,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=lvMain,lvMain,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=lvMain,lvMain,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=lvMain,lvMain,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=lvMain,lvMain,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lvMain,lvMain,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lvMain,lvMain,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lvMain,lvMain,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event AfterLabelEdit(Cancel As Integer, NewString As String) 'MappingInfo=lvMain,lvMain,-1,AfterLabelEdit
Attribute AfterLabelEdit.VB_Description = "Occurs after a user edits the label of the currently selected Node or ListItem object."
Event BeforeLabelEdit(Cancel As Integer) 'MappingInfo=lvMain,lvMain,-1,BeforeLabelEdit
Attribute BeforeLabelEdit.VB_Description = "Occurs when a user attempts to edit the label of the currently selected ListItem or Node object."
Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event ColumnClick(ByVal ColumnHeader As ColumnHeader) 'MappingInfo=lvMain,lvMain,-1,ColumnClick
Attribute ColumnClick.VB_Description = "Occurs when a ColumnHeader object in a ListView control is clicked."
Event ItemCheck(ByVal Item As ListItem) 'MappingInfo=lvMain,lvMain,-1,ItemCheck
Attribute ItemCheck.VB_Description = "Occurs when a ListSubItem object is checked"
Event ItemClick(ByVal Item As ListItem) 'MappingInfo=lvMain,lvMain,-1,ItemClick
Attribute ItemClick.VB_Description = "Occurs when a ListItem object is clicked or selected"

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = lvMain.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lvMain.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    ForeColor = lvMain.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lvMain.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lvMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lvMain.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
    lvMain.Refresh
End Sub

Private Sub lvMain_Click()
    RaiseEvent Click
End Sub

Private Sub lvMain_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub lvMain_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub lvMain_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub lvMain_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub lvMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lvMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lvMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lvMain_AfterLabelEdit(Cancel As Integer, NewString As String)
    RaiseEvent AfterLabelEdit(Cancel, NewString)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,AllowColumnReorder
Public Property Get AllowColumnReorder() As Boolean
Attribute AllowColumnReorder.VB_Description = "Returns/sets whether a user can reorder columns in report view."
    AllowColumnReorder = lvMain.AllowColumnReorder
End Property

Public Property Let AllowColumnReorder(ByVal New_AllowColumnReorder As Boolean)
    lvMain.AllowColumnReorder() = New_AllowColumnReorder
    PropertyChanged "AllowColumnReorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,Arrange
Public Property Get Arrange() As ListArrangeConstants
Attribute Arrange.VB_Description = "Returns/sets how the icons in a ListView control's Icon or SmallIcon view are arranged."
    Arrange = lvMain.Arrange
End Property

Public Property Let Arrange(ByVal New_Arrange As ListArrangeConstants)
    lvMain.Arrange() = New_Arrange
    PropertyChanged "Arrange"
End Property

Private Sub lvMain_BeforeLabelEdit(Cancel As Integer)
    RaiseEvent BeforeLabelEdit(Cancel)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,Checkboxes
Public Property Get Checkboxes() As Boolean
Attribute Checkboxes.VB_Description = "Returns/sets a value which determines if the control displays a checkbox next to each item in the list."
    Checkboxes = lvMain.Checkboxes
End Property

Public Property Let Checkboxes(ByVal New_Checkboxes As Boolean)
    lvMain.Checkboxes() = New_Checkboxes
    PropertyChanged "Checkboxes"
End Property

Private Sub lvMain_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    RaiseEvent ColumnClick(ColumnHeader)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,ColumnHeaderIcons
Public Property Get ColumnHeaderIcons() As Object
Attribute ColumnHeaderIcons.VB_Description = "Returns/sets the ImageList control to be used for ColumnHeader icons."
    Set ColumnHeaderIcons = lvMain.ColumnHeaderIcons
End Property

Public Property Set ColumnHeaderIcons(ByVal New_ColumnHeaderIcons As Object)
    Set lvMain.ColumnHeaderIcons = New_ColumnHeaderIcons
    PropertyChanged "ColumnHeaderIcons"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,ColumnHeaders
Public Property Get ColumnHeaders() As IColumnHeaders
Attribute ColumnHeaders.VB_Description = "Returns a reference to a collection of ColumnHeader objects."
    Set ColumnHeaders = lvMain.ColumnHeaders
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,FullRowSelect
Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Returns/sets whether selecting a column highlights the entire row."
    FullRowSelect = lvMain.FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
    lvMain.FullRowSelect() = New_FullRowSelect
    PropertyChanged "FullRowSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,GridLines
Public Property Get GridLines() As Boolean
Attribute GridLines.VB_Description = "Returns/sets whether grid lines appear between rows and columns"
    GridLines = lvMain.GridLines
End Property

Public Property Let GridLines(ByVal New_GridLines As Boolean)
    lvMain.GridLines() = New_GridLines
    PropertyChanged "GridLines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,HideColumnHeaders
Public Property Get HideColumnHeaders() As Boolean
Attribute HideColumnHeaders.VB_Description = "Returns/sets whether or not a ListView control's column headers are hidden in Report view."
    HideColumnHeaders = lvMain.HideColumnHeaders
End Property

Public Property Let HideColumnHeaders(ByVal New_HideColumnHeaders As Boolean)
    lvMain.HideColumnHeaders() = New_HideColumnHeaders
    PropertyChanged "HideColumnHeaders"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Determines whether the selected item will display as selected when the ListView loses focus"
    HideSelection = lvMain.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    lvMain.HideSelection() = New_HideSelection
    PropertyChanged "HideSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,HotTracking
Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets whether hot tracking is enabled."
    HotTracking = lvMain.HotTracking
End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
    lvMain.HotTracking() = New_HotTracking
    PropertyChanged "HotTracking"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,HoverSelection
Public Property Get HoverSelection() As Boolean
Attribute HoverSelection.VB_Description = "Returns/sets whether hover selection is enabled."
    HoverSelection = lvMain.HoverSelection
End Property

Public Property Let HoverSelection(ByVal New_HoverSelection As Boolean)
    lvMain.HoverSelection() = New_HoverSelection
    PropertyChanged "HoverSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,Icons
Public Property Get Icons() As Object
Attribute Icons.VB_Description = "Returns/sets the images associated with the Icon properties of a ListView control."
    Set Icons = lvMain.Icons
End Property

Public Property Set Icons(ByVal New_Icons As Object)
    Set lvMain.Icons = New_Icons
    PropertyChanged "Icons"
End Property

Private Sub lvMain_ItemCheck(ByVal Item As ListItem)
    RaiseEvent ItemCheck(Item)
End Sub

Private Sub lvMain_ItemClick(ByVal Item As ListItem)
    RaiseEvent ItemClick(Item)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,LabelEdit
Public Property Get LabelEdit() As ListLabelEditConstants
Attribute LabelEdit.VB_Description = "Returns/sets a value that determines if a user can edit the label of a ListItem or Node object."
    LabelEdit = lvMain.LabelEdit
End Property

Public Property Let LabelEdit(ByVal New_LabelEdit As ListLabelEditConstants)
    lvMain.LabelEdit() = New_LabelEdit
    PropertyChanged "LabelEdit"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,LabelWrap
Public Property Get LabelWrap() As Boolean
Attribute LabelWrap.VB_Description = "Returns or sets a value that determines if labels are wrapped when the ListView is in Icon view."
    LabelWrap = lvMain.LabelWrap
End Property

Public Property Let LabelWrap(ByVal New_LabelWrap As Boolean)
    lvMain.LabelWrap() = New_LabelWrap
    PropertyChanged "LabelWrap"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PopupMenu
Public Sub PopupMenu(ByVal Menu As Object, Optional ByVal flags As Variant, Optional ByVal X As Variant, Optional ByVal Y As Variant, Optional ByVal DefaultMenu As Variant)
Attribute PopupMenu.VB_Description = "Displays a pop-up menu on an MDIForm or Form object."
    UserControl.PopupMenu Menu, flags, X, Y, DefaultMenu
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,SmallIcons
Public Property Get SmallIcons() As Object
Attribute SmallIcons.VB_Description = "Returns/sets the images associated with the SmallIcons property of a ListView control."
    Set SmallIcons = lvMain.SmallIcons
End Property

Public Property Set SmallIcons(ByVal New_SmallIcons As Object)
    Set lvMain.SmallIcons = New_SmallIcons
    PropertyChanged "SmallIcons"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,Sorted
Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = lvMain.Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
    lvMain.Sorted() = New_Sorted
    PropertyChanged "Sorted"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,SortOrder
Public Property Get SortOrder() As ListSortOrderConstants
Attribute SortOrder.VB_Description = "Returns/sets whether or not the ListItems will be sorted in ascending or descending order."
    SortOrder = lvMain.SortOrder
End Property

Public Property Let SortOrder(ByVal New_SortOrder As ListSortOrderConstants)
    lvMain.SortOrder() = New_SortOrder
    PropertyChanged "SortOrder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Path() As String
Attribute Path.VB_Description = "AdvListView Path "
    Path = m_Path
End Property

Public Property Let Path(ByVal New_Path As String)
    m_Path = New_Path
    PropertyChanged "Path"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Populate() As Boolean
Attribute Populate.VB_Description = "Use it to fill the AdvListView Control with Files and Folders of the selected Path. Returns False on Error."
    PopulateList (Path)
    GetIcons (Path)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,1,0
Public Property Get ItemSelected() As String
Attribute ItemSelected.VB_Description = "AdvListView selected item"
    ItemSelected = lvMain.SelectedItem
End Property

Private Sub UserControl_Resize()
    lvMain.Width = lvMain.Width + (UserControl.Width - UserControlWidth)
    lvMain.Height = lvMain.Height + (UserControl.Height - UserControlHeight)
    UserControlWidth = UserControl.Width
    UserControlHeight = UserControl.Height
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Path = App.Path
    m_ItemSelected = m_def_ItemSelected
End Sub

Private Sub UserControl_Initialize()
Dim LWidth
UserControlHeight = UserControl.Height
UserControlWidth = UserControl.Width
'---------Set the ListView Headers-----------------------
LWidth = lvMain.Width - 5 * Screen.TwipsPerPixelX
lvMain.ColumnHeaders.Add 1, , "File Name", 0.3 * LWidth
lvMain.ColumnHeaders.Add 2, , "Size", 0.2 * LWidth, lvwColumnRight
lvMain.ColumnHeaders.Add 3, , "Created", 0.25 * LWidth
lvMain.ColumnHeaders.Add 4, , "Modified", 0.25 * LWidth

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lvMain.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    lvMain.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lvMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lvMain.AllowColumnReorder = PropBag.ReadProperty("AllowColumnReorder", False)
    lvMain.Arrange = PropBag.ReadProperty("Arrange", 0)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    lvMain.Checkboxes = PropBag.ReadProperty("Checkboxes", False)
    Set ColumnHeaderIcons = PropBag.ReadProperty("ColumnHeaderIcons", Nothing)
    lvMain.FullRowSelect = PropBag.ReadProperty("FullRowSelect", True)
    lvMain.GridLines = PropBag.ReadProperty("GridLines", False)
    lvMain.HideColumnHeaders = PropBag.ReadProperty("HideColumnHeaders", False)
    lvMain.HideSelection = PropBag.ReadProperty("HideSelection", False)
    lvMain.HotTracking = PropBag.ReadProperty("HotTracking", False)
    lvMain.HoverSelection = PropBag.ReadProperty("HoverSelection", False)
    Set Icons = PropBag.ReadProperty("Icons", Nothing)
    lvMain.LabelEdit = PropBag.ReadProperty("LabelEdit", 0)
    lvMain.LabelWrap = PropBag.ReadProperty("LabelWrap", True)
    Set SmallIcons = PropBag.ReadProperty("SmallIcons", Nothing)
    lvMain.Sorted = PropBag.ReadProperty("Sorted", False)
    lvMain.SortOrder = PropBag.ReadProperty("SortOrder", 0)
    m_Path = PropBag.ReadProperty("Path", m_def_Path)
    'm_ItemSelected = PropBag.ReadProperty("ItemSelected", m_def_ItemSelected)
    lvMain.View = PropBag.ReadProperty("View", 3)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", lvMain.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", lvMain.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lvMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("AllowColumnReorder", lvMain.AllowColumnReorder, False)
    Call PropBag.WriteProperty("Arrange", lvMain.Arrange, 0)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("Checkboxes", lvMain.Checkboxes, False)
    Call PropBag.WriteProperty("ColumnHeaderIcons", ColumnHeaderIcons, Nothing)
    Call PropBag.WriteProperty("FullRowSelect", lvMain.FullRowSelect, True)
    Call PropBag.WriteProperty("GridLines", lvMain.GridLines, False)
    Call PropBag.WriteProperty("HideColumnHeaders", lvMain.HideColumnHeaders, False)
    Call PropBag.WriteProperty("HideSelection", lvMain.HideSelection, False)
    Call PropBag.WriteProperty("HotTracking", lvMain.HotTracking, False)
    Call PropBag.WriteProperty("HoverSelection", lvMain.HoverSelection, False)
    Call PropBag.WriteProperty("Icons", Icons, Nothing)
    Call PropBag.WriteProperty("LabelEdit", lvMain.LabelEdit, 0)
    Call PropBag.WriteProperty("LabelWrap", lvMain.LabelWrap, True)
    Call PropBag.WriteProperty("SmallIcons", SmallIcons, Nothing)
    Call PropBag.WriteProperty("Sorted", lvMain.Sorted, False)
    Call PropBag.WriteProperty("SortOrder", lvMain.SortOrder, 0)
    Call PropBag.WriteProperty("Path", m_Path, m_def_Path)
    'Call PropBag.WriteProperty("ItemSelected", m_ItemSelected, m_def_ItemSelected)
    Call PropBag.WriteProperty("View", lvMain.View, 3)
End Sub

Private Sub PopulateList(Path As String)
'----------------------------------------------------------------
'Populate ListView With Files and Folders of "Path"
'----------------------------------------------------------------
Dim thisFolder As Folder
Dim subFolder As Folder
Dim allFolders As Folders
Dim allFiles As Files
Dim thisFile As File
Dim thisItem As ListItem

Set FSys = CreateObject("Scripting.FileSystemObject")
Set thisFolder = FSys.GetFolder(Path)
Set allFiles = thisFolder.Files
Set allFolders = thisFolder.SubFolders
'------------------Clear All Items In Memory-----------------------
lvMain.ListItems.Clear
lvMain.Icons = Nothing
lvMain.SmallIcons = Nothing
iml32.ListImages.Clear
iml16.ListImages.Clear
'----------------Get the Folders-----------------------------------
If allFolders.Count > 0 Then
    On Error Resume Next
    For Each subFolder In allFolders
        Set thisItem = lvMain.ListItems.Add(, , subFolder.Name)
        thisItem.SubItems(1) = FormatKB(subFolder.Size)
        thisItem.SubItems(2) = Left(subFolder.DateCreated, 9)
        thisItem.SubItems(3) = Left(subFolder.DateLastModified, 9)
    Next
End If
'----------------Get the Files-------------------------------------
If allFiles.Count > 0 Then
    On Error Resume Next
    For Each thisFile In allFiles
        Set thisItem = lvMain.ListItems.Add(, , thisFile.Name)
        thisItem.SubItems(1) = FormatKB(thisFile.Size)
        thisItem.SubItems(2) = Left(thisFile.DateCreated, 9)
        thisItem.SubItems(3) = Left(thisFile.DateLastModified, 9)
    Next
End If
End Sub

Private Sub GetIcons(Path As String)
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim FileName As String

'Size the picture boxes containing the icons
pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX
pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY

On Local Error Resume Next

For Each Item In lvMain.ListItems
  FileName = CheckPath(Path) & Item.Text
  GetIcon FileName, Item.Index
Next

On Error Resume Next

With lvMain
  '.ListItems.Clear
  .Icons = iml32        'Large
  .SmallIcons = iml16   'Small
  For Each Item In .ListItems
    Item.Icon = Item.Index
    Item.SmallIcon = Item.Index
  Next
End With

End Sub

Private Function GetIcon(FileName As String, Index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long


'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then
  'Large Icon
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hLIcon, ShInfo.iIcon, pic32.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  'Small Icon
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  Set imgObj = iml32.ListImages.Add(Index, , pic32.Image)
  Set imgObj = iml16.ListImages.Add(Index, , pic16.Image)
End If
End Function

Private Function CheckPath(ByVal Path As String) As String
'--------------------------------------------------
'Checks if path ends with "\". If not, add it.
'--------------------------------------------------
If Right(Path, 1) <> "\" Then
  CheckPath = Path & "\"
Else
  CheckPath = Path
End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvMain,lvMain,-1,View
Public Property Get View() As ListViewConstants
Attribute View.VB_Description = "Returns/sets the current view of the ListView control."
    View = lvMain.View
End Property

Public Property Let View(ByVal New_View As ListViewConstants)
    lvMain.View() = New_View
    PropertyChanged "View"
End Property


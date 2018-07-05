VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FileLV 
   ClientHeight    =   2115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   ScaleHeight     =   2115
   ScaleWidth      =   3285
   Begin MSComctlLib.ImageList imgSMALL 
      Left            =   1320
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvLIST 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2355
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.PictureBox picICON 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   600
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   480
      Width           =   240
   End
End
Attribute VB_Name = "FileLV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''Be sure to extract with path names.
''''
''''Note, since the OCX is not included and
''''(of course) not on your system.  If you open
''''the group project it will error on the test
''''project because you do not have the control.
''''You will need to make the control first.
''''
''''
''''Also note, since I have binary compatability set
''''on the control project you will get a nag message
''''about not being able to set compatability when you
''''open, but just ignore that.
''''
''''Final note: I made this in under an hour mainly
''''using the Active X Control Wizard.  I just
''''added the code for getting the icon for a file
''''to the project and a couple other properties.
''''I only made this because a little while ago somebody
''''else posted something similar but it did not work
''''100% but I thought it was a good idea, so I made this
''''real quick in case somebody wants to add to it.
''''
''''This is by no means 100% perfect nor does it have
''''all the functionality that you may want.  But of course
''''feel free to add to it and use it as you wish.
''''
''''-Clint LaFever
''''http://vbasic.iscool.net
'''lafeverc@home.com



Option Explicit
'Event Declarations:
Event Click() 'MappingInfo=lvLIST,lvLIST,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=lvLIST,lvLIST,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvLIST,lvLIST,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvLIST,lvLIST,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvLIST,lvLIST,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event AfterLabelEdit(Cancel As Integer, NewString As String) 'MappingInfo=lvLIST,lvLIST,-1,AfterLabelEdit
Attribute AfterLabelEdit.VB_Description = "Occurs after a user edits the label of the currently selected Node or ListItem object."
Event BeforeLabelEdit(Cancel As Integer) 'MappingInfo=lvLIST,lvLIST,-1,BeforeLabelEdit
Attribute BeforeLabelEdit.VB_Description = "Occurs when a user attempts to edit the label of the currently selected ListItem or Node object."
Event ColumnClick(ByVal ColumnHeader As ColumnHeader) 'MappingInfo=lvLIST,lvLIST,-1,ColumnClick
Attribute ColumnClick.VB_Description = "Occurs when a ColumnHeader object in a ListView control is clicked."
Event ItemCheck(ByVal Item As ListItem) 'MappingInfo=lvLIST,lvLIST,-1,ItemCheck
Attribute ItemCheck.VB_Description = "Occurs when a ListSubItem object is checked"
Event ItemClick(ByVal Item As ListItem) 'MappingInfo=lvLIST,lvLIST,-1,ItemClick
Attribute ItemClick.VB_Description = "Occurs when a ListItem object is clicked or selected"
'Default Property Values:
Const m_def_Path = ""
'Property Variables:
Dim m_Path As String
Public Enum itItemType
    itFOLDER = 0
    itFILE = 1
End Enum
Public ItemType As itItemType
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = lvLIST.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lvLIST.BackColor() = New_BackColor
    picICON.BackColor = lvLIST.BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    ForeColor = lvLIST.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lvLIST.ForeColor() = New_ForeColor
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
'MappingInfo=lvLIST,lvLIST,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lvLIST.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lvLIST.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = lvLIST.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    lvLIST.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
    lvLIST.Refresh
End Sub

Private Sub lvLIST_Click()
    RaiseEvent Click
End Sub

Private Sub lvLIST_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    lvLIST.View = lvwReport
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub lvLIST_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub lvLIST_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub lvLIST_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub lvLIST_AfterLabelEdit(Cancel As Integer, NewString As String)
    RaiseEvent AfterLabelEdit(Cancel, NewString)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,AllowColumnReorder
Public Property Get AllowColumnReorder() As Boolean
Attribute AllowColumnReorder.VB_Description = "Returns/sets whether a user can reorder columns in report view."
    AllowColumnReorder = lvLIST.AllowColumnReorder
End Property

Public Property Let AllowColumnReorder(ByVal New_AllowColumnReorder As Boolean)
    lvLIST.AllowColumnReorder() = New_AllowColumnReorder
    PropertyChanged "AllowColumnReorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not controls, Forms or an MDIForm are painted at run time with 3-D effects."
    Appearance = lvLIST.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    lvLIST.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,Arrange
Public Property Get Arrange() As ListArrangeConstants
Attribute Arrange.VB_Description = "Returns/sets how the icons in a ListView control's Icon or SmallIcon view are arranged."
    Arrange = lvLIST.Arrange
End Property

Public Property Let Arrange(ByVal New_Arrange As ListArrangeConstants)
    lvLIST.Arrange() = New_Arrange
    PropertyChanged "Arrange"
End Property

Private Sub lvLIST_BeforeLabelEdit(Cancel As Integer)
    RaiseEvent BeforeLabelEdit(Cancel)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,CausesValidation
Public Property Get CausesValidation() As Boolean
Attribute CausesValidation.VB_Description = "Returns/sets whether validation occurs on the control which lost focus."
    CausesValidation = lvLIST.CausesValidation
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
    lvLIST.CausesValidation() = New_CausesValidation
    PropertyChanged "CausesValidation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,Checkboxes
Public Property Get Checkboxes() As Boolean
Attribute Checkboxes.VB_Description = "Returns/sets a value which determines if the control displays a checkbox next to each item in the list."
    Checkboxes = lvLIST.Checkboxes
End Property

Public Property Let Checkboxes(ByVal New_Checkboxes As Boolean)
    lvLIST.Checkboxes() = New_Checkboxes
    PropertyChanged "Checkboxes"
End Property

Private Sub lvLIST_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    RaiseEvent ColumnClick(ColumnHeader)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,ColumnHeaderIcons
Public Property Get ColumnHeaderIcons() As Object
Attribute ColumnHeaderIcons.VB_Description = "Returns/sets the ImageList control to be used for ColumnHeader icons."
    Set ColumnHeaderIcons = lvLIST.ColumnHeaderIcons
End Property

Public Property Set ColumnHeaderIcons(ByVal New_ColumnHeaderIcons As Object)
    Set lvLIST.ColumnHeaderIcons = New_ColumnHeaderIcons
    PropertyChanged "ColumnHeaderIcons"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,ColumnHeaders
Public Property Get ColumnHeaders() As IColumnHeaders
Attribute ColumnHeaders.VB_Description = "Returns a reference to a collection of ColumnHeader objects."
    Set ColumnHeaders = lvLIST.ColumnHeaders
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,FullRowSelect
Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Returns/sets whether selecting a column highlights the entire row."
    FullRowSelect = lvLIST.FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
    lvLIST.FullRowSelect() = New_FullRowSelect
    PropertyChanged "FullRowSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,GetFirstVisible
Public Function GetFirstVisible() As IListItem
Attribute GetFirstVisible.VB_Description = "Retrieves a reference of the first item visible in the client area."
    GetFirstVisible = lvLIST.GetFirstVisible()
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,GridLines
Public Property Get GridLines() As Boolean
Attribute GridLines.VB_Description = "Returns/sets whether grid lines appear between rows and columns"
    GridLines = lvLIST.GridLines
End Property

Public Property Let GridLines(ByVal New_GridLines As Boolean)
    lvLIST.GridLines() = New_GridLines
    PropertyChanged "GridLines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,HideColumnHeaders
Public Property Get HideColumnHeaders() As Boolean
Attribute HideColumnHeaders.VB_Description = "Returns/sets whether or not a ListView control's column headers are hidden in Report view."
    HideColumnHeaders = lvLIST.HideColumnHeaders
End Property

Public Property Let HideColumnHeaders(ByVal New_HideColumnHeaders As Boolean)
    lvLIST.HideColumnHeaders() = New_HideColumnHeaders
    PropertyChanged "HideColumnHeaders"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Determines whether the selected item will display as selected when the ListView loses focus"
    HideSelection = lvLIST.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    lvLIST.HideSelection() = New_HideSelection
    PropertyChanged "HideSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,Icons
Public Property Get Icons() As Object
Attribute Icons.VB_Description = "Returns/sets the images associated with the Icon properties of a ListView control."
    Set Icons = lvLIST.Icons
End Property

Public Property Set Icons(ByVal New_Icons As Object)
    Set lvLIST.Icons = New_Icons
    PropertyChanged "Icons"
End Property

Private Sub lvLIST_ItemCheck(ByVal Item As ListItem)
    RaiseEvent ItemCheck(Item)
End Sub

Private Sub lvLIST_ItemClick(ByVal Item As ListItem)
    If Item.Tag = "FOLDER" Then
        Me.ItemType = itFOLDER
    Else
        Me.ItemType = itFILE
    End If
    RaiseEvent ItemClick(Item)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,LabelEdit
Public Property Get LabelEdit() As ListLabelEditConstants
Attribute LabelEdit.VB_Description = "Returns/sets a value that determines if a user can edit the label of a ListItem or Node object."
    LabelEdit = lvLIST.LabelEdit
End Property

Public Property Let LabelEdit(ByVal New_LabelEdit As ListLabelEditConstants)
    lvLIST.LabelEdit() = New_LabelEdit
    PropertyChanged "LabelEdit"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,LabelWrap
Public Property Get LabelWrap() As Boolean
Attribute LabelWrap.VB_Description = "Returns or sets a value that determines if labels are wrapped when the ListView is in Icon view."
    LabelWrap = lvLIST.LabelWrap
End Property

Public Property Let LabelWrap(ByVal New_LabelWrap As Boolean)
    lvLIST.LabelWrap() = New_LabelWrap
    PropertyChanged "LabelWrap"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,MultiSelect
Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Returns/sets a value indicating whether a user can make multiple selections in the ListView control and how the multiple selections can be made."
    MultiSelect = lvLIST.MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
    lvLIST.MultiSelect() = New_MultiSelect
    PropertyChanged "MultiSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets the background picture for the control."
    Set Picture = lvLIST.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set lvLIST.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,PictureAlignment
Public Property Get PictureAlignment() As ListPictureAlignmentConstants
Attribute PictureAlignment.VB_Description = "Returns/sets the picture alignment."
    PictureAlignment = lvLIST.PictureAlignment
End Property

Public Property Let PictureAlignment(ByVal New_PictureAlignment As ListPictureAlignmentConstants)
    lvLIST.PictureAlignment() = New_PictureAlignment
    PropertyChanged "PictureAlignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,SmallIcons
Public Property Get SmallIcons() As Object
Attribute SmallIcons.VB_Description = "Returns/sets the images associated with the SmallIcons property of a ListView control."
    Set SmallIcons = lvLIST.SmallIcons
End Property

Public Property Set SmallIcons(ByVal New_SmallIcons As Object)
    Set lvLIST.SmallIcons = New_SmallIcons
    PropertyChanged "SmallIcons"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,Sorted
Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = lvLIST.Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
    lvLIST.Sorted() = New_Sorted
    PropertyChanged "Sorted"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,SortKey
Public Property Get SortKey() As Integer
Attribute SortKey.VB_Description = "Returns/sets the current sort key."
    SortKey = lvLIST.SortKey
End Property

Public Property Let SortKey(ByVal New_SortKey As Integer)
    lvLIST.SortKey() = New_SortKey
    PropertyChanged "SortKey"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,SortOrder
Public Property Get SortOrder() As ListSortOrderConstants
Attribute SortOrder.VB_Description = "Returns/sets whether or not the ListItems will be sorted in ascending or descending order."
    SortOrder = lvLIST.SortOrder
End Property

Public Property Let SortOrder(ByVal New_SortOrder As ListSortOrderConstants)
    lvLIST.SortOrder() = New_SortOrder
    PropertyChanged "SortOrder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvLIST,lvLIST,-1,StartLabelEdit
Public Sub StartLabelEdit()
Attribute StartLabelEdit.VB_Description = "Begins a label editing operation on a ListItem or Node object."
    lvLIST.StartLabelEdit
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lvLIST.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    picICON.BackColor = lvLIST.BackColor
    lvLIST.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lvLIST.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lvLIST.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    lvLIST.AllowColumnReorder = PropBag.ReadProperty("AllowColumnReorder", False)
    lvLIST.Appearance = PropBag.ReadProperty("Appearance", 1)
    lvLIST.Arrange = PropBag.ReadProperty("Arrange", 0)
    lvLIST.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
    lvLIST.Checkboxes = PropBag.ReadProperty("Checkboxes", False)
    Set ColumnHeaderIcons = PropBag.ReadProperty("ColumnHeaderIcons", Nothing)
    lvLIST.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
    lvLIST.GridLines = PropBag.ReadProperty("GridLines", False)
    lvLIST.HideColumnHeaders = PropBag.ReadProperty("HideColumnHeaders", False)
    lvLIST.HideSelection = PropBag.ReadProperty("HideSelection", True)
    Set Icons = PropBag.ReadProperty("Icons", Nothing)
    lvLIST.LabelEdit = PropBag.ReadProperty("LabelEdit", 0)
    lvLIST.LabelWrap = PropBag.ReadProperty("LabelWrap", True)
    lvLIST.MultiSelect = PropBag.ReadProperty("MultiSelect", False)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    lvLIST.PictureAlignment = PropBag.ReadProperty("PictureAlignment", 0)
    lvLIST.Sorted = PropBag.ReadProperty("Sorted", False)
    lvLIST.SortKey = PropBag.ReadProperty("SortKey", 0)
    lvLIST.SortOrder = PropBag.ReadProperty("SortOrder", 0)
    m_Path = PropBag.ReadProperty("Path", m_def_Path)
    If Me.Path <> "" Then
        FillFiles
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With UserControl
        .lvLIST.Move 0, 0, .ScaleWidth, .ScaleHeight
    End With
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", lvLIST.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", lvLIST.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lvLIST.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", lvLIST.BorderStyle, 1)
    Call PropBag.WriteProperty("AllowColumnReorder", lvLIST.AllowColumnReorder, False)
    Call PropBag.WriteProperty("Appearance", lvLIST.Appearance, 1)
    Call PropBag.WriteProperty("Arrange", lvLIST.Arrange, 0)
    Call PropBag.WriteProperty("CausesValidation", lvLIST.CausesValidation, True)
    Call PropBag.WriteProperty("Checkboxes", lvLIST.Checkboxes, False)
    Call PropBag.WriteProperty("ColumnHeaderIcons", ColumnHeaderIcons, Nothing)
    Call PropBag.WriteProperty("FullRowSelect", lvLIST.FullRowSelect, False)
    Call PropBag.WriteProperty("GridLines", lvLIST.GridLines, False)
    Call PropBag.WriteProperty("HideColumnHeaders", lvLIST.HideColumnHeaders, False)
    Call PropBag.WriteProperty("HideSelection", lvLIST.HideSelection, True)
    Call PropBag.WriteProperty("Icons", Icons, Nothing)
    Call PropBag.WriteProperty("LabelEdit", lvLIST.LabelEdit, 0)
    Call PropBag.WriteProperty("LabelWrap", lvLIST.LabelWrap, True)
    Call PropBag.WriteProperty("MultiSelect", lvLIST.MultiSelect, False)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("PictureAlignment", lvLIST.PictureAlignment, 0)
    Call PropBag.WriteProperty("Sorted", lvLIST.Sorted, False)
    Call PropBag.WriteProperty("SortKey", lvLIST.SortKey, 0)
    Call PropBag.WriteProperty("SortOrder", lvLIST.SortOrder, 0)
    Call PropBag.WriteProperty("Path", m_Path, m_def_Path)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Path() As String
Attribute Path.VB_Description = "Sets the path to files to display."
    Path = m_Path
End Property

Public Property Let Path(ByVal New_Path As String)
    m_Path = New_Path
    While Right(m_Path, 1) = "\"
        m_Path = Left(m_Path, Len(m_Path) - 1)
    Wend
    PropertyChanged "Path"
    If Me.Path <> "" Then
        FillFiles
    Else
        lvLIST.ListItems.Clear
    End If
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Path = m_def_Path
End Sub

Private Sub FillFiles()
    On Error GoTo ErrorFillFiles
    Dim obj As Scripting.FileSystemObject, f As Scripting.Folder, i As Scripting.File
    Dim sf As Scripting.Folder, itm As ListItem, iIMG As CICON, x As Long
    FreezeWindow lvLIST.hWnd
    lvLIST.ListItems.Clear
    lvLIST.ColumnHeaders.Clear
    lvLIST.ColumnHeaders.Add , , "Please wait...", 1440
    lvLIST.ListItems.Add , , "Loading..."
    FreezeWindow
    lvLIST.Refresh
    DoEvents
    FreezeWindow lvLIST.hWnd
    lvLIST.ListItems.Clear
    lvLIST.ColumnHeaders.Clear
    lvLIST.ColumnHeaders.Add , , "NAME"
    lvLIST.ColumnHeaders.Add , , "SIZE"
    lvLIST.ColumnHeaders.Add , , "TYPE"
    lvLIST.ColumnHeaders.Add , , "MODIFIED"
    lvLIST.SmallIcons = Nothing
    imgSMALL.ListImages.Clear
    AddImage imgSMALL, icon_FOLDER_CLOSED, IMG_SIXTEEN
    Set obj = New Scripting.FileSystemObject
    Set f = obj.GetFolder(Me.Path)
    For Each i In f.Files
        Set iIMG = New CICON
        picICON.Picture = LoadPicture()
        iIMG.ExtractIconToHDC picICON.hdc, Me.Path & "\" & i.Name
        Set iIMG = Nothing
        picICON.Picture = picICON.Image
        imgSMALL.ListImages.Add , , picICON.Picture
        picICON.Picture = LoadPicture()
    Next
    Set lvLIST.SmallIcons = imgSMALL
    For Each sf In f.SubFolders
        Set itm = lvLIST.ListItems.Add(, , sf.Name, , 1)
        itm.SubItems(2) = sf.Type
        itm.SubItems(3) = sf.DateLastModified
        itm.Tag = "FOLDER"
    Next
    x = 2
    For Each i In f.Files
        Set itm = lvLIST.ListItems.Add(, , i.Name, , x)
        x = x + 1
        itm.SubItems(1) = i.Size
        itm.SubItems(2) = i.Type
        itm.SubItems(3) = i.DateLastModified
        itm.Tag = "FILE"
    Next
    FreezeWindow
    Exit Sub
ErrorFillFiles:
    FreezeWindow
    MsgBox Err & ":Error in FillFiles.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Public Property Get ListItems() As IListItems
    On Error Resume Next
    Set ListItems = lvLIST.ListItems
End Property
Public Property Get SelectedItem() As IListItem
    On Error Resume Next
    Set SelectedItem = lvLIST.SelectedItem
End Property

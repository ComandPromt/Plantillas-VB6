VERSION 5.00
Begin VB.UserControl ctxOutlookBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   3576
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1512
   MousePointer    =   99  'Custom
   OLEDropMode     =   1  'Manual
   PropertyPages   =   "ctxOutlookBar.ctx":0000
   ScaleHeight     =   3576
   ScaleWidth      =   1512
   ToolboxBitmap   =   "ctxOutlookBar.ctx":003C
End
Attribute VB_Name = "ctxOutlookBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
'   You are free to use this source as long as this copyright message
'     appears on your program's "About" dialog:
'
'   Outlook Bar Project
'   Copyright (c) 2002 Vlad Vissoultchev (wqweto@myrealbox.com)
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "ctxOutlookBar"
Implements ISubclassingSink
Implements IHookingSink

'=========================================================================
' Public Events
'=========================================================================

'Purpose: Occurs when a <b>cButton</b> object has been clicked.
Event ButtonClick(ByVal oBtn As cButton)
Attribute ButtonClick.VB_Description = "Occurs when a cButton object has been clicked."
Attribute ButtonClick.VB_HelpID = 102
Attribute ButtonClick.VB_MemberFlags = "200"
'Purpose: Occurs when a <b>cButton</b> object has been double clicked.
Event ButtonDblClick(ByVal oBtn As cButton)
Attribute ButtonDblClick.VB_Description = "Occurs when a cButton object has been double clicked."
Attribute ButtonDblClick.VB_HelpID = 103
'Purpose: Occurs when the user presses and then releases a mouse button over an <b>Outlook Bar</b> control.
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an Outlook Bar control."
Attribute Click.VB_HelpID = 104
'Purpose: Occurs when the user presses and releases a mouse button two times in succession over an <b>Outlook Bar</b> control.
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button two times in succession over an Outlook Bar control."
Attribute DblClick.VB_HelpID = 105
'Purpose: Occur when the user presses a mouse button.
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occur when the user presses a mouse button."
Attribute MouseDown.VB_HelpID = 132
'Purpose: Occurs when the user moves the mouse over the control, or outside the boundaries of the control if any of the mouse buttons are pressed.
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse over the control, or outside the boundaries of the control if any of the mouse buttons are pressed."
Attribute MouseMove.VB_HelpID = 133
'Purpose: Occur when the user releases a mouse button.
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occur when the user releases a mouse button."
Attribute MouseUp.VB_HelpID = 134
'Purpose: Occurs when a source component is dropped onto a target component, informing the source component that a drag action was either performed or canceled.
Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs when a source component is dropped onto a target component, informing the source component that a drag action was either performed or canceled."
Attribute OLECompleteDrag.VB_HelpID = 135
'Purpose: Occurs when a source component is dropped onto a target component when the source component determines that a drop can occur.
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when a source component is dropped onto a target component when the source component determines that a drop can occur."
Attribute OLEDragDrop.VB_HelpID = 137
'Purpose: Occurs when one component is dragged over another.
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when one component is dragged over another."
Attribute OLEDragOver.VB_HelpID = 139
'Purpose: Occurs after every <b>OLEDragOver</b> event. <b>OLEGiveFeedback</b> allows the source component to provide visual feedback to the user, such as changing the mouse cursor to indicate what will happen if the user drops the object, or provide visual feedback on the selection (in the source component) to indicate what will happen.
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs after every OLEDragOver event. OLEGiveFeedback allows the source component to provide visual feedback to the user, such as changing the mouse cursor to indicate what will happen if the user drops the object, or provide visual feedback on the selec"
Attribute OLEGiveFeedback.VB_HelpID = 141
'Purpose: Occurs on a source component when a target component performs the <b>GetData</b> method on the source’s <b>DataObject</b> object, but the data for the specified format has not yet been loaded.
Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs on a source component when a target component performs the GetData method on the source’s DataObject object, but the data for the specified format has not yet been loaded."
Attribute OLESetData.VB_HelpID = 142
'Purpose: Occurs when an <b>Outlook Bar</b> control's OLEDrag method is performed.<p>This event specifies the data formats and drop effects that the source component supports. It can also be used to insert data into the <b>DataObject</b> object.
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when an Outlook Bar control's OLEDrag method is performed. This event specifies the data formats and drop effects that the source component supports. It can also be used to insert data into the DataObject object."
Attribute OLEStartDrag.VB_HelpID = 143
'Purpose: Occur when the user presses a key while an object has the focus.
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occur when the user presses a key while an object has the focus."
Attribute KeyDown.VB_HelpID = 128
'Purpose: Occurs when the user presses and releases an ANSI key.
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_HelpID = 129
'Purpose: Occur when the user releases a key while an object has the focus.
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occur when the user releases a key while an object has the focus."
Attribute KeyUp.VB_HelpID = 130
'Purpose: Occurs when the currently selected item has changed.
Event SelItemChange()
Attribute SelItemChange.VB_Description = "Occurs when the currently selected item has changed."
Attribute SelItemChange.VB_HelpID = 149
'Purpose: Occurs when the currently selected group has changed.
Event SelGroupChange()
Attribute SelGroupChange.VB_Description = "Occurs when the currently selected group has changed."
Attribute SelGroupChange.VB_HelpID = 148
'Purpose: Occurs when windows appearance settings has changed.
Event WindowsSettingsChanged()
Attribute WindowsSettingsChanged.VB_Description = "Occurs when windows appearance settings has changed."
Attribute WindowsSettingsChanged.VB_HelpID = 151
'Purpose: Occurs when the user presses a mouse button and drags the mouse.
Event MouseDragged(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDragged.VB_Description = "Occurs when the user presses a mouse button and drags the mouse."
Attribute MouseDragged.VB_HelpID = 155
'Purpose: Occurs before an item is moved on an automatic drag operations.
Event OLEBeforeMove(ByVal oBtn As cButton, ByVal NewIndex As Long, Cancel As Boolean)
Attribute OLEBeforeMove.VB_Description = "Occurs before an item is moved on an automatic drag operations."
Attribute OLEBeforeMove.VB_HelpID = 156
'Purpose: Occurs when a user attempts to edit the caption of the currently selected <b>cButton</b> object
Event BeforeLabelEdit(Cancel As Boolean)
Attribute BeforeLabelEdit.VB_Description = "Occurs when a user attempts to edit the caption of the currently selected cButton object"
Attribute BeforeLabelEdit.VB_HelpID = 159
'Purpose: Occurs after a user edits the caption of the currently selected <b>cButton</b> object.
Event AfterLabelEdit(Cancel As Boolean, NewCaption As String)
Attribute AfterLabelEdit.VB_Description = "Occurs after a user edits the caption of the currently selected cButton object."
Attribute AfterLabelEdit.VB_HelpID = 157

'=========================================================================
' Public Enums
'=========================================================================

Public Enum UcsOleDragModeEnum
    ucsOleDragManual = 0        ' Manual. The programmer handles all OLE drag/drop operations.
    ucsOleDragAutomatic = 1     ' Automatic. The component handles all OLE drag/drop operations.
End Enum

Public Enum UcsOleDropModeEnum
    ucsOleDropNone = 0          ' (Default) None. The <b>Outlook Bar</b> control does not accept OLE drops and displays the No Drop cursor.
    ucsOleDropManual = 1        ' Manual. The <b>Outlook Bar</b> control triggers the OLE drop events, allowing the programmer to handle the OLE drop operation in code.
End Enum

Public Enum UcsHitTestEnum
    ucsHitNoWhere               ' The point is outside the client area of the Outlook Bar control.
    ucsHitGroupButton           ' The point is in a group button.
    ucsHitItemButton            ' The point is in an item button.
    ucsHitScrollupArrow         ' The point is in scroll up arrow.
    ucsHitScrolldownArrow       ' The point is in scroll down arrow.
    ucsHitBackground            ' The point on the background of the control.
    ucsHitGroupBackground       ' The point on the background of the current group, outside any item.
End Enum

Public Enum UcsLabelEditEnum
    ucsLbeAutomatic             ' (Default) Automatic. The BeforeLabelEdit event is generated when the user clicks the caption of a selected <b>cButton</b>.
    ucsLbeManual                ' Manual. The BeforeLabelEdit event is never generated. StartLabelEdit function is used to control when the caption of a selected <b>cButton</b> needs to be edited.
End Enum

Private Enum TrackMouseEventFlags
    TME_HOVER = &H1
    TME_LEAVE = &H2
    TME_NONCLIENT = &H10
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

'=========================================================================
' API
'=========================================================================

'--- for mouse_event
Private Const MOUSEEVENTF_MOVE          As Long = &H1
Private Const MOUSEEVENTF_LEFTDOWN      As Long = &H2
'--- for InitCommonControlsEx
Private Const ICC_TAB_CLASSES           As Long = &H8
'--- for CreateWindowEx (tooltip window)
Private Const TOOLTIPS_CLASS            As String = "tooltips_class32"
Private Const WS_POPUP                  As Long = &H80000000
Private Const WS_EX_TOPMOST             As Long = &H8
Private Const TTS_ALWAYSTIP             As Long = &H1
Private Const TTS_NOPREFIX              As Long = &H2
'--- for tooltip window messages
Private Const TTM_ADDTOOL               As Long = (&H400 + 4)
Private Const TTM_DELTOOL               As Long = (&H400 + 5)
Private Const TTM_NEWTOOLRECT           As Long = (&H400 + 6)
Private Const TTM_UPDATETIPTEXT         As Long = (&H400 + 12)
Private Const TTM_SETMAXTIPWIDTH        As Long = (&H400 + 24)
'--- for TOOLINFO.uFlags
Private Const TTF_SUBCLASS              As Long = &H10
'--- for LoadCursor
Private Const IDC_HAND                  As Long = 32649
'--- for GetSystemMetrics
Private Const SM_CXVSCROLL              As Long = 2
'--- for subclassing
Private Const WM_SYSCOLORCHANGE         As Long = &H15
Private Const WM_SETTINGCHANGE          As Long = &H1A
Private Const WM_CANCELMODE             As Long = &H1F
Private Const WM_COMMAND                As Long = &H111
Private Const WM_MOUSEWHEEL             As Long = &H20A
Private Const WM_MOUSEHOVER             As Long = &H2A1
Private Const WM_MOUSELEAVE             As Long = &H2A3
'--- for GetWindowLong
Private Const GWL_STYLE                 As Long = (-16)
Private Const WS_CAPTION                As Long = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
'--- for SetStretchBltMode
Private Const HALFTONE                  As Long = 4
'--- for hooked multi-line TextBox creation
Private Const ES_MULTILINE              As Long = &H4&
'--- for textbox notifications
Private Const EN_KILLFOCUS              As Long = &H200
'--- for SystemParametersInfo
Private Const SPI_GETMOUSEHOVERWIDTH    As Long = &H62
Private Const SPI_GETMOUSEHOVERHEIGHT   As Long = &H64
Private Const SPI_GETMOUSEHOVERTIME     As Long = &H66

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ApiInitCommonControlsEx Lib "comctl32" Alias "InitCommonControlsEx" (ICCE As INITCOMMONCONTROLSEX) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function CopyCursor Lib "user32" Alias "CopyIcon" (ByVal hcur As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function TrackMouseEvent Lib "comctl32.dll" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSESTRUCT) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Type TRACKMOUSESTRUCT
    cbSize      As Long
    dwFlags     As TrackMouseEventFlags
    hwndTrack   As Long
    dwHoverTime As Long
End Type

Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Type TOOLINFO
    cbSize      As Long
    uFlags      As Long
    hwnd        As Long
    uId         As Long
    RECT        As RECT
    hinst       As Long
    lpszText    As Long
End Type

Private Type INITCOMMONCONTROLSEX
   dwSize       As Long
   dwICC        As Long
End Type

Private Type POINTAPI
    x           As Long
    y           As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const CAP_MSG               As String = "ButtonBar Control"
Private Const DEF_MASKCOLOR         As Long = &HFF00FF '--- magenta
Private Const DEF_ANIMSTEPS         As Long = 5
Private Const DEF_OLEDRAGMODE       As Long = ucsOleDragManual
Private Const DEF_OLEDROPMODE       As Long = ucsOleDropNone
Private Const DEF_SCROLLITEMSCOUNT  As Long = 2
Private Const DEF_USESYSTEMFONT     As Boolean = True
Private Const DEF_FLATSCROLLARROWS  As Boolean = True
Private Const DEF_WRAPTEXT          As Boolean = False
Private Const DEF_ALLOWGROUPDRAG    As Boolean = False
Private Const DEF_LABELEDIT         As Long = ucsLbeManual
Private Const STR_CONTROL           As String = "Control"
Private Const STR_GROUP             As String = "Group"
Private Const STR_ITEM              As String = "Item"
Private Const STR_ITEM_LARGE_ICONS  As String = "Item (in Large Icons)"
Private Const STR_SMALL_ICON        As String = "Small Icon"
Private Const STR_LARGE_ICON        As String = "Large Icon"
Private Const STR_HOVER             As String = "Hover"
Private Const STR_PRESSED           As String = "Pressed"
Private Const STR_SELECTED          As String = "Selected"
Private Const STR_VB_TEXTBOX        As String = "VB.TextBox"
Private Const STR_TEXTBOX_NAME      As String = "LabelEdit"
Private Const STR_VB_TIMER          As String = "VB.Timer"
Private Const STR_TIMER_NAME        As String = "LabelTimer"
Private Const INNER_BORDER          As Long = 4     '--- in px
Private Const ITEM_VERT_DISTANCE    As Long = 0     '--- in px
Private Const ANIM_SPEED            As Long = 10    '--- in milisecs
Private Const LABELEDIT_SPEED       As Long = 500   '--- in milisecs (1/2 sec)

Private WithEvents m_oFmtContrl As cFormatDef
Attribute m_oFmtContrl.VB_VarHelpID = -1
Private WithEvents m_oFmtGrpNrm As cFormatDef
Attribute m_oFmtGrpNrm.VB_VarHelpID = -1
Private WithEvents m_oFmtGrpHvr As cFormatDef
Attribute m_oFmtGrpHvr.VB_VarHelpID = -1
Private WithEvents m_oFmtGrpPrs As cFormatDef
Attribute m_oFmtGrpPrs.VB_VarHelpID = -1
Private WithEvents m_oFmtGrpSel As cFormatDef
Attribute m_oFmtGrpSel.VB_VarHelpID = -1
Private WithEvents m_oFmtItmNrm As cFormatDef
Attribute m_oFmtItmNrm.VB_VarHelpID = -1
Private WithEvents m_oFmtItmLrg As cFormatDef
Attribute m_oFmtItmLrg.VB_VarHelpID = -1
Private WithEvents m_oFmtItmHvr As cFormatDef
Attribute m_oFmtItmHvr.VB_VarHelpID = -1
Private WithEvents m_oFmtItmPrs As cFormatDef
Attribute m_oFmtItmPrs.VB_VarHelpID = -1
Private WithEvents m_oFmtItmSel As cFormatDef
Attribute m_oFmtItmSel.VB_VarHelpID = -1
Private WithEvents m_oFmtSIcNrm As cFormatDef
Attribute m_oFmtSIcNrm.VB_VarHelpID = -1
Private WithEvents m_oFmtSIcHvr As cFormatDef
Attribute m_oFmtSIcHvr.VB_VarHelpID = -1
Private WithEvents m_oFmtSIcPrs As cFormatDef
Attribute m_oFmtSIcPrs.VB_VarHelpID = -1
Private WithEvents m_oFmtSIcSel As cFormatDef
Attribute m_oFmtSIcSel.VB_VarHelpID = -1
Private WithEvents m_oFmtLIcNrm As cFormatDef
Attribute m_oFmtLIcNrm.VB_VarHelpID = -1
Private WithEvents m_oFmtLIcHvr As cFormatDef
Attribute m_oFmtLIcHvr.VB_VarHelpID = -1
Private WithEvents m_oFmtLIcPrs As cFormatDef
Attribute m_oFmtLIcPrs.VB_VarHelpID = -1
Private WithEvents m_oFmtLIcSel As cFormatDef
Attribute m_oFmtLIcSel.VB_VarHelpID = -1
Private m_oRendContrl           As cFormatDef
Attribute m_oRendContrl.VB_VarHelpID = -1
Private m_oRendGrpNrm           As cFormatDef
Attribute m_oRendGrpNrm.VB_VarHelpID = -1
Private m_oRendGrpHvr           As cFormatDef
Attribute m_oRendGrpHvr.VB_VarHelpID = -1
Private m_oRendGrpPrs           As cFormatDef
Attribute m_oRendGrpPrs.VB_VarHelpID = -1
Private m_oRendGrpSel           As cFormatDef
Attribute m_oRendGrpSel.VB_VarHelpID = -1
Private m_oRendItmNrm           As cFormatDef
Attribute m_oRendItmNrm.VB_VarHelpID = -1
Private m_oRendItmNrmLrg        As cFormatDef
Private m_oRendItmHvr           As cFormatDef
Attribute m_oRendItmHvr.VB_VarHelpID = -1
Private m_oRendItmHvrLrg        As cFormatDef
Private m_oRendItmPrs           As cFormatDef
Attribute m_oRendItmPrs.VB_VarHelpID = -1
Private m_oRendItmPrsLrg        As cFormatDef
Private m_oRendItmSel           As cFormatDef
Attribute m_oRendItmSel.VB_VarHelpID = -1
Private m_oRendItmSelLrg        As cFormatDef
Private m_oRendSIcNrm           As cFormatDef
Attribute m_oRendSIcNrm.VB_VarHelpID = -1
Private m_oRendSIcHvr           As cFormatDef
Attribute m_oRendSIcHvr.VB_VarHelpID = -1
Private m_oRendSIcPrs           As cFormatDef
Attribute m_oRendSIcPrs.VB_VarHelpID = -1
Private m_oRendSIcSel           As cFormatDef
Attribute m_oRendSIcSel.VB_VarHelpID = -1
Private m_oRendLIcNrm           As cFormatDef
Private m_oRendLIcHvr           As cFormatDef
Private m_oRendLIcPrs           As cFormatDef
Private m_oRendLIcSel           As cFormatDef
Private WithEvents m_oTop       As cButton
Attribute m_oTop.VB_VarHelpID = -1
Private WithEvents m_oFont      As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_oDrawFont             As StdFont
Private m_clrMask               As OLE_COLOR
Private m_oSelectedItem         As cButton
Private m_oSelectedGroup        As cButton
Private m_oOver                 As cButton
Private m_oPressed              As cButton
Private m_oClicked              As cButton
Private m_lItemHeight           As Long
Private m_lGroupHeight          As Long
Private m_lAnimationSteps       As Long
Private m_eOleDragMode          As UcsOleDragModeEnum
Private m_eOleDropMode          As UcsOleDropModeEnum
Private m_lGroupOffset          As Long
Private m_lScrollItemsCount     As Long
Private m_bUseSystemFont        As Boolean
Private m_hWndTooltip           As Long
Private m_lScrollBtnState       As Long
Private m_oHandIcon             As StdPicture
Private m_oSubclassTop          As cSubclassingThunk
Private m_oSubclassControl      As cSubclassingThunk
Private m_bFlatScrollArrows     As Boolean
Private m_lScrollArrowSize      As Long
Private m_bWrapText             As Boolean
Private m_bAllowGroupDrag       As Boolean
Private m_lDropHighlightIdx     As Long
Private m_lGroupHighlightIdx    As Long
Private m_eLabelEdit            As UcsLabelEditEnum
Private m_sDownX                As Single
Private m_sDownY                As Single
Private m_oOleDragged           As cButton
Private m_oLabelItem            As cButton
Private WithEvents m_oLabelEdit As VB.TextBox
Attribute m_oLabelEdit.VB_VarHelpID = -1
Private m_oLabelHook            As cHookingThunk
Private m_lLabelAlign           As UcsFormatHorAlignmentStyle
Private m_sLabelCaption         As String
Private m_sClickX               As Single
Private m_sClickY               As Single
Private m_sHoverWidth           As Single
Private m_sHoverHeight          As Single
Private m_lHoverTime            As Long
Private m_bInSet                As Boolean
#If DebugMode Then
    Private m_sDebugID          As String
#End If

Private Type UcsHsbColor
    Hue                 As Double
    Sat                 As Double
    Bri                 As Double
End Type

Private Type UcsRgbQuad
    R                   As Byte
    G                   As Byte
    B                   As Byte
    A                   As Byte
End Type

Private Enum UcsScrollArrowEnum
    ucsScrollArrowUp = &H1
    ucsScrollArrowDown = &H2
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub RaiseError(sFunc As String)
    PushError sFunc, MODULE_NAME
    PopRaiseError
End Sub

Private Function ShowError(sFunc As String) As VbMsgBoxResult
    PushError sFunc, MODULE_NAME
    ShowError = PopShowError(CAP_MSG)
End Function

'=========================================================================
' Properties
'=========================================================================

'= design-time ===========================================================

'Purpose: Returns or sets a <b>Font</b> object used in an <b>Outlook Bar</b> control.
Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns or sets a Font object used in an Outlook Bar control."
Attribute Font.VB_HelpID = 107
    Set Font = m_oFont
End Property

Private Property Get DrawFont() As StdFont
    If m_oDrawFont Is Nothing Then
        If UseSystemFont Then
            With New cMemDC
                Set m_oDrawFont = .SystemIconFont
            End With
        Else
            Set m_oDrawFont = Font
        End If
    End If
    Set DrawFont = m_oDrawFont
End Property

Property Set Font(ByVal oValue As StdFont)
    If Not oValue Is Nothing Then
        Set m_oFont = CloneFont(oValue)
        pvPropertyChanged
    End If
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of on <b>Outlook Bar</b> control.
Property Get FormatControl() As cFormatDef
Attribute FormatControl.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of on Outlook Bar control."
Attribute FormatControl.VB_HelpID = 108
    Set FormatControl = m_oFmtContrl
    Set FormatControl.ParentFont = DrawFont
End Property

Property Set FormatControl(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_CONTROL_FORMAT
    End If
    m_oFmtContrl.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of a group of an <b>Outlook Bar</b> control.
Property Get FormatGroup() As cFormatDef
Attribute FormatGroup.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of a group of an Outlook Bar control."
Attribute FormatGroup.VB_HelpID = 109
    Set FormatGroup = m_oFmtGrpNrm
End Property

Property Set FormatGroup(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_GROUP_FORMAT
    End If
    m_oFmtGrpNrm.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of a group of an <b>Outlook Bar</b> control when the mouse is hovering over it.
Property Get FormatGroupHover() As cFormatDef
Attribute FormatGroupHover.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of a group of an Outlook Bar control when the mouse is hovering over it."
Attribute FormatGroupHover.VB_HelpID = 110
    Set FormatGroupHover = m_oFmtGrpHvr
End Property

Property Set FormatGroupHover(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_GROUP_FORMAT_HOVER
    End If
    m_oFmtGrpHvr.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of a group of an <b>Outlook Bar</b> control when the left mouse button is pressed over it.
Property Get FormatGroupPressed() As cFormatDef
Attribute FormatGroupPressed.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of a group of an Outlook Bar control when the left mouse button is pressed over it."
Attribute FormatGroupPressed.VB_HelpID = 111
    Set FormatGroupPressed = m_oFmtGrpPrs
End Property

Property Set FormatGroupPressed(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_GROUP_FORMAT_PRESSED
    End If
    m_oFmtGrpPrs.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of a group of an <b>Outlook Bar</b> control when the group is selected.
Property Get FormatGroupSelected() As cFormatDef
Attribute FormatGroupSelected.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of a group of an Outlook Bar control when the group is selected."
Attribute FormatGroupSelected.VB_HelpID = 112
    Set FormatGroupSelected = m_oFmtGrpSel
End Property

Property Set FormatGroupSelected(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_GROUP_FORMAT_SELECTED
    End If
    m_oFmtGrpSel.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of an item of an <b>Outlook Bar</b> control.
Property Get FormatItem() As cFormatDef
Attribute FormatItem.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of an item of an Outlook Bar control."
Attribute FormatItem.VB_HelpID = 113
    Set FormatItem = m_oFmtItmNrm
End Property

Property Set FormatItem(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_ITEM_FORMAT
    End If
    m_oFmtItmNrm.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of an item of an <b>Outlook Bar</b> control when group is displaying large icons.
Property Get FormatItemLargeIcons() As cFormatDef
Attribute FormatItemLargeIcons.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of an item of an Outlook Bar control when group is displaying large icons."
Attribute FormatItemLargeIcons.VB_HelpID = 115
    Set FormatItemLargeIcons = m_oFmtItmLrg
End Property

Property Set FormatItemLargeIcons(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_ITEM_FORMAT_LARGE_ICONS
    End If
    m_oFmtItmLrg.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of an item of an <b>Outlook Bar</b> control when mouse is hovering over it.
Property Get FormatItemHover() As cFormatDef
Attribute FormatItemHover.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of an item of an Outlook Bar control when mouse is hovering over it."
Attribute FormatItemHover.VB_HelpID = 114
    Set FormatItemHover = m_oFmtItmHvr
End Property

Property Set FormatItemHover(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_ITEM_FORMAT_HOVER
    End If
    m_oFmtItmHvr.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of an item of an <b>Outlook Bar</b> control when the left mouse button is pressed over it.
Property Get FormatItemPressed() As cFormatDef
Attribute FormatItemPressed.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of an item of an Outlook Bar control when the left mouse button is pressed over it."
Attribute FormatItemPressed.VB_HelpID = 116
    Set FormatItemPressed = m_oFmtItmPrs
End Property

Property Set FormatItemPressed(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_ITEM_FORMAT_PRESSED
    End If
    m_oFmtItmPrs.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of an item of an <b>Outlook Bar</b> control when the item is selected.
Property Get FormatItemSelected() As cFormatDef
Attribute FormatItemSelected.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of an item of an Outlook Bar control when the item is selected."
Attribute FormatItemSelected.VB_HelpID = 117
    Set FormatItemSelected = m_oFmtItmSel
End Property

Property Set FormatItemSelected(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_ITEM_FORMAT_SELECTED
    End If
    m_oFmtItmSel.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of the small icon of an item.
Property Get FormatSmallIcon() As cFormatDef
Attribute FormatSmallIcon.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of the small icon of an item."
Attribute FormatSmallIcon.VB_HelpID = 122
    Set FormatSmallIcon = m_oFmtSIcNrm
End Property

Property Set FormatSmallIcon(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_SMALL_ICON_FORMAT
    End If
    m_oFmtSIcNrm.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of the small icon of an item when the mouse is hovering over it.
Property Get FormatSmallIconHover() As cFormatDef
Attribute FormatSmallIconHover.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of the small icon of an item when the mouse is hovering over it."
Attribute FormatSmallIconHover.VB_HelpID = 123
    Set FormatSmallIconHover = m_oFmtSIcHvr
End Property

Property Set FormatSmallIconHover(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_SMALL_ICON_FORMAT_HOVER
    End If
    m_oFmtSIcHvr.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of the small icon of an item when the left mouse button is pressed over it.
Property Get FormatSmallIconPressed() As cFormatDef
Attribute FormatSmallIconPressed.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of the small icon of an item when the left mouse button is pressed over it."
Attribute FormatSmallIconPressed.VB_HelpID = 124
    Set FormatSmallIconPressed = m_oFmtSIcPrs
End Property

Property Set FormatSmallIconPressed(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_SMALL_ICON_FORMAT_PRESSED
    End If
    m_oFmtSIcPrs.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of the small icon of an item when the item is selected.
Property Get FormatSmallIconSelected() As cFormatDef
Attribute FormatSmallIconSelected.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of the small icon of an item when the item is selected."
Attribute FormatSmallIconSelected.VB_HelpID = 125
    Set FormatSmallIconSelected = m_oFmtSIcSel
End Property

Property Set FormatSmallIconSelected(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_SMALL_ICON_FORMAT_SELECTED
    End If
    m_oFmtSIcSel.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of the large icon of an item.
Property Get FormatLargeIcon() As cFormatDef
Attribute FormatLargeIcon.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of the large icon of an item."
Attribute FormatLargeIcon.VB_HelpID = 118
    Set FormatLargeIcon = m_oFmtLIcNrm
End Property

Property Set FormatLargeIcon(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_LARGE_ICON_FORMAT
    End If
    m_oFmtLIcNrm.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of the large icon of an item when the mouse is hovering over it.
Property Get FormatLargeIconHover() As cFormatDef
Attribute FormatLargeIconHover.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of the large icon of an item when the mouse is hovering over it."
Attribute FormatLargeIconHover.VB_HelpID = 119
    Set FormatLargeIconHover = m_oFmtLIcHvr
End Property

Property Set FormatLargeIconHover(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_LARGE_ICON_FORMAT_HOVER
    End If
    m_oFmtLIcHvr.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of the large icon of an item when left mouse button is pressed over it.
Property Get FormatLargeIconPressed() As cFormatDef
Attribute FormatLargeIconPressed.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of the large icon of an item when left mouse button is pressed over it."
Attribute FormatLargeIconPressed.VB_HelpID = 120
    Set FormatLargeIconPressed = m_oFmtLIcPrs
End Property

Property Set FormatLargeIconPressed(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_LARGE_ICON_FORMAT_PRESSED
    End If
    m_oFmtLIcPrs.Contents = oValue.Contents
End Property

'Purpose: Returns or sets a <b>cFormatDef</b> object used to format the appearance of the large icon of an item when the item is selected.
Property Get FormatLargeIconSelected() As cFormatDef
Attribute FormatLargeIconSelected.VB_Description = "Returns or sets a cFormatDef object used to format the appearance of the large icon of an item when the item is selected."
Attribute FormatLargeIconSelected.VB_HelpID = 121
    Set FormatLargeIconSelected = m_oFmtLIcSel
End Property

Property Set FormatLargeIconSelected(ByVal oValue As cFormatDef)
    If oValue Is Nothing Then
        Set oValue = DEF_LARGE_ICON_FORMAT_SELECTED
    End If
    m_oFmtLIcSel.Contents = oValue.Contents
End Property

'Purpose: Returns or sets the color used to create masks for images in an <b>Outlook Bar</b> control.
Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns or sets the color used to create masks for images in an Outlook Bar control."
Attribute MaskColor.VB_HelpID = 131
    MaskColor = m_clrMask
End Property

Property Let MaskColor(ByVal clrValue As OLE_COLOR)
    m_clrMask = clrValue
    pvPropertyChanged
End Property

'Purpose: Returns or sets the number of steps the animation of group unfolding will take.
Property Get AnimationSteps() As Long
Attribute AnimationSteps.VB_Description = "Returns or sets the number of steps the animation of group unfolding will take."
Attribute AnimationSteps.VB_HelpID = 101
    AnimationSteps = m_lAnimationSteps
End Property

Property Let AnimationSteps(ByVal lValue As Long)
    m_lAnimationSteps = IIf(lValue < 1 Or lValue > 100, DEF_ANIMSTEPS, lValue)
    pvPropertyChanged
End Property

'Purpose: Returns or sets how an <b>Outlook Bar</b> control handles drag operations.
Property Get OleDragMode() As UcsOleDragModeEnum
Attribute OleDragMode.VB_Description = "Returns or sets how an Outlook Bar control handles drag operations."
Attribute OleDragMode.VB_HelpID = 138
    OleDragMode = m_eOleDragMode
End Property

Property Let OleDragMode(ByVal eValue As UcsOleDragModeEnum)
    m_eOleDragMode = eValue
    PropertyChanged
End Property

'Purpose: Returns or sets how an <b>Outlook Bar</b> control handles drop operations.
Property Get OleDropMode() As UcsOleDropModeEnum
Attribute OleDropMode.VB_Description = "Returns or sets how an Outlook Bar control handles drop operations."
Attribute OleDropMode.VB_HelpID = 140
    OleDropMode = m_eOleDropMode
End Property

Property Let OleDropMode(ByVal eValue As UcsOleDropModeEnum)
    m_eOleDropMode = eValue
    UserControl.OleDropMode = eValue
    PropertyChanged
End Property

'Purpose: Returns or sets the number of items that will be scrolled when item scroll arrows are clicked.
Property Get ScrollItemsCount() As Long
Attribute ScrollItemsCount.VB_Description = "Returns or sets the number of items that will be scrolled when item scroll arrows are clicked."
Attribute ScrollItemsCount.VB_HelpID = 145
    ScrollItemsCount = m_lScrollItemsCount
End Property

Property Let ScrollItemsCount(ByVal lValue As Long)
    m_lScrollItemsCount = IIf(lValue < 1 Or lValue > 100, DEF_SCROLLITEMSCOUNT, lValue)
    PropertyChanged
End Property

'Purpose: Returns or sets a value that determines whether <b>Outlook Bar</b> control default font will be built-in system font or custom one supplied by <b>Font</b> property.
Property Get UseSystemFont() As Boolean
Attribute UseSystemFont.VB_Description = "Returns or sets a value that determines whether Outlook Bar control default font will be built-in system font or custom one supplied by Font property."
Attribute UseSystemFont.VB_HelpID = 150
    UseSystemFont = m_bUseSystemFont
End Property

Property Let UseSystemFont(ByVal bValue As Boolean)
    m_bUseSystemFont = bValue
    pvPropertyChanged
End Property

'Purpose: Returns or sets a value that determines whether scroll arrows on groups are displayed using XP flat style or are normal ones.
Property Get FlatScrollArrows() As Boolean
Attribute FlatScrollArrows.VB_Description = "Returns or sets a value that determines whether scroll arrows on groups are displayed using XP flat style or are normal ones."
Attribute FlatScrollArrows.VB_HelpID = 154
    FlatScrollArrows = m_bFlatScrollArrows
End Property

Property Let FlatScrollArrows(ByVal bValue As Boolean)
    m_bFlatScrollArrows = bValue
    pvPropertyChanged
End Property

'Purpose: Returns or sets a value that determines whether captions of items are wrapped on multiple lines if bigger than control's width or terminated with ellipses.
Property Get WrapText() As Boolean
Attribute WrapText.VB_Description = "Returns or sets a value that determines whether captions of items are wrapped on multiple lines if bigger than control's width or terminated with ellipses."
Attribute WrapText.VB_HelpID = 152
    WrapText = m_bWrapText
End Property

Property Let WrapText(ByVal bValue As Boolean)
    m_bWrapText = bValue
    pvPropertyChanged
End Property

'Purpose: Returns or sets a value that determines whether group captions can be dragged when OLE dragging is enabled.
Property Get AllowGroupDrag() As Boolean
Attribute AllowGroupDrag.VB_Description = "Returns or sets a value that determines whether group captions can be dragged when OLE dragging is enabled."
Attribute AllowGroupDrag.VB_HelpID = 158
    AllowGroupDrag = m_bAllowGroupDrag
End Property

Property Let AllowGroupDrag(ByVal bValue As Boolean)
    m_bAllowGroupDrag = bValue
    pvPropertyChanged
End Property

'Purpose: Returns or sets a value that determines if a user can edit captions of <b>cButton</b> objects in an <b>Outlook Bar</b> control.
Property Get LabelEdit() As UcsLabelEditEnum
Attribute LabelEdit.VB_Description = "Returns or sets a value that determines if a user can edit captions of cButton objects in an Outlook Bar control."
Attribute LabelEdit.VB_HelpID = 170
    LabelEdit = m_eLabelEdit
End Property

Property Let LabelEdit(ByVal eValue As UcsLabelEditEnum)
    m_eLabelEdit = eValue
    pvPropertyChanged
End Property

'= run-time ==============================================================

'Purpose: Returns the <b>cButtons</b> collection in an <b>Outlook Bar</b> control.
Property Get Groups() As cButtons
Attribute Groups.VB_Description = "Returns the cButtons collection in an Outlook Bar control."
Attribute Groups.VB_HelpID = 126
    Set Groups = m_oTop.GroupItems
End Property

Property Set Groups(ByVal oValue As cButtons)
    m_oTop.GroupItems.Contents = oValue.Contents
    pvPropertyChanged
End Property

'Purpose: Returns or sets a reference to a selected <b>cButton</b> object representing a group.
Property Get SelectedGroup() As cButton
Attribute SelectedGroup.VB_Description = "Returns or sets a reference to a selected cButton object representing a group."
Attribute SelectedGroup.VB_HelpID = 146
    Set SelectedGroup = m_oSelectedGroup
End Property

Property Set SelectedGroup(ByVal oValue As cButton)
    If Not m_oSelectedGroup Is oValue Then
        Set m_oSelectedGroup = oValue
        RefreshControl
        RaiseEvent SelGroupChange
        RefreshControl
    End If
End Property

'Purpose: Returns or sets a reference to a selected <b>cButton</b> object representing an item.
Property Get SelectedItem() As cButton
Attribute SelectedItem.VB_Description = "Returns or sets a reference to a selected cButton object representing an item."
Attribute SelectedItem.VB_HelpID = 147
    Set SelectedItem = m_oSelectedItem
End Property

Property Set SelectedItem(ByVal oValue As cButton)
    If Not m_oSelectedItem Is oValue Then
        Set m_oSelectedItem = oValue
        RefreshControl
        RaiseEvent SelItemChange
        RefreshControl
    End If
End Property

'Purpose: Returns or sets the drop highlight insertion position in currently selected group.
Property Get DropHighlightIdx() As Long
Attribute DropHighlightIdx.VB_Description = "Returns or sets the drop highlight insertion position in currently selected group."
Attribute DropHighlightIdx.VB_HelpID = 153
Attribute DropHighlightIdx.VB_MemberFlags = "400"
    DropHighlightIdx = m_lDropHighlightIdx
End Property

Property Let DropHighlightIdx(ByVal lValue As Long)
    If m_lDropHighlightIdx <> lValue Then
        m_lDropHighlightIdx = lValue
        RefreshControl
    End If
End Property

'Purpose: Returns or sets the drop highlight insertion position of the dragged group.
Property Get GroupHighlightIdx() As Long
Attribute GroupHighlightIdx.VB_Description = "Returns or sets the drop highlight insertion position in currently selected group."
Attribute GroupHighlightIdx.VB_HelpID = 160
Attribute GroupHighlightIdx.VB_MemberFlags = "400"
    GroupHighlightIdx = m_lGroupHighlightIdx
End Property

Property Let GroupHighlightIdx(ByVal lValue As Long)
    If m_lGroupHighlightIdx <> lValue Then
        m_lGroupHighlightIdx = lValue
        RefreshControl
    End If
End Property

'= private ===============================================================

Private Property Get DEF_FONT() As StdFont
    Set DEF_FONT = New StdFont
    With DEF_FONT
        .Name = "Tahoma"
        .Size = 8
    End With
End Property

Private Property Get DEF_CONTROL_FORMAT() As cFormatDef
    Static oFmt         As cFormatDef
    
    If pvInitFormat(oFmt, STR_CONTROL, , ucsFbdSingle3D, ucsGrdAlphaBlend, ucsFhaCenter, ucsFvaMiddle, , vbButtonFace, vbWindowBackground, 214) Then
        With oFmt
            .ForeColor = vbWindowText
            .OffsetX = 0
            .OffsetY = 0
            .Padding = 2
            .BorderSunken = ucsTriFalse
            Set .ParentFont = DrawFont
        End With
    End If
    Set DEF_CONTROL_FORMAT = oFmt
End Property

Private Property Get DEF_GROUP_FORMAT() As cFormatDef
    Static oFmt         As cFormatDef
    
    Call pvInitFormat(oFmt, STR_GROUP, FormatControl, ucsFbdSingle3D, ucsGrdSolid, ucsFhaCenter)
    Set DEF_GROUP_FORMAT = oFmt
End Property

Private Property Get DEF_GROUP_FORMAT_HOVER() As cFormatDef
    Static oFmt         As cFormatDef
        
    Call pvInitFormat(oFmt, STR_HOVER, FormatGroup, ucsFbdDouble3D)
    Set DEF_GROUP_FORMAT_HOVER = oFmt
End Property

Private Property Get DEF_GROUP_FORMAT_PRESSED() As cFormatDef
    Static oFmt         As cFormatDef

    If pvInitFormat(oFmt, STR_PRESSED, FormatGroup, ucsFbdDouble3D) Then
        With oFmt
            .BorderSunken = ucsTriTrue
        End With
    End If
    Set DEF_GROUP_FORMAT_PRESSED = oFmt
End Property

Private Property Get DEF_GROUP_FORMAT_SELECTED() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_SELECTED, FormatGroup)
    Set DEF_GROUP_FORMAT_SELECTED = oFmt
End Property

Private Property Get DEF_ITEM_FORMAT() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_ITEM, FormatControl, ucsFbdNone, ucsGrdTransparent, ucsFhaLeft, ucsFvaMiddle)
    Set DEF_ITEM_FORMAT = oFmt
End Property

Private Property Get DEF_ITEM_FORMAT_LARGE_ICONS() As cFormatDef
    Static oFmt         As cFormatDef

    If pvInitFormat(oFmt, STR_ITEM_LARGE_ICONS, FormatControl, ucsFbdNone, ucsGrdTransparent, ucsFhaCenter, ucsFvaMiddle) Then
        oFmt.Padding = 7
    End If
    Set DEF_ITEM_FORMAT_LARGE_ICONS = oFmt
End Property

Private Property Get DEF_ITEM_FORMAT_HOVER() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_HOVER, FormatItem)
    Set DEF_ITEM_FORMAT_HOVER = oFmt
End Property

Private Property Get DEF_ITEM_FORMAT_PRESSED() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_PRESSED, FormatItem)
    Set DEF_ITEM_FORMAT_PRESSED = oFmt
End Property

Private Property Get DEF_ITEM_FORMAT_SELECTED() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_SELECTED, FormatItem)
    Set DEF_ITEM_FORMAT_SELECTED = oFmt
End Property

Private Property Get DEF_SMALL_ICON_FORMAT() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_SMALL_ICON, FormatControl, ucsFbdNone, ucsGrdTransparent, ucsFhaLeft, ucsFvaMiddle)
    Set DEF_SMALL_ICON_FORMAT = oFmt
End Property

Private Property Get DEF_SMALL_ICON_FORMAT_HOVER() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_HOVER, FormatSmallIcon, ucsFbdFixed, ucsGrdAlphaBlend, , , vbHighlight, vbHighlight, vbWindowBackground, 70)
    Set DEF_SMALL_ICON_FORMAT_HOVER = oFmt
End Property

Private Property Get DEF_SMALL_ICON_FORMAT_PRESSED() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_PRESSED, FormatSmallIcon, ucsFbdFixed, ucsGrdAlphaBlend, , , vbHighlight, vbHighlight, vbWindowBackground, 121)
    Set DEF_SMALL_ICON_FORMAT_PRESSED = oFmt
End Property

Private Property Get DEF_SMALL_ICON_FORMAT_SELECTED() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_SELECTED, FormatSmallIcon, ucsFbdFixed, ucsGrdAlphaBlend, , , vbHighlight, vbHighlight, vbWindowBackground, 70)
    Set DEF_SMALL_ICON_FORMAT_SELECTED = oFmt
End Property

Private Property Get DEF_LARGE_ICON_FORMAT() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_LARGE_ICON, FormatControl, ucsFbdNone, ucsGrdTransparent, ucsFhaCenter, ucsFvaTop)
    Set DEF_LARGE_ICON_FORMAT = oFmt
End Property

Private Property Get DEF_LARGE_ICON_FORMAT_HOVER() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_HOVER, FormatLargeIcon, ucsFbdFixed, ucsGrdAlphaBlend, , , vbHighlight, vbHighlight, vbWindowBackground, 70)
    Set DEF_LARGE_ICON_FORMAT_HOVER = oFmt
End Property

Private Property Get DEF_LARGE_ICON_FORMAT_PRESSED() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_PRESSED, FormatLargeIcon, ucsFbdFixed, ucsGrdAlphaBlend, , , vbHighlight, vbHighlight, vbWindowBackground, 121)
    Set DEF_LARGE_ICON_FORMAT_PRESSED = oFmt
End Property

Private Property Get DEF_LARGE_ICON_FORMAT_SELECTED() As cFormatDef
    Static oFmt         As cFormatDef

    Call pvInitFormat(oFmt, STR_SELECTED, FormatLargeIcon, ucsFbdFixed, ucsGrdAlphaBlend, , , vbHighlight, vbHighlight, vbWindowBackground, 70)
    Set DEF_LARGE_ICON_FORMAT_SELECTED = oFmt
End Property

Private Property Let ApiTooltipText(sValue As String)
    Static ti           As TOOLINFO
    Static sText        As String
    Dim hr              As Long
    Dim bNew            As Boolean
    
    If m_hWndTooltip = 0 Then
        Exit Property
    End If
    With ti
        bNew = (.cbSize = 0)
        .cbSize = LenB(ti)
        .uFlags = TTF_SUBCLASS
        .hwnd = UserControl.hwnd
        .uId = 0
        .hinst = App.hInstance
        sText = StrConv(sValue, vbFromUnicode)
        .lpszText = StrPtr(sText)
        With .RECT
            .Right = ScaleWidth \ Screen.TwipsPerPixelX
            .Bottom = ScaleHeight \ Screen.TwipsPerPixelY
        End With
    End With
    If bNew Then
        hr = SendMessage(m_hWndTooltip, TTM_ADDTOOL, 0, ti)
    Else
        If Len(sValue) > 0 Then
            hr = SendMessage(m_hWndTooltip, TTM_UPDATETIPTEXT, 0, ti)
            hr = SendMessage(m_hWndTooltip, TTM_NEWTOOLRECT, 0, ti)
        Else
            hr = SendMessage(m_hWndTooltip, TTM_DELTOOL, 0, ti)
            ti.cbSize = 0
        End If
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

'Purpose: Forces a complete repaint of an <b>Outlook Bar</b> control.
Public Sub RefreshControl()
Attribute RefreshControl.VB_Description = "Forces a complete repaint of an Outlook Bar control."
Attribute RefreshControl.VB_HelpID = 144
    Dim rc              As RECT
    
    On Error Resume Next
    '--- refresh UI
    AutoRedraw = False
    GetClientRect hwnd, rc
    InvalidateRect hwnd, rc, 1
End Sub

'Purpose: Returns a value that represents the part of an <b>Outlook Bar</b> control whose position matches the specified coordinates.
Public Function HitTest( _
            ByVal x As Single, _
            ByVal y As Single, _
            oBtn As cButton) As UcsHitTestEnum
Attribute HitTest.VB_Description = "Returns a value that represents the part of an Outlook Bar control whose position matches the specified coordinates."
Attribute HitTest.VB_HelpID = 127
    Const FUNC_NAME     As String = "HitTest"
    Dim lIdx            As Long
    Dim lTop            As Long
    Dim lBottom         As Long
    Dim lGroupHeight    As Long
    Dim lItemHeight     As Long
    Dim lWidth          As Long
    Dim oGroup          As cButtons
    
    On Error GoTo EH
    '--- check if outside control bounds
    If x < 0 Or x > ScaleWidth Or y < 0 Or y > ScaleHeight Then
        HitTest = ucsHitNoWhere
        Exit Function
    End If
    '--- if inside the pt is at least on the background
    HitTest = ucsHitBackground
    x = x \ Screen.TwipsPerPixelX
    y = y \ Screen.TwipsPerPixelY
    lWidth = ScaleWidth \ Screen.TwipsPerPixelX
    lGroupHeight = m_lGroupHeight
    lTop = 0
    lBottom = ScaleHeight \ Screen.TwipsPerPixelY
    For lIdx = 1 To Groups.Count
        If Groups(lIdx).Visible Then
            If x >= 1 And y >= lTop And x <= lWidth - 1 And y <= lTop + lGroupHeight Then
                Set oBtn = Groups(lIdx)
                HitTest = ucsHitGroupButton
                Exit Function
            End If
            lTop = lTop + lGroupHeight + 1
        End If
        If Groups(lIdx).Selected Then
            Exit For
        End If
    Next
    For lIdx = Groups.Count To lIdx Step -1
        If Groups(lIdx).Selected Then
            Exit For
        End If
        If Groups(lIdx).Visible Then
            lBottom = lBottom - lGroupHeight - 1
            If x >= 1 And y >= lBottom And x < lWidth - 1 And y <= lBottom + lGroupHeight Then
                Set oBtn = Groups(lIdx)
                HitTest = ucsHitGroupButton
                Exit Function
            End If
        End If
    Next
    If lBottom > lTop Then
        If lIdx >= 1 And lIdx <= Groups.Count Then
            lTop = lTop + INNER_BORDER
            lBottom = lBottom - INNER_BORDER
            If x >= INNER_BORDER And x < lWidth - INNER_BORDER And _
                    y >= lTop And y < lBottom Then
                '--- if inside -> the pt is at least on the group background
                HitTest = ucsHitGroupBackground
                pvGetItemInfo Groups(lIdx), lItemHeight, 0, 0
                Set oGroup = Groups(lIdx).GroupItems
                '--- if scroll arrows are potentailly visible and near the right edge
                If lTop + (lItemHeight + ITEM_VERT_DISTANCE) * oGroup.Count > lBottom _
                        And x >= lWidth - INNER_BORDER - m_lScrollArrowSize Then
                    '--- if upper scroll arrow visible and over it
                    If m_lGroupOffset > 0 And _
                            y < lTop + m_lScrollArrowSize Then
                        Set oBtn = Nothing
                        HitTest = ucsHitScrollupArrow
                        Exit Function
                    End If
                    '--- if lower scroll arrow visible and over it
                    If oGroup.Count * (lItemHeight + ITEM_VERT_DISTANCE) - m_lGroupOffset + 1 >= lBottom - lTop And _
                            y >= lBottom - m_lScrollArrowSize Then
                        Set oBtn = Nothing
                        HitTest = ucsHitScrolldownArrow
                        Exit Function
                    End If
                End If
                '--- check if on an item
                For lIdx = 1 To oGroup.Count
                    If oGroup(lIdx).Visible Then
                        If y + m_lGroupOffset >= lTop And y + m_lGroupOffset < lTop + lItemHeight Then
                            Set oBtn = oGroup(lIdx)
                            HitTest = ucsHitItemButton
                            Exit Function
                        End If
                        lTop = lTop + lItemHeight + ITEM_VERT_DISTANCE
                    End If
                Next
            End If
        End If
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'Purpose: Initiates an OLE drag/drop operation.
Public Sub OleDrag()
Attribute OleDrag.VB_Description = "Initiates an OLE drag/drop operation."
Attribute OleDrag.VB_HelpID = 136
    UserControl.OleDrag
End Sub

'Purpose: Ensures visibility of a particular item. If necessary, this method scrolls <b>Outlook Bar</b> control.
Public Sub EnsureVisible( _
            ByVal oBtn As cButton, _
            ByVal bAnimate As Boolean)
Attribute EnsureVisible.VB_Description = "Ensures visibility of a particular item. If necessary, this method scrolls Outlook Bar control."
Attribute EnsureVisible.VB_HelpID = 106
    Const FUNC_NAME     As String = "EnsureVisible"
    Dim lItemHeight     As Long
    Dim lGroupVisibleHeight As Long
    Dim lOffset         As Long
    Dim lGroupVisibleCount As Long
    Dim lBtnVisibleIdx  As Long
    Dim lIdx            As Long
    
    On Error GoTo EH
    '--- argument check
    If oBtn Is Nothing Then
        Exit Sub
    End If
    If oBtn.Class <> ucsBtnClassItem Then
        Exit Sub
    End If
    If oBtn.Parent Is Nothing Then
        Exit Sub
    End If
    '--- make sure button is a control item
    Set oBtn = Groups(oBtn.Parent.Index).GroupItems(oBtn.Index)
    If Not SelectedGroup Is oBtn.Parent Then
        Set SelectedGroup = oBtn.Parent
    End If
    '--- ensure formats are rendered
    pvRenderFormats
    '--- count visible groups
    For lIdx = 1 To Groups.Count
        If Groups(lIdx).Visible Then
            lGroupVisibleCount = lGroupVisibleCount + 1
        End If
    Next
    '--- count visible items before oBtn
    For lIdx = 1 To oBtn.Index
        If oBtn.Parent.GroupItems(lIdx).Visible Then
            lBtnVisibleIdx = lBtnVisibleIdx + 1
        End If
    Next
    '--- get more info for button's group
    pvGetItemInfo oBtn.Parent, lItemHeight, 0, 0
    lGroupVisibleHeight = ScaleHeight \ Screen.TwipsPerPixelY - lGroupVisibleCount * (m_lGroupHeight + 1) - 2 * (INNER_BORDER + m_oRendContrl.BorderSize) + 1
    '--- ensure enough offset
    lOffset = pvMax(lBtnVisibleIdx * lItemHeight - lGroupVisibleHeight + 1, 0)
    If m_lGroupOffset < lOffset Then
        If bAnimate Then
            pvAnimateItems lOffset
        Else
            m_lGroupOffset = lOffset
        End If
    End If
    lOffset = (lBtnVisibleIdx - 1) * (lItemHeight + ITEM_VERT_DISTANCE)
    If m_lGroupOffset > lOffset Then
        If bAnimate Then
            pvAnimateItems lOffset
        Else
            m_lGroupOffset = lOffset
        End If
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'Purpose: Enables a user to edit a caption of a <b>cButton</b> object.
Public Sub StartLabelEdit(Optional oBtn As cButton)
Attribute StartLabelEdit.VB_Description = "Enables a user to edit a caption of a cButton object."
Attribute StartLabelEdit.VB_HelpID = 171
    '--- finish if anything remains to be renamed
    If Not m_oLabelItem Is Nothing Then
        pvFinishLabelEdit m_oLabelEdit.Text
    End If
    '--- figure out which button to rename
    Set m_oLabelItem = oBtn
    If m_oLabelItem Is Nothing Then
        Set m_oLabelItem = m_oSelectedItem
    End If
    If m_oLabelItem Is Nothing Then
        Set m_oLabelItem = m_oSelectedGroup
    End If
    If m_oLabelItem Is Nothing Then
        Exit Sub
    End If
    '--- cant edit invisible buttons
    If m_oLabelItem.Visible = False Then
        Set m_oLabelItem = Nothing
        Exit Sub
    End If
    '--- mouse events should be captured by the edit control
    ReleaseCapture
    '--- turn off current tooltip
    ApiTooltipText = ""
    '--- store label caption
    m_sLabelCaption = m_oLabelItem.Caption
    m_oLabelItem.Caption = ""
    '--- scroll nicely
    EnsureVisible m_oLabelItem, True
    '--- scavange previous textbox control
    If Not m_oLabelEdit Is Nothing Then
        Controls.Remove m_oLabelEdit
        Set m_oLabelEdit = Nothing
    End If
    '--- figure out textbox alignment
    '--- note: ES_LEFT etc. constants do coincide by value with ucsFhaXXX enum
    If m_oLabelItem.Class = ucsBtnClassItem Then
        m_lLabelAlign = IIf(m_oSelectedGroup.IconsType = ucsIcsSmallIcons, m_oRendItmNrm.HorAlignment, m_oRendItmNrmLrg.HorAlignment)
    Else
        m_lLabelAlign = ucsFhaCenter
    End If
    '--- add mutli-line textbox control (by using CBT hook)
    Set m_oLabelHook = New cHookingThunk
    m_oLabelHook.Hook WH_CBT, Me
    Set m_oLabelEdit = Controls.Add(STR_VB_TEXTBOX, STR_TEXTBOX_NAME)
    Set m_oLabelHook = Nothing
    '--- set textbox properties
    With m_oLabelEdit
        Set .Font = CloneFont(DrawFont)
        Set UserControl.Font = CloneFont(DrawFont)
        .Appearance = 0
        .Text = m_sLabelCaption
        '--- fix position
        m_oLabelEdit_Change
        .ZOrder
        .SelStart = 0
        .SelLength = Len(m_oLabelEdit.Text)
        .Visible = True
        .SetFocus
    End With
End Sub

'= Private ===============================================================

Private Sub pvPropertyChanged()
    On Error Resume Next
    frGetMeasures
    RefreshControl
    UserControl.PropertyChanged
End Sub
    
Private Sub pvFixGroupOffset()
'--- fix current group offset
    Dim lItemHeight     As Long
    Dim lMaxOffset      As Long

    On Error Resume Next
    '--- state check
    If m_oSelectedGroup Is Nothing Then
        Exit Sub
    End If
    If m_lGroupOffset < 0 Then
        m_lGroupOffset = 0
    Else
        pvGetItemInfo m_oSelectedGroup, lItemHeight, 0, 0
        lMaxOffset = Groups.Count * (m_lGroupHeight + 1) + m_oSelectedGroup.GroupItems.Count * (lItemHeight + ITEM_VERT_DISTANCE) - (ScaleHeight \ Screen.TwipsPerPixelY) + 2 * INNER_BORDER + 2 * m_oRendContrl.BorderSize
        If m_lGroupOffset > lMaxOffset Then
            m_lGroupOffset = pvMax(lMaxOffset, 0)
        End If
    End If
End Sub

Private Sub pvPaintBorder( _
            ByVal oMemDC As cMemDC, _
            LeftX As Long, _
            TopY As Long, _
            RightX As Long, _
            BottomY As Long, _
            ByVal Border As UcsFormatBorderStyle, _
            Optional ByVal bSunken As Boolean = True, _
            Optional DblOffset As Long, _
            Optional clrBorder As OLE_COLOR)
    Const FUNC_NAME     As String = "pvPaintBorder"
    Dim lOffset         As Long
    Dim clrLeftTop      As OLE_COLOR
    Dim clrRightBottom  As OLE_COLOR
    
    With oMemDC
        .NormalizeRect LeftX, TopY, RightX, BottomY
        Select Case Border
        Case ucsFbdFixed
            .FrameRect LeftX, TopY, RightX, BottomY, clrBorder
            lOffset = 1
        Case ucsFbdSingle3D
'            .DrawEdge LeftX, TopY, RightX, BottomY, IIf(bSunken, BDR_SUNKENOUTER, BDR_RAISEDINNER)
            If bSunken Then
                clrLeftTop = clrBorder
                clrRightBottom = vb3DHighlight
            Else
                clrLeftTop = vb3DHighlight
                clrRightBottom = clrBorder
            End If
            .DrawLine LeftX, TopY, RightX, TopY, clrLeftTop
            .DrawLine LeftX, TopY, LeftX, BottomY, clrLeftTop
            .DrawLine LeftX, BottomY - 1, RightX, BottomY - 1, clrRightBottom
            .DrawLine RightX - 1, TopY, RightX - 1, BottomY, clrRightBottom
            lOffset = 1
        Case ucsFbdDouble3D
            .DrawEdge LeftX - DblOffset, TopY - DblOffset, RightX, BottomY, IIf(bSunken, EDGE_SUNKEN, EDGE_RAISED)
            lOffset = 2 - DblOffset
        End Select
    End With
    '--- deflate paint rect
    LeftX = LeftX + lOffset
    TopY = TopY + lOffset
    RightX = RightX - lOffset
    BottomY = BottomY - lOffset
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvPaintGroup( _
            ByVal oMemDC As cMemDC, _
            ByVal LeftX As Long, _
            ByVal TopY As Long, _
            ByVal RightX As Long, _
            ByVal BottomY As Long, _
            oBtn As cButton)
    Const FUNC_NAME     As String = "pvPaintGroup"
    Dim oFmt            As cFormatDef
    Dim lAlign          As Long
    Dim oIcon           As StdPicture
    Dim lIconX          As Long
    Dim lIconY          As Long
    Dim lIconW          As Long
    Dim lIconH          As Long
    
    On Error GoTo EH
    With oMemDC
        If m_oPressed Is oBtn And oBtn.Enabled Then
            If m_oOver Is oBtn Then
                Set oFmt = m_oRendGrpPrs
            Else
                Set oFmt = m_oRendGrpHvr
            End If
        ElseIf m_oPressed Is Nothing And m_oOver Is oBtn And oBtn.Enabled Then
            Set oFmt = m_oRendGrpHvr
        ElseIf m_oSelectedGroup Is oBtn Then
            Set oFmt = m_oRendGrpSel
        Else
            Set oFmt = m_oRendGrpNrm
        End If
        pvFillGradient oMemDC, LeftX, TopY, RightX, BottomY, oFmt.BackGradient
        pvPaintBorder oMemDC, LeftX, TopY, RightX, BottomY, oFmt.Border, oFmt.BorderSunken, 1, oFmt.BorderColor
        Set oIcon = IIf(oBtn.IconsType = ucsIcsLargeIcons, oBtn.LargeIcon, oBtn.SmallIcon)
        If Not oIcon Is Nothing Then
            lIconW = pvHM2Pix(oIcon.Width)
            lIconH = pvHM2Pix(oIcon.Height)
        End If
        Select Case oFmt.HorAlignment
        Case ucsFhaRight
            lIconX = RightX - 1 - lIconW
            RightX = RightX - lIconW - 2
        Case Else
            lIconX = LeftX + 1
            LeftX = LeftX + lIconW + 2
        End Select
        If Not oIcon Is Nothing Then
            lIconY = (TopY + BottomY - pvHM2Pix(oIcon.Height)) \ 2
            If oBtn.Enabled Then
                .PaintPicture oIcon, lIconX + oFmt.OffsetX, _
                        lIconY + oFmt.OffsetY, clrMask:=MaskColor
            Else
                .PaintDisabledPicture oIcon, lIconX + oFmt.OffsetX, _
                        lIconY + oFmt.OffsetY, clrMask:=MaskColor
            End If
        End If
        '--- always single line
        lAlign = Array(DT_LEFT, DT_CENTER, DT_RIGHT)(oFmt.HorAlignment) _
              Or Array(DT_TOP, DT_VCENTER, DT_BOTTOM)(oFmt.VertAlignment) _
              Or DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_WORDBREAK
        '--- draw text (enabled or disabled)
        LeftX = LeftX + oFmt.Padding + oFmt.OffsetX
        TopY = TopY + oFmt.Padding + oFmt.OffsetY
        RightX = RightX - oFmt.Padding + oFmt.OffsetX
        BottomY = BottomY - oFmt.Padding + oFmt.OffsetY
        Set .Font = oFmt.Font
        If oBtn.Enabled Then
            .ForeColor = oFmt.ForeColor
            .DrawText oBtn.Caption, LeftX, TopY, RightX, BottomY, lAlign
        Else
            .ForeColor = vb3DHighlight
            .DrawText oBtn.Caption, LeftX + 1, TopY + 1, RightX + 1, BottomY + 1, lAlign
            .ForeColor = vbButtonShadow
            .DrawText oBtn.Caption, LeftX, TopY, RightX, BottomY, lAlign
        End If
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvPaintItem( _
            ByVal oMemDC As cMemDC, _
            ByVal LeftX As Long, _
            ByVal TopY As Long, _
            ByVal RightX As Long, _
            ByVal BottomY As Long, _
            oBtn As cButton, _
            ByVal lIconW As Long, _
            ByVal lIconH As Long, _
            oIcon As StdPicture)
    Const FUNC_NAME     As String = "pvPaintItem"
    Dim oFmt            As cFormatDef
    Dim oFmtIcon        As cFormatDef
    Dim lAlign          As Long
    Dim lIconX          As Long
    Dim lIconY          As Long
    Dim bResizeNeeded   As Boolean
    Dim bSmall          As Boolean
    Dim lHeight         As Long
    
    On Error GoTo EH
    bSmall = oBtn.Parent.IconsType = ucsIcsSmallIcons
    With oMemDC
        If m_oPressed Is oBtn And oBtn.Enabled Then
            If m_oOver Is oBtn Then
                Set oFmt = IIf(bSmall, m_oRendItmPrs, m_oRendItmPrsLrg)
                Set oFmtIcon = IIf(bSmall, m_oRendSIcPrs, m_oRendLIcPrs)
            Else
                Set oFmt = IIf(bSmall, m_oRendItmHvr, m_oRendItmHvrLrg)
                Set oFmtIcon = IIf(bSmall, m_oRendSIcHvr, m_oRendLIcHvr)
            End If
        ElseIf m_oPressed Is Nothing And m_oOver Is oBtn And oBtn.Enabled Then
            Set oFmt = IIf(bSmall, m_oRendItmHvr, m_oRendItmHvrLrg)
            Set oFmtIcon = IIf(bSmall, m_oRendSIcHvr, m_oRendLIcHvr)
        ElseIf m_oSelectedItem Is oBtn Then
            Set oFmt = IIf(bSmall, m_oRendItmSel, m_oRendItmSelLrg)
            Set oFmtIcon = IIf(bSmall, m_oRendSIcSel, m_oRendLIcSel)
        Else
            Set oFmt = IIf(bSmall, m_oRendItmNrm, m_oRendItmNrmLrg)
            Set oFmtIcon = IIf(bSmall, m_oRendSIcNrm, m_oRendLIcNrm)
        End If
        
        pvFillGradient oMemDC, LeftX, TopY, RightX, BottomY, oFmt.BackGradient
        pvPaintBorder oMemDC, (LeftX), (TopY), (RightX), (BottomY), oFmt.Border, oFmt.BorderSunken, 0, oFmt.BorderColor
        bResizeNeeded = oFmt.VertAlignment <> oFmtIcon.VertAlignment And oFmt.HorAlignment = oFmtIcon.HorAlignment Or oFmtIcon.HorAlignment = ucsFhaCenter
        Select Case oFmtIcon.HorAlignment
        Case ucsFhaRight
            lIconX = RightX - oFmt.BorderSize - oFmt.Padding - oFmtIcon.Padding - lIconW
            If bResizeNeeded Or oFmt.VertAlignment = oFmtIcon.VertAlignment Then
                RightX = RightX - 2 * (oFmt.BorderSize + oFmt.Padding + oFmtIcon.Padding) - lIconW
            End If
        Case ucsFhaCenter
            lIconX = (RightX + LeftX - lIconW) \ 2
            '--- vert align will modify top or bottom
        Case Else ' ucsFhaLeft
            lIconX = LeftX + oFmt.BorderSize + oFmt.Padding + oFmtIcon.Padding
            If bResizeNeeded Or oFmt.VertAlignment = oFmtIcon.VertAlignment Then
                LeftX = LeftX + oFmt.BorderSize + 2 * oFmt.Padding + 2 * oFmtIcon.Padding + lIconW
            End If
        End Select
        Select Case oFmtIcon.VertAlignment
        Case ucsFvaBottom
            lIconY = BottomY - oFmt.BorderSize - oFmtIcon.Padding - lIconH
            If bResizeNeeded Then
                BottomY = BottomY - oFmt.BorderSize - lIconH - 2 * oFmtIcon.Padding
            End If
        Case ucsFvaMiddle
            lIconY = (BottomY + TopY - lIconH) \ 2
        Case Else ' ucsFvaTop
            lIconY = TopY + oFmt.BorderSize + oFmt.Padding + oFmtIcon.Padding
            If bResizeNeeded Then
                TopY = TopY + oFmt.BorderSize + lIconH + 2 * oFmtIcon.Padding  '+ oFmt.Padding
            End If
        End Select
        pvFillGradient oMemDC, lIconX - oFmtIcon.Padding, _
                lIconY - oFmtIcon.Padding, _
                lIconX + lIconW + oFmtIcon.Padding, _
                lIconY + lIconH + oFmtIcon.Padding, _
                oFmtIcon.BackGradient
        pvPaintBorder oMemDC, lIconX - oFmtIcon.Padding, _
                lIconY - oFmtIcon.Padding, _
                lIconX + lIconW + oFmtIcon.Padding, _
                lIconY + lIconH + oFmtIcon.Padding, oFmtIcon.Border, oFmtIcon.BorderSunken, 0, oFmtIcon.BorderColor
        If Not oIcon Is Nothing Then
            If oBtn.Enabled Then
                .PaintPicture oIcon, lIconX + oFmtIcon.OffsetX + (lIconW - pvHM2Pix(oIcon.Width)) \ 2, _
                        lIconY + oFmtIcon.OffsetY + (lIconH - pvHM2Pix(oIcon.Height)) \ 2, clrMask:=MaskColor
            Else
                .PaintDisabledPicture oIcon, lIconX + oFmtIcon.OffsetX + (lIconW - pvHM2Pix(oIcon.Width)) \ 2, _
                        lIconY + oFmtIcon.OffsetY + (lIconH - pvHM2Pix(oIcon.Height)) \ 2, clrMask:=MaskColor
            End If
        End If
        '--- draw text (enabled or disabled)
        LeftX = LeftX + oFmt.Padding + oFmt.BorderSize + oFmt.OffsetX
        TopY = TopY + oFmt.Padding + oFmt.BorderSize + oFmt.OffsetY
        RightX = RightX - oFmt.Padding - oFmt.BorderSize + oFmt.OffsetX
        BottomY = BottomY - oFmt.Padding - oFmt.BorderSize + oFmt.OffsetY
        '--- Hack: adhoc array, much like <C++> "0123456789"[Idx] </C++>
        lAlign = Array(DT_LEFT, DT_CENTER, DT_RIGHT)(oFmt.HorAlignment) _
              Or Array(DT_TOP, DT_VCENTER, DT_BOTTOM)(oFmt.VertAlignment)
        '--- check if multiline (or not)
        If InStr(oBtn.Caption, vbCrLf) > 0 Or WrapText Then
            lAlign = lAlign Or DT_WORD_ELLIPSIS Or DT_WORDBREAK
            '--- vert align not available if not DT_SINGLELINE so calc TopY manually
            .DrawText oBtn.Caption, (LeftX), (0), (RightX), lHeight, lAlign Or DT_CALCRECT
            Select Case oFmt.VertAlignment
            Case ucsFvaMiddle
                TopY = TopY + (BottomY - TopY - lHeight) \ 2
            Case ucsFvaBottom
                TopY = BottomY - lHeight
            End Select
        Else
            lAlign = lAlign Or DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_WORDBREAK
        End If
        Set .Font = oFmt.Font
        If oBtn.Enabled Then
            .ForeColor = oFmt.ForeColor
            .DrawText oBtn.Caption, LeftX, TopY, RightX, BottomY, lAlign
        Else
            .ForeColor = vb3DHighlight
            .DrawText oBtn.Caption, LeftX + 1, TopY + 1, RightX + 1, BottomY + 1, lAlign
            .ForeColor = vbButtonShadow
            .DrawText oBtn.Caption, LeftX, TopY, RightX, BottomY, lAlign
        End If
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvPaintArrow( _
            ByVal oMemDC As cMemDC, _
            ByVal LeftX As Long, _
            ByVal TopY As Long, _
            ByVal RightX As Long, _
            ByVal BottomY As Long, _
            ByVal lArrowType As UcsScrollArrowEnum, _
            ByVal eIconsType As UcsIconStyle)
    Const FUNC_NAME     As String = "pvPaintArrow"
    Dim oFmt            As cFormatDef
    
    On Error GoTo EH
    If Not m_bFlatScrollArrows Then
        '--- paint normal 3D arrow
        oMemDC.DrawFrameControl LeftX, TopY, RightX, BottomY, DFC_SCROLL, _
                    DFCS_SCROLLDOWN Or _
                    IIf((m_lScrollBtnState And lArrowType) <> 0, _
                                DFCS_FLAT Or DFCS_PUSHED, 0)
    Else
        '--- takes small/large hover/pressed format and draws the button
        If eIconsType = ucsIcsSmallIcons Then
            Set oFmt = IIf((m_lScrollBtnState And lArrowType) <> 0, m_oRendSIcPrs, m_oRendSIcHvr)
        Else
            Set oFmt = IIf((m_lScrollBtnState And lArrowType) <> 0, m_oRendLIcPrs, m_oRendLIcHvr)
        End If
        '--- fill background and draw edge
        pvFillGradient oMemDC, LeftX, TopY, RightX, BottomY, oFmt.BackGradient
        pvPaintBorder oMemDC, (LeftX), (TopY), (RightX), (BottomY), oFmt.Border, oFmt.BorderSunken, 0, oFmt.BorderColor
        '--- prepare to transparently blit arrow
        With New cMemDC
            .Init RightX - LeftX - 2, BottomY - TopY - 2
            .DrawFrameControl -1, -1, .Width + 1, .Height + 1, DFC_SCROLL, _
                        IIf(lArrowType = 1, DFCS_SCROLLUP, DFCS_SCROLLDOWN) Or DFCS_FLAT ' Or IIf((m_lScrollBtnState And lArrowType) <> 0, DFCS_PUSHED, 0)
            '--- if pushed -> invert image
            If (m_lScrollBtnState And lArrowType) <> 0 Then
                '--- invert
                .DrawMode = vbXorPen
                .Pen = .CreatePen(&HFFFFFF)
                .Rectangle clrFill:=&HFFFFFF
            End If
            '--- blit with mask color = pixel (0,0) color
            .TransBlt oMemDC.hDC, LeftX + 1, TopY + 1, clrMask:=.GetPixel(0, 0)
        End With
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvPaintDropHilight( _
            ByVal oMemDC As cMemDC, _
            ByVal lLeft As Long, _
            ByVal lTop As Long, _
            ByVal lRight As Long, _
            ByVal clrFill As OLE_COLOR)
    Const FUNC_NAME     As String = "pvPaintDropHilight"
    Dim lI              As Long
    
    On Error GoTo EH
    With oMemDC
        .FillRect lLeft, lTop, lRight, lTop + 1, vbBlack
        For lI = 0 To 4
            .DrawLine lLeft + lI, lTop - (4 - lI), lLeft + lI, lTop + (5 - lI), vbBlack
            .DrawLine lRight - lI - 1, lTop - (4 - lI), lRight - lI - 1, lTop + (5 - lI), vbBlack
        Next
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvPaintGroupItems( _
            ByVal oMemDC As cMemDC, _
            ByVal lLeft As Long, _
            ByVal lTop As Long, _
            ByVal lRight As Long, _
            ByVal lBottom As Long, _
            ByVal oGroup As cButton, _
            ByVal lOffsetTop As Long, _
            ByVal bShowArrows As Boolean, _
            ByVal bShowDropHilight As Boolean)
    Const FUNC_NAME     As String = "pvPaintGroupItems"
    Dim lItemHeight     As Long
    Dim lIconW          As Long
    Dim lIconH          As Long
    Dim oBtn            As cButton
    Dim lItmTop         As Long
    Dim lI              As Long
    Dim lY              As Long
    
    On Error GoTo EH
    pvGetItemInfo oGroup, lItemHeight, lIconW, lIconH
    With oMemDC
        lTop = lTop + INNER_BORDER
        lLeft = lLeft + INNER_BORDER
        lRight = lRight - INNER_BORDER
        lBottom = lBottom - INNER_BORDER
        If lTop < lBottom Then
            .SetClipRect lLeft, lTop, lRight, lBottom
            lItmTop = lTop - lOffsetTop
            For Each oBtn In oGroup.GroupItems
                If oBtn.Visible Then
                    If lItmTop + lItemHeight > lTop Then
                        pvPaintItem oMemDC, lLeft, lItmTop, lRight, lItmTop + lItemHeight, oBtn, _
                            lIconW, lIconH, IIf(oGroup.IconsType = ucsIcsSmallIcons, oBtn.SmallIcon, oBtn.LargeIcon)
                    End If
                    lItmTop = lItmTop + lItemHeight + ITEM_VERT_DISTANCE
                    '--- beyond bottom of clipping area
                    If lItmTop > lBottom Then
                        Exit For
                    End If
                End If
            Next
            If bShowDropHilight Then
                If DropHighlightIdx >= 0 And DropHighlightIdx <= oGroup.GroupItems.Count Then
                    lY = lTop - lOffsetTop + lItemHeight * DropHighlightIdx
                    pvPaintDropHilight oMemDC, lLeft, lY, lRight, vbBlack
                End If
            End If
            If bShowArrows Then
                '--- down arrow
                If lItmTop >= lBottom Then
                    pvPaintArrow oMemDC, lRight - m_lScrollArrowSize, lBottom - m_lScrollArrowSize, _
                        lRight, lBottom, ucsScrollArrowDown, oGroup.IconsType
                End If
                '--- up arrow
                If lOffsetTop > 0 Then
                    pvPaintArrow oMemDC, lRight - m_lScrollArrowSize, lTop, lRight, _
                        lTop + m_lScrollArrowSize, ucsScrollArrowUp, oGroup.IconsType
                End If
            End If
        End If
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvRenderFormats()
    If m_oRendContrl Is Nothing Then
        Set m_oRendContrl = m_oFmtContrl.Render
        Set m_oRendGrpNrm = m_oFmtGrpNrm.Render(m_oRendContrl)
        Set m_oRendGrpHvr = m_oFmtGrpHvr.Render(m_oRendGrpNrm)
        Set m_oRendGrpPrs = m_oFmtGrpPrs.Render(m_oRendGrpNrm)
        Set m_oRendGrpSel = m_oFmtGrpSel.Render(m_oRendGrpNrm)
        Set m_oRendItmNrm = m_oFmtItmNrm.Render(m_oRendContrl)
        Set m_oRendItmNrmLrg = m_oFmtItmLrg.Render(m_oRendContrl)
        Set m_oRendItmHvr = m_oFmtItmHvr.Render(m_oRendItmNrm)
        Set m_oRendItmHvrLrg = m_oFmtItmHvr.Render(m_oRendItmNrmLrg)
        Set m_oRendItmPrs = m_oFmtItmPrs.Render(m_oRendItmNrm)
        Set m_oRendItmPrsLrg = m_oFmtItmPrs.Render(m_oRendItmNrmLrg)
        Set m_oRendItmSel = m_oFmtItmSel.Render(m_oRendItmNrm)
        Set m_oRendItmSelLrg = m_oFmtItmSel.Render(m_oRendItmNrmLrg)
        Set m_oRendSIcNrm = m_oFmtSIcNrm.Render(m_oRendContrl)
        Set m_oRendSIcHvr = m_oFmtSIcHvr.Render(m_oRendSIcNrm)
        Set m_oRendSIcPrs = m_oFmtSIcPrs.Render(m_oRendSIcNrm)
        Set m_oRendSIcSel = m_oFmtSIcSel.Render(m_oRendSIcNrm)
        Set m_oRendLIcNrm = m_oFmtLIcNrm.Render(m_oRendContrl)
        Set m_oRendLIcHvr = m_oFmtLIcHvr.Render(m_oRendLIcNrm)
        Set m_oRendLIcPrs = m_oFmtLIcPrs.Render(m_oRendLIcNrm)
        Set m_oRendLIcSel = m_oFmtLIcSel.Render(m_oRendLIcNrm)
    End If
End Sub

Private Sub pvPaintControl( _
            Optional ByVal lStep As Long, _
            Optional ByVal oPrevSlected As cButton)
    Const FUNC_NAME     As String = "pvPaintControl"
    Dim oMemDC          As cMemDC
    Dim lGroupHeight    As Long
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lRight          As Long
    Dim lBottom         As Long
    Dim lMiddle         As Long
    Dim lIdx            As Long
    Dim lBtmIdx         As Long
    Dim oGroup          As cButtons
    Dim hDC             As Long
    Dim oBtn            As cButton
    Dim oIcon           As StdPicture
    Dim lGroupBorder    As Long
    Dim lVisibleAtTop   As Long
    Dim lVisibleInBetween As Long
    
    On Error GoTo EH
    '--- ensure formats are rendered
    pvRenderFormats
    '--- calc group height accounting for group icons
    lGroupBorder = pvMax(m_oRendGrpNrm.BorderSize + m_oRendGrpNrm.Padding, _
                pvMax(m_oRendGrpHvr.BorderSize + m_oRendGrpHvr.Padding, _
                pvMax(m_oRendGrpPrs.BorderSize + m_oRendGrpPrs.Padding, _
                pvMax(m_oRendGrpSel.BorderSize + m_oRendGrpSel.Padding, 0))))
    For Each oBtn In Groups
        Set oIcon = IIf(oBtn.IconsType = ucsIcsLargeIcons, oBtn.LargeIcon, oBtn.SmallIcon)
        If Not oIcon Is Nothing Then
            lIdx = pvHM2Pix(oIcon.Height) + 2 * lGroupBorder
            If m_lGroupHeight < lIdx Then
                m_lGroupHeight = lIdx
            End If
        End If
    Next
    lGroupHeight = m_lGroupHeight
    Set oMemDC = New cMemDC
    With oMemDC
        '--- init
        .Init ScaleWidth \ Screen.TwipsPerPixelX, ScaleHeight \ Screen.TwipsPerPixelY, , UserControl.hDC
        lRight = .Width
        lBottom = .Height
        '--- back and border
        pvPaintBorder oMemDC, lLeft, lTop, lRight, lBottom, m_oRendContrl.Border, True, 0, m_oRendContrl.BorderColor
        .FillRect lLeft, lTop, lRight, lBottom, Ambient.BackColor
        pvFillGradient oMemDC, lLeft, lTop, lRight, lBottom, m_oRendContrl.BackGradient
        '--- set dc parameters
        .SetClipRect lLeft, lTop, lRight, lBottom
        .BackStyle = BS_TRANSPARENT
        '--- loop groups on bottom
        For lIdx = Groups.Count To 2 Step -1
            If Groups(lIdx).Selected Or Groups(lIdx) Is oPrevSlected Then
                Exit For
            End If
            If Groups(lIdx).Visible Then
                lBottom = lBottom - lGroupHeight - 1
                pvFillGradient oMemDC, lLeft, lBottom, lRight, lBottom + 1, FormatControl.BackGradient
                pvPaintGroup oMemDC, lLeft, lBottom + 1, lRight, lBottom + lGroupHeight + 1, Groups(lIdx)
            End If
        Next
        lBtmIdx = lIdx
        '--- loop groups on top
        lTop = lTop - 1
        For lIdx = 1 To lBtmIdx
            If Groups(lIdx).Visible Then
                lVisibleAtTop = lVisibleAtTop + 1
                pvFillGradient oMemDC, lLeft, lTop, lRight, lTop + 1, FormatControl.BackGradient
                pvPaintGroup oMemDC, lLeft, lTop + 1, lRight, lTop + lGroupHeight + 1, Groups(lIdx)
                lTop = lTop + lGroupHeight + 1
            End If
            If Groups(lIdx).Selected Or Groups(lIdx) Is oPrevSlected Then
                Exit For
            End If
        Next
        If lIdx > lBtmIdx Then
            lIdx = lBtmIdx
        End If
        If lIdx < lBtmIdx Then
            For lMiddle = lIdx + 1 To lBtmIdx
                lVisibleInBetween = lVisibleInBetween + Abs(Groups(lMiddle).Visible)
            Next
            lMiddle = lTop + ((lBottom - lTop - lVisibleInBetween * (lGroupHeight + 1)) * lStep) \ (AnimationSteps + 1)
        Else
            lMiddle = lBottom
        End If
        '--- paint upper visible group items
        If lBottom > lTop Then
            If lIdx > 0 And lIdx <= Groups.Count Then
                pvPaintGroupItems oMemDC, lLeft, lTop, lRight, lMiddle, Groups(lIdx), lBottom - lMiddle - lVisibleInBetween * (lGroupHeight + 1) + m_lGroupOffset, lStep = 0, lIdx >= lBtmIdx
            End If
        End If
        If lIdx < lBtmIdx Then
            lTop = lMiddle
            .SetClipRect lLeft, lTop, lRight, lBottom
            For lIdx = lIdx + 1 To lBtmIdx
                If Groups(lIdx).Visible Then
                    pvPaintGroup oMemDC, lLeft, lTop + 1, lRight, lTop + lGroupHeight + 1, Groups(lIdx)
                    lTop = lTop + lGroupHeight + 1
                End If
            Next
            '--- paint lower visible group items
            If lBottom > lTop Then
                If lBtmIdx > 0 And lBtmIdx <= Groups.Count Then
                    pvPaintGroupItems oMemDC, lLeft, lTop, lRight, lBottom, Groups(lBtmIdx), 0, False, False
                End If
            End If
        ElseIf GroupHighlightIdx >= 0 Then
            .SetClipRect -1
            '--- paint group hilight
            lTop = (m_lGroupHeight + 1) * GroupHighlightIdx
            If GroupHighlightIdx >= lVisibleAtTop Then
                lTop = lTop + lBottom - (m_lGroupHeight + 1) * lVisibleAtTop
            End If
            pvPaintDropHilight oMemDC, lLeft, lTop, lRight, vbBlack
        End If
        .SetClipRect -1
        '--- blit to control dc
        If .IsMemoryDC Then
            .BitBlt UserControl.hDC
        End If
    End With
    If AutoRedraw Then
        Refresh
    End If
    Exit Sub
EH:
    Debug.Print MODULE_NAME & "." & FUNC_NAME & ": " & Err.Description & vbCrLf & Err.Source
    Resume Next
'    RaiseError FUNC_NAME
End Sub

Private Sub pvFillGradient( _
            ByVal oMemDC As cMemDC, _
            ByVal LeftX As Long, _
            ByVal TopY As Long, _
            ByVal RightX As Long, _
            ByVal BottomY As Long, _
            ByVal oGrad As cGradientDef)
    Const FUNC_NAME     As String = "pvFillGradient"
    Dim hsbColor        As UcsHsbColor
    Dim rgbColor        As UcsRgbQuad
    Dim rgbSecondColor  As UcsRgbQuad
  
    On Error GoTo EH
    With oMemDC
        Select Case oGrad.GradientType
        Case ucsGrdSolid
            .FillRect LeftX, TopY, RightX, BottomY, oGrad.Color
        Case ucsGrdVertical, ucsGrdHorizontal
            .FillGradient LeftX, TopY, RightX, BottomY, oGrad.Color, oGrad.SecondColor, oGrad.GradientType = ucsGrdVertical
        Case ucsGrdBlend
            .ForeColor = oGrad.Color
            .BackColor = oGrad.SecondColor
            .FillRect LeftX, TopY, RightX, BottomY, , .DotBrush
        Case ucsGrdColorOffset
            hsbColor = pvRGBToHSB(oGrad.Color)
            '--- offset hue
            hsbColor.Hue = hsbColor.Hue + oGrad.OffsetHue \ 100
            '--- calculate saturation
            If oGrad.PercentSaturation >= 0 Then
                hsbColor.Sat = hsbColor.Sat * (100 - oGrad.PercentSaturation) / 100
            Else
                hsbColor.Sat = 100 - ((100 - hsbColor.Sat) * (100 + oGrad.PercentSaturation) / 100)
            End If
            '--- calculate brightness
            If oGrad.PercentBrightness >= 0 Then
                hsbColor.Bri = hsbColor.Bri * (100 - oGrad.PercentBrightness) / 100
            Else
                hsbColor.Bri = 100 - ((100 - hsbColor.Bri) * (100 + oGrad.PercentBrightness) / 100)
            End If
            '--- do fill
            .FillRect LeftX, TopY, RightX, BottomY, pvHSBToRGB(hsbColor)
        Case ucsGrdAlphaBlend
            Call OleTranslateColor(oGrad.Color, 0, rgbColor)
            Call OleTranslateColor(oGrad.SecondColor, 0, rgbSecondColor)
            .FillRect LeftX, TopY, RightX, BottomY, RGB( _
                (rgbColor.R * oGrad.Alpha + rgbSecondColor.R * (255 - oGrad.Alpha)) \ 255, _
                (rgbColor.G * oGrad.Alpha + rgbSecondColor.G * (255 - oGrad.Alpha)) \ 255, _
                (rgbColor.B * oGrad.Alpha + rgbSecondColor.B * (255 - oGrad.Alpha)) \ 255)
        Case ucsGrdStretchBitmap
            If Not pvIsPictureEmpty(oGrad.Picture) Then
                With New cMemDC
                    .Init pvHM2Pix(oGrad.Picture.Width), pvHM2Pix(oGrad.Picture.Height)
                    oMemDC.StretchBlt .hDC, 0, 0, .Width, .Height, LeftX, TopY, RightX - LeftX, BottomY - TopY
                    .PaintPicture oGrad.Picture, clrMask:=MaskColor
                    Call SetStretchBltMode(.hDC, HALFTONE)
                    .StretchBlt oMemDC.hDC, LeftX, TopY, RightX - LeftX, BottomY - TopY
                End With
            Else
                .FillRect LeftX, TopY, RightX, BottomY, vbButtonFace
            End If
        Case ucsGrdTileBitmap
            If Not pvIsPictureEmpty(oGrad.Picture) Then
                Dim lX              As Long
                Dim lY              As Long
                Dim lWidth          As Long
                Dim lHeight         As Long
                Dim lLeftOffset     As Long
                Dim lTopOffset      As Long
                
                With New cMemDC
                    .Init pvHM2Pix(oGrad.Picture.Width), pvHM2Pix(oGrad.Picture.Height)
                    '--- if absolute tiling -> align LeftX and TopY
                    If oGrad.TileAbsolutePosition Then
                        lLeftOffset = LeftX Mod .Width
                        lTopOffset = TopY Mod .Height
                        LeftX = LeftX - lLeftOffset
                        TopY = TopY - lTopOffset
                    End If
                    For lY = 0 To (BottomY - TopY + .Height - 1) \ .Height - 1
                        For lX = 0 To (RightX - LeftX + .Width - 1) \ .Width - 1
                            lWidth = pvMin(.Width, RightX - LeftX - lX * .Width)
                            lHeight = pvMin(.Height, BottomY - TopY - lY * .Height)
                            oMemDC.BitBlt .hDC, 0, 0, lWidth, lHeight, LeftX + lX * .Width, TopY + lY * .Height
                            .PaintPicture oGrad.Picture, clrMask:=MaskColor
                            .BitBlt oMemDC.hDC, _
                                        LeftX + lX * .Width + Abs(lX = 0) * lLeftOffset, _
                                        TopY + lY * .Height + Abs(lY = 0) * lTopOffset, _
                                        lWidth - Abs(lX = 0) * lLeftOffset, _
                                        lHeight - Abs(lY = 0) * lTopOffset, _
                                        Abs(lX = 0) * lLeftOffset, _
                                        Abs(lY = 0) * lTopOffset
                        Next
                    Next
                End With
            Else
                .FillRect LeftX, TopY, RightX, BottomY, vbButtonFace
            End If
        End Select
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvGetItemInfo( _
            ByVal oGroup As cButton, _
            lItemHeight As Long, _
            lIconWidth As Long, _
            lIconHeight As Long)
    Const FUNC_NAME     As String = "pvGetItemInfo"
    Dim oIcnFmt         As cFormatDef
    Dim oBtn            As cButton
    Dim oIcon           As StdPicture
    Dim lIconBorder     As Long
    Dim lItemBorder     As Long
    Dim lItemWidth      As Long
    Dim lTextLines      As Long
    Dim lAlign          As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    
    On Error GoTo EH
    '--- state check
    If oGroup Is Nothing Then
        Exit Sub
    End If
    '--- ensure formats are rendered
    pvRenderFormats
    '--- figure out current icon format
    If oGroup.IconsType = ucsIcsSmallIcons Then
        Set oIcnFmt = m_oRendSIcNrm
    Else
        Set oIcnFmt = m_oRendLIcNrm
    End If
    '--- get icon max dimensions
    For Each oBtn In oGroup.GroupItems
        If oGroup.IconsType = ucsIcsSmallIcons Then
            Set oIcon = oBtn.SmallIcon
        Else
            Set oIcon = oBtn.LargeIcon
        End If
        If Not oIcon Is Nothing Then
            If oIcon.handle <> 0 Then
                lIconHeight = pvMax(lIconHeight, pvHM2Pix(oIcon.Height))
                lIconWidth = pvMax(lIconWidth, pvHM2Pix(oIcon.Width))
            End If
        End If
    Next
    '--- get max icon border size & max item border size
    If oGroup.IconsType = ucsIcsSmallIcons Then
        lIconBorder = pvMax(m_oRendSIcNrm.BorderSize + m_oRendSIcNrm.Padding, _
                    pvMax(m_oRendSIcHvr.BorderSize + m_oRendSIcHvr.Padding, _
                    pvMax(m_oRendSIcPrs.BorderSize + m_oRendSIcPrs.Padding, _
                    pvMax(m_oRendSIcSel.BorderSize + m_oRendSIcSel.Padding, lIconBorder))))
        lItemBorder = pvMax(m_oRendItmNrm.BorderSize + m_oRendItmNrm.Padding, _
                    pvMax(m_oRendItmHvr.BorderSize + m_oRendItmHvr.Padding, _
                    pvMax(m_oRendItmPrs.BorderSize + m_oRendItmPrs.Padding, _
                    pvMax(m_oRendItmSel.BorderSize + m_oRendItmSel.Padding, 0))))
    Else
        lIconBorder = pvMax(m_oRendLIcNrm.BorderSize + m_oRendLIcNrm.Padding, _
                    pvMax(m_oRendLIcHvr.BorderSize + m_oRendLIcHvr.Padding, _
                    pvMax(m_oRendLIcPrs.BorderSize + m_oRendLIcPrs.Padding, _
                    pvMax(m_oRendLIcSel.BorderSize + m_oRendLIcSel.Padding, lIconBorder))))
        lItemBorder = pvMax(m_oRendItmNrmLrg.BorderSize + m_oRendItmNrmLrg.Padding, _
                    pvMax(m_oRendItmHvrLrg.BorderSize + m_oRendItmHvrLrg.Padding, _
                    pvMax(m_oRendItmPrsLrg.BorderSize + m_oRendItmPrsLrg.Padding, _
                    pvMax(m_oRendItmSelLrg.BorderSize + m_oRendItmSelLrg.Padding, 0))))
    End If
    '--- Hack: adhoc array, much like <C++> "0123456789"[Idx] </C++>
    lAlign = Array(DT_LEFT, DT_CENTER, DT_RIGHT)(m_oRendItmNrm.HorAlignment) _
            Or Array(DT_TOP, DT_VCENTER, DT_BOTTOM)(m_oRendItmNrm.VertAlignment) _
            Or DT_WORD_ELLIPSIS Or DT_WORDBREAK
    '--- get text max lines
    With New cMemDC
        .Init ScaleWidth \ Screen.TwipsPerPixelX, ScaleHeight \ Screen.TwipsPerPixelY, , UserControl.hDC
        Set .Font = m_oRendItmNrm.Font
        '--- calc item width
        lItemWidth = .Width - 2 * (1 + INNER_BORDER + lItemBorder + lIconBorder) - 2
        If oIcnFmt.VertAlignment <> m_oRendItmNrm.VertAlignment And oIcnFmt.HorAlignment = m_oRendItmNrm.HorAlignment Or oIcnFmt.HorAlignment = ucsFhaCenter Then
        Else
            lItemWidth = lItemWidth - lIconWidth
        End If
        For Each oBtn In oGroup.GroupItems
            If InStr(oBtn.Caption, vbCrLf) > 0 Or WrapText Then
                '--- calc number of lines
                lWidth = lItemWidth
                .DrawText oBtn.Caption, 0, 0, lWidth, lHeight, lAlign Or DT_CALCRECT
                lTextLines = pvMax(lHeight \ .TextHeight("ABCDH"), lTextLines)
            Else
                lTextLines = pvMax(1, lTextLines)
            End If
        Next
    End With
    '--- calc item height based on composition of the icon
    If oIcnFmt.VertAlignment <> m_oRendItmNrm.VertAlignment And oIcnFmt.HorAlignment = m_oRendItmNrm.HorAlignment Or oIcnFmt.HorAlignment = ucsFhaCenter Then
        lItemHeight = lIconHeight + 2 * lIconBorder + lTextLines * m_lItemHeight + 2 * lItemBorder
    Else
        lItemHeight = pvMax(lIconHeight + 2 * lIconBorder, lTextLines * m_lItemHeight) + 2 * lItemBorder
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvLngUndef(ByVal lValue As Long) As Long
    On Error Resume Next
    If lValue = LNG_UNDEFINED Then
        pvLngUndef = 0
    Else
        pvLngUndef = lValue
    End If
End Function

Friend Sub frGetMeasures()
    Dim lValue          As Long
    
    On Error Resume Next
    '--- state check
    If m_bInSet Then
        Exit Sub
    End If
    Set m_oRendContrl = Nothing
    With New cMemDC
        .Init , , , 0
        Set .Font = FormatControl.Font
        m_lItemHeight = .TextHeight("ABCH") '+ 6 + 2 * pvLngUndef(m_oFmtContrl.Padding)
'        Set .Font = FormatGroup.Font
        m_lGroupHeight = .TextHeight("ABCH") + 4 + 2 * FormatGroup.Render.Padding
    End With
    Set m_oDrawFont = Nothing
    m_lScrollArrowSize = GetSystemMetrics(SM_CXVSCROLL)
    If SystemParametersInfo(SPI_GETMOUSEHOVERWIDTH, 0, lValue, 0) = 0 Then
        lValue = 4
    End If
    m_sHoverWidth = lValue * Screen.TwipsPerPixelX \ 2
    If SystemParametersInfo(SPI_GETMOUSEHOVERHEIGHT, 0, lValue, 0) = 0 Then
        lValue = 4
    End If
    m_sHoverHeight = lValue * Screen.TwipsPerPixelY \ 2
    If SystemParametersInfo(SPI_GETMOUSEHOVERTIME, 0, lValue, 0) = 0 Then
        lValue = 400
    End If
    m_lHoverTime = lValue * 2
End Sub

Private Sub pvAnimateToGroup(ByVal oNewGroup As cButton)
    Dim lStep           As Long
    Dim lTimer          As Long
    Dim oPrev           As cButton
    
    On Error Resume Next
    UpdateWindow UserControl.hwnd
    Set oPrev = SelectedGroup
    Set m_oSelectedGroup = oNewGroup
    RaiseEvent SelGroupChange
    m_lGroupOffset = 0
    If Not m_oSelectedItem Is Nothing Then
        If m_oSelectedGroup Is m_oSelectedItem.Parent Then
            EnsureVisible m_oSelectedItem, False
        End If
    End If
    For lStep = 1 To AnimationSteps
        lTimer = timeGetTime
        '--- figure out direction of the animation
        pvPaintControl IIf(oPrev.Index < oNewGroup.Index, AnimationSteps + 1 - lStep, lStep), oPrev
        Do While timeGetTime() < lTimer + ANIM_SPEED
            Call Sleep(1)
        Loop
    Next
    pvPaintControl 0, oNewGroup
End Sub

Private Sub pvAnimateItems(ByVal lToGroupOffset)
    Dim lStep           As Long
    Dim lGroupOffset    As Long
    Dim lTimer          As Long
    
    On Error Resume Next
    UpdateWindow UserControl.hwnd
    lGroupOffset = m_lGroupOffset
    For lStep = 1 To AnimationSteps
        lTimer = timeGetTime
        m_lGroupOffset = lGroupOffset + (lToGroupOffset - lGroupOffset) * lStep \ AnimationSteps
        pvFixGroupOffset
        pvPaintControl
        Do While timeGetTime() < lTimer + ANIM_SPEED
            Call Sleep(1)
        Loop
    Next
End Sub

Private Function pvInitFormat( _
            oFmt As cFormatDef, _
            sName As String, _
            Optional oParent As cFormatDef, _
            Optional eBorder As UcsFormatBorderStyle = ucsFbd_Undefined, _
            Optional eGradientType As UcsGradientType = ucsGrd_Undefined, _
            Optional eHorAlign As UcsFormatHorAlignmentStyle = ucsFha_Undefined, _
            Optional eVertAlign As UcsFormatVertAlignmentStyle = ucsFva_Undefined, _
            Optional clrBorder As OLE_COLOR = vbButtonShadow, _
            Optional clrBackColor As OLE_COLOR = vbButtonFace, _
            Optional clrSecondBackColor As OLE_COLOR = vbButtonFace, _
            Optional lAlpha As Long) As Boolean
    On Error Resume Next
    If oFmt Is Nothing Then
        Set oFmt = New cFormatDef
        With oFmt
            .Name = sName
            Set .ParentFmt = oParent
            .Border = eBorder
            .BorderColor = clrBorder
            .BackGradient.GradientType = eGradientType
            .BackGradient.Color = clrBackColor
            .BackGradient.SecondColor = clrSecondBackColor
            .BackGradient.Alpha = lAlpha
            .HorAlignment = eHorAlign
            .VertAlignment = eVertAlign
        End With
        '--- format just created
        pvInitFormat = True
    End If
End Function

Private Function pvFindNextEnabled(ByVal oBtn As cButton, Optional ByVal IncludeCurrent As Boolean) As cButton
    If Not (IncludeCurrent And oBtn.Enabled And oBtn.Visible) Then
        Do
            If oBtn.Index = oBtn.Parent.GroupItems.Count Then
                Exit Function
            End If
            Set oBtn = oBtn.Parent.GroupItems(oBtn.Index + 1)
        Loop While Not (oBtn.Enabled And oBtn.Visible)
    End If
    Set pvFindNextEnabled = oBtn
End Function

Private Function pvFindPrevEnabled(ByVal oBtn As cButton, Optional ByVal IncludeCurrent As Boolean) As cButton
    If Not (IncludeCurrent And oBtn.Enabled And oBtn.Visible) Then
        Do
            If oBtn.Index = 1 Then
                Exit Function
            End If
            Set oBtn = oBtn.Parent.GroupItems(oBtn.Index - 1)
        Loop While Not (oBtn.Enabled And oBtn.Visible)
    End If
    Set pvFindPrevEnabled = oBtn
End Function

Private Sub pvCreateTooltipWindow()
    On Error Resume Next
    '--- create the tooltip window
    m_hWndTooltip = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, vbNullString, WS_POPUP Or TTS_NOPREFIX Or TTS_ALWAYSTIP, 0, 0, 0, 0, UserControl.hwnd, 0, App.hInstance, ByVal 0)
    '--- make tooltips multi-line
    Call SendMessage(m_hWndTooltip, TTM_SETMAXTIPWIDTH, 0&, ByVal &H7FFF&)
End Sub

Private Function pvHSBToRGB(hsbColor As UcsHsbColor) As Long
'--- based on *cool* code by Branco Medeiros (http://www.geocities.com/branco_medeiros)
'--- Converts an HSB value to the RGB color model. Adapted from Java.awt.Color.java
    Dim nH              As Double
    Dim nS              As Double
    Dim nL              As Double
    Dim nF              As Double
    Dim nP              As Double
    Dim nQ              As Double
    Dim nT              As Double
    Dim lH              As Long
    Dim clrConv         As UcsRgbQuad

    With clrConv
        If hsbColor.Sat > 0 Then
            nH = hsbColor.Hue / 60
            nL = hsbColor.Bri / 100
            nS = hsbColor.Sat / 100
            lH = Int(nH)
            nF = nH - lH
            nP = nL * (1 - nS)
            nQ = nL * (1 - nS * nF)
            nT = nL * (1 - nS * (1 - nF))
            Select Case lH
            Case 0
                .R = nL * 255
                .G = nT * 255
                .B = nP * 255
            Case 1
                .R = nQ * 255
                .G = nL * 255
                .B = nP * 255
            Case 2
                .R = nP * 255
                .G = nL * 255
                .B = nT * 255
            Case 3
                .R = nP * 255
                .G = nQ * 255
                .B = nL * 255
            Case 4
                .R = nT * 255
                .G = nP * 255
                .B = nL * 255
            Case 5
                .R = nL * 255
                .G = nP * 255
                .B = nQ * 255
            End Select
        Else
            .R = (hsbColor.Bri * 255) / 100
            .G = .R
            .B = .R
        End If
    End With
    '--- return long
    CopyMemory lH, clrConv, 4
    pvHSBToRGB = lH
End Function

Private Function pvRGBToHSB(ByVal clrValue As OLE_COLOR) As UcsHsbColor
'--- based on *cool* code by Branco Medeiros (http://www.geocities.com/branco_medeiros)
'--- Converts an RGB value to the HSB color model. Adapted from Java.awt.Color.java
    Dim nTemp           As Double
    Dim lMin            As Long
    Dim lMax            As Long
    Dim lDelta          As Long
    Dim rgbValue        As UcsRgbQuad
  
    Call OleTranslateColor(clrValue, 0, rgbValue)
    lMax = pvMax(pvMax(rgbValue.R, rgbValue.G), rgbValue.B)
    lMin = pvMin(pvMin(rgbValue.R, rgbValue.G), rgbValue.B)
    lDelta = lMax - lMin
    pvRGBToHSB.Bri = (lMax * 100) / 255
    If lMax > 0 Then
        pvRGBToHSB.Sat = (lDelta / lMax) * 100
        If lDelta > 0 Then
            If lMax = rgbValue.R Then
                nTemp = (CLng(rgbValue.G) - rgbValue.B) / lDelta
            ElseIf lMax = rgbValue.G Then
                nTemp = 2 + (CLng(rgbValue.B) - rgbValue.R) / lDelta
            Else
                nTemp = 4 + (CLng(rgbValue.R) - rgbValue.G) / lDelta
            End If
            pvRGBToHSB.Hue = nTemp * 60
            If pvRGBToHSB.Hue < 0 Then
                pvRGBToHSB.Hue = pvRGBToHSB.Hue + 360
            End If
        End If
    End If
End Function

Private Function pvHM2Pix(ByVal Value As Double) As Double
   pvHM2Pix = Value * 1440 / 2540 / Screen.TwipsPerPixelX
End Function

Private Function pvMax(ByVal lA As Long, ByVal lB As Long) As Long
    pvMax = Abs(lB > lA) * lB + Abs(Not lB > lA) * lA
End Function

Private Function pvMin(ByVal lA As Long, ByVal lB As Long) As Long
    pvMin = Abs(lB < lA) * lB + Abs(Not lB < lA) * lA
End Function

Private Sub pvSubclass()
    Dim hwndTopMost     As Long
    
    If Ambient.UserMode Then
        hwndTopMost = UserControl.hwnd
        Do While (GetWindowLong(hwndTopMost, GWL_STYLE) And WS_CAPTION) = 0 And hwndTopMost <> 0
            hwndTopMost = GetParent(hwndTopMost)
        Loop
        If hwndTopMost <> 0 Then
            Set m_oSubclassTop = New cSubclassingThunk
            With m_oSubclassTop
                .AddBeforeMsgs WM_SETTINGCHANGE ' , WM_MOUSEWHEEL
                .Subclass hwndTopMost, Me
            End With
        End If
        Set m_oSubclassControl = New cSubclassingThunk
        With m_oSubclassControl
            .AddBeforeMsgs WM_SYSCOLORCHANGE, WM_CANCELMODE, MouseWheelFwdMsg, WM_COMMAND, WM_MOUSELEAVE, WM_MOUSEHOVER
            .Subclass UserControl.hwnd, Me
        End With
    End If
End Sub

Private Sub pvBeginLabelEdit()
    Dim bCancel         As Boolean
    
    '--- first check if anything got capture (like popup menu)
    If GetCapture() = 0 Then
        RaiseEvent BeforeLabelEdit(bCancel)
        If Not bCancel Then
            StartLabelEdit m_oClicked
        End If
    End If
End Sub

Private Sub pvFinishLabelEdit(sText As String)
    Dim bCancel As Boolean
    
    On Error Resume Next
    If Not m_oLabelEdit Is Nothing Then
        bCancel = Not m_oLabelEdit.Visible
        m_oLabelEdit.Visible = False
        '--- BUG: hanging-up in MS Access
'        Controls.Remove m_oLabelEdit
'        '--- a constituent control can't be unloaded on UserControl_Resize!!!
'        '---    "Unable to unload within this context"
'        If Err.Number = 0 Then
'            Set m_oLabelEdit = Nothing
'        End If
        If bCancel Then
            Exit Sub
        End If
    End If
    If sText <> m_sLabelCaption Then
        RaiseEvent AfterLabelEdit(bCancel, sText)
    End If
    If Not bCancel Then
        m_oLabelItem.Caption = sText
    End If
    Set m_oLabelItem = Nothing
End Sub

Private Function pvIsPictureEmpty(ByVal oPic As StdPicture) As Boolean
    If Not oPic Is Nothing Then
        If oPic.handle <> 0 Then
            Exit Function
        End If
    End If
    '--- empty
    pvIsPictureEmpty = True
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub UserControl_Initialize()
    Dim icc             As INITCOMMONCONTROLSEX
    
    On Error Resume Next
    Set m_oTop = New cButton
    m_oTop.Class = ucsBtnClassControl
    m_sDownX = -1
    m_sClickX = -1
    m_lDropHighlightIdx = -1
    m_lGroupHighlightIdx = -1
    Set m_oFont = DEF_FONT
    Set m_oFmtContrl = DEF_CONTROL_FORMAT
    Set m_oFmtGrpNrm = DEF_GROUP_FORMAT
    Set m_oFmtGrpHvr = DEF_GROUP_FORMAT_HOVER
    Set m_oFmtGrpPrs = DEF_GROUP_FORMAT_PRESSED
    Set m_oFmtGrpSel = DEF_GROUP_FORMAT_SELECTED
    Set m_oFmtItmNrm = DEF_ITEM_FORMAT
    Set m_oFmtItmLrg = DEF_ITEM_FORMAT_LARGE_ICONS
    Set m_oFmtItmHvr = DEF_ITEM_FORMAT_HOVER
    Set m_oFmtItmPrs = DEF_ITEM_FORMAT_PRESSED
    Set m_oFmtItmSel = DEF_ITEM_FORMAT_SELECTED
    Set m_oFmtSIcNrm = DEF_SMALL_ICON_FORMAT
    Set m_oFmtSIcHvr = DEF_SMALL_ICON_FORMAT_HOVER
    Set m_oFmtSIcPrs = DEF_SMALL_ICON_FORMAT_PRESSED
    Set m_oFmtSIcSel = DEF_SMALL_ICON_FORMAT_SELECTED
    Set m_oFmtLIcNrm = DEF_LARGE_ICON_FORMAT
    Set m_oFmtLIcHvr = DEF_LARGE_ICON_FORMAT_HOVER
    Set m_oFmtLIcPrs = DEF_LARGE_ICON_FORMAT_PRESSED
    Set m_oFmtLIcSel = DEF_LARGE_ICON_FORMAT_SELECTED
    With New cMemDC
        Set m_oHandIcon = .IconToPicture(CopyCursor(LoadCursor(0, IDC_HAND)))
    End With
    m_bFlatScrollArrows = True
    '--- initialize common controls
    icc.dwSize = LenB(icc)
    icc.dwICC = ICC_TAB_CLASSES
    Call ApiInitCommonControlsEx(icc)
    '--- initialize global mouse wheel hook
    If g_oWheelHook Is Nothing Then
        Set g_oWheelHook = New cWheelHook
    End If
    #If DebugMode Then
        DebugInit m_sDebugID, MODULE_NAME
    #End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
    If Not m_oOver Is Nothing Then
        RaiseEvent ButtonDblClick(m_oOver)
    End If
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Const FUNC_NAME     As String = "UserControl_KeyDown"
    Dim oGrp            As cButton
    Dim oBtn            As cButton
    Dim lKey            As Long
    Dim lShift          As Long
    Dim lItemHeight     As Long
    
    On Error GoTo EH
    '--- give user a chance to modify keys
    RaiseEvent KeyDown(KeyCode, Shift)
    '--- state check
    If m_oSelectedGroup Is Nothing Then
        Exit Sub
    End If
    '--- don't touch byref arguments
    lKey = KeyCode
    lShift = Shift
    '--- translate pageup/pagedown to group movement
    If lKey = vbKeyPageDown Then
        lKey = vbKeyDown: lShift = vbCtrlMask
    ElseIf lKey = vbKeyPageUp Then
        lKey = vbKeyUp: lShift = vbCtrlMask
    End If
'    lKey = 1 / 2
    If lKey = vbKeyUp Then
        If lShift = 0 Then
            If m_oSelectedItem Is Nothing Then
                If Not m_oSelectedGroup Is Nothing Then
                    If m_oSelectedGroup.GroupItems.Count > 0 Then
                        Set oBtn = m_oSelectedGroup.GroupItems(1)
                    End If
                End If
            Else
                '--- find prev enabled
                Set oBtn = pvFindPrevEnabled(m_oSelectedItem)
                If oBtn Is Nothing Then
                    '--- fall-through to group movement
                    lShift = vbCtrlMask
                End If
            End If
        End If
        If lShift = vbCtrlMask And Not m_oSelectedGroup Is Nothing Then
            Set oGrp = pvFindPrevEnabled(m_oSelectedGroup)
            If Not oGrp Is Nothing Then
                oGrp.Selected = True
                If m_oSelectedGroup.GroupItems.Count > 0 Then
                    '--- find last enabled element
                    Set oBtn = pvFindPrevEnabled(m_oSelectedGroup.GroupItems(m_oSelectedGroup.GroupItems.Count), True)
                    If Not oBtn Is Nothing Then
                        oBtn.Selected = True
                        '--- position offset to oBtn.Index (no animation)
                        pvGetItemInfo oBtn.Parent, lItemHeight, 0, 0
                        m_lGroupOffset = oBtn.Parent.GroupItems.Count * (lItemHeight + ITEM_VERT_DISTANCE)
                    End If
                End If
            Else
                If m_oSelectedGroup.GroupItems.Count > 0 Then
                    '--- find first enabled element
                    Set oBtn = pvFindNextEnabled(m_oSelectedGroup.GroupItems(1), True)
                End If
            End If
        End If
    ElseIf lKey = vbKeyDown Then
        If lShift = 0 Then
            If m_oSelectedItem Is Nothing Then
                If Not m_oSelectedGroup Is Nothing Then
                    If m_oSelectedGroup.GroupItems.Count > 0 Then
                        Set oBtn = m_oSelectedGroup.GroupItems(1)
                    End If
                End If
            Else
                '--- find next enabled
                Set oBtn = pvFindNextEnabled(m_oSelectedItem)
                If oBtn Is Nothing Then
                    '--- fall-through to group movement
                    lShift = vbCtrlMask
                End If
            End If
        End If
        If lShift = vbCtrlMask And Not m_oSelectedGroup Is Nothing Then
            Set oGrp = pvFindNextEnabled(m_oSelectedGroup)
            If Not oGrp Is Nothing Then
                oGrp.Selected = True
                If m_oSelectedGroup.GroupItems.Count > 0 Then
                    '--- find first enabled element
                    Set oBtn = pvFindNextEnabled(m_oSelectedGroup.GroupItems(1), True)
                    If Not oBtn Is Nothing Then
                        oBtn.Selected = True
                        '--- position offset to beginning (no animation)
                        m_lGroupOffset = 0
                    End If
                End If
            Else
                If m_oSelectedGroup.GroupItems.Count > 0 Then
                    '--- find last enabled element
                    Set oBtn = pvFindPrevEnabled(m_oSelectedGroup.GroupItems(m_oSelectedGroup.GroupItems.Count), True)
                End If
            End If
        End If
    Else
        Exit Sub
    End If
    If Not oBtn Is Nothing Then
        oBtn.Selected = True
        EnsureVisible oBtn, True
    Else
        '--- dont draw scrollup arrow :-))
        m_lGroupOffset = 0
    End If
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lItemHeight     As Long
    Dim eHit            As UcsHitTestEnum
    
    On Error Resume Next
    If Not m_oLabelItem Is Nothing Then
        pvFinishLabelEdit m_oLabelEdit.Text
    End If
    m_sDownX = x: m_sDownY = y
    RaiseEvent MouseDown(Button, Shift, x, y)
    If (Button And vbLeftButton) <> 0 Then
        eHit = HitTest(x, y, m_oPressed)
        If m_oPressed.Enabled Then
            Set m_oOver = m_oPressed
            Select Case eHit
            Case ucsHitScrollupArrow, ucsHitScrolldownArrow
                m_lScrollBtnState = IIf(eHit = ucsHitScrollupArrow, ucsScrollArrowUp, ucsScrollArrowDown)
                pvGetItemInfo m_oSelectedGroup, lItemHeight, 0, 0
                pvAnimateItems m_lGroupOffset + IIf(eHit = ucsHitScrollupArrow, -ScrollItemsCount, ScrollItemsCount) * (lItemHeight + ITEM_VERT_DISTANCE)
                m_lScrollBtnState = 0
            Case Else
                RefreshControl
            End Select
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim uTrackMouse         As TRACKMOUSESTRUCT
    
    On Error Resume Next
    If Not m_oLabelItem Is Nothing Then
        pvFinishLabelEdit m_oLabelEdit.Text
    End If
    m_sDownX = -1
    RaiseEvent MouseUp(Button, Shift, x, y)
    If (Button And vbLeftButton) <> 0 Then
        Call HitTest(x, y, m_oOver)
        If Not m_oOver Is Nothing Then
            If m_oOver.Enabled And m_oOver Is m_oPressed Then
                RaiseEvent ButtonClick(m_oOver)
                '--- check for labeledit
                If m_oClicked Is m_oPressed Or m_oPressed Is m_oSelectedGroup Then
                    Set m_oClicked = m_oPressed
                    If LabelEdit = ucsLbeAutomatic Then
                        m_sClickX = x
                        m_sClickY = y
                        With uTrackMouse
                            .cbSize = Len(uTrackMouse)
                            .dwFlags = TME_HOVER
                            .dwHoverTime = m_lHoverTime
                            .hwndTrack = UserControl.hwnd
                        End With
                        TrackMouseEvent uTrackMouse
                    End If
                Else
                    Set m_oClicked = m_oPressed
                End If
                Set m_oPressed = Nothing
                If m_oOver.Class = ucsBtnClassItem Then
                    Set SelectedItem = m_oOver
                    EnsureVisible SelectedItem, True
                Else
                    '--- if selected group changed
                    If Not m_oSelectedGroup Is m_oOver Then
                        pvAnimateToGroup m_oOver
                    End If
                End If
            Else '--- m_oOver.Enabled And m_oOver Is m_oPressed
                Set m_oPressed = Nothing
            End If
        End If
        RefreshControl
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim oBtn            As cButton
    Dim eDropMode       As UcsOleDropModeEnum
    Dim uTrackMouse     As TRACKMOUSESTRUCT
    
    On Error Resume Next
    '--- test if waiting hover
    If m_sClickX >= 0 Then
        If Abs(m_sClickX - x) > m_sHoverWidth Or Abs(m_sClickY - y) > m_sHoverHeight Then
            m_sClickX = -1
        End If
    End If
    '--- test for ole dragging
    If m_sDownX >= 0 Then
        If Abs(m_sDownX - x) > ScaleX(2, vbPixels) Or Abs(m_sDownY - y) > ScaleY(2, vbPixels) Then
            RaiseEvent MouseDragged(Button, Shift, m_sDownX, m_sDownY)
            m_sDownX = -1
            If OleDragMode = ucsOleDragAutomatic Then
                If Not m_oPressed Is Nothing Then
                    '--- if group check for AllowGroupDrag
                    If m_oPressed.Class = ucsBtnClassGroup And AllowGroupDrag _
                            Or m_oPressed.Class = ucsBtnClassItem Then
                        Set m_oOleDragged = m_oPressed
                        Set m_oPressed = Nothing
                        RefreshControl
                        '--- temporarily manual drop mode
                        UserControl.OleDropMode = ucsOleDropManual
                        OleDrag
                        '--- restore drop mode
                        OleDropMode = m_eOleDropMode
                    End If
                End If
            End If
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)
    '--- if lost capture -> re-sync
    If (Button And vbLeftButton) = 0 And Not m_oPressed Is Nothing Then
        Set m_oPressed = Nothing
        RefreshControl
    End If
    '--- if label edit don't hilight (because no mouse capture)
    If Not m_oLabelItem Is Nothing Then
        Exit Sub
    End If
    '--- track over which item is mouse pointer
    Call HitTest(x, y, oBtn)
    If Not oBtn Is m_oOver Then
        Set m_oOver = oBtn
        RefreshControl
        '--- sync tooltip
        If Not m_oOver Is Nothing Then
            ApiTooltipText = m_oOver.TooltipText
        Else
            ApiTooltipText = ""
        End If
    End If
    '--- change mouseicon
    If oBtn.Class <> ucsBtnClassGroup Then
        Set MouseIcon = Nothing
    Else
        Set MouseIcon = m_oHandIcon
    End If
    '--- manage mouse capture
    If x >= 0 And x < ScaleWidth And y >= 0 And y < ScaleHeight Then
        With uTrackMouse
            .cbSize = Len(uTrackMouse)
            .dwFlags = TME_LEAVE
            .hwndTrack = UserControl.hwnd
        End With
        TrackMouseEvent uTrackMouse
    End If
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
    DropHighlightIdx = -1
    GroupHighlightIdx = -1
    Set m_oOleDragged = Nothing
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Const FUNC_NAME     As String = "UserControl_OLEDragDrop"
    
    On Error GoTo EH
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
    If Not m_oOleDragged Is Nothing And (DropHighlightIdx >= 0 Or GroupHighlightIdx >= 0) Then
        Dim oBtn            As cButton
        Dim bCancel         As Boolean

        RaiseEvent OLEBeforeMove(m_oOleDragged, _
                IIf(m_oOleDragged.Class = ucsBtnClassItem, _
                        DropHighlightIdx, GroupHighlightIdx), _
                bCancel)
        If Not bCancel Then
            If m_oOleDragged.Class = ucsBtnClassItem Then
                '--- item dropped
                If DropHighlightIdx > 0 Then
                    Set oBtn = SelectedGroup.GroupItems.ItemByPosition(DropHighlightIdx)
                End If
                If Not m_oOleDragged Is oBtn Then
                    m_oOleDragged.Parent.GroupItems.frRemove m_oOleDragged.Index
                    With SelectedGroup.GroupItems
                        If Not oBtn Is Nothing Then
                            .frAdd m_oOleDragged, oBtn.Index + 1
                        Else
                            .frAdd m_oOleDragged, 1
                        End If
                    End With
                End If
            Else
                '--- group dropped
                If GroupHighlightIdx > 0 Then
                    Set oBtn = Groups.ItemByPosition(GroupHighlightIdx)
                End If
                If Not m_oOleDragged Is oBtn Then
                    With Groups
                        .frRemove m_oOleDragged.Index
                        If Not oBtn Is Nothing Then
                            .frAdd m_oOleDragged, oBtn.Index + 1
                        Else
                            .frAdd m_oOleDragged, 1
                        End If
                    End With
                End If
            End If
        End If
    End If
    Exit Sub
EH:
    Debug.Print MODULE_NAME & "." & FUNC_NAME & ": " & Err.Description & vbCrLf & Err.Source
    Resume Next
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Const FUNC_NAME     As String = "UserControl_OLEDragOver"
    Dim oBtn            As cButton
    Dim oPrev           As cButton
    Dim oNext           As cButton
    Dim eHit            As UcsHitTestEnum
    Dim ePrev           As UcsHitTestEnum
    Dim eNext           As UcsHitTestEnum
    Dim lItemHeight     As Long
    
    On Error GoTo EH
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
    If Not m_oOleDragged Is Nothing Then
        eHit = HitTest(x, y, oBtn)
        If m_oOleDragged.Class = ucsBtnClassItem Then
            Select Case eHit
            Case ucsHitScrollupArrow
                DropHighlightIdx = -1
                pvGetItemInfo m_oSelectedGroup, lItemHeight, 0, 0
                pvAnimateItems m_lGroupOffset - ScrollItemsCount * (lItemHeight + ITEM_VERT_DISTANCE)
                RefreshControl
                Effect = vbDropEffectMove Or vbDropEffectScroll
            Case ucsHitScrolldownArrow
                DropHighlightIdx = -1
                pvGetItemInfo m_oSelectedGroup, lItemHeight, 0, 0
                pvAnimateItems m_lGroupOffset + ScrollItemsCount * (lItemHeight + ITEM_VERT_DISTANCE)
                RefreshControl
                Effect = vbDropEffectMove Or vbDropEffectScroll
            Case Else
                Call HitTest(x, y - ScaleY(8, vbPixels), oPrev)
                Call HitTest(x, y + ScaleY(7, vbPixels), oNext)
                If Not oPrev Is Nothing Then
                    If oPrev.Class <> ucsBtnClassItem Then
                        Set oPrev = Nothing
                    End If
                End If
                If Not oNext Is Nothing Then
                    If oNext.Class <> ucsBtnClassItem Then
                        Set oNext = Nothing
                    End If
                End If
                If Not oBtn Is Nothing Then
                    If oBtn.Class = ucsBtnClassGroup Then
                        If Not oBtn.Selected Then
                            DropHighlightIdx = -1
                            Effect = vbDropEffectMove
                            pvAnimateToGroup oBtn
                        End If
                        Exit Sub
                    End If
                End If
                If oPrev Is oNext Then
                    '--- first check if empty group -> allow add if on background
                    If eHit = ucsHitGroupBackground And SelectedGroup.GroupItems.Count = 0 Then
                        DropHighlightIdx = 0
                        Effect = vbDropEffectMove
                    Else
                        DropHighlightIdx = -1
                        Effect = vbDropEffectNone
                    End If
                    If Not oBtn Is m_oOver Then
                        Set m_oOver = oBtn
                        RefreshControl
                    End If
                Else
                    If Not oPrev Is Nothing Then
                        Set m_oOver = Nothing
                        DropHighlightIdx = oPrev.Position
                        Effect = vbDropEffectMove
                    ElseIf Not oNext Is Nothing Then
                        Set m_oOver = Nothing
                        DropHighlightIdx = oNext.Position - 1
                        Effect = vbDropEffectMove
                    Else
                        DropHighlightIdx = -1
                        Effect = vbDropEffectNone
                    End If
                End If
            End Select
        Else
            ePrev = HitTest(x, y - ScaleY(8, vbPixels), oPrev)
            eNext = HitTest(x, y + ScaleY(7, vbPixels), oNext)
            If Not oPrev Is Nothing Then
                If oPrev.Class <> ucsBtnClassGroup Then
                    Set oPrev = Nothing
                End If
            End If
            If Not oNext Is Nothing Then
                If oNext.Class <> ucsBtnClassGroup Then
                    Set oNext = Nothing
                End If
            End If
            '--- special case: opened is last group
            If eNext = ucsHitNoWhere And (ePrev = ucsHitGroupBackground Or ePrev = ucsHitItemButton) Then
                Set oPrev = SelectedGroup
            End If
            If oNext Is m_oOleDragged Or oPrev Is m_oOleDragged Or oNext Is oPrev Then
                GroupHighlightIdx = -1
                Effect = vbDropEffectNone
            Else
                If Not oNext Is Nothing Then
                    GroupHighlightIdx = oNext.Position - 1
                    Effect = vbDropEffectMove
                ElseIf Not oPrev Is Nothing Then
                    If eNext = ucsHitNoWhere Then
                        GroupHighlightIdx = oPrev.Position
                        Effect = vbDropEffectMove
                    Else
                        GroupHighlightIdx = -1
                        Effect = vbDropEffectNone
                    End If
                Else
                    GroupHighlightIdx = -1
                    Effect = vbDropEffectNone
                End If
            End If
        End If
        RefreshControl
    End If
    Exit Sub
EH:
    Debug.Print MODULE_NAME & "." & FUNC_NAME & ": " & Err.Description & vbCrLf & Err.Source
    Resume Next
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    If OleDragMode = ucsOleDragAutomatic Then
        Data.SetData "", vbCFText
        AllowedEffects = vbDropEffectMove Or vbDropEffectScroll
    End If
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_Paint()
    Const FUNC_NAME     As String = "UserControl_Paint"
    Dim pt              As POINTAPI
    
    On Error GoTo EH
    '--- sync current mouse position
    Set m_oOver = Nothing
    If m_oLabelItem Is Nothing Then
        GetCursorPos pt
        If WindowFromPoint(pt.x, pt.y) = UserControl.hwnd Then
            ScreenToClient UserControl.hwnd, pt
            Call HitTest(pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY, m_oOver)
        End If
    End If
    '--- sync current group offset
    pvFixGroupOffset
    '--- VB's redraw bitmap hack
    AutoRedraw = True
    pvPaintControl
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    If Not m_oLabelItem Is Nothing Then
        pvFinishLabelEdit m_oLabelEdit.Text
    End If
    If Not m_oSelectedItem Is Nothing Then
        If m_oSelectedGroup Is m_oSelectedItem.Parent Then
            EnsureVisible m_oSelectedItem, False
        End If
    End If
    RefreshControl
End Sub

Private Sub UserControl_Show()
    pvSubclass
End Sub

Private Sub UserControl_Hide()
    '--- unsubclass
    Set m_oSubclassControl = Nothing
    Set m_oSubclassTop = Nothing
End Sub

Private Sub UserControl_Terminate()
    PushError
    On Error Resume Next
    '--- destroy tooltip window
    If m_hWndTooltip <> 0 Then
        Call DestroyWindow(m_hWndTooltip)
        m_hWndTooltip = 0
    End If
    '--- clean up
    Set m_oSelectedGroup = Nothing
    Set m_oSelectedItem = Nothing
    If Not m_oTop.Items Is Nothing Then
        m_oTop.Items.Clear
        Set m_oTop.Items = Nothing
    End If
    Set m_oTop = Nothing
    Set m_oOver = Nothing
    Set m_oPressed = Nothing
    Set m_oOleDragged = Nothing
    Set m_oClicked = Nothing
    Set m_oLabelItem = Nothing
    Set m_oLabelHook = Nothing
    Set m_oLabelEdit = Nothing
    #If DebugMode Then
        DebugTerm m_sDebugID
    #End If
    PopError
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    m_bInSet = True
    Extender.Align = vbAlignLeft
    Groups.Add "Button Group"
    MaskColor = DEF_MASKCOLOR
    AnimationSteps = DEF_ANIMSTEPS
    OleDragMode = DEF_OLEDRAGMODE
    OleDropMode = DEF_OLEDROPMODE
    ScrollItemsCount = DEF_SCROLLITEMSCOUNT
    UseSystemFont = DEF_USESYSTEMFONT
    FlatScrollArrows = DEF_FLATSCROLLARROWS
    WrapText = DEF_WRAPTEXT
    AllowGroupDrag = DEF_ALLOWGROUPDRAG
    LabelEdit = DEF_LABELEDIT
    m_bInSet = False
    frGetMeasures
    pvCreateTooltipWindow
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"
    
    On Error GoTo EH
    m_bInSet = True
    With PropBag
        Set Font = .ReadProperty("Font", DEF_FONT)
        FormatControl.Contents = .ReadProperty("FormatControl", DEF_GROUP_FORMAT.Contents)
        FormatGroup.Contents = .ReadProperty("FormatGroup", DEF_GROUP_FORMAT.Contents)
        FormatGroupHover.Contents = .ReadProperty("FormatGroupHover", DEF_GROUP_FORMAT_HOVER.Contents)
        FormatGroupPressed.Contents = .ReadProperty("FormatGroupPressed", DEF_GROUP_FORMAT_PRESSED.Contents)
        FormatGroupSelected.Contents = .ReadProperty("FormatGroupSelected", DEF_GROUP_FORMAT_SELECTED.Contents)
        FormatItem.Contents = .ReadProperty("FormatItem", DEF_ITEM_FORMAT.Contents)
        FormatItemLargeIcons.Contents = .ReadProperty("FormatItemLargeIcons", DEF_ITEM_FORMAT_LARGE_ICONS.Contents)
        FormatItemHover.Contents = .ReadProperty("FormatItemHover", DEF_ITEM_FORMAT_HOVER.Contents)
        FormatItemPressed.Contents = .ReadProperty("FormatItemPressed", DEF_ITEM_FORMAT_PRESSED.Contents)
        FormatItemSelected.Contents = .ReadProperty("FormatItemSelected", DEF_ITEM_FORMAT_SELECTED.Contents)
        FormatSmallIcon.Contents = .ReadProperty("FormatSmallIcon", DEF_SMALL_ICON_FORMAT.Contents)
        FormatSmallIconHover.Contents = .ReadProperty("FormatSmallIconHover", DEF_SMALL_ICON_FORMAT_HOVER.Contents)
        FormatSmallIconPressed.Contents = .ReadProperty("FormatSmallIconPressed", DEF_SMALL_ICON_FORMAT_PRESSED.Contents)
        FormatSmallIconSelected.Contents = .ReadProperty("FormatSmallIconSelected", DEF_SMALL_ICON_FORMAT_SELECTED.Contents)
        FormatLargeIcon.Contents = .ReadProperty("FormatLargeIcon", DEF_LARGE_ICON_FORMAT.Contents)
        FormatLargeIconHover.Contents = .ReadProperty("FormatLargeIconHover", DEF_LARGE_ICON_FORMAT_HOVER.Contents)
        FormatLargeIconPressed.Contents = .ReadProperty("FormatLargeIconPressed", DEF_LARGE_ICON_FORMAT_PRESSED.Contents)
        FormatLargeIconSelected.Contents = .ReadProperty("FormatLargeIconSelected", DEF_LARGE_ICON_FORMAT_SELECTED.Contents)
        MaskColor = .ReadProperty("MaskColor", DEF_MASKCOLOR)
        Groups.Contents = .ReadProperty("Groups", "")
        AnimationSteps = .ReadProperty("AnimationSteps", DEF_ANIMSTEPS)
        OleDragMode = .ReadProperty("OleDragMode", DEF_OLEDRAGMODE)
        OleDropMode = .ReadProperty("OleDropMode", DEF_OLEDROPMODE)
        ScrollItemsCount = .ReadProperty("ScrollItemsCount", DEF_SCROLLITEMSCOUNT)
        UseSystemFont = .ReadProperty("UseSystemFont", DEF_USESYSTEMFONT)
        FlatScrollArrows = .ReadProperty("FlatScrollArrows", DEF_FLATSCROLLARROWS)
        WrapText = .ReadProperty("WrapText", DEF_WRAPTEXT)
        AllowGroupDrag = .ReadProperty("AllowGroupDrag", DEF_ALLOWGROUPDRAG)
        LabelEdit = .ReadProperty("LabelEdit", DEF_LABELEDIT)
    End With
    m_bInSet = False
    frGetMeasures
    pvCreateTooltipWindow
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_WriteProperties"
    
    On Error GoTo EH
    With PropBag
        Call .WriteProperty("Font", Font, DEF_FONT)
        Call .WriteProperty("FormatControl", FormatControl.Contents, DEF_GROUP_FORMAT.Contents)
        Call .WriteProperty("FormatGroup", FormatGroup.Contents, DEF_GROUP_FORMAT.Contents)
        Call .WriteProperty("FormatGroupHover", FormatGroupHover.Contents, DEF_GROUP_FORMAT_HOVER.Contents)
        Call .WriteProperty("FormatGroupPressed", FormatGroupPressed.Contents, DEF_GROUP_FORMAT_PRESSED.Contents)
        Call .WriteProperty("FormatGroupSelected", FormatGroupSelected.Contents, DEF_GROUP_FORMAT_SELECTED.Contents)
        Call .WriteProperty("FormatItem", FormatItem.Contents, DEF_ITEM_FORMAT.Contents)
        Call .WriteProperty("FormatItemLargeIcons", FormatItemLargeIcons.Contents, DEF_ITEM_FORMAT_LARGE_ICONS.Contents)
        Call .WriteProperty("FormatItemHover", FormatItemHover.Contents, DEF_ITEM_FORMAT_HOVER.Contents)
        Call .WriteProperty("FormatItemPressed", FormatItemPressed.Contents, DEF_ITEM_FORMAT_PRESSED.Contents)
        Call .WriteProperty("FormatItemSelected", FormatItemSelected.Contents, DEF_ITEM_FORMAT_SELECTED.Contents)
        Call .WriteProperty("FormatSmallIcon", FormatSmallIcon.Contents, DEF_SMALL_ICON_FORMAT.Contents)
        Call .WriteProperty("FormatSmallIconHover", FormatSmallIconHover.Contents, DEF_SMALL_ICON_FORMAT_HOVER.Contents)
        Call .WriteProperty("FormatSmallIconPressed", FormatSmallIconPressed.Contents, DEF_SMALL_ICON_FORMAT_PRESSED.Contents)
        Call .WriteProperty("FormatSmallIconSelected", FormatSmallIconSelected.Contents, DEF_SMALL_ICON_FORMAT_SELECTED.Contents)
        Call .WriteProperty("FormatLargeIcon", FormatLargeIcon.Contents, DEF_LARGE_ICON_FORMAT.Contents)
        Call .WriteProperty("FormatLargeIconHover", FormatLargeIconHover.Contents, DEF_LARGE_ICON_FORMAT_HOVER.Contents)
        Call .WriteProperty("FormatLargeIconPressed", FormatLargeIconPressed.Contents, DEF_LARGE_ICON_FORMAT_PRESSED.Contents)
        Call .WriteProperty("FormatLargeIconSelected", FormatLargeIconSelected.Contents, DEF_LARGE_ICON_FORMAT_SELECTED.Contents)
        Call .WriteProperty("MaskColor", MaskColor, DEF_MASKCOLOR)
        Call .WriteProperty("Groups", Groups.Contents)
        Call .WriteProperty("AnimationSteps", AnimationSteps, DEF_ANIMSTEPS)
        Call .WriteProperty("OleDragMode", OleDragMode, DEF_OLEDRAGMODE)
        Call .WriteProperty("OleDropMode", OleDropMode, DEF_OLEDROPMODE)
        Call .WriteProperty("ScrollItemsCount", ScrollItemsCount, DEF_SCROLLITEMSCOUNT)
        Call .WriteProperty("UseSystemFont", UseSystemFont, DEF_USESYSTEMFONT)
        Call .WriteProperty("FlatScrollArrows", FlatScrollArrows, DEF_FLATSCROLLARROWS)
        Call .WriteProperty("WrapText", WrapText, DEF_WRAPTEXT)
        Call .WriteProperty("AllowGroupDrag", AllowGroupDrag, DEF_ALLOWGROUPDRAG)
        Call .WriteProperty("LabelEdit", LabelEdit, DEF_LABELEDIT)
    End With
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub m_oLabelEdit_Change()
    Dim lItemHeight         As Long
    Dim lIconW              As Long
    Dim lIconH              As Long
    Dim oFmt                As cFormatDef
    Dim oFmtIcon            As cFormatDef
    Dim bResizeNeeded       As Boolean
    Dim LeftX               As Long
    Dim TopY                As Long
    Dim RightX              As Long
    Dim BottomY             As Long
    Dim lLeft               As Long
    Dim lTop                As Long
    Dim lWidth              As Long
    Dim lHeight             As Long
    Dim lAlign              As Long
    
    If m_oLabelItem Is Nothing Then
        Exit Sub
    End If
    Select Case m_oLabelItem.Class
    Case ucsBtnClassItem
        '--- get item position
        pvGetItemInfo m_oLabelItem.Parent, lItemHeight, lIconW, lIconH
        TopY = m_oRendContrl.BorderSize + INNER_BORDER + (m_oLabelItem.Parent.Position * (m_lGroupHeight + 1) + (m_oLabelItem.Position - 1) * (lItemHeight + ITEM_VERT_DISTANCE) - m_lGroupOffset) - 1
        BottomY = TopY + lItemHeight
        LeftX = m_oRendContrl.BorderSize + INNER_BORDER
        RightX = ScaleWidth / Screen.TwipsPerPixelX - LeftX
        '--- figure out formts to use in calcs
        Set oFmt = IIf(m_oLabelItem.Parent.IconsType = ucsIcsSmallIcons, m_oRendItmNrm, m_oRendItmNrmLrg)
        Set oFmtIcon = IIf(m_oLabelItem.Parent.IconsType = ucsIcsSmallIcons, m_oRendSIcNrm, m_oRendLIcNrm)
        '--- apply the way the icon affects the item
        bResizeNeeded = oFmt.VertAlignment <> oFmtIcon.VertAlignment And oFmt.HorAlignment = oFmtIcon.HorAlignment Or oFmtIcon.HorAlignment = ucsFhaCenter
        Select Case oFmtIcon.HorAlignment
        Case ucsFhaRight
            If bResizeNeeded Or oFmt.VertAlignment = oFmtIcon.VertAlignment Then
                RightX = RightX - 2 * (oFmt.BorderSize + oFmt.Padding + oFmtIcon.Padding) - lIconW
            End If
        Case ucsFhaCenter
        Case Else ' ucsFhaLeft
            If bResizeNeeded Or oFmt.VertAlignment = oFmtIcon.VertAlignment Then
                LeftX = LeftX + oFmt.BorderSize + 2 * oFmt.Padding + 2 * oFmtIcon.Padding + lIconW
            End If
        End Select
        Select Case oFmtIcon.VertAlignment
        Case ucsFvaBottom
            If bResizeNeeded Then
                BottomY = BottomY - oFmt.BorderSize - lIconH - 2 * oFmtIcon.Padding
            End If
        Case ucsFvaMiddle
        Case Else ' ucsFvaTop
            If bResizeNeeded Then
                TopY = TopY + oFmt.BorderSize + lIconH + 2 * oFmtIcon.Padding  '+ oFmt.Padding
            End If
        End Select
        '--- apply item format properties and account for the textbox border of 2px
        LeftX = LeftX + oFmt.Padding + oFmt.BorderSize + oFmt.OffsetX - 2
        TopY = TopY + oFmt.Padding + oFmt.BorderSize + oFmt.OffsetY - 2
        RightX = RightX - oFmt.Padding - oFmt.BorderSize + oFmt.OffsetX + 2
        BottomY = BottomY - oFmt.Padding - oFmt.BorderSize + oFmt.OffsetY + 2
        '--- calc width
        lWidth = TextWidth(m_oLabelEdit.Text) / Screen.TwipsPerPixelX + 16
        If lWidth > RightX - LeftX Then
            lWidth = RightX - LeftX
        End If
        '--- calc left
        Select Case m_lLabelAlign
        Case ucsFhaLeft
            lLeft = LeftX
        Case ucsFhaCenter
            lLeft = (RightX + LeftX - lWidth) \ 2
        Case ucsFhaRight
            lLeft = RightX - lWidth
        End Select
        '--- calc height
        If oFmt.HorAlignment = ucsFhaCenter Then
            With New cMemDC
                .Init
                Set .Font = m_oLabelEdit.Font
                '--- Hack: adhoc array, much like <C++> "0123456789"[Idx] </C++>
                lAlign = Array(DT_LEFT, DT_CENTER, DT_RIGHT)(oFmt.HorAlignment) _
                      Or Array(DT_TOP, DT_VCENTER, DT_BOTTOM)(oFmt.VertAlignment)
                .DrawText m_oLabelEdit.Text & " ", (LeftX + 4), (0), (RightX - 4), lHeight, lAlign Or DT_CALCRECT Or DT_WORDBREAK Or DT_EDITCONTROL
            End With
        Else
            lHeight = TextHeight(m_oLabelEdit.Text) / Screen.TwipsPerPixelY
        End If
        lHeight = lHeight + 4
        '--- calc top
        Select Case oFmt.VertAlignment
        Case ucsFvaTop
            lTop = TopY
        Case ucsFvaMiddle
            lTop = (TopY + BottomY - lHeight) \ 2
        Case ucsFvaBottom
            lTop = BottomY - lHeight
        End Select
        If lTop < TopY Then
            lTop = TopY
        End If
    Case ucsBtnClassGroup
        If m_oLabelItem.Position <= m_oSelectedGroup.Position Then
            lTop = m_oRendContrl.BorderSize + (m_oLabelItem.Position - 1) * (m_lGroupHeight + 1)
        Else
            lTop = ScaleHeight / Screen.TwipsPerPixelY - m_oRendContrl.BorderSize - (Groups(Groups.Count).Position - m_oLabelItem.Position + 1) * (m_lGroupHeight + 1) + 1
        End If
        lLeft = m_oRendContrl.BorderSize
        lWidth = ScaleWidth / Screen.TwipsPerPixelX - 2 * lLeft
        '--- calc height
        With New cMemDC
            .Init
            Set .Font = m_oLabelEdit.Font
            .DrawText m_oLabelEdit.Text & " ", lLeft + 4, (0), lLeft + lWidth - 4, lHeight, DT_CALCRECT Or DT_WORDBREAK
            lHeight = lHeight + 4
        End With
        If lHeight < m_lGroupHeight Then
            lHeight = m_lGroupHeight
        End If
    End Select
    '--- move textbox (and skip automagic height bug)
    If m_oLabelEdit.Left <> lLeft * Screen.TwipsPerPixelX _
                Or m_oLabelEdit.Top <> lTop * Screen.TwipsPerPixelY Then
        m_oLabelEdit.Font.Size = 1
        m_oLabelEdit.Move lLeft * Screen.TwipsPerPixelX, lTop * Screen.TwipsPerPixelY
    End If
    If m_oLabelEdit.Width <> lWidth * Screen.TwipsPerPixelX Then
        m_oLabelEdit.Font.Size = 1
        m_oLabelEdit.Width = lWidth * Screen.TwipsPerPixelX
    End If
    If m_oLabelEdit.Height <> lHeight * Screen.TwipsPerPixelY Then
        m_oLabelEdit.Font.Size = 1
        m_oLabelEdit.Height = lHeight * Screen.TwipsPerPixelY
    End If
    If m_oLabelEdit.Font.Size <> UserControl.Font.Size Then
        m_oLabelEdit.Font.Size = UserControl.Font.Size
    End If
End Sub

Private Sub m_oLabelEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 0 Then
        pvFinishLabelEdit m_oLabelEdit.Text
    ElseIf KeyCode = vbKeyEscape And Shift = 0 Then
        pvFinishLabelEdit m_sLabelCaption
    End If
End Sub

Private Sub m_oFmtContrl_ControlFont(oFont As stdole.StdFont)
    Set oFont = DrawFont
End Sub

Private Sub m_oFmtContrl_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtGrpNrm_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtGrpHvr_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtGrpPrs_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtGrpSel_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtItmNrm_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtItmLrg_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtItmHvr_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtItmPrs_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtItmSel_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtSIcNrm_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtSIcHvr_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtSIcPrs_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFmtSIcSel_Changed()
    pvPropertyChanged
End Sub

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    pvPropertyChanged
End Sub

Private Sub m_oTop_GetControl(oValue As ctxOutlookBar)
    Set oValue = Me
End Sub

'=========================================================================
' Subclasser interface
'=========================================================================

Private Sub ISubclassingSink_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
    Const FUNC_NAME     As String = "ISubclassingSink_Before"
    Dim lItemHeight     As Long
    Dim rc              As RECT
    Dim x               As Single
    Dim y               As Single
    
    On Error GoTo EH
    Select Case uMsg
    Case WM_CANCELMODE
        Set m_oPressed = Nothing
        RefreshControl
    Case WM_SETTINGCHANGE, WM_SYSCOLORCHANGE
        If Not m_oLabelItem Is Nothing Then
            pvFinishLabelEdit m_oLabelEdit.Text
        End If
        frGetMeasures
        RefreshControl
        RaiseEvent WindowsSettingsChanged
        '--- fix the tooltip (if changed)
        mouse_event MOUSEEVENTF_MOVE, 1, 0, 0, 0
        mouse_event MOUSEEVENTF_MOVE, -1, 0, 0, 0
    Case MouseWheelFwdMsg ' WM_MOUSEWHEEL
        If Not m_oLabelItem Is Nothing Then
            pvFinishLabelEdit m_oLabelEdit.Text
        End If
        Call GetWindowRect(UserControl.hwnd, rc)
        x = ScaleX(((lParam And &HFFFF&) - rc.Left), vbPixels, vbTwips)
        y = ScaleY((lParam \ &H10000) - rc.Top, vbPixels, vbTwips)
        Select Case HitTest(x, y, Nothing)
        Case ucsHitGroupBackground, ucsHitItemButton
            pvGetItemInfo m_oSelectedGroup, lItemHeight, 0, 0
            pvAnimateItems m_lGroupOffset + IIf((wParam \ &H1000&) > 0, -ScrollItemsCount, ScrollItemsCount) * (lItemHeight + ITEM_VERT_DISTANCE)
            '--- fix the tooltip (if changed)
            mouse_event MOUSEEVENTF_MOVE, 1, 0, 0, 0
            mouse_event MOUSEEVENTF_MOVE, -1, 0, 0, 0
        End Select
        bHandled = True
    Case WM_COMMAND
        If wParam \ &H10000 = EN_KILLFOCUS Then
            If Not m_oLabelEdit Is Nothing Then
                If m_oLabelEdit.Visible Then
                    pvFinishLabelEdit m_oLabelEdit.Text
                End If
            End If
        End If
    Case WM_MOUSELEAVE
        UserControl_MouseMove 0, 0, -1, -1
    Case WM_MOUSEHOVER
        If m_sClickX >= 0 Then
            pvBeginLabelEdit
        End If
    End Select
    Exit Sub
EH:
    Debug.Print MODULE_NAME & "." & FUNC_NAME & ": " & Err.Description & vbCrLf & Err.Source
    Resume Next
End Sub

Private Sub ISubclassingSink_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    
End Sub

'=========================================================================
' Hooking interface
'=========================================================================

Private Sub IHookingSink_Before(bHandled As Boolean, lReturn As Long, nCode As SubclassingSink.HookCode, wParam As Long, lParam As Long)
    Dim cs              As CREATESTRUCT
    
    If nCode = HCBT_CREATEWND Then
        cs = m_oLabelHook.CREATESTRUCT(m_oLabelHook.CBT_CREATEWND(lParam).lpcs)
        SetWindowLong wParam, GWL_STYLE, (cs.Style And Not 3) Or ES_MULTILINE Or m_lLabelAlign
    End If
End Sub

Private Sub IHookingSink_After(lReturn As Long, ByVal nCode As SubclassingSink.HookCode, ByVal wParam As Long, ByVal lParam As Long)

End Sub

'=========================================================================
' End of file
'=========================================================================


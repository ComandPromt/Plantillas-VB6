VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTaskBar"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------
'This class encapsulates the 32-bit API functions for using
'the Windows 95/NT4 TaskBar Notification Area (also know as the
'Tray).  With this class you may add icons to the Notification
'Area, replace them (for signaling), change their ToolTip text,
'and delete them when they are no longer needed.
'----------------------------------------------------------
'Copyright 1995/1996 by Don Bradner.  May be freely distributed
'Author contact: CIS 76130,1007; internet dbirdman@redshift.com
'http://www.redshift.com/~arcatpet.  Support available as time
'allows, including the VBPJ and MSBASIC forums on Compuserve.
'==========================================================
'                       PRIVATE VARIABLES
'==========================================================

' -----------------------------------------------------------
' NOTIFYICONDATA type is needed for Shell_NotifyIcon function
'------------------------------------------------------------

Private Type NOTIFYICONDATA
    lStructureSize    As Long
    hWnd   As Long
    lID As Long
    lFlags As Long
    lCallBackMessage As Long
    hIcon As Long
    sTip As String * 64
End Type

' -----------------------------------------------------------
' lRect type and AppBarData are needed for SHAppBarMessage function
'------------------------------------------------------------
Private Type lRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type APPBARDATA
    lStructureSize As Long
    hWnd As Long
    lCallBackMessage As Long
    lEdge As Long
    rc As lRect
    lparam As Long
End Type

Private Declare Function Shell_NotifyIcon& Lib "shell32.dll" (ByVal lMessage&, NID As NOTIFYICONDATA)
Private Declare Function SHAppBarMessage& Lib "shell32.dll" (ByVal dwMessage&, pData As APPBARDATA)

Private Const NIM_ADD = 0&
Private Const NIM_DELETE = 2&
Private Const NIM_MODIFY = 1&
Private Const NIF_ICON = 2&
Private Const NIF_MESSAGE = 1&
Private Const NIF_TIP = 4&

Private Const ABM_GETTASKBARPOS = &H5&
Private structNotify As NOTIFYICONDATA
Private structBarData As APPBARDATA

Private Const WM_USER = &H400
Private Const UM_TASKBARMESSAGE = WM_USER + &H201

Private lTempLong&

'==========================================================
'                          PROPERTIES
'==========================================================

Public Property Let hWnd(Client As Form)
    structNotify.hWnd = Client.hWnd
End Property

Public Property Let Message(MsgValue&)
    '---------------------------------
    'Sets the message number to a value
    'different than the startup value.
    'useful if more than one instance of
    'the class is used, since each should
    'return separate messages.
    '-----------------------------------
    structNotify.lCallBackMessage = MsgValue
End Property

Public Property Get Message&()
    '---------------------------------
    'Returns the message number that
    'has been set at initialization or
    'changed by the client app via the
    'Property Let Message Proc
    '---------------------------------
    Message = structNotify.lCallBackMessage
End Property

'==========================================================
'                            METHODS
'==========================================================

Public Sub AddIcon(lID&, sTip$, hIcon)
    structNotify.lID = lID
    structNotify.hIcon = hIcon
    structNotify.sTip = sTip & Chr$(0)
    structNotify.lFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    lTempLong = Shell_NotifyIcon(NIM_ADD, structNotify)
End Sub

Public Function AppBarExists&()
    '--------------------------------------------------------
    'Since a user may be running 32-bit Windows but not be using
    'the Explorer shell, we use the SHAppBarMessage call to determine
    'whether there is a system taskbar.
    '---------------------------------------------------------
    structBarData.lStructureSize = 36&
    AppBarExists = SHAppBarMessage(ABM_GETTASKBARPOS, structBarData)
End Function

Public Sub ChangeIcon(lID&, hIcon)
    structNotify.lID = lID
    structNotify.hIcon = hIcon
    structNotify.lFlags = NIF_ICON
    lTempLong = Shell_NotifyIcon(NIM_MODIFY, structNotify)
End Sub

Public Sub ChangeMessage(lID&, sNewMessage$)
    structNotify.lID = lID
    structNotify.sTip = sNewMessage & Chr$(0)
    structNotify.lFlags = NIF_TIP
    lTempLong = Shell_NotifyIcon(NIM_MODIFY, structNotify)
End Sub

Public Sub DeleteIcon(lID&)
    structNotify.lID = lID
    lTempLong = Shell_NotifyIcon(NIM_DELETE, structNotify)
End Sub

Public Sub RemoveAllIcons(iTotal&)
Dim lRet&
For lTempLong = 1 To iTotal
    structNotify.lID = lTempLong
    lRet = Shell_NotifyIcon(NIM_DELETE, structNotify)
Next lTempLong
End Sub

'==========================================================
'                     INITIALIZE/TERMINATE
'==========================================================
Private Sub Class_Initialize()
    '------------------------------------------
    'Preset portions of the data structure
    '------------------------------------------
    structNotify.lStructureSize = 88&
    structNotify.lFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    structNotify.lCallBackMessage = UM_TASKBARMESSAGE
    structNotify.sTip = ""
End Sub


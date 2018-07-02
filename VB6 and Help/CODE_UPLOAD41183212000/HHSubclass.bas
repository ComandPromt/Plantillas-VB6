Attribute VB_Name = "modHHSubclass"
' *****************************************************
' Code to subclass a Visual Basic form for
' WM_HELP messaging for HTML Help purposes
' Version 3.0c
' (c)August 1999, Delmar Computing Services
'
' Developed by David Liske, Tipton, Michigan, USA
' Microsoft HTML Help MVP
' http://www.vbexplorer.com/htmlhelp.asp
'
' ATTENTION:
' Due to the use of the AddressOf operator, this code
' *will* crash the Visual Basic IDE in Debug mode.
' If debugging of the application is necessary,
' uncomment the first Exit Sub in the HHSubclass routine.
' Again, this routine *cannot* be run in Debug mode.
'
' To use this module, the subclassed calling form
' needs to have the following methods included:
'
' Public Sub OnContextMenu(hWndControl As Long)
' Public Sub OnHelp(hWndControl As Long)
' Public Sub OnNavComplete(phhnt As Long)
' Public Sub OnTCard(wParam As Long, lParam As Long)
' Public Sub OnTrack(phhnn As Long)
' Public Sub OnWindowCreate(phhnt As Long)
'
' Please send any performance or functionality
' modifications of this file to delmar@tc3net.com
' *****************************************************

Option Explicit

' Notification codes
Private Const HHN_FIRST = -860
Private Const HHN_LAST = -879

Private Const HHN_NAVCOMPLETE = HHN_FIRST
Private Const HHN_TRACK = HHN_FIRST - 1
Private Const HHN_WINDOW_CREATE = HHN_FIRST - 2

Private Const HH_MAX_TABS = 19

'Windows messaging
Private Const WM_CONTEXTMENU = &H7B
Private Const WM_HELP = &H53
Private Const WM_NCDESTROY = &H82
Private Const WM_NOTIFY = &H4E
Private Const WM_TCARD = &H52

Private Const GWL_WNDPROC = (-4)

'Keyboard API
Public Const VK_F1 = &H70
Public Const VK_NUMLOCK = &H90
Public Const VK_CAPITAL = &H14
Public Const VK_SCROLL = &H91

' UDT for mouse cursor position
Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Type HELPINFO
  cbSize As Long
  iContextType As Long
  iCtrlId As Long
  hItemHandle As Long
  dwContextId As Long
  MousePos As POINTAPI
End Type

Private Type NMHDR
  hwndFrom As Long
  idfrom As Long
  code As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

' UDT for keyboard API
Private Type KeyboardBytes
  kbByte(0 To 255) As Byte
End Type

Public kbArray As KeyboardBytes

Private Type HH_WINTYPE
  cbStruct As Integer                         ' IN: size of this structure including all
                                             ' Information Types
  fUniCodeStrings As Boolean                  ' IN/OUT: TRUE if all strings are in UNICODE
  pszType As String                           ' IN/OUT: Name of a type of window
  fsValidMembers As Variant                   ' IN: Bit flag of valid members
                                             ' (HHWIN_PARAM_)
  fsWinProperties As Variant                  ' IN/OUT: Properties/attributes of the window
                                             ' (HHWIN_)
  pszCaption As String                        ' IN/OUT: Window title
  dwStyles As Variant                         ' IN/OUT: Window styles
  dwExStyles As Variant                       ' IN/OUT: Extended Window styles
  rcWindowPos As RECT                         ' IN: Starting position, OUT: current
                                             ' position
  nShowState As Integer                       ' IN: show state (e.g., SW_SHOW)
  hwndHelp As Variant                         ' OUT: window handle
  hwndCaller As Variant                       ' OUT: who called this window
                                             ' The following members are only valid if
                                             ' HHWIN_PROP_TRI_PANE is set
  hwndToolBar As Variant                      ' OUT: toolbar window in tri-pane window
  hwndNavigation As Variant                   ' OUT: navigation window in tri-pane window
  hwndHTML As Variant                         ' OUT: window displaying HTML in tri-pane
                                             ' window
  iNavWidth As Integer                        ' IN/OUT: width of navigation window
  rcHTML As RECT                              ' OUT: HTML window coordinates
  pszToc As String                            ' IN: Location of the table of contents file
  pszIndex As String                           ' IN: Location of the index file
  pszFile As String                           ' IN: Default location of the html file
  pszHome As String                           ' IN/OUT: html file to display when Home
                                             ' button is clicked
  fsToolBarFlags As Variant                   ' IN: flags controling the appearance of the
                                             ' toolbar
  fNotExpanded As Boolean                     ' IN: TRUE/FALSE to contract or expand, OUT:
                                             ' current state
  curNavType As Integer                       ' IN/OUT: UI to display in the navigational
                                             ' pane
  tabpos As Integer                           ' IN/OUT: HHWIN_NAVTAB_TOP, HHWIN_NAVTAB_LEFT,
                                             ' or HHWIN_NAVTAB_BOTTOM
  idNotify As Integer                         ' IN: ID to use for WM_NOTIFY messages
  tabOrder(HH_MAX_TABS + 1) As Byte           ' IN/OUT: tab order: Contents, Index,
                                             ' Search, History, Favorites, Reserved 1-5,
                                             ' Custom tabs
  cHistory As Integer                         ' IN/OUT: number of history items to keep
                                             ' (default is 30)
  pszJump1 As String                          ' Text for HHWIN_BUTTON_JUMP1
  pszJump2 As String                          ' Text for HHWIN_BUTTON_JUMP2
  pszUrlJump1 As String                       ' URL for HHWIN_BUTTON_JUMP1
  pszUrlJump2 As String                       ' URL for HHWIN_BUTTON_JUMP2
  rcMinSize As RECT                           ' Minimum size for window (ignored in version
                                             ' 1 of the Workshop)
  cbInfoTypes As Integer                      ' size of paInfoTypes;
End Type

'UDT for the HHN_TRACK message
Private Type tagHHNTRACK
  hdr As NMHDR
  pszCurUrl As String
  idAction As Integer
  phhWinType As HH_WINTYPE
End Type

'UDT for the HHN_NAVCOMPLETE and HHN_WINDOW_CREATE messages
Private Type tagHHN_NOTIFY
  hdr As NMHDR
  pszUrl As String
End Type

Private Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal msgWinMessage As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" (Dest As Any, _
    Source As Any, _
    ByVal nLen As Long)

Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
    
Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Public Declare Function GetKeyState Lib "user32" _
    (ByVal nVirtKey As Long) As Integer

Public Declare Function SetKeyboardState Lib "user32" _
    (kbArray As KeyboardBytes) As Long

Private colHTMLHelp As New Collection

Private Function HHSubclassWndProc(ByVal hwnd As Long, _
  ByVal msgWinMessage As Long, ByVal wParam As Long, _
  ByVal lParam As Long) As Long

  Dim colhHelp As Object
    
  ' Loop through all the forms in the collection and use
  ' the handle to determing the message the form belongs to
  For Each colhHelp In colHTMLHelp
    If (colhHelp.hwnd = hwnd) Then
      Exit For
    End If
  Next colhHelp
  
  ' Track down which message was sent and run the
  ' appropriate procedure on the calling form
  Select Case (msgWinMessage)
  Case WM_CONTEXTMENU
    ' The HELP_CONTEXTMENU command causes Help to
    ' display a menu, which is system defined. The
    ' menu contains a What's This command and allows
    ' users to display Help for the control.
    Call colhHelp.frm.OnContextMenu(wParam)
        
  Case WM_HELP
    ' The WM_HELP message is sent whenever the user
    ' presses the F1 key.  It also occurs in response
    ' to What's This Help requests.
    Dim hlpHelpInfo As HELPINFO
    Call CopyMemory(hlpHelpInfo, ByVal lParam, Len(hlpHelpInfo))
    Call colhHelp.frm.OnHelp(hlpHelpInfo.hItemHandle)
        
  Case WM_NOTIFY
    Dim nmhHeader As NMHDR
    Call CopyMemory(nmhHeader, ByVal lParam, Len(nmhHeader))
    
    Select Case (nmhHeader.code)
    Case HHN_NAVCOMPLETE
      ' Sent when the user successfully navigates to a
      ' topic in a compiled HTML Help (.chm) file.
      ' Uses the UDT tagHHN_NOTIFY.
      Call colhHelp.frm.OnNavComplete(lParam)
                
    Case HHN_TRACK
      ' Sent when a user clicks a button on the toolbar
      ' or a tab in the Navigation pane of the HTML Help
      ' Viewer. The message is sent before the action is
      ' started by the viewer.  Uses the UDT tagHHNTRACK.
      Call colhHelp.frm.OnTrack(lParam)
    
    Case HHN_WINDOW_CREATE
      ' Sent right before an HTML Help window is created.
      ' Uses the UDT tagHHN_NOTIFY.
      Call colhHelp.frm.OnWindowCreate(lParam)

    Case Else
      ' Let the message continue on its way
      HHSubclassWndProc = CallWindowProc _
          (colhHelp.lpPrevWndFunc, hwnd, msgWinMessage, _
          wParam, ByVal lParam)

    End Select
                
  Case WM_TCARD
    ' The WM_TCARD message is sent to a program that
    ' has initiated a training card based on Windows
    ' Help technology.  Does not apply to training cards
    ' created via embedded HTML Help.
    Call colhHelp.frm.OnTCard(wParam, lParam)
        
  Case Else
    ' Let the message continue on its way
    HHSubclassWndProc = CallWindowProc _
        (colhHelp.lpPrevWndFunc, hwnd, msgWinMessage, _
        wParam, ByVal lParam)
    
  End Select

  If (msgWinMessage = WM_NCDESTROY) Then
    ' If the window no longer exists,
    ' get it out of the HHSubclass collection
    Dim intCount As Integer
    For intCount = 1 To colHTMLHelp.Count
      If (colhHelp Is colHTMLHelp(intCount)) Then
        Call colHTMLHelp.Remove(intCount)
        Exit For
      End If
    Next intCount
  End If

End Function

Public Sub HHSubclass(frm As Object)
    
  ' Uncomment this line in Debug mode (see the
  ' "Attention" section of the comment block for
  ' this module):
  ' Exit Sub
  
  Dim hHelp As New HTMLHelp
    
  ' Create the object as a form
  Set hHelp.frm = frm
  hHelp.hwnd = frm.hwnd
  hHelp.lpPrevWndFunc = GetWindowLong _
      (frm.hwnd, _
      GWL_WNDPROC)
    
  ' Replace the basic window procedure of the
  ' form calling this procedure
  Call SetWindowLong(frm.hwnd, _
      GWL_WNDPROC, _
      AddressOf HHSubclassWndProc)
    
  ' Put this form into the subclass collection
  ' we created in the Declarations section
  colHTMLHelp.Add hHelp

End Sub

Public Sub HHUnSubClass(frm As Object)
  
  Dim hHelp As New HTMLHelp
  
  ' Release the subclassed form
  Call SetWindowLong(frm.hwnd, _
      GWL_WNDPROC, hHelp.lpPrevWndFunc)
  
End Sub

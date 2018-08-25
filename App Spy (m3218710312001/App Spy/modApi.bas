Attribute VB_Name = "modActTsk"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'           Author: Muhammad Abubakar

'           http://go.to/abubakar

'           <joehacker@yahoo.com>

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit
'APIs : WHERE THE REAL POWER IS
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
'Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

'APIs for Spying Menus:
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CascadeWindows Lib "user32" (ByVal hwndParent As Long, ByVal wHow As Long, lpRect As RECT, ByVal cKids As Long, lpkids As Long) As Integer

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String '* 255
    cch As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Const WM_COMMAND = &H111
Public Const MIIM_TYPE = &H10
Public Const MFT_STRING = &H0&


'Public Const WM_SETFOCUS = &H7     Messages for:
Public Const WM_CLOSE = &H10                    'Closing window
Public Const SW_SHOW = 5                        'showing window
Public Const WM_SETTEXT = &HC                   'Setting text of child window
Public Const WM_GETTEXT = &HD                   'Getting text of child window
Public Const WM_GETTEXTLENGTH = &HE
Public Const EM_GETPASSWORDCHAR = &HD2          'Checking if its a password field or not
Public Const BM_CLICK = &HF5                    'Clicking a button
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const WM_MDICASCADE = &H227              'Cascading windows
Public Const MDITILE_HORIZONTAL = &H1
Public Const MDITILE_SKIPDISABLED = &H2
Public Const WM_MDITILE = &H226

Public VCount As Integer, ICount As Integer
Public SpyHwnd As Long
'the following functions "WndEnumProc" & "WndEnumChildProc"
'are application defined, which means that they are defined by the
'programmer him/her self. They are actually required by the 'callback'
'function, EnumWindows & EnumChildWindows in this case. One more point
'to note is that we pass the 'address' of our app-defined functions
'to the call back function through 'AddressOf' operator. Its a condition
'of AddressOf operator that the function whose address is passed should
'be in a '*.bas' file which is module ofcourse.

Public Function WndEnumProc(ByVal hwnd As Long, ByVal lParam As ListView) As Long
    Dim WText As String * 512
    Dim bRet As Long, WLen As Long
    Dim WClass As String * 50
        
    WLen = GetWindowTextLength(hwnd)
    bRet = GetWindowText(hwnd, WText, WLen + 1)
    GetClassName hwnd, WClass, 50

    With frmSpy
        If (.chkCap.Value = vbUnchecked) Then
            Insert hwnd, lParam, WText, WClass
        ElseIf (.chkCap.Value = vbChecked And WLen <> 0) Then
            Insert hwnd, lParam, WText, WClass
        End If
    End With
    
    WndEnumProc = 1
End Function
Private Sub Insert(iHwnd As Long, lParam As ListView, iText As String, iClass As String)
    lParam.ListItems.Add.Text = Str(iHwnd)
    lParam.ListItems.Item(VCount).SubItems(1) = iClass
    lParam.ListItems.Item(VCount).SubItems(2) = iText
    VCount = VCount + 1
End Sub
Public Function WndEnumChildProc(ByVal hwnd As Long, ByVal lParam As ListView) As Long
    Dim bRet As Long
    Dim myStr As String * 50

    bRet = GetClassName(hwnd, myStr, 50)
    'if you want the text for only Edit class then use the if statement:
    'If (Left(myStr, 4) = "Edit") Then
    'lParam.Sorted = False

    With lParam.ListItems
        .Add.Text = Str(hwnd)
        .Item(ICount).SubItems(1) = myStr
        .Item(ICount).SubItems(2) = GetText(hwnd)
        If SendMessage(hwnd, EM_GETPASSWORDCHAR, 0, 0) = 0 Then
            .Item(ICount).SubItems(3) = "No"
        Else
            .Item(ICount).SubItems(3) = "Yes"
            .Item(ICount).ForeColor = vbRed
            .Item(ICount).Bold = True
            .Item(ICount).ListSubItems.Item(1).ForeColor = vbRed
            .Item(ICount).ListSubItems.Item(1).Bold = True
            .Item(ICount).ListSubItems.Item(2).ForeColor = vbRed
            .Item(ICount).ListSubItems.Item(2).Bold = True
            .Item(ICount).ListSubItems.Item(3).ForeColor = vbRed
            .Item(ICount).ListSubItems.Item(3).Bold = True
        End If
    End With
    
    ICount = ICount + 1

    'lParam.Sorted = True
    'End If
    WndEnumChildProc = 1

End Function

Function GetText(iHwnd As Long) As String
    Dim Textlen As Long
    Dim Text As String

    Textlen = SendMessage(iHwnd, WM_GETTEXTLENGTH, 0, 0)
    If Textlen = 0 Then
        GetText = ">No text for this class<"
        Exit Function
    End If
    Textlen = Textlen + 1
    Text = Space(Textlen)
    Textlen = SendMessage(iHwnd, WM_GETTEXT, Textlen, ByVal Text)
    'The 'ByVal' keyword is necessary or you'll get an invalid page fault
    'and the app crashes, and takes VB with it.
    GetText = Left(Text, Textlen)

End Function

Attribute VB_Name = "Task_Modifier"

Option Explicit
Private Type WNDCLASSEX     ' Same as WNDCLASS but has a few advanced values
    cbSize As Long
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long               ' Handle to large icon (Alt-Tab icon)
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long             ' Handle to Small icon (Top Left Icon/Taskbar Icon)
End Type

Private Type WNDCLASS
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long               ' Handle to icon (only 1 size)
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

'API Types

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long
Public Declare Function GetAncestor Lib "user32.dll" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Integer, ByVal Y As Integer, ByVal hIcon As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetClassInfoEx Lib "user32" Alias "GetClassInfoExA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASSEX) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const WM_DESTROY As Long = &H2
Private Const WM_CLOSE As Long = &H10
Private Const WM_SYSCOMMAND As Long = &H112
Private Const SC_CLOSE As Long = &HF060
Public Const WM_GETICON As Long = &H7F
Public Const GWL_HINSTANCE As Long = -6
Public Const GCL_HICON As Long = -14
Public Const GA_ROOT As Long = 2
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_NOTOPMOST As Long = -2
Public Const SWP_NOMOVE As Long = 2
Public Const SWP_NOSIZE As Long = 1
Public Const AOT_Flags As Long = SWP_NOMOVE Or SWP_NOSIZE

Public curhwnd As Long 'Our Current Task hwnd
Public TreeX As New MainNode
Public SelectedNodeKey As String
Public SelectedNodeHwnd As Long
Public TaskMenuID As Long 'used to show which option is open and to make sure form doesnt close until =0
Public Showicons As Boolean 'used on TaskTree(treeview) to show the icons
Public AuthorMode As Boolean 'used for read-only and authormode
Public SearchForWindows As Boolean 'this is used for the search timer
Public LaunchPar As Byte
Public LaunchFile As String

'i use this to make a delay in my apps
Public Sub Delay(Seconds As Integer)

  Dim S As Single

    S = Timer
    Do
        DoEvents
    Loop Until Timer - S > Seconds

End Sub

'my method of finding the best icon to use for my treeview
Public Function DetermineBestIcon(hwnd) As Long

  Dim iconh As Long
  Dim RetLen As Integer
  Dim sysdirbuff As String

    iconh = GetIconHandle(GetAncestor(hwnd, GA_ROOT))
    If iconh = 0 Then
        iconh = GrabIcon("t" & hwnd)
    End If
    If iconh = 0 Then
        sysdirbuff = String$(255, 0)
        RetLen = GetSystemDirectory(sysdirbuff, 255)
        sysdirbuff = Left$(sysdirbuff, RetLen)
        iconh = GrabIconFromFile(sysdirbuff & "\shell32.dll", 2)
    End If
    DetermineBestIcon = iconh

End Function

' a function i created that ive seen many people use in their apps.(variables and all)
'a loop to get a list of the child controls from each parent control
Public Function GetAllChildren(curhwnd As Long) As Long 'Called By LoadTaskList

  Dim Curhwn As Long
  Dim tmpcounter As Long

    Curhwn = GetWindow(curhwnd, 5)
    Do
        If Curhwn Then
            tmpcounter = tmpcounter + 1
            TreeX.AddNode Curhwn
            GetAllChildren Curhwn
            Curhwn = GetWindow(Curhwn, 2)
          Else
            Exit Do
        End If
    Loop
    GetAllChildren = tmpcounter

End Function

'my way of either showing the text or classname of a control to make it userfriendly
Public Function GetFriendlyName(Curhwn As Long) As String

  Dim SClassName As String * 255

    GetFriendlyName = GetText(Curhwn)
    If GetFriendlyName = "" Then
        GetClassName CLng(Curhwn), SClassName, 255
        GetFriendlyName = Left$(SClassName, InStr(SClassName, Chr$(0)) - 1)
    End If

End Function

'find a handle of a icon used by a form
Private Function GetIconHandle(hwnd As Long) As Long
  
  Dim ClassName As String
  Dim WCX As WNDCLASSEX
  Dim hInstance As Long
  Dim hIcon As Long
  Dim X As Long   ' temp variable
  Dim WC As WNDCLASS

    'Method: SendMessage (Small Icon)
    hIcon = SendMessage(hwnd, WM_GETICON, CLng(0), CLng(0))
  
    If hIcon > 0 Then ' found it
        GetIconHandle = hIcon
        Exit Function
    End If
    'Method: SendMessage (Large Icon)
    hIcon = SendMessage(hwnd, WM_GETICON, CLng(1), CLng(0))
    If hIcon > 0 Then ' found it
        GetIconHandle = hIcon
        Exit Function
    End If
    'Method: GetClassInfoEx (Small or Large with Small Pref.)
    hInstance = GetWindowLong(hwnd, GWL_HINSTANCE)
    WCX.cbSize = Len(WCX)
    ClassName = Space$(255)
    X = GetClassName(hwnd, ClassName, 255)
    X = GetClassInfoEx(hInstance, ClassName, WCX)
    If X > 0 Then
        If WCX.hIconSm = 0 Then 'No small icon
            hIcon = WCX.hIcon ' No small icon.. Windows should have given default.. weird
          Else
            hIcon = WCX.hIconSm ' Small Icon is better
        End If
        GetIconHandle = hIcon   ' found it =]
        Exit Function
    End If
 
    '*************************************
    'Method: GetClassInfo (Large Icon)
    '*************************************
    X = GetClassInfo(hInstance, ClassName, WC)
    If X > 0 Then
        hIcon = WC.hIcon
        GetIconHandle = hIcon
        Exit Function    ' Found it
    End If
        
    '*************************************
    'Method: GetClassLong (Large Icon)
    '*************************************
    X = GetClassLong(hwnd, GCL_HICON)
    If X > 0 Then
        hIcon = X
      Else
        hIcon = 0
    End If

    If hIcon < 0 Then
        hIcon = 0
    End If
    GetIconHandle = hIcon

End Function

Private Function GrabIcon(Optional ay = "") As Long

  Dim cc As Long
  Dim iconmod As String, numicons As Long
  Dim hModule As Long, iconh As Long
  Dim mainhwnd As Long

    If ay = "" Then
        cc = mainhwnd
      Else
        cc = CLng(Mid$(ay, 2, Len(ay) - 1))
    End If
    hModule = GetModuleHandle(0)
    iconmod$ = GetExeFromHandle(cc) + Chr$(0)  'prepares filename
    iconh = ExtractIcon(hModule, iconmod, -1) 'gets number of icons
    numicons = iconh - 1 'puts it into a variable
    If numicons > 0 Then
        iconh = ExtractIcon(hModule, iconmod, 0)     'Extracts the first icon
    End If
    GrabIcon = iconh

End Function

'extracts an icon from a file
Private Function GrabIconFromFile(File_name As String, IconNumber As Long) As Long

    GrabIconFromFile = ExtractIcon(GetModuleHandle(0), File_name, IconNumber)

End Function

'checks to see if a string consists of only numbers
Public Function IsStringNumeric(iString As String) As Boolean

  'Since I couldnt find a "IsCharNumeric" API i decided to manipulate
  'the 2 other functions.
  
  Dim iByteArray() As Byte, i As Long

    iByteArray = StrConv(iString, vbFromUnicode)
    For i = 0 To UBound(iByteArray)
        IsStringNumeric = (IsCharAlpha(iByteArray(i)) = 0) And IsCharAlphaNumeric(iByteArray(i))
        If IsStringNumeric = False Then
            Exit Function
        End If
    Next i

End Function

'a few attepts to close a window
Public Function KillWindow(hwnd As Long) As Integer

  'Close A Window. After 1 second then Send a Destroy Command.
  'Note:  Im not sure if its done properly, but it doesnt seem to hurt anything.
  'Also you should not destroy an Explorer Window. (that is the reason i use Close)
  '0=failed to end
  '1=closed
  '2=destroyed
  '3=terminated process

  Dim ProcID As Long
  Dim CloseState As Integer

    CloseState = 1 'Close
    SendMessage hwnd, WM_SYSCOMMAND, SC_CLOSE, 0&   'For Certain Window Types that should not be Destroyed
    Delay 1
    If IsWindow(hwnd) Then
        CloseState = 2 'Destroy
        SendMessage hwnd, WM_DESTROY, 0, 0
        Delay 1
    End If

    If IsWindow(hwnd) Then
        CloseState = 3 'Terminate
        Get_Thread_ProcessID hwnd, ProcID
        EndProcess ProcID
        Delay 1
    End If

    If IsWindow(hwnd) Then
        CloseState = 0 'didnt close
    End If
    KillWindow = CloseState

End Function

'i use this to set the values for my controls on my form
Public Sub SetProps(MyObj As Object, Optional iTop As Long, Optional iLeft As Long, Optional iWidth As Long, Optional iHeight As Long, Optional iVisible As Boolean, Optional iLocked As Boolean, Optional iText As String)

    With MyObj
        If IsMissing(iTop) = False Then
            .Top = iTop
        End If
        If IsMissing(iLeft) = False Then
            .Left = iLeft
        End If
        If IsMissing(iWidth) = False Then
            .Width = iWidth
        End If
        If IsMissing(iVisible) = False Then
            .Visible = iVisible
        End If
        If (TypeOf MyObj Is ComboBox) = False Then
            If IsMissing(iHeight) = False Then
                .Height = iHeight
            End If
          Else
            If IsMissing(iLocked) = False Then
                .Enabled = Not iLocked
            End If
        End If
        If (TypeOf MyObj Is CheckBox) Then
            If IsMissing(iLocked) = False Then
                .Enabled = Not iLocked
            End If
        End If
        If TypeOf MyObj Is TextBox Then
            If IsMissing(iText) = False Then
                .Text = iText
            End If
            If IsMissing(iLocked) = False Then
                .Locked = iLocked
            End If
        End If
        If (TypeOf MyObj Is Label) Or (TypeOf MyObj Is CommandButton) Or (TypeOf MyObj Is CheckBox) Then
            If IsMissing(iText) = False Then
                .Caption = iText
            End If
        End If
    End With

End Sub

'used to set the textbox properties on my form
Public Sub SetPropsText(MyObj As TextBox, Optional tBold As Boolean, Optional tColor As Long)

    With MyObj
        If IsMissing(tBold) = False Then
            .FontBold = tBold
        End If
        If IsMissing(tColor) = False Then
            .ForeColor = tColor
        End If
    End With

End Sub

'callback procedure used to sort thru nodes and change the node color to red if control no longer exists
Public Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)

  Dim i As Long

    For i = 2 To frmMain.TaskTree.Nodes.Count
        DoEvents
        If IsWindow(CLng(Mid$(frmMain.TaskTree.Nodes.item(i).Key, 2))) = 0 Then
            frmMain.TaskTree.Nodes.item(i).ForeColor = RGB(255, 0, 0)
        End If
    Next i

End Sub

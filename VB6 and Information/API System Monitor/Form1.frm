VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   5505
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   1560
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MemBarP 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   360
      ScaleHeight     =   2445
      ScaleWidth      =   120
      TabIndex        =   2
      ToolTipText     =   "Processor Usage (%)"
      Top             =   0
      Width           =   120
   End
   Begin VB.PictureBox MemBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   180
      ScaleHeight     =   2445
      ScaleWidth      =   120
      TabIndex        =   1
      ToolTipText     =   "Processor Usage (%)"
      Top             =   0
      Width           =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   180
      Top             =   2760
   End
   Begin VB.PictureBox cpuBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   0
      ScaleHeight     =   2445
      ScaleWidth      =   120
      TabIndex        =   0
      ToolTipText     =   "Processor Usage (%)"
      Top             =   0
      Width           =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   180
      Top             =   3420
   End
   Begin VB.Menu mMnu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mDelay 
         Caption         =   "Delay"
         Begin VB.Menu m5 
            Caption         =   "5 sec"
         End
         Begin VB.Menu m4 
            Caption         =   "4 sec"
         End
         Begin VB.Menu m3 
            Caption         =   "3 sec"
         End
         Begin VB.Menu m2 
            Caption         =   "2 sec"
         End
         Begin VB.Menu m1 
            Caption         =   "1 sec"
         End
         Begin VB.Menu m750 
            Caption         =   "750ms"
         End
         Begin VB.Menu m500 
            Caption         =   "500ms"
         End
         Begin VB.Menu m250 
            Caption         =   "250ms"
         End
         Begin VB.Menu m100 
            Caption         =   "100ms"
         End
         Begin VB.Menu m50 
            Caption         =   "50ms"
         End
      End
      Begin VB.Menu mPos 
         Caption         =   "Position"
         Begin VB.Menu mTL 
            Caption         =   "Top-Left"
         End
         Begin VB.Menu mTR 
            Caption         =   "Top-Right"
         End
         Begin VB.Menu mBL 
            Caption         =   "Bottom-Left"
         End
         Begin VB.Menu mBR 
            Caption         =   "Bottom-Right"
         End
      End
      Begin VB.Menu mSize 
         Caption         =   "Size"
         Begin VB.Menu mWM 
            Caption         =   "W"
            Begin VB.Menu mW 
               Caption         =   "1"
               Index           =   1
            End
            Begin VB.Menu mW 
               Caption         =   "2"
               Index           =   2
            End
            Begin VB.Menu mW 
               Caption         =   "3"
               Index           =   3
            End
            Begin VB.Menu mW 
               Caption         =   "4"
               Index           =   4
            End
         End
         Begin VB.Menu mHM 
            Caption         =   "H"
            Begin VB.Menu mH 
               Caption         =   "1"
               Index           =   1
            End
            Begin VB.Menu mH 
               Caption         =   "2"
               Index           =   2
            End
            Begin VB.Menu mH 
               Caption         =   "3"
               Index           =   3
            End
            Begin VB.Menu mH 
               Caption         =   "4"
               Index           =   4
            End
         End
      End
      Begin VB.Menu mOnTop 
         Caption         =   "On Top"
         Checked         =   -1  'True
      End
      Begin VB.Menu mASm 
         Caption         =   "AutoStart"
         Begin VB.Menu mAS 
            Caption         =   "Run under User"
            Index           =   0
         End
         Begin VB.Menu mAS 
            Caption         =   "Run as Service"
            Index           =   1
         End
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'for the always on top function
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40

'Registry API
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

'unused
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

'reg constants
Private Const REG_SZ = 1        'a null terminated string
Private Const REG_DWORD = 4     'a double word is 4 bytes (A.K.A. Long variable)
Private Const HKEY_DYN_DATA = &H80000006    'reg key root
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_SET_VALUE = &H2
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&

'for regcreatekey, but unused
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Dim regkey() As Long  'used to keep open a reg key throughout the program
Dim last As Long      'last processor usage %
Dim lastavg As Long   'last avg position
Dim sum As Double     'used to calc the avg
Dim cnt As Double     'used to calc the avg
Dim stime As Single   'start time
Dim WinDir As String  'windows directory

'for the memory
Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type
Dim memoryInfo As MEMORYSTATUS
Dim lastpcent As Single, lastTot As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

'Close the CPU meter stats
Public Function CloseCPU() As Long
    Dim data As Long
    Dim hret As Long

    hret = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StopStat", regkey(4))
    hret = RegQueryValueEx(regkey(4), "KERNEL\CPUUsage", 0&, REG_DWORD, data, 4)
    hret = RegCloseKey(regkey(4))
    hret = RegCloseKey(regkey(0))
End Function

'Initialize the CPU meter stats
Public Function InitializeCPU() As Long
    Dim data As Long
    Dim hret As Long

    hret = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", regkey(0))
    hret = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartSrv", regkey(1))
    hret = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StopSrv", regkey(2))
    hret = RegQueryValueEx(regkey(1), "KERNEL", 0&, REG_DWORD, data, 4)
    hret = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", regkey(3))
    hret = RegQueryValueEx(regkey(3), "KERNEL\CPUUsage", 0&, REG_DWORD, data, 4)
    hret = RegCloseKey(regkey(3))
    hret = RegQueryValueEx(regkey(2), "KERNEL", 0&, REG_DWORD, data, 4)
    hret = RegCloseKey(regkey(1))
    hret = RegCloseKey(regkey(2))
    
End Function

'Get the cpu info via gfx meter
Public Function GetCPUUsage() As Long
    Dim data As Long
    Dim hret As Long

    hret = RegQueryValueEx(regkey(0), "KERNEL\CPUUsage", 0&, REG_DWORD, data, 4)
    GetCPUUsage = data
End Function

'exit by double clicking
Private Sub cpuBar_DblClick()
  Unload Me
  End
End Sub

Public Sub Form_Resize()
  cpuBar.Top = Screen.TwipsPerPixelY
  cpuBar.Left = Screen.TwipsPerPixelX
  MemBar.Top = Screen.TwipsPerPixelY
  MemBar.Left = cpuBar.Left + cpuBar.width + Screen.TwipsPerPixelX
  MemBarP.Top = Screen.TwipsPerPixelY
  MemBarP.Left = MemBar.Left + MemBar.width + Screen.TwipsPerPixelX
  Me.Move 0, 0, MemBarP.Left + MemBarP.width + Screen.TwipsPerPixelX, cpuBar.height
  'if the left is >0 then bit 1 is set, if the top is >0 then bit 0 is set
  pos = GetSetting("SYSMON", "SETTINGS", "POSITION", 0)
  Select Case pos
    Case 0:  Me.Move 0, 0
             mTL.Checked = True
    Case 1:  Me.Move 0, Screen.height - Me.height
             mBL.Checked = True
    Case 2:  Me.Move Screen.width - Me.width, 0
             mTR.Checked = True
    Case 3:  Me.Move Screen.width - Me.width, Screen.height - Me.height
             mBR.Checked = True
  End Select
  Call Form_Paint
End Sub

Private Sub Form_Terminate()
  Call CloseCPU
  SaveSetting "SYSMON", "SETTINGS", "DELAY", Timer1.Interval
  'if the left is >0 then bit 1 is set, if the top is >0 then bit 0 is set
  SaveSetting "SYSMON", "SETTINGS", "POSITION", IIf(Me.Left > 0, 2, 0) + IIf(Me.Top > 0, 1, 0)
  SaveSetting "SYSMON", "SETTINGS", "ONTOP", mOnTop.Checked
  SaveSetting "SYSMON", "SETTINGS", "AUTOSTART_R", mAS(0).Checked
  SaveSetting "SYSMON", "SETTINGS", "AUTOSTART_S", mAS(1).Checked

End Sub

Private Sub MemBarP_DblClick()
  Unload Me
  End
End Sub

Private Sub MemBar_DblClick()
  Unload Me
  End
End Sub

Private Sub Form_DblClick()
  Unload Me
  End
End Sub

'bring up popup menu
Private Sub cpuBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    Me.PopupMenu mMnu
  End If
End Sub

Private Sub MemBarP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    Me.PopupMenu mMnu
  End If
End Sub

Private Sub MemBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    Me.PopupMenu mMnu
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    Me.PopupMenu mMnu
  End If
End Sub

Private Sub Form_Load()
  Dim use As Long, pos As Byte, index As Integer, hKey As Long
    
  'get the windows directory
  Dim lResult As Long
  Dim lValueType As Long
  Dim lDataBufSize As Long
  Call RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", hKey)
  ' Get length/data type
  lResult = RegQueryValueEx(hKey, "SystemRoot", 0&, lValueType, ByVal 0&, lDataBufSize)
  If lResult = ERROR_SUCCESS Then
      If lValueType = REG_SZ Then
          WinDir = String(lDataBufSize, " ")
          lResult = RegQueryValueEx(hKey, "SystemRoot", 0&, 0&, ByVal WinDir, lDataBufSize)
          If lResult = ERROR_SUCCESS Then
              RegQueryStringValue = True
              If InStr(WinDir, Chr(0)) > 1 Then
                WinDir = Left(WinDir, InStr(WinDir, Chr(0)) - 1)
              Else
                WinDir = "C:\WINDOWS"
              End If
          End If
      End If
  End If
  Call RegCloseKey(hKey)
  
  'set the processor monitor's update interval
  Timer1.Interval = GetSetting("SYSMON", "SETTINGS", "DELAY", 250)
  
  'set the width of the meters
  index = GetSetting("SYSMON", "SETTINGS", "WIDTH", 2)
  For i% = 1 To 4
    mW(i%).Checked = False
  Next i%
  mW(index).Checked = True
  Call setbarwidth(index * 3)
  
  'set the height of the meters
  index = GetSetting("SYSMON", "SETTINGS", "HEIGHT", 3)
  For i% = 1 To 4
    mH(i%).Checked = False
  Next i%
  mH(index).Checked = True
  Call setbarheight(index ^ 2 * 15 + 30)
  
  'check the proper delay menu item
  Select Case Timer1.Interval
    Case 50: m50.Checked = True
    Case 100: m100.Checked = True
    Case 250: m250.Checked = True
    Case 500: m500.Checked = True
    Case 750: m750.Checked = True
    Case 1000: m1.Checked = True
    Case 2000: m2.Checked = True
    Case 3000: m3.Checked = True
    Case 4000: m4.Checked = True
    Case 5000: m5.Checked = True
  End Select
  ReDim regkey(4)     'we use 5 different reg keys
  Call InitializeCPU  'tell windows we want it to monitor the processor
  use = GetCPUUsage() 'get the % usage of the processor
  sum = use * (Timer1.Interval \ 50)
  cnt = Timer1.Interval \ 50
  
  'get the average position
  lastavg = cpuBar.height - cpuBar.height * (sum / cnt / 100)
  last = use
  stime = Time
  
  'get memory info
  GlobalMemoryStatus memoryInfo
  lastpcent = Int((Int(memoryInfo.dwAvailPhys / 1044032 * 10 + 0.5) / 10) / (Int(memoryInfo.dwTotalPhys / 1044032 * 10 + 0.5) / 10) * 100)
  lastTot = memoryInfo.dwMemoryLoad
  
  'set tooltip text's
  MemBarP.ToolTipText = "Physical Mem Free: " & Format(Int(memoryInfo.dwTotalPhys / 1044032 * 10 + 0.5) / 10) & " MB of " & Format(Int(memoryInfo.dwAvailPhys / 1044032 * 10 + 0.5) / 10) & " MB (" & Format(lastpcent) & "%)"
  MemBar.ToolTipText = "Total Mem Free: " & Format(100 - lastTot) & "%"
  
  Timer1.Enabled = True  'turn on the timer
  
  'get some settings
  mAS(0).Checked = GetSetting("SYSMON", "SETTINGS", "AUTOSTART_R", False)
  mAS(1).Checked = GetSetting("SYSMON", "SETTINGS", "AUTOSTART_S", False)
  mOnTop.Checked = GetSetting("SYSMON", "SETTINGS", "ONTOP", True)
  Call Form_Resize
  Call AlwaysOnTop(Me, mOnTop.Checked)
End Sub

'repaints the meters (bars) by clearing them and redrawing to the full values
Private Sub Form_Paint()
  On Error Resume Next
  cpuBar.Cls
  cpuBar.Line (0, cpuBar.height)-(cpuBar.width, cpuBar.height - cpuBar.height * (last / 100)), &HC000&, BF
  cpuBar.Line (0, cpuBar.height - cpuBar.height * (sum / cnt / 100) - Screen.TwipsPerPixelY)-(cpuBar.width, cpuBar.height - cpuBar.height * (sum / cnt / 100) + Screen.TwipsPerPixelY), &HFFFFFF, BF
  MemBarP.Cls
  MemBarP.Line (0, MemBarP.height)-(MemBarP.width, MemBarP.height - MemBarP.height * lastpcent / 100), &HC0&, BF
  MemBar.Cls
  MemBar.Line (0, MemBar.height)-(MemBar.width, MemBar.height * lastTot / 100), &HC00000, BF
End Sub

'save the settings before exiting
Private Sub Form_Unload(Cancel As Integer)
  Call CloseCPU
  SaveSetting "SYSMON", "SETTINGS", "DELAY", Timer1.Interval
  'if the left is >0 then bit 1 is set, if the top is >0 then bit 0 is set
  SaveSetting "SYSMON", "SETTINGS", "POSITION", IIf(Me.Left > 0, 2, 0) + IIf(Me.Top > 0, 1, 0)
  SaveSetting "SYSMON", "SETTINGS", "ONTOP", mOnTop.Checked
  SaveSetting "SYSMON", "SETTINGS", "AUTOSTART_R", mAS(0).Checked
  SaveSetting "SYSMON", "SETTINGS", "AUTOSTART_S", mAS(1).Checked
End Sub

'delays...
Private Sub m50_Click()
  Call uncheckall
  m50.Checked = True
  Timer1.Enabled = False
  Timer1.Interval = 50
  Timer1.Enabled = True
End Sub

Private Sub m100_Click()
  Call uncheckall
  m100.Checked = True
  Timer1.Enabled = False
  Timer1.Interval = 100
  Timer1.Enabled = True
End Sub

Private Sub m250_Click()
  Call uncheckall
  m250.Checked = True
  Timer1.Enabled = False
  Timer1.Interval = 250
  Timer1.Enabled = True
End Sub

Private Sub m500_Click()
  Call uncheckall
  m500.Checked = True
  Timer1.Enabled = False
  Timer1.Interval = 500
  Timer1.Enabled = True
End Sub

Private Sub m750_Click()
  Call uncheckall
  m750.Checked = True
  Timer1.Enabled = False
  Timer1.Interval = 750
  Timer1.Enabled = True
End Sub

Private Sub m1_Click()
  Call uncheckall
  m1.Checked = True
  Timer1.Enabled = False
  Timer1.Interval = 1000
  Timer1.Enabled = True
End Sub

Private Sub m2_Click()
  Call uncheckall
  m2.Checked = True
  Timer1.Enabled = False
  Timer1.Interval = 2000
  Timer1.Enabled = True
End Sub

Private Sub m3_Click()
  Call uncheckall
  m3.Checked = True
  Timer1.Enabled = False
  Timer1.Interval = 3000
  Timer1.Enabled = True
End Sub

Private Sub m4_Click()
  Call uncheckall
  m4.Checked = True
  Timer1.Enabled = False
  Timer1.Interval = 4000
  Timer1.Enabled = True
End Sub

Private Sub m5_Click()
  Call uncheckall
  m5.Checked = True
  Timer1.Enabled = False
  Timer1.Interval = 5000
  Timer1.Enabled = True
End Sub

'autostart
Private Sub mAS_Click(index As Integer)
  Dim hKey As Long, data As String, length As Long, retval As Long
  Dim secure As SECURITY_ATTRIBUTES
  secure.nLength = 0
  If mAS(index).Checked = True Then
    'disable the autostart by deleting the registry value
    mAS(index).Checked = False
    If index = 0 Then
      'open run under user key
      Call RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", hKey)
    Else
      'open run as service key
      Call RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", hKey)
    End If
    Call RegDeleteValue(hKey, "System_Monitor")
    Call RegCloseKey(hKey)
      'we could can't delete the file because it's in use if the autostart has been run
      'Call DeleteFile(WinDir & "\SysMon2.exe")
  Else
    'enable the autostart by copying the exe and telling windows to run it as a service
    mAS(index).Checked = True
    'delete the other one if it exists
    If mAS(index Xor 1).Checked = True Then
      mAS(index Xor 1).Checked = False
      If index Xor 1 = 0 Then
        'open run under user key
        Call RegOpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", hKey)
      Else
        'open run as service key
        Call RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", hKey)
      End If
      Call RegDeleteValue(hKey, "System_Monitor")
      Call RegCloseKey(hKey)
    End If
    data = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & App.EXEName & ".exe"
    'copy the exe
    retval = CopyFile(data, WinDir & "\SysMon2.exe", False)
    If retval = 1 Then    'if the copy was successful, i think
      data = WinDir & "\SysMon2.exe"
      length = Len(data) + 1
      'tell windows to run the exe every time it starts
      If index = 0 Then
        Call RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", 0, REG_SZ, 0, KEY_ALL_ACCESS, secure, hKey, retval)
      Else
        Call RegCreateKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", 0, REG_SZ, 0, KEY_ALL_ACCESS, secure, hKey, retval)
      End If
      Call RegSetValueEx(hKey, "System_Monitor", 0, REG_SZ, ByVal data, length)
      Call RegCloseKey(hKey)
    Else
      mAS(index).Checked = False
      MsgBox "unable to copy " & data, vbCritical, "error"
    End If
  End If
End Sub

'set window positions (bottom ones)
Private Sub mBL_Click()
  Call uncheckall2
  mBL.Checked = True
  Me.Move 0, Screen.height - Me.height
End Sub

Private Sub mBR_Click()
  Call uncheckall2
  mBR.Checked = True
  Me.Move Screen.width - Me.width, Screen.height - Me.height
End Sub

'exit
Private Sub mExit_Click()
  Unload Me
  End
End Sub

'change the height of the meters
Private Sub mH_Click(index As Integer)
  For i% = 1 To 4
    mH(i%).Checked = False
  Next i%
  mH(index).Checked = True
  Call setbarheight(index ^ 2 * 15 + 30)
  SaveSetting "SYSMON", "SETTINGS", "HEIGHT", index
End Sub

'set or take off ontop
Private Sub mOnTop_Click()
  If mOnTop.Checked = True Then
    mOnTop.Checked = False
    Call AlwaysOnTop(Me, False)
  Else
    mOnTop.Checked = True
    Call AlwaysOnTop(Me, True)
  End If
End Sub

'set window positions (top ones)
Private Sub mTL_Click()
  Call uncheckall2
  mTL.Checked = True
  Me.Move 0, 0
End Sub

Private Sub mTR_Click()
  Call uncheckall2
  mTR.Checked = True
  Me.Move Screen.width - Me.width, 0
End Sub

'change the width of the meters
Private Sub mW_Click(index As Integer)
  For i% = 1 To 4
    mW(i%).Checked = False
  Next i%
  mW(index).Checked = True
  Call setbarwidth(index * 3)
  SaveSetting "SYSMON", "SETTINGS", "WIDTH", index
End Sub

'updates only the CPU meter
Private Sub Timer1_Timer()
  Dim use As Long, avg As Long, t1 As Long, t2 As Long
  use = GetCPUUsage()  'returns a percentage
  sum = sum + use * (Timer1.Interval \ 50)
  cnt = cnt + Timer1.Interval \ 50
  t1 = cpuBar.height - cpuBar.height * (use / 100)
  avg = cpuBar.height - cpuBar.height * (sum / cnt / 100)
  'only draw or erase the changes to the meter
  If last > use Then
    cpuBar.Line (0, cpuBar.height - cpuBar.height * (last / 100))-(cpuBar.width, t1), &H404040, BF
  Else
    cpuBar.Line (0, cpuBar.height - cpuBar.height * (last / 100))-(cpuBar.width, cpuBar.height - cpuBar.height * (use / 100)), &HC000&, BF
  End If
  
  'erase the old average with the proper color
  If lastavg <> avg And lastavg > t1 Then
    cpuBar.Line (0, lastavg - Screen.TwipsPerPixelY)-(cpuBar.width, lastavg + Screen.TwipsPerPixelY), &HC000&, BF
  ElseIf lastavg <> avg Then
    cpuBar.Line (0, lastavg - Screen.TwipsPerPixelY)-(cpuBar.width, lastavg + Screen.TwipsPerPixelY), &H404040, BF
  End If
  'draw the new average
  cpuBar.Line (0, avg - Screen.TwipsPerPixelY)-(cpuBar.width, avg + Screen.TwipsPerPixelY), &HFFFFFF, BF
  
  'set the last values for the meter
  last = use
  lastavg = avg
End Sub

'this timer updates the time, tooltips, physical mem & total mem bars
Private Sub Timer2_Timer()
  Dim Totp1 As Single, Availp1 As Single, pcent As Single

  GlobalMemoryStatus memoryInfo   'get the memory info
  Totp1 = Int(memoryInfo.dwTotalPhys / 1044032 * 10 + 0.5) / 10
  Availp1 = Int(memoryInfo.dwAvailPhys / 1044032 * 10 + 0.5) / 10
  pcent = Int(Availp1 / Totp1 * 100)
  'only draw or erase what has changed in the physical memory bar
  'the physical mem percent is in % unused, so we calculate from the bottom
  If lastpcent <= pcent Then
    MemBarP.Line (0, MemBarP.height - MemBarP.height * lastpcent / 100)-(MemBarP.width, MemBarP.height - MemBarP.height * pcent / 100), &HC0&, BF
  Else
    MemBarP.Line (0, MemBarP.height - MemBarP.height * lastpcent / 100)-(MemBarP.width, MemBarP.height - MemBarP.height * pcent / 100), &H404040, BF
  End If
  
  'only draw or erase what has changed in the total memory bar
  'the total mem percent is in % used, so we calculate from the top
  If lastTot <= memoryInfo.dwMemoryLoad Then
    MemBar.Line (0, MemBar.height * lastTot / 100)-(MemBar.width, MemBar.height * memoryInfo.dwMemoryLoad / 100), &H404040, BF
  Else
    MemBar.Line (0, MemBar.height * lastTot / 100)-(MemBar.width, MemBar.height * memoryInfo.dwMemoryLoad / 100), &HC00000, BF
  End If
  
  'update last values
  lastpcent = pcent
  lastTot = memoryInfo.dwMemoryLoad
  
  'tool tip text's
  MemBarP.ToolTipText = "Physical Mem Free: " & Format(Availp1) & " MB of " & Format(Totp1) & " MB (" & Format(lastpcent) & "%)"
  MemBar.ToolTipText = "Total Mem Free: " & Format(100 - lastTot) & "%"
  cpuBar.ToolTipText = "Processor Usage (avg: " & Format(Int(sum / cnt)) & "%) - Time: " & Format(Time - stime, "hh:mm:ss")
End Sub

'uncheck all the delay menu settings
Function uncheckall()
  m50.Checked = False
  m100.Checked = False
  m250.Checked = False
  m500.Checked = False
  m750.Checked = False
  m1.Checked = False
  m2.Checked = False
  m3.Checked = False
  m4.Checked = False
  m5.Checked = False
End Function

'uncheck all the position menu settings
Function uncheckall2()
  mBL.Checked = False
  mBR.Checked = False
  mTL.Checked = False
  mTR.Checked = False
End Function

'sets the proper width for all the bars
Function setbarwidth(width As Integer)   'width is in pixels
  width = width * Screen.TwipsPerPixelX  'width now in twips
  cpuBar.width = width
  MemBar.width = width
  MemBarP.width = width
  Call Form_Resize          'redraws the form based on meters
End Function

'sets the proper height for all the bars
Function setbarheight(height As Integer)   'height is in pixels
  height = height * Screen.TwipsPerPixelX  'height now in twips
  cpuBar.height = height
  MemBar.height = height
  MemBarP.height = height
  Call Form_Resize          'redraws the form based on meters
End Function

'place the form on top of everything (if true)
Function AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    'lflag will set on top or not
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.width / Screen.TwipsPerPixelX, _
    myfrm.height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Function



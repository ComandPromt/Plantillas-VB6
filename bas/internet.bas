Attribute VB_Name = "Internet"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOW = 5
Private Const conSwNormal = 1

Private Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
'
Private Const RAS95_MaxEntryName = 256
Private Const RAS95_MaxDeviceType = 16
Private Const RAS95_MaxDeviceName = 32
'
Public Type RASCONN95
    dwSize As Long
    hRasCon As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
'
Public Type RASCONNSTATUS95
    dwSize As Long
    RasConnState As Long
    dwError As Long
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Private Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Long


Public Sub DownloadFile(FileURL As String)
   Dim sDownload As String
   
   sDownload = StrConv(FileURL, vbUnicode)
   Call DoFileDownload(sDownload)
End Sub


Public Property Get IsConnected() As Boolean
Dim TRasCon(255) As RASCONN95
Dim lg As Long
Dim lpcon As Long
Dim RetVal As Long
Dim Tstatus As RASCONNSTATUS95
'
TRasCon(0).dwSize = 412
lg = 256 * TRasCon(0).dwSize
'
RetVal = RasEnumConnections(TRasCon(0), lg, lpcon)
If RetVal <> 0 Then
                    MsgBox "ERROR"
                    Exit Property
                    End If
'
Tstatus.dwSize = 160
RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)
If Tstatus.RasConnState = &H2000 Then
                         IsConnected = True
                         Else
                         IsConnected = False
                         End If

End Property


Public Sub GotoURL(Form As Object, URL As String)
ShellExecute Form.hwnd, "open", URL, vbNullString, vbNullString, conSwNormal
End Sub

    Public Sub SendEmail(Form As Object, EMailAddress As String)
ShellExecute Form.hwnd, "open", "mailto:" & EMailAddress, vbNullString, vbNullString, SW_SHOW
    End Sub
    Public Function IsNewVersion(VersionOfProgram As String, iNetControl As Object, URLofFile As String) As String

On Error GoTo 10

a$ = iNetControl.OpenURL(URLofFile, icString)
Debug.Print a$

If Not a$ = VersionOfProgram Then
IsNewVersion = a$
Else
IsNewVersion = ""
End If
10
    End Function
    
    

Attribute VB_Name = "Winsock"
Option Explicit
Public Const NCBASTAT = &H33
Public Const NCBNAMSZ = 16
Public Const HEAP_ZERO_MEMORY = &H8
Public Const HEAP_GENERATE_EXCEPTIONS = &H4
Public Const NCBRESET = &H32

Public Type NCB
  ncb_command As Byte
  ncb_retcode As Byte
  ncb_lsn As Byte
  ncb_num As Byte
  ncb_buffer As Long
  ncb_length As Integer
  ncb_callname As String * NCBNAMSZ
  ncb_name As String * NCBNAMSZ
  ncb_rto As Byte
  ncb_sto As Byte
  ncb_post As Long
  ncb_lana_num As Byte
  ncb_cmd_cplt As Byte
  ncb_reserve(9) As Byte ' Reserve, doit etre 0
  ncb_event As Long
End Type

Public Type ADAPTER_STATUS
  adapter_address(5) As Byte
  rev_major As Byte
  reserved0 As Byte
  adapter_type As Byte
  rev_minor As Byte
  duration As Integer
  frmr_recv As Integer
  frmr_xmit As Integer
  iframe_recv_err As Integer
  xmit_aborts As Integer
  xmit_success As Long
  recv_success As Long
  iframe_xmit_err As Integer
  recv_buff_unavail As Integer
  t1_timeouts As Integer
  ti_timeouts As Integer
  Reserved1 As Long
  free_ncbs As Integer
  max_cfg_ncbs As Integer
  max_ncbs As Integer
  xmit_buf_unavail As Integer
  max_dgram_size As Integer
  pending_sess As Integer
  max_cfg_sess As Integer
  max_sess As Integer
  max_sess_pkt_size As Integer
  name_count As Integer
End Type

Public Type NAME_BUFFER
  name As String * NCBNAMSZ
  name_num As Integer
  name_flags As Integer
End Type

Public Type ASTAT
  adapt As ADAPTER_STATUS
  NameBuff(30) As NAME_BUFFER
End Type

Const INADDR_NONE = -1
Const PF_INET = 2
Public IPAddress(3) As Byte

Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_DISCARDED = &H4000
Public Const GMEM_FIXED = &H0
Public Const GMEM_INVALID_HANDLE = &H8000
Public Const GMEM_LOCKCOUNT = &HFF
Public Const GMEM_MODIFY = &H80
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Public Const RESOURCE_CONNECTED = &H1
Public Const RESOURCE_GLOBALNET = &H2
Public Const RESOURCE_REMEMBERED = &H3
Public Const RESOURCE_CONTEXT = &H4
Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEDISPLAYTYPE_SERVER = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_UNKNOWN = &HFFFF
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const RESOURCEUSAGE_CONTAINER = &H2
Public Const RESOURCEUSAGE_RESERVED = &H80000000

Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type

Public Type NETRES2 ' NETRESOURCE compatible VB
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Public Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Public Type T_WSA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public WSAData As T_WSA

Public Type SHARE_INFO_0
    shi0_netname As String
End Type

Public Type SHARE_INFO_2
    shi2_netname As String
    shi2_type As Long
    shi2_remark As String
    shi2_permissions As Long
    shi2_max_uses As Long
    shi2_current_uses As Long
    shi2_path As String
    shi2_passwd As String
End Type

Public Declare Function Netbios Lib "netapi32.dll" (pncb As NCB) As Byte
Public Declare Function GetProcessHeap Lib "kernel32" () As Long
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Public Declare Function WSAStartup Lib "wsock32.dll" _
   (ByVal wVersionRequired As Long, lpWSADATA As T_WSA) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function gethostname Lib "wsock32.dll" _
   (ByVal szHost As String, ByVal dwHostLen As Long) As Long

Public Declare Function GlobalAlloc Lib "kernel32" _
    (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function CopyPointer2String Lib "kernel32" Alias "lstrcpyA" _
    (ByVal NewString As String, ByVal OldString As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Public Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" _
    (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, _
    lpNetResource As Any, lphEnum As Long) As Long
Public Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" _
    (ByVal hEnum As Long, lpcCount As Long, ByVal lpBuffer As Long, lpBufferSize As Long) As Long
Public Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Public Declare Function NetShareGetInfo Lib "netapi32.dll" _
    (ByRef servername As String, ByRef netname As String, ByVal level As Long, ByRef bufptr As SHARE_INFO_0) As Long

Public Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal HostName As String) As Long
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal addr As String) As Long

Public Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Public Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpName As String, ByVal bfoce As Boolean) As Long

Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long 'username only

Public Function EthernetAddress(LanaNumber As Long) As String
Dim udtNCB       As NCB
Dim bytResponse  As Byte
Dim udtASTAT     As ASTAT
Dim udtTempASTAT As ASTAT
Dim lngASTAT     As Long
Dim strOut       As String
Dim i            As Integer

    udtNCB.ncb_command = NCBRESET
    bytResponse = Netbios(udtNCB)
    udtNCB.ncb_command = NCBASTAT
    udtNCB.ncb_lana_num = LanaNumber
    udtNCB.ncb_callname = "* "
    udtNCB.ncb_length = Len(udtASTAT)
    lngASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, udtNCB.ncb_length)
    strOut = ""
    If lngASTAT Then
        udtNCB.ncb_buffer = lngASTAT
        bytResponse = Netbios(udtNCB)
        CopyMemory udtASTAT, udtNCB.ncb_buffer, Len(udtASTAT)
        With udtASTAT.adapt
            For i = 0 To 5
                strOut = strOut & Right$("00" & Hex$(.adapter_address(i)), 2)
            Next i
        End With
        HeapFree GetProcessHeap(), 0, lngASTAT
    End If
    EthernetAddress = strOut
End Function

Public Function GetCurrentUser()
Dim lpBuff As String * 25
Dim ret As Long
ret = GetUserName(lpBuff, 25)
GetCurrentUser = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
Utilisateur = GetCurrentUser
End Function

Function WinsockInit() As Boolean
    WinsockInit = Not WSAStartup(&H101, WSAData)
End Function

Public Function GetIPAddress() As String
   Dim sHostName    As String * 256
   Dim lpHost    As Long
   Dim Host      As HOSTENT
   Dim dwIPAddr  As Long
   Dim tmpIPAddr() As Byte
   Dim i         As Integer
   Dim sIPAddr  As String
   
   If Not WinsockInit Then
      GetIPAddress = ""
      Exit Function
   End If
   If gethostname(sHostName, 256) = SOCKET_ERROR Then
      GetIPAddress = ""
      MsgBox "Erreur Socket " & Str$(WSAGetLastError()) & _
              " . Get Host Name impossible."
      WSACleanup
      Exit Function
   End If
   sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
    
   If lpHost = 0 Then
      GetIPAddress = ""
      MsgBox "Socket Windows ne repond pas. " & _
              "Get Host by Name impossible."
      WSACleanup
      Exit Function
   End If
   CopyMemory Host, lpHost, Len(Host)
   CopyMemory dwIPAddr, Host.hAddrList, 4
   ReDim tmpIPAddr(1 To Host.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, Host.hLen
   For i = 1 To Host.hLen
      sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next
   GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   
   WSACleanup
    
End Function
Public Function GetIPHostName() As String

    Dim sHostName As String * 256
    
    If Not WinsockInit Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Erreur Socket Windows " & Str$(WSAGetLastError()) & _
                " . Get Host Name impossible."
        WSACleanup
        Exit Function
    End If
    
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    WSACleanup

End Function

Public Function DottedIPToDNS(ByVal sAddress As String) As String
    Dim lAddress As Long
    Dim p As Long
    Dim HostName As String
    Dim Host As HOSTENT

    lAddress = inet_addr(sAddress)
    p = gethostbyaddr(lAddress, 4, PF_INET)
    If p <> 0 Then
        CopyMemory Host, ByVal p, Len(Host)
        HostName = String(256, 0)
        CopyMemory ByVal HostName, ByVal Host.hName, 256
        If HostName = "" Then
            DottedIPToDNS = "DNS erreur : resolution impossible " & Str$(WSAGetLastError())
        Else
            DottedIPToDNS = Left(HostName, InStr(HostName, Chr(0)) - 1)
        End If
    Else
        DottedIPToDNS = "Pas de nom"
    End If
End Function

Function DNSToDottedIP(sHost As String) As String
    Dim p As Long
    Dim Host As HOSTENT
    Dim lpListAddress As Long
    Dim FirstAddress As Long
    Dim Address As Long

    sHost = sHost & String(64 - Len(sHost), 0)
    p = gethostbyname(sHost)
    If p = SOCKET_ERROR Then
        Exit Function
    Else
        If p <> 0 Then
            CopyMemory Host, ByVal p, Len(Host)
            lpListAddress = Host.hAddrList
            DNSToDottedIP = ""
' Next three strings allow receive only first IP address
 CopyMemory FirstAddress, ByVal lpListAddress, 4
 CopyMemory Address, ByVal FirstAddress, 4
 DNSToDottedIP = LongIPToDotted(FirstAddress)
' *************
' One DNS can contain a list of IP. If you need all
' addresses, you can make a loop here, increasing
' lpListAddres by 4 every time
' If You want to do so, comment three strings above my comments
' and uncomment Loop below. You must change Multiline
' property of Text1(1) to True and add vertical scrollbar
' to see all addresses in this text box
' ************
' Do
'  CopyMemory FirstAddress, ByVal lpListAddress, 4
'  CopyMemory Address, ByVal FirstAddress, 4
'  If Address = 0 Then Exit Do
'  DNSToDottedIP = DNSToDottedIP & LongIPToDotted(Address) & vbCrLf
'  lpListAddress = lpListAddress + 4
' Loop
        Else
            DNSToDottedIP = "Pas de nom"
        End If
    End If
End Function

Function DottedIPToLong(Address As String) As String
  Dim lTmp As Long
  lTmp = inet_addr(Address)
  If lTmp = INADDR_NONE Then
     DottedIPToLong = "Adresse incorrecte"
  Else
     DottedIPToLong = CStr(lTmp)
  End If
End Function

Function LongIPToDotted(Address As Long) As String
  Dim i As Integer, sTmp As String
  sTmp = ""
  WSACleanup
  CopyMemory IPAddress(0), Address, 4
  For i = 0 To 3
    sTmp = sTmp & CStr(IPAddress(i)) & "."
  Next i
  LongIPToDotted = Left$(sTmp, Len(sTmp) - 1)
End Function




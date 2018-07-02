Attribute VB_Name = "PingIP"
Public Const IP_STATUS_BASE = 11000
Public Const IP_SUCCESS = 0
Public Const IP_BUF_TOO_SMALL = (11000 + 1)
Public Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Public Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Public Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Public Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Public Const IP_NO_RESOURCES = (11000 + 6)
Public Const IP_BAD_OPTION = (11000 + 7)
Public Const IP_HW_ERROR = (11000 + 8)
Public Const IP_PACKET_TOO_BIG = (11000 + 9)
Public Const IP_REQ_TIMED_OUT = (11000 + 10)
Public Const IP_BAD_REQ = (11000 + 11)
Public Const IP_BAD_ROUTE = (11000 + 12)
Public Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Public Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Public Const IP_PARAM_PROBLEM = (11000 + 15)
Public Const IP_SOURCE_QUENCH = (11000 + 16)
Public Const IP_OPTION_TOO_BIG = (11000 + 17)
Public Const IP_BAD_DESTINATION = (11000 + 18)
Public Const IP_ADDR_DELETED = (11000 + 19)
Public Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Public Const IP_MTU_CHANGE = (11000 + 21)
Public Const IP_UNLOAD = (11000 + 22)
Public Const IP_ADDR_ADDED = (11000 + 23)
Public Const IP_GENERAL_FAILURE = (11000 + 50)
Public Const MAX_IP_STATUS = 11000 + 50
Public Const IP_PENDING = (11000 + 255)
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
'Public Const PING_TIMEOUT = 500
Public timeout_ping As Integer

Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Integer
    Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Public Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long
   
Public Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Integer, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal TimeOut As Long) As Long

Public Function Ping(szAddress As String, ECHO As ICMP_ECHO_REPLY) As Long

   Dim hPort As Long
   Dim dwAddress As Long
   Dim sDataToSend As String
   Dim iOpt As Long
   
   sDataToSend = "kikoo"
   dwAddress = AddressStringToLong(szAddress)
   
   hPort = IcmpCreateFile()
   
    If IcmpSendEcho(hPort, _
                   dwAddress, _
                   sDataToSend, _
                   Len(sDataToSend), _
                   0, _
                   ECHO, _
                   Len(ECHO), _
                   timeout_ping) Then
        Ping = ECHO.RoundTripTime
   Else
        Ping = ECHO.status * -1
   End If
                       
   Call IcmpCloseHandle(hPort)
   
End Function

Public Function AddressStringToLong(ByVal tmp As String) As Long

   Dim i As Integer
   Dim parts(1 To 4) As String
   
   i = 0
   
   While InStr(tmp, ".") > 0
      i = i + 1
      parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
      tmp = Mid(tmp, InStr(tmp, ".") + 1)
   Wend
   
   i = i + 1
   parts(i) = tmp
   
   If i <> 4 Then
      AddressStringToLong = 0
      Exit Function
   End If
   
  'build the long value out of the
  'hex of the extracted strings
   AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & _
                         Right("00" & Hex(parts(3)), 2) & _
                         Right("00" & Hex(parts(2)), 2) & _
                         Right("00" & Hex(parts(1)), 2))
   
End Function

Public Function HiByte(ByVal wParam As Integer)

    HiByte = wParam \ &H1 And &HFF&

End Function

Public Function LoByte(ByVal wParam As Integer)

    LoByte = wParam And &HFF&

End Function

Public Function SocketsCleanup() As Boolean

    Dim X As Long
    
    X = WSACleanup()
    
    If X <> 0 Then
        MsgBox "Windows Sockets error " & Trim$(Str$(X)) & _
               " occurred in Cleanup.", vbExclamation
        SocketsCleanup = False
    Else
        SocketsCleanup = True
    End If
    
End Function

Public Function SocketsInitialize() As Boolean

    'Dim WSAD As WSAData
    Dim WSAData As T_WSA
    Dim X As Integer
    Dim szLoByte As String, szHiByte As String, szBuf As String
    
    X = WSAStartup(WS_VERSION_REQD, WSAData)
    
    If X <> 0 Then
        MsgBox "Erreur initialisation de socket !"
        SocketsInitialize = False
        Exit Function
    End If
    
    If LoByte(WSAData.wVersion) < WS_VERSION_MAJOR Or _
       (LoByte(WSAData.wVersion) = WS_VERSION_MAJOR And _
        HiByte(WSAData.wVersion) < WS_VERSION_MINOR) Then
        
        szHiByte = Trim$(Str$(HiByte(WSAData.wVersion)))
        szLoByte = Trim$(Str$(LoByte(WSAData.wVersion)))
        szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
        szBuf = szBuf & " non supporte par Windows "
        MsgBox szBuf, vbExclamation
        SocketsInitialize = False
        Exit Function
        
    End If
    
    If WSAData.iMaxSockets < MIN_SOCKETS_REQD Then
        szBuf = "Cette application necessite un minimum de " & _
                 Trim$(Str$(MIN_SOCKETS_REQD)) & " sockets."
        MsgBox szBuf, vbExclamation
        SocketsInitialize = False
        Exit Function
    End If
    
    SocketsInitialize = True
        
End Function

Public Function GetStatusCode(status As Long) As String

   Dim msg As String

   Select Case status
      Case IP_SUCCESS:                  msg = "ip trouvee... ping OK"
      Case IP_BUF_TOO_SMALL:            msg = "ip buffer trop petit"
      Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net non trouve"
      Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host non trouve"
      Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot non trouve"
      Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port non trouve"
      Case IP_NO_RESOURCES:             msg = "ip pas de ressource"
      Case IP_BAD_OPTION:               msg = "ip option invalide"
      Case IP_HW_ERROR:                 msg = "ip hw_error"
      Case IP_PACKET_TOO_BIG:           msg = "ip packet trop gros"
      Case IP_REQ_TIMED_OUT:            msg = "ip req timed out"
      Case IP_BAD_REQ:                  msg = "ip mauvaise req"
      Case IP_BAD_ROUTE:                msg = "ip mauvaise route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expire transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expire reassem"
      Case IP_PARAM_PROBLEM:            msg = "ip param_problem"
      Case IP_SOURCE_QUENCH:            msg = "ip source quench"
      Case IP_OPTION_TOO_BIG:           msg = "ip option too_big"
      Case IP_BAD_DESTINATION:          msg = "ip bad destination"
      Case IP_ADDR_DELETED:             msg = "ip addr deleted"
      Case IP_SPEC_MTU_CHANGE:          msg = "ip spec mtu change"
      Case IP_MTU_CHANGE:               msg = "ip mtu_change"
      Case IP_UNLOAD:                   msg = "ip unload"
      Case IP_ADDR_ADDED:               msg = "ip addr added"
      Case IP_GENERAL_FAILURE:          msg = "ip general failure"
      Case IP_PENDING:                  msg = "ip pending"
      Case PING_TIMEOUT:                msg = "ping timeout"
      Case Else:                        msg = "unknown msg returned"
   End Select
   
   GetStatusCode = CStr(status) & "   [ " & msg & " ]"
   
End Function

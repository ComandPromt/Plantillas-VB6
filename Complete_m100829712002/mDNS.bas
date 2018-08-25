Attribute VB_Name = "mDNS"
'this code is the magic of the mx control

Option Explicit

Private Const NS_ALL = 0
Private Const AF_INET = 2
Private Const IPPROTO_TCP = 6
Private Const IPPROTO_UDP = 17
Private Const LUP_RETURN_ALL = &HFF0
Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYS_STATUS_LEN = 128
Private Const SOCK_STREAM = 1                   ' stream socket
Private Const SOCK_DGRAM = 2                    ' datagram socket
Private Const SOCK_RAW = 3                      ' raw-protocol interface
Private Const SOCK_RDM = 4                      ' reliably-delivered message
Private Const SOCK_SEQPACKET = 5                ' sequenced packet stream

Private Type GUID     '  size is 16
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type afProtocols
    iAddressFamily As Long
    iProtocol As Long
End Type

Private Type sockaddr2
    sa_family As Integer
    sa_data(13) As Byte
End Type

Private Type SOCKET_ADDRESS
    lpSockaddr As Long
    iSockaddrLength  As Long
End Type

Private Type CSADDR_INFO
    LocalAddr As SOCKET_ADDRESS
    RemoteAddr As SOCKET_ADDRESS
    iSocketType As Long
    iProtocol As Long
End Type

Private Type WSAQuerySetW
    dwSize As Long
    lpszServiceInstanceName As Long
    lpServiceClassId As Long
    lpVersion As Long
    lpszComment As Long
    dwNameSpace As Long
    lpNSProviderId As Long
    lpszContext As Long
    dwNumberOfProtocols As Long
    lpafpProtocols As Long
    lpszQueryString As Long
    dwNumberOfCsAddrs As Long
    lpcsaBuffer As Long
    dwOutputFlags As Long
    lpBlob As Long
End Type

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(WSADESCRIPTION_LEN) As Byte
    szSystemStatus(WSASYS_STATUS_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Declare Function WSALookupServiceBegin Lib "ws2_32.dll" Alias "WSALookupServiceBeginA" (ByVal lpqsRestrictions As Long, ByVal dwControlFlags As Long, lphLookup As Long) As Long
Private Declare Function WSALookupServiceNext Lib "ws2_32.dll" Alias "WSALookupServiceNextA" (ByVal lphLookup As Long, ByVal dwControlFlags As Long, lpdwBufferLength As Long, lpqsResults As Byte) As Long
Private Declare Function WSALookupServiceEnd Lib "ws2_32.dll" (ByVal lphLookup As Long) As Long
'Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, lpWSAData As WSADATA) As Long
'Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Private Declare Function WSAAddressToString Lib "ws2_32.dll" Alias "WSAAddressToStringA" (lpsaAddress As sockaddr, ByVal dwAddressLength As Long, ByVal lpProtocolInfo As Long, ByVal lpszAddressString As String, lpdwAddressStringLength As Long) As Long
'Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long

 
 
Public Const DNS_RECURSION As Byte = 1

Public Type DNS_HEADER
    qryID As Integer
    options As Byte
    response As Byte
    qdcount As Integer
    ancount As Integer
    nscount As Integer
    arcount As Integer
End Type

' Registry data types
Const REG_NONE = 0&
Const REG_SZ = 1&
Const REG_EXPAND_SZ = 2&
Const REG_BINARY = 3&
Const REG_DWORD = 4&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

' Registry access types
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ

' Registry keys
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006

' The only registry error that I care about =)
Const ERROR_SUCCESS = 0&

' Registry access functions
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

' Variant (string array) that holds all the DNS servers found in the registry
Global sDNS As Variant
Global sMX As Variant
Global sPref As Variant ' holds the preferences
Global sBestMX As String ' Holds the "best" MX record (the one with the lowest preference)
Public ms_Domain As String
Public mi_DNSCount As Integer
Public mi_MXCount As Integer

Type IP_ADDRESS_STRING
    IpAddressString(4 * 4 - 1) As Byte
End Type
 
Type IP_MASK_STRING
    IpMaskString(4 * 4 - 1) As Byte
End Type
 
Type IP_ADDR_STRING
    Next      As Long
    IpAddress As IP_ADDRESS_STRING
    IpMask    As IP_MASK_STRING
    Context   As Long
End Type
 
Public Const MAX_HOSTNAME_LEN = 128
Public Const MAX_DOMAIN_NAME_LEN = 128
Public Const MAX_SCOPE_ID_LEN = 256
 
Type FIXED_INFO
    HostName(MAX_HOSTNAME_LEN + 4 - 1) As Byte
    DomainName(MAX_DOMAIN_NAME_LEN + 4 - 1) As Byte
    CurrentDnsServer As Long
    DnsServerList    As IP_ADDR_STRING
    NodeType         As Long
    ScopeId(MAX_SCOPE_ID_LEN + 4 - 1) As Byte
    EnableRouting    As Long
    EnableProxy      As Long
    EnableDns        As Long
End Type
 
Declare Function GetNetworkParams Lib "iphlpapi.dll" _
   (pFixedInfo As Any, _
    pOutBufLen As Long) As Long
 
Public Const ERROR_NOT_SUPPORTED = 50
Public Const ERROR_BUFFER_OVERFLOW = 111
Public Const ERROR_INVALID_PARAMETER = 87
Public Const ERROR_NO_DATA = 232
 
Declare Sub MoveMemory Lib "kernel32.dll" _
    Alias "RtlMoveMemory" _
   (Destination As Any, _
    Source As Any, _
    ByVal Length As Long)
    
' Remove the NULL character from the end of a string
Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Sub GetDNSInfo()
    Dim hKey As Long
    Dim hError As Long
    Dim sdhcpBuffer As String
    Dim sBuffer As String
    Dim sFinalBuff As String
    Dim lngFixedInfoNeeded      As Long
    Dim bytFixedInfoBuffer()    As Byte
    Dim udtFixedInfo            As FIXED_INFO
    Dim lngIpAddrStringPtr      As Long
    Dim udtIpAddrString         As IP_ADDR_STRING
    Dim strDnsIpAddress         As String
    Dim lngWin32apiResultCode   As Long
    Dim guidServiceClass        As GUID
    Dim qs                      As WSAQuerySetW
    Dim csa()                   As CSADDR_INFO
    Dim dwFlags                 As Long
    Dim dwLen                   As Long
    Dim hLookup                 As Long
    Dim afProtocols(1)          As afProtocols
    Dim nRet                    As Long
    Dim WSVersion               As Integer
    Dim uData                   As WSADataType
    Dim bBuffer()               As Byte
    Dim lSize                   As Long
    Dim sBuffer2                As String
    Dim i                       As Integer
    Dim ptr                     As Long
    Dim remSockAddr             As sockaddr2
    Dim sText                   As String
    
    With guidServiceClass
        .Data1 = &H90035    ' last two digits are the port number(53) in hex
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    sdhcpBuffer = Space(1000)
    sBuffer = Space(1000)
    sDNS = vbNullString
    
    If (RegOpenKeyEx(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\VxD\MSTCP", 0, KEY_READ, hKey) = ERROR_SUCCESS) Then
        ' DNS servers configured through Network control panel applet (95/98)
        RegQueryValueEx hKey, "NameServer", 0, REG_SZ, sBuffer, 1000
        RegCloseKey hKey
        If Trim(StripTerminator(sBuffer)) <> "" Then
            sFinalBuff = Trim(StripTerminator(sBuffer)) & ","
        End If
    End If
 
    If (RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", 0, KEY_READ, hKey) = ERROR_SUCCESS) Then
        ' DNS servers configured through Network control panel applet (NT)
        RegQueryValueEx hKey, "NameServer", 0, REG_SZ, sBuffer, 1000
        RegCloseKey hKey
        If Trim(StripTerminator(sBuffer)) <> "" Then
            If InStr(1, sFinalBuff, Trim(StripTerminator(sBuffer))) = 0 Then
                If sFinalBuff <> "" Then
                    sFinalBuff = sFinalBuff & Trim(StripTerminator(sBuffer)) & ","
                Else
                    sFinalBuff = Trim(StripTerminator(sBuffer)) & ","
                End If
            End If
        End If
    End If
 
    If (RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", 0, KEY_READ, hKey) = ERROR_SUCCESS) Then
        ' DNS servers configured dhcp (NT)
        RegQueryValueEx hKey, "DhcpNameServer", 0, REG_SZ, sBuffer, 1000
        RegCloseKey hKey
        If Trim(StripTerminator(sBuffer)) <> "" Then
            If InStr(1, sFinalBuff, Trim(StripTerminator(sBuffer))) = 0 Then
                If sFinalBuff <> "" Then
                    sFinalBuff = sFinalBuff & Trim(StripTerminator(sBuffer)) & ","
                Else
                    sFinalBuff = Trim(StripTerminator(sBuffer)) & ","
                End If
            End If
        End If
    End If
 
    If Is98 Or Is98se Or IsME Or Is2000 Or IsNT4 Or Is95 Or Is95B Then
        ' get dns servers with the new GetNetworkParams call
        ' only works on 98/ME/2000
        ' use the WSALookupService calls for 2000/nt4
        lngWin32apiResultCode = _
            GetNetworkParams(ByVal vbNullString, _
                             lngFixedInfoNeeded)
        If lngWin32apiResultCode = _
           ERROR_BUFFER_OVERFLOW Then
            ReDim _
                bytFixedInfoBuffer _
                   (lngFixedInfoNeeded)
        Else
            GoTo TerminateGetNetworkParams
        End If
        lngWin32apiResultCode = _
            GetNetworkParams(bytFixedInfoBuffer(0), _
                             lngFixedInfoNeeded)
        MoveMemory _
            udtFixedInfo, _
            bytFixedInfoBuffer(0), _
            Len(udtFixedInfo)
        With udtFixedInfo
            lngIpAddrStringPtr = _
                VarPtr(.DnsServerList)
            Do While lngIpAddrStringPtr
                MoveMemory _
                    udtIpAddrString, _
                    ByVal lngIpAddrStringPtr, _
                    Len(udtIpAddrString)
                With udtIpAddrString
                    strDnsIpAddress = _
                        StrConv(.IpAddress _
                                    .IpAddressString, _
                                vbUnicode)
                    If sFinalBuff = vbNullString Then
                        sFinalBuff = Left(strDnsIpAddress, InStr(strDnsIpAddress, vbNullChar) - 1) & ","
                    Else
                        If InStr(1, sFinalBuff, Left(strDnsIpAddress, InStr(strDnsIpAddress, vbNullChar) - 1) & ",") = 0 Then
                            sFinalBuff = sFinalBuff & Left(strDnsIpAddress, InStr(strDnsIpAddress, vbNullChar) - 1) & ","
                        End If
                    End If
                    lngIpAddrStringPtr = .Next
                End With
            Loop
        End With
    
         ' WSALookupService calls
        qs.dwSize = Len(qs)
        qs.lpszServiceInstanceName = 0
        qs.lpServiceClassId = VarPtr(guidServiceClass.Data1)
        qs.dwNameSpace = NS_ALL
        qs.dwNumberOfProtocols = 2
        qs.lpafpProtocols = afProtocols(0).iAddressFamily
    
        afProtocols(0).iAddressFamily = AF_INET
        afProtocols(0).iProtocol = IPPROTO_TCP
        afProtocols(1).iAddressFamily = AF_INET
        afProtocols(1).iProtocol = IPPROTO_UDP
    
        dwFlags = LUP_RETURN_ALL
        WSVersion = &H202   ' just assume we can handle up to winsock version 2.2
    
        nRet = WSAStartup(WSVersion, uData)
        If nRet = 0 Then
            nRet = WSALookupServiceBegin(VarPtr(qs.dwSize), dwFlags, hLookup)
            If nRet = 0 Then
                lSize = 2048
                ReDim bBuffer(lSize - 1)
                
                While WSALookupServiceNext(hLookup, dwFlags, lSize, bBuffer(0)) = 0
                    Call CopyMemory(qs.dwSize, bBuffer(0), Len(qs))
                    ReDim csa(qs.dwNumberOfCsAddrs - 1)
                    For i = 0 To qs.dwNumberOfCsAddrs - 1
                        ptr = qs.lpcsaBuffer + (i * Len(csa(i)))
                        Call CopyMemory(csa(i).LocalAddr, ByVal ptr, Len(csa(i)))
                        Call CopyMemory(remSockAddr.sa_family, ByVal csa(i).RemoteAddr.lpSockaddr, Len(remSockAddr))
                        sText = remSockAddr.sa_data(2) & "." & remSockAddr.sa_data(3) & "." & remSockAddr.sa_data(4) & "." & remSockAddr.sa_data(5)
                        If sFinalBuff = vbNullString Then
                            sFinalBuff = sText & ","
                        Else
                            sFinalBuff = sFinalBuff & sText & ","
                        End If
                    Next
                    lSize = 2048
                    ReDim bBuffer(lSize - 1)
                Wend
                nRet = WSALookupServiceEnd(hLookup)
            Else
                nRet = WSAGetLastError
' this will error out every time on win98/98se.
' seems the call's parameters have changed. with no documentation of the change.
'                MsgBox "Socket Error : " & nRet
            End If

        End If
        nRet = WSACleanup
    
    End If
    
    If Is95 Or Is95B Then
        ' get dns servers the old way
        ' anyone wanna tell me how to do this?
        
    End If
 
    If Right(sFinalBuff, 1) = "," Then sFinalBuff = Left(sFinalBuff, Len(sFinalBuff) - 1)
     
    sDNS = Split(sFinalBuff, ",")
    
    mi_DNSCount = UBound(sDNS)
    
TerminateGetNetworkParams:
    
End Sub

' Parse the server name out of the MX record, returns it in variable sName, iNdx is also
' modified to point to the end of the parsed structure.
Private Sub ParseName(dnsReply() As Byte, iNdx As Integer, sName As String)
    Dim iCompress As Integer        ' Compression index (index to original buffer)
    Dim iChCount As Integer         ' Character count (number of chars to read from buffer)
        
    ' While we dont encounter a null char (end-of-string specifier)
    While (dnsReply(iNdx) <> 0)
        ' Read the next character in the stream (length specifier)
        iChCount = dnsReply(iNdx)
        ' If our length specifier is 192 (0xc0) we have a compressed string
        If (iChCount = 192) Then
            ' Read the location of the rest of the string (offset into buffer)
            iCompress = dnsReply(iNdx + 1)
            ' Call ourself again, this time with the offset of the compressed string
            ParseName dnsReply(), iCompress, sName
            ' Step over the compression indicator and compression index
            iNdx = iNdx + 2
            ' After a compressed string, we are done
            Exit Sub
        End If
        
        ' Move to next char
        iNdx = iNdx + 1
        ' While we should still be reading chars
        While (iChCount)
            ' add the char to our string
            sName = sName + Chr(dnsReply(iNdx))
            iChCount = iChCount - 1
            iNdx = iNdx + 1
        Wend
        ' If the next char isn't null then the string continues, so add the dot
        If (dnsReply(iNdx) <> 0) Then sName = sName + "."
    Wend
End Sub

' Parses the buffer returned by the DNS server, returns the best MX server (lowest preference
' number), iNdx is modified to point to the current buffer position (should be the end of the buffer
' by the end, unless a record other than MX is found)
Private Function GetMXName(dnsReply() As Byte, iNdx As Integer, iAnCount As Integer) As String
    Dim iChCount As Integer     ' Character counter
    Dim sTemp As String         ' Holds the original query string
    
    Dim iBestPref As Integer    ' Holds the "best" preference number (lowest)
    Dim iMXCount As Integer
    ReDim sMX(0) As Variant
    ReDim sPref(0) As Variant
    iMXCount = 0
    iBestPref = -1
    sBestMX = vbNullString
    
    ParseName dnsReply(), iNdx, sTemp
    ' Step over null
    iNdx = iNdx + 2
    
    ' Step over 6 bytes (not sure what the 6 bytes are, but all other
    '   documentation shows steping over these 6 bytes)
    iNdx = iNdx + 6
    
    While (iAnCount)
        ' Check to make sure we received an MX record
        If (dnsReply(iNdx) = 15) Then
            Dim sName As String
            Dim iPref As Integer
            
            sName = ""
            
            ' Step over the last half of the integer that specifies the record type (1 byte)
            ' Step over the RR Type, RR Class, TTL (3 integers - 6 bytes)
            iNdx = iNdx + 1 + 6
            
            ' Step over the MX data length specifier (1 integer - 2 bytes)
            iNdx = iNdx + 2
            
            MemCopy iPref, dnsReply(iNdx), 2
            iPref = ntohs(iPref)
            ' Step over the MX preference value (1 integer - 2 bytes)
            iNdx = iNdx + 2
            
            ' Have to step through the byte-stream, looking for 0xc0 or 192 (compression char)
            ParseName dnsReply(), iNdx, sName
            
            If Trim(sName) <> "" Then
                iMXCount = iMXCount + 1
                ReDim Preserve sMX(iMXCount - 1) As Variant
                ReDim Preserve sPref(iMXCount - 1) As Variant
                sMX(iMXCount - 1) = sName
                sPref(iMXCount - 1) = iPref
                mi_MXCount = iMXCount - 1
                If (iBestPref = -1 Or iPref < iBestPref) Then
                    iBestPref = iPref
                    sBestMX = sName
                End If
            End If
            ' Step over 3 useless bytes
            iNdx = iNdx + 3
        Else
            GetMXName = sBestMX
            Exit Function
        End If
        iAnCount = iAnCount - 1
    Wend
    
    GetMXName = sBestMX
End Function

' Takes sDomain and converts it to the QNAME-type string, returns that. QNAME is how a
' DNS server expects the string.
'
'    Ex...    Pass -        mail.com
'             Returns -     &H4mail&H3com
'                            ^      ^
'                            |______|____ These two are character counters, they count the
'                                         number of characters appearing after them
Private Function MakeQName(sDomain As String) As String
    Dim iQCount As Integer      ' Character count (between dots)
    Dim iNdx As Integer         ' Index into sDomain string
    Dim iCount As Integer       ' Total chars in sDomain string
    Dim sQName As String        ' QNAME string
    Dim sDotName As String      ' Temp string for chars between dots
    Dim sChar As String         ' Single char from sDomain string
    
    iNdx = 1
    iQCount = 0
    iCount = Len(sDomain)
    ' While we haven't hit end-of-string
    While (iNdx <= iCount)
        ' Read a single char from our domain
        sChar = Mid(sDomain, iNdx, 1)
        ' If the char is a dot, then put our character count and the part of the string
        If (sChar = ".") Then
            sQName = sQName & Chr(iQCount) & sDotName
            iQCount = 0
            sDotName = ""
        Else
            sDotName = sDotName + sChar
            iQCount = iQCount + 1
        End If
        iNdx = iNdx + 1
    Wend
    
    sQName = sQName & Chr(iQCount) & sDotName
    
    MakeQName = sQName
End Function

' Performs the actual IP work to contact the DNS server, calls the other functions to parse
' and return the best server to send email through
Public Function MX_Query() As String
    Dim StartupData As WSADataType
    Dim SocketBuffer As sockaddr
    Dim IpAddr As Long
    Dim iRC As Integer
    Dim dnsHead As DNS_HEADER
    Dim iSock As Integer
    Dim dnsQuery() As Byte
    Dim sQName As String
    Dim dnsQueryNdx As Integer
    Dim iTemp As Integer
    Dim iNdx As Integer
    Dim dnsReply(2048) As Byte
    Dim iAnCount As Integer

    ' check for properties being set
    ''
    If Len(ms_Domain) < 5 Then
        Err.Raise 0, "MXQuery", "No Valid Domain Specified"
        Exit Function
    End If
    ''
    
    ' Initialize the Winsocket
    iRC = WSAStartup(&H101, StartupData)
    iRC = WSAStartup(&H101, StartupData)
    If iRC = SOCKET_ERROR Then Exit Function
    
    
    ' Create a socket
    iSock = socket(AF_INET, SOCK_DGRAM, 0)
    If iSock = SOCKET_ERROR Then Exit Function
    
    GetDNSInfo
    
    ' check to see that we found a dns server
    If UBound(sDNS) <= 0 Then
        ' problem
        'Err.Raise , , "No DNS Entries found. Cannont complete MX Lookup."
        Exit Function
    End If
    
    IpAddr = GetHostByNameAlias(sDNS(0))
    If IpAddr = -1 Then Exit Function
    
    ' get dns info
    
    ' Setup the connnection parameters
    SocketBuffer.sin_family = AF_INET
    SocketBuffer.sin_port = htons(53)
    SocketBuffer.sin_addr = IpAddr
    SocketBuffer.sin_zero = String$(8, 0)
    
    ' Set the DNS parameters
    dnsHead.qryID = htons(&H11DF)
    dnsHead.options = DNS_RECURSION
    dnsHead.qdcount = htons(1)
    dnsHead.ancount = 0
    dnsHead.nscount = 0
    dnsHead.arcount = 0
    
    dnsQueryNdx = 0
    
    ReDim dnsQuery(4000)
    
    ' Setup the dns structure to send the query in
    ' First goes the DNS header information
    MemCopy dnsQuery(dnsQueryNdx), dnsHead, 12
    dnsQueryNdx = dnsQueryNdx + 12
    
    ' Then the domain name (as a QNAME)
    sQName = MakeQName(ms_Domain)
    iNdx = 0
    While (iNdx < Len(sQName))
        dnsQuery(dnsQueryNdx + iNdx) = Asc(Mid(sQName, iNdx + 1, 1))
        iNdx = iNdx + 1
    Wend

    dnsQueryNdx = dnsQueryNdx + Len(sQName)
    
    ' Null terminate the string
    dnsQuery(dnsQueryNdx) = &H0
    dnsQueryNdx = dnsQueryNdx + 1
    
    ' The type of query (15 means MX query)
    iTemp = htons(15)
    MemCopy dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
    dnsQueryNdx = dnsQueryNdx + Len(iTemp)
    
    ' The class of query (1 means INET)
    iTemp = htons(1)
    MemCopy dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
    dnsQueryNdx = dnsQueryNdx + Len(iTemp)
    
    ReDim Preserve dnsQuery(dnsQueryNdx - 1)
    ' Send the query to the DNS server
    iRC = sendto(iSock, dnsQuery(0), dnsQueryNdx + 1, 0, SocketBuffer, Len(SocketBuffer))
    If (iRC = SOCKET_ERROR) Then
        MsgBox "Problem sending"
        Exit Function
    End If
    
    ' Wait for answer from the DNS server
    iRC = recvfrom(iSock, dnsReply(0), 2048, 0, SocketBuffer, Len(SocketBuffer))
    If (iRC = SOCKET_ERROR) Then
        MsgBox "Problem receiving"
        Exit Function
    End If
    
    ' Get the number of answers
    MemCopy iAnCount, dnsReply(6), 2
    iAnCount = ntohs(iAnCount)
    ' Parse the answer buffer
    MX_Query = GetMXName(dnsReply(), 12, iAnCount)
    
End Function

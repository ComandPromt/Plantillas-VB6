Attribute VB_Name = "NetStat"
'-------------------------------------------------------------------------------
' Copyright © 2001 Gregory Kirk. All rights reserved.
'
' You have a royalty-free right to use, modify, reproduce and distribute the
' Application Files (and/or any modified version) in any way you find useful,
' provided that you agree that Gregory Kirk has no warranty, obligations or
' liability for any Application Files.
'
' This module was at first based upon the original netstat file (netstat.c) by
' Mark Russinovich. However, I have not included UDP or SNMP and have made
' significant alterations.
' I am still experimenting with SMTP/POP3, UDP, and IPX/SPX.
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' This is basically a combined module of TCP/ICMP and required winsock functions.
'-------------------------------------------------------------------------------

Option Explicit

'-------------------------------------------------------------------------------
'Types and function for the ICMP table:

Public MIBICMPSTATS As MIBICMPSTATS
Public Type MIBICMPSTATS
    dwEchos As Long
    dwEchoReps As Long
End Type

Public MIBICMPINFO As MIBICMPINFO
Public Type MIBICMPINFO
    icmpOutStats As MIBICMPSTATS
End Type

Public MIB_ICMP As MIB_ICMP
Public Type MIB_ICMP
    stats As MIBICMPINFO
End Type

Public Declare Function GetIcmpStatistics Lib "iphlpapi.dll" (pStats As MIBICMPINFO) As Long
Public Last_ICMP_Cnt As Integer 'ICMP count

'-------------------------------------------------------------------------------
'Types and functions for the TCP table:

Type MIB_TCPROW
  dwState As Long
  dwLocalAddr As Long
  dwLocalPort As Long
  dwRemoteAddr As Long
  dwRemotePort As Long
End Type

Type MIB_TCPTABLE
  dwNumEntries As Long
  table(100) As MIB_TCPROW
End Type
Public MIB_TCPTABLE As MIB_TCPTABLE

Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As MIB_TCPTABLE, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Public Declare Function SetTcpEntry Lib "IPhlpAPI" (pTcpRow As MIB_TCPROW) As Long 'This is used to close an open port.
Public IP_States(13) As String
Private Last_Tcp_Cnt As Integer 'TCP connection count

'-------------------------------------------------------------------------------
'Types and functions for winsock:

Private Const AF_INET = 2
Private Const IP_SUCCESS As Long = 0
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const SOCKET_ERROR As Long = -1
Private Const WS_VERSION_REQD As Long = &H101

Type HOSTENT
    h_name As Long        ' official name of host
    h_aliases As Long     ' alias list
    h_addrtype As Integer ' host address type
    h_length As Integer   ' length of address
    h_addr_list As Long   ' list of addresses
End Type

Type servent
  s_name As Long            ' (pointer to string) official service name
  s_aliases As Long         ' (pointer to string) alias list (might be null-seperated with 2null terminated)
  s_port As Long            ' port #
  s_proto As Long           ' (pointer to) protocol to use
End Type

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Public Declare Function ntohs Lib "WSOCK32.DLL" (ByVal netshort As Long) As Long
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal CP As String) As Long
Private Declare Function inet_ntoa Lib "WSOCK32.DLL" (ByVal inn As Long) As Long
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" (Addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal host_name As String) As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Declare Function lstrlen Lib "kernel32" (ByVal lpString As Any) As Integer
Private Blocked As Boolean

'-------------------------------------------------------------------------------
'Function for checking for new connections and blocking them if specified:
Sub CheckTcp()
Dim Item As ListItem, LTmp As Long
Dim x As Integer, i As Integer, n As Integer
Dim RemA As String, LocP As String, RemP As String
Dim tcpt As MIB_TCPTABLE

Blocked = False
LTmp = Len(MIB_TCPTABLE)  'Size of the TCP table
GetTcpTable tcpt, LTmp, 0 'Load the TCP table data.
x = tcpt.dwNumEntries     'Number of TCP entries.

If x > Last_Tcp_Cnt Or x < Last_Tcp_Cnt Then '+ or - an entry detected.
frmMain.RefreshNS

For i = 0 To tcpt.dwNumEntries - 1

RemA = GetAscIP(tcpt.table(i).dwRemoteAddr) 'Retrieve the IP address
RemP = ntohs(tcpt.table(i).dwRemotePort) 'Retrieve the remote port
LocP = ntohs(tcpt.table(i).dwLocalPort) 'Retrieve the local port

    If frmMain.Filtering = False Then Exit For 'Exit the loop if filtering is off.
        
        '//Address blocking
        If frmMain.chkAct(0).Value = 1 Then
            For n = 1 To frmMain.lvwFilter(0).ListItems.Count
                If frmMain.lvwFilter(0).ListItems.Item(n).Checked = False Then GoTo NextLoop
                If RemA = frmMain.lvwFilter(0).ListItems.Item(n).Key And tcpt.table(i).dwState <> 2 Then
                    If frmMain.Logging = True And frmMain.chkLog(2).Value = 1 Then rLog RemA, LocP, RemP, "Blocked Address", Time, True
                    Blocked = True
                    tcpt.table(i).dwState = 12
                    SetTcpEntry tcpt.table(i)
                    DoEvents
                    GoTo EndLp
                End If
NextLoop:
            Next n
        End If
        
        '//Remote port blocking
        If frmMain.chkAct(1).Value = 1 Then
            For n = 1 To frmMain.lvwFilter(1).ListItems.Count
                If frmMain.lvwFilter(1).ListItems.Item(n).Checked = False Then GoTo NextLoop2
                If RemP = frmMain.lvwFilter(1).ListItems.Item(n).Text And tcpt.table(i).dwState <> 2 Then
                    If frmMain.Logging = True And frmMain.chkLog(2).Value = 1 Then rLog RemA, LocP, RemP, "Blocked Remote Port", Time, True
                    Blocked = True
                    tcpt.table(i).dwState = 12
                    SetTcpEntry tcpt.table(i)
                    DoEvents
                    GoTo EndLp
                End If
NextLoop2:
            Next n
        End If
        
        '//Local port blocking
        If frmMain.chkAct(2).Value = 1 Then
            For n = 1 To frmMain.lvwFilter(2).ListItems.Count
                If frmMain.lvwFilter(2).ListItems.Item(n).Checked = False Then GoTo NextLoop3
                If LocP = frmMain.lvwFilter(2).ListItems.Item(n).Text And tcpt.table(i).dwState <> 2 Then
                    If frmMain.Logging = True And frmMain.chkLog(2).Value = 1 Then rLog RemA, LocP, RemP, "Blocked Local Port", Time, True
                    Blocked = True
                    tcpt.table(i).dwState = 12
                    SetTcpEntry tcpt.table(i)
                    DoEvents
                    GoTo EndLp
                End If
NextLoop3:
            Next n
        End If
EndLp:
    Next i
End If
Last_Tcp_Cnt = tcpt.dwNumEntries 'Update the TCP count

'//ICMP Statistics
If GetIcmpStatistics(MIBICMPINFO) <> 0 Then
    frmMain.SBar.Panels(3).Text = "ICMP failure"
    rLog "ICMP", "ICMP failure", "", "", Time
Else
    With MIBICMPINFO.icmpOutStats
        If Last_ICMP_Cnt <> .dwEchoReps + .dwEchos Then
            frmMain.SBar.Panels(3).Text = "ICMP Echo Requests: " & .dwEchoReps & ", Echo Replies: " & .dwEchos
            rLog "ICMP", "Echo Requests: " & .dwEchoReps, "Echo Replies: " & .dwEchos, "", Time
            Last_ICMP_Cnt = .dwEchoReps + .dwEchos 'Update the ICMP count
        End If
    End With
End If

If Blocked = True Then frmMain.RefreshNS
End Sub

'-------------------------------------------------------------------------------
'Sub for defining IP state constants:
Sub InitStates()
  IP_States(0) = "UNKNOWN"
  IP_States(1) = "CLOSED"
  IP_States(2) = "LISTENING"
  IP_States(3) = "SYN_SENT"
  IP_States(4) = "SYN_RCVD"
  IP_States(5) = "ESTABLISHED"
  IP_States(6) = "FIN_WAIT1"
  IP_States(7) = "FIN_WAIT2"
  IP_States(8) = "CLOSE_WAIT"
  IP_States(9) = "CLOSING"
  IP_States(10) = "LAST_ACK"
  IP_States(11) = "TIME_WAIT"
  IP_States(12) = "DELETE_TCB"
End Sub

'-------------------------------------------------------------------------------
'Function for obtaining the IP number of a hostname:
Public Function GetIPFromHostName(HostName$) As Long
Dim phe&, heDestHost As HOSTENT
Dim addrList&, retIP&
    retIP = inet_addr(HostName$)
    If retIP = &HFFFF Then
        phe = gethostbyname(HostName$)
        If phe <> 0 Then
            CopyMemory heDestHost, ByVal phe, Len(heDestHost)
            CopyMemory addrList, ByVal heDestHost.h_addr_list, 4
            CopyMemory retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = &HFFFF
        End If
    End If
    GetIPFromHostName = retIP
End Function

'-------------------------------------------------------------------------------
'Function for obtaining the hostname of an IP number:
Public Function GetHostNameFromIP(ByVal sAddress As String) As String
   Dim ptrHosent As Long
   Dim hAddress As Long
   Dim nbytes As Long
   
If SocketsInitialize1() Then
    hAddress = inet_addr(sAddress) 'Convert string address to long, this was the cause of meny errors, so do not mess with this.
    If hAddress <> SOCKET_ERROR Then
        DoEvents
        ptrHosent = gethostbyaddr(hAddress, 4, AF_INET) 'Obtain a pointer to the HOSTENT structure.
        DoEvents
        If ptrHosent <> 0 Then
            CopyMemory ptrHosent, ByVal ptrHosent, 4 'Convert address and get resolved hostname.
            nbytes = lstrlen(ByVal ptrHosent)
            If nbytes > 0 Then
                sAddress = Space$(nbytes)
                CopyMemory ByVal sAddress, ByVal ptrHosent, nbytes
                GetHostNameFromIP = sAddress
            End If
        Else
            GetHostNameFromIP = sAddress 'No DNS entry, so set it back to the IP.
        End If
        SocketsCleanup
    Else 'SOCKET_ERROR
        GetHostNameFromIP = "Invalid IP."
    End If
Else 'Sockets failed to initialize.
    Exit Function
End If
End Function

'-------------------------------------------------------------------------------
'Function for obtaining the IP number:
Public Function GetAscIP(ByVal inn As Long) As String
  Dim nStr&
    Dim lpStr As Long
    Dim retString As String
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr Then
        nStr = lstrlen(lpStr)
        If nStr > 32 Then nStr = 32
        CopyMemory ByVal retString, ByVal lpStr, nStr
        retString = Left(retString, nStr)
        GetAscIP = retString
    Else
        GetAscIP = "Unable to get IP"
    End If
End Function

'-------------------------------------------------------------------------------
'Function for Initializing a socket:
Private Function SocketsInitialize1() As Boolean
   Dim WSAD As WSADATA
   Dim success As Long
   SocketsInitialize1 = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
End Function

'-------------------------------------------------------------------------------
'Sub for socket clean up:
Private Sub SocketsCleanup()
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
End Sub


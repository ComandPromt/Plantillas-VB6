VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ashleys Mailserver"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File2 
      Height          =   480
      Left            =   2970
      Pattern         =   "*.txt"
      TabIndex        =   1
      Top             =   1590
      Visible         =   0   'False
      Width           =   1350
   End
   Begin MSWinsockLib.Winsock http 
      Index           =   0
      Left            =   2280
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin MSWinsockLib.Winsock pop3 
      Left            =   3390
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   110
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   2445
      Top             =   1920
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   0
      Pattern         =   "*.txt"
      TabIndex        =   0
      Top             =   330
      Width           =   1395
   End
   Begin Ashleys_Mailserver.MX mx 
      Left            =   3885
      Top             =   2655
      _ExtentX        =   714
      _ExtentY        =   450
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   4185
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   25
   End
   Begin VB.Label Label3 
      Caption         =   "Alternativly, you can send a webbrowser to you ip address, or connect via telnet to ports 110 or 25."
      Height          =   630
      Left            =   1455
      TabIndex        =   4
      Top             =   675
      Width           =   3150
   End
   Begin VB.Label Label2 
      Caption         =   "Check out the file 'readme.txt'. It explains a quickstart quide that will educate you about all the features of this mailserver."
      Height          =   630
      Left            =   1500
      TabIndex        =   3
      Top             =   15
      Width           =   3150
   End
   Begin VB.Label Label1 
      Caption         =   "Outgoing Email Que"
      Height          =   225
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public obj As New Dictionary

Private sentemails As New Collection
Private pop3connections As New Collection

Private Sub Form_Load()
    App.Title = "Ashley's MailServer"
    On Error GoTo portopenfailure
    
    ws(0).listen
    pop3.listen
    http(0).listen
    
    On Error GoTo 0
    File1.Path = subfolder("out")
    Exit Sub
portopenfailure:
    MsgBox "An error has occured while opening the required ports, Ashley's mailserver requires the following ports:" & vbCrLf & "Http: (port 80) for webmail" & vbCrLf & "Smtp: (port 25) for mail exchange between other servers and clients." & vbCrLf & "Pop3: (port 110) for mail download in clients" & vbCrLf & vbCrLf & "Close any webservers, mail servers, proxy servers, or any other apps that may be using these ports. if in doubt, try a restart", vbCritical
    End
End Sub

Private Sub http_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'whenever someone requests a webpage, move their request to the side, so
    'we can accept more. (Just I've notice IE will request about 10 files at once)
    'The number of silmultanious connections is dependant only on memory for buffer space.
    
    'Is there are free one perchance?
tryagain:
    For a = 1 To http.UBound
        If http(a).State = 0 Or http(a).State = 8 Then
            http(a).Close
            http(a).accept requestID
            Exit Sub
        End If
        DoEvents
    Next a
    'This puts a limit on the number of silmultanius connections
    'to remove, comment out the next line
    'GoTo tryagain
    DoEvents
    'guess not, well, make one!
    Num = http.UBound + 1
    Load http(Num)
    http(Num).accept requestID
End Sub


Private Sub http_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'All my fans will probibly notice that this entire sub has been copy/pasted from
    'my webserver that ALLMOST won code of the month on PSC (FULL perl support, acceptable
    'ASP support, some PHP support, resume downloads, etc.) Damn hacker taking out
    'the voting charts. This mail server is my second attempt, please vote for it.
    
    If docroot = "\" Then docroot = App.Path
    
    Dim a As String, filedata As String, headers As Dictionary
    If http(Index).State <> 7 Then http(Index).Close: Exit Sub
    http(Index).getdata a
    'Debug.Print a
    
    'ok, the way I've done this is wrong, because, say someone uploads a file via a cgi that's, say 500kb.
    'the browser will split it into packets of (say) 8kb each, and will send them here.
    'which means this function will be called first with the request and the first 7.5kb of the file.
    'then again, with the second chunk of the file, but no headers, which will crash this because theres
    'no headers to parse. I was stuck, and, cause it was 2AM in the morning, I thought of this stupid
    'and unreliable solution. This is a FIXME! (actually, it works rather good, and prob follows the protocol)
    If bytesTotal = 8192 Or InStr(1, a, "multipart/form-data", TextCompare) Then
        http(Index).Tag = http(Index).Tag & a
        stats.List(Index) = "Incomming File..."
        Exit Sub
    Else
        a = http(Index).Tag & a
        http(Index).Tag = ""
    End If
    
    
    If a = "" Then
        http(Index).Close
        Exit Sub
    End If
    
    otherheaders = Mid(a & vbNewLine, InStr(1, a, vbNewLine) + 2)
    otherheaders = Mid(otherheaders, 1, InStr(1, otherheaders, vbNewLine & vbNewLine) - 1)
    
    Set headers = parseheaders(CStr(otherheaders))
    
    If InStr(1, a, vbNewLine & vbNewLine) > 0 Then postdata = Mid(a, InStr(1, a, vbNewLine & vbNewLine) + 4) Else postdata = ""
    
    If CLng(headers("Content-length")) > Len(postdata) Then
        'ok, there are more packets comming, were executing too early
        'ie5 for mac is the cause of this code
        stats.List(Index) = "Awaiting POST data"
        http(Index).Tag = http(Index).Tag & a
        Exit Sub
    End If
    
    If headers("Content-type") = "application/x-www-form-urlencoded" And IsEmpty(headers("Content-length")) Then
        'my mac did this while posting the feedback form. makes no sense, but, it works.
        ' (it splits the request in 2)
        'stats.List(Index) = "Awaiting POST data"
        http(Index).Tag = http(Index).Tag & a
        Exit Sub
    End If
    
    'get the request, and then take the first line of it, then take just the request page
    a = Left(a, InStr(1, a, vbNewLine) - 1)
    a = Mid(a, InStr(1, a, " ") + 1)
    a = Left(a, InStr(1, a, " ") - 1)

    
    While Mid(a, 1, 3) = "/.."
        a = Mid(a, 4)
    Wend
    
    If Right(a, 1) = "/" Then a = a & Default
    
    'seperated the request string into filename and GET data
    If Not CBool(InStr(1, a, "?")) Then
        a = a & "?"
    End If
    cmd = Left(a, InStr(1, a, "?") - 1)
    data = Mid(a, InStr(1, a, "?") + 1)
    cmd = Replace(cmd, "/", "\")
    cmd = Replace(cmd, "%20", " ")
    
    header = "HTTP/1.0 200 OK" & vbNewLine & "Server: Ashleys opensource mailserver" & vbNewLine & "Host: " & _
    http(Index).LocalIP & vbNewLine & "Connection: close" & vbNewLine
    
    Select Case LCase(fso.GetParentFolderName(cmd))
    Case "\img"
        Open fso.BuildPath(fso.BuildPath(App.Path, "graphics"), fso.GetFileName(cmd)) For Binary As #1
        Dim I As String
        I = Space(LOF(1))
        Get #1, , I
        Close #1
        back = "Content-type: image/png" & vbCrLf & vbCrLf & I
    Case "\mail"
        back = dowebsite(fso.GetFileName(cmd), tophpvariables(CStr(data), CStr(postdata), headers("cookie")), headers, header)
    Case Else
        header = ""
        back = "HTTP/1.0 302 FOUND" & vbNewLine & "Server: Ashleys opensource mailserver" & vbNewLine & "Host: " & _
    http(Index).LocalIP & vbNewLine & "Url: /mail/inbox.webmail" & vbNewLine & "Location: /mail/inbox.webmail" & vbNewLine & "Connection: close" & vbNewLine & _
    vbNewLine
    End Select
    
    back = header & back
    On Error GoTo outofhere
    While Len(back) > 0
        http(Index).SendData Mid(back, 1, 10000)
        back = Mid(back, 10001)
        t = Timer + 3
        While t > Timer
            DoEvents
        Wend
    Wend
outofhere:
    t = Timer + 1
    While t > Timer
        DoEvents
    Wend
    
    On Error Resume Next
    http(Index).Close
    
End Sub

Private Sub pop3_ConnectionRequest(ByVal requestID As Long)
    Dim a As pop3session
    
    Set a = New pop3session
    
    a.ws.accept requestID
    Randomize
    
    ts = "<" & Int(Rnd() * 100000000000#) & Int(Rnd() * 100000000000#) & Int(Rnd() * 100000000000#) & Int(Rnd() * 100000000000#) & ">"
    a.timestamp = ts
    a.ws.SendData "+OK Hello and welcome to Ashley's POP3 server " & ts & vbCrLf
    
    a.Show
    
    
End Sub

Private Sub pop3_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    pop3.Close
    pop3.listen
End Sub

Private Sub Timer1_Timer()
    If ws(ws.UBound).State = 8 And ws.UBound > 0 Then Unload ws(ws.UBound)
    If http(http.UBound).State = 8 And http.UBound > 0 Then Unload http(http.UBound)
    
    ws(0).Close
    ws(0).listen
        
    On Error Resume Next
    http(0).Close
    http(0).listen
        
    On Error GoTo 0
    pop3.Close
    pop3.listen
    Dim s As sendemail
    
    For a = 0 To File1.ListCount - 1
        If Mid(File1.List(a), 1, 1) <> "!" Then
            Set s = New sendemail
            s.setup fso.BuildPath(File1.Path, File1.List(a))
            sentemails.Add s
        End If
    Next a
    
    If File1.ListCount = 0 And sentemails.Count > 0 Then
        Set sentemails = New Collection
    End If
    
    File1.Refresh
    
End Sub

Private Sub ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    For a = 1 To ws.UBound
        If ws(a).State = 0 Or ws(a).State = 8 Or ws(a).State = 9 Then Exit For
    Next
    On Error Resume Next
    Load ws(a)
    On Error GoTo 0
    ws(a).Close
    ws(a).accept requestID
    ws(a).SendData "220 Ashleys VB6 mailserver ready for mail" & vbCrLf
    Set c = New inmail
    Set obj(a) = c
End Sub

Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim a As String
    ws(Index).getdata a
    obj(Index).moreincomming a
    obj(Index).parsebuffer
    c = obj(Index).outbuffer
    If Len(c) > 0 Then ws(Index).SendData c: obj(Index).outbuffer = ""
    If Mid(c, 1, 3) = "221" Then ws(Index).Close
End Sub

Private Sub ws_Error(Index As Integer, ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ws(Index).Close
    If Index = 0 Then ws(0).listen
End Sub

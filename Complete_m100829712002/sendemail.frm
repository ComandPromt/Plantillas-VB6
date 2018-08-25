VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form sendemail 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sending Email"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1
   ScaleMode       =   0  'User
   ScaleWidth      =   1
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock ws 
      Left            =   3975
      Top             =   510
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape pb 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   320
      Left            =   0
      Top             =   1278
      Width           =   4680
   End
   Begin VB.Label txt 
      Caption         =   "Label1"
      Height          =   1170
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   4575
   End
End
Attribute VB_Name = "sendemail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private afrom As String
Private ato As String
Private data As String
Private attemptcount As Long
Private datafilename As String

Private stage As Long

Public Sub setup(FileName As String)
    If Not fso.FileExists(FileName) Then
        Unload Me
        Exit Sub
    End If
    datafilename = FileName
    Dim ts As TextStream
    
    Set ts = fso.OpenTextFile(FileName)
    afrom = ts.ReadLine
    ato = ts.ReadLine
    attemptcount = Val(ts.ReadLine) + 1
    donttrytill = ts.ReadLine
    data = ts.ReadAll
    ts.Close
    
    If CDate(donttrytill) > Now Then
        Me.Hide
        Unload Me
        Exit Sub
    End If
    
    If attemptcount = 8 Then
        Set ts = fso.OpenTextFile(fso.BuildPath(subfolder("out"), Int(Rnd * 100000000) & ".txt"), ForWriting, True)
        ts.WriteLine "mailsubsystem@" & HostName
        ts.WriteLine afrom
        ts.WriteLine "0"
        ts.WriteLine Now
        ts.WriteLine "from: mailsubsystem@" & HostName
        ts.WriteLine "to: " & afrom
        ts.WriteLine "subject: unable to deliver your message for the last 4 hours"
        ts.WriteBlankLines 2
        ts.WriteLine "*****THIS IS A WARNING MESSAGE ONLY*****"
        ts.WriteLine "***** DO NOT RESEND YOUR MESSAGE *******"
        ts.WriteBlankLines 2
        ts.WriteLine "I was unable to deliver your message."
        ts.WriteLine ""
        ts.WriteLine "I will continue trying for a total of 4 days"
        ts.WriteLine
        ts.WriteLine "Message content follows:"
        ts.WriteLine
        ts.Write data
        ts.Close
    End If
    
    If attemptcount = 12 Then
        Set ts = fso.OpenTextFile(fso.BuildPath(subfolder("out"), Int(Rnd * 100000000) & ".txt"), ForWriting, True)
        ts.WriteLine "mailsubsystem@" & HostName
        ts.WriteLine afrom
        ts.WriteLine "0"
        ts.WriteLine Now
        ts.WriteLine "from: mailsubsystem@" & HostName
        ts.WriteLine "to: " & afrom
        ts.WriteLine "subject: FAILURE in sending your message"
        ts.WriteBlankLines 2
        ts.WriteLine "******YOUR MESSAGE WASN'T DELIVERED*****"
        ts.WriteLine "*****THE SERVER REFUSES TO TRY AGAIN****"
        ts.WriteBlankLines 2
        ts.WriteLine "I was unable to deliver your message."
        ts.WriteLine ""
        ts.WriteLine "After 4 days, I have been unsucessful in delivering your message"
        ts.WriteLine
        ts.WriteLine "Message content follows:"
        ts.WriteLine
        ts.Write data
        ts.Close
        fso.DeleteFile FileName
        Exit Sub
    End If
    
    Set ts = fso.OpenTextFile(FileName, ForWriting, False)
    ts.WriteLine afrom
    ts.WriteLine ato
    ts.WriteLine attemptcount
    ts.WriteLine Now + (2 ^ (attemptcount - 1)) / 1440
    ts.Write data
    ts.Close
    
    If mxlookup(Mid(ato, InStr(1, ato, "@") + 1)) = "" Then
        If attemptcount = 1 Then
            Set ts = fso.OpenTextFile(fso.BuildPath(subfolder("out"), Int(Rnd * 100000000) & ".txt"), ForWriting, True)
            ts.WriteLine "mailsubsystem@" & HostName
            ts.WriteLine afrom
            ts.WriteLine "0"
            ts.WriteLine Now
            ts.WriteLine "from: mailsubsystem@" & HostName
            ts.WriteLine "to: " & afrom
            ts.WriteLine "subject: Warning of possible future failure regarding your message"
            ts.WriteBlankLines 2
            ts.WriteLine "*****THIS IS A WARNING MESSAGE ONLY*****"
            ts.WriteLine "***** DO NOT RESEND YOUR MESSAGE *******"
            ts.WriteBlankLines 2
            ts.WriteLine "My first attempt to deliver your message failed. The server was unable to locate the host: " & Mid(ato, InStr(1, ato, "@") + 1)
            ts.WriteLine ""
            ts.WriteLine "I will try again 7 times over the next 4 hours, and notify you if your message still remains undelivered. If it is still undelivered at this date, I will then try another 4 times over the next 4 days. After 4 days, if your message still remains undelivered, I will notify you, and the email will be removed from my job sheet."
            ts.WriteLine ""
            ts.WriteLine "Message content follows:"
            ts.WriteLine
            ts.Write data
            ts.Close
        End If
        Me.Hide
        Unload Me
        Exit Sub
    End If
    txt = "From: " & afrom & vbNewLine & "To: " & ato & vbNewLine & "Size: " & Len(data)
    pb.Width = 0
    Me.Show
    ws.Close
    
    If InStr(1, ato, "@" & HostName) Then
        'we're trying to email outselves, usually a mailer deamon gone wrong
        'dont actually connect to ourselves, cause it sometimes doesn't work
        'and more and more mailer deamons stack up, just fake the connection
        'then get the hell out of here
        Dim sendit As New inmail
        sendit.moreincomming "HELO " & HostName & " webmail" & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming "MAIL FROM: " & afrom & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming "RCPT TO: " & ato & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming "DATA" & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming data & vbCrLf & "." & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming "QUIT" & vbCrLf
        sendit.parsebuffer
        fso.DeleteFile FileName
        Me.Hide
        Unload Me
        Exit Sub
    End If
    
    ws.connect mxlookup(Mid(ato, InStr(1, ato, "@") + 1)), 25
    t = Timer + 60
    While ws.State <> 7 And t < Timer
        DoEvents
    Wend
    If t > Timer Then
        If attemptcount = 1 Then
            Set ts = fso.OpenTextFile(fso.BuildPath(subfolder("out"), Int(Rnd * 100000000) & ".txt"), ForWriting, True)
            ts.WriteLine "mailsubsystem@" & HostName
            ts.WriteLine afrom
            ts.WriteLine "0"
            ts.WriteLine Now
            ts.WriteLine "from: mailsubsystem@" & HostName
            ts.WriteLine "to: " & afrom
            ts.WriteLine "subject: Warning of possible future failure regarding your message"
            ts.WriteBlankLines 2
            ts.WriteLine "*****THIS IS A WARNING MESSAGE ONLY*****"
            ts.WriteLine "***** DO NOT RESEND YOUR MESSAGE *******"
            ts.WriteBlankLines 2
            ts.WriteLine "My first attempt to deliver your message failed. The server was unable to contact the host: " & Mid(toaddr, InStr(1, toaddr, "@") + 1) & ". The host may be down, crowded, congested, or just plain gone walk abouts."
            ts.WriteLine ""
            ts.WriteLine "I will try again 7 times over the next 4 hours, and notify you if your message still remains undelivered. If it is still undelivered at this date, I will then try another 4 times over the next 4 days. After 4 days, if your message still remains undelivered, I will notify you, and the email will be removed from my job sheet."
            ts.WriteLine ""
            ts.WriteLine "Message content follows:"
            ts.WriteLine
            ts.Write data
            ts.Close
        End If
        Me.Hide
        Unload Me
        Exit Sub
    End If
    stage = 0
    data = data & vbCrLf & "." & vbCrLf
    'ws.SendData "HELO " & ws.LocalIP & vbCrLf
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim a As String
    On Error Resume Next
    ws.getdata a
    On Error GoTo 0
    If Err.Description <> "" Then a = "220 connect ok" & vbCrLf
    If Mid(a, 1, 3) = "250" Or Mid(a, 1, 3) = "220" Or Mid(a, 1, 3) = "351" Or Mid(a, 1, 3) = "354" Then
        stage = stage + 1
        pb.Width = stage * 0.1
    Else
        If attemptcount = 1 Then
            Dim ts As TextStream
            Set ts = fso.OpenTextFile(fso.BuildPath(subfolder("out"), Int(Rnd * 100000000) & ".txt"), ForWriting, True)
            ts.WriteLine "mailsubsystem@" & HostName
            ts.WriteLine afrom
            ts.WriteLine "0"
            ts.WriteLine Now
            ts.WriteLine "from: mailsubsystem@" & HostName
            ts.WriteLine "to: " & afrom
            ts.WriteLine "subject: Warning of possible future failure regarding your message"
            ts.WriteBlankLines 2
            ts.WriteLine "*****THIS IS A WARNING MESSAGE ONLY*****"
            ts.WriteLine "***** DO NOT RESEND YOUR MESSAGE *******"
            ts.WriteBlankLines 2
            ts.WriteLine "My first attempt to deliver your message failed. The server " & ws.RemoteHost & " returned the following line:"
            ts.WriteLine ""
            ts.WriteLine a
            ts.WriteLine ""
            ts.WriteLine "I will try again 7 times over the next 4 hours, and notify you if your message still remains undelivered. If it is still undelivered at this date, I will then try another 4 times over the next 4 days. After 4 days, if your message still remains undelivered, I will notify you, and the email will be removed from my job sheet."
            ts.WriteLine ""
            ts.WriteLine "Message content follows:"
            ts.WriteLine
            ts.Write data
            ts.Close
        End If
        Hide
        Unload Me
        Exit Sub
    End If
    
    Debug.Print stage & ":" & a
    Select Case stage
    Case 1
        ws.SendData "HELO " & HostName & vbCrLf
    Case 2
        ws.SendData "MAIL FROM: " & afrom & vbCrLf
    Case 3
        ws.SendData "RCPT TO: " & ato & vbCrLf
    Case 4
        ws.SendData "DATA" & vbCrLf
    Case 5
        While Len(data) > 1
            ws.SendData Mid(data, 1, InStr(1, data, vbCrLf) + 1)
            data = Mid(data, InStr(1, data, vbCrLf) + 2)
        Wend
    Case 6
        ws.SendData "QUIT" & vbCrLf
        fso.DeleteFile datafilename, True
        Me.Hide
        Unload Me
    End Select
End Sub

Private Sub ws_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ws.Close
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "inmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class handles 1 incomming smtp session per instance.
'either from another smtp server, a client ie outlook express, or a telnet session.

'it is also used internally (in some occasions), for instance, in the SMS notification
'section, this class will fake an smtp session with another instance of itself, just to simplify things.

Public inbuffer As String
Public outbuffer As String

Public State As smtpinstate

Public fromaddr As String
Public toaddr As New Collection
Public body As String
Public helo As String

Public Sub moreincomming(incoming As String)
    On Error GoTo 0
    inbuffer = inbuffer & incoming
    parsebuffer
    parsebuffer
End Sub

Public Sub parsebuffer()
    If InStr(1, inbuffer, vbCrLf) = 0 Then Exit Sub
    If previewtopline = "" Then pulltopline: parsebuffer: Exit Sub
    Select Case State
    Case smtpinstate.aWanthelo
        If LCase(Left(previewtopline, 4)) = "helo" Then
            State = bWantfrom
            outbuffer = outbuffer & "250 Please to meet you!" & vbCrLf
            helo = Mid(previewtopline, 6)
        ElseIf LCase(Left(previewtopline, 4)) = "ehlo" Then
            outbuffer = outbuffer & "502 No ESMTP support please" & vbCrLf
        ElseIf LCase(Left(previewtopline, 4)) = "help" Then
            outbuffer = outbuffer & "250 Step 1, state your ip (ie HELO 142.152.168.3)" & vbCrLf
        Else
            outbuffer = outbuffer & "503 It's polite to say HELO first." & vbCrLf
        End If
        pulltopline
    Case smtpinstate.bWantfrom
        If LCase(Left(previewtopline, 4)) = "mail" And extractemail(previewtopline) <> "" Then
            fromaddr = extractemail(previewtopline)
            outbuffer = outbuffer & "250 " & fromaddr & " sender ok." & vbCrLf
            State = cGettingrcpt
        ElseIf LCase(Left(previewtopline, 4)) = "help" Then
            outbuffer = outbuffer & "250 specify email sender, (ie MAIL FROM: me@myemail.com)" & vbCrLf
        Else
            outbuffer = outbuffer & "503 Expected 'MAIL', see help for example" & vbCrLf
        End If
        pulltopline
    Case smtpinstate.cGettingrcpt
        If LCase(Left(previewtopline, 4)) = "rcpt" And extractemail(previewtopline) <> "" Then
            If toaddr.Count > 60 Then
                outbuffer = outbuffer & "452 Too many recipients" & vbCrLf
            Else
                If InStr(1, extractemail(previewtopline), "@" & HostName) Then
                    'ok, it's to one of us, check first to see whether they exist (over size, too big, etc. would all come here)
                    cl = Mid(extractemail(previewtopline), 1, InStr(1, extractemail(previewtopline), "@") - 1)
                    If fso.FolderExists(fso.BuildPath(subfolder(""), cl)) Then
                        If getmailboxsize(CStr(cl)) < accountsize Then
                            outbuffer = outbuffer & "250 " & extractemail(previewtopline) & " Yeah, they have an account here" & vbCrLf
                            toaddr.Add extractemail(previewtopline)
                        Else
                            outbuffer = outbuffer & "552 " & extractemail(previewtopline) & " account is full. I will attempt to SMS them." & vbCrLf
                            quickmail getaccountinfo(cl, "sms"), "account overflowing", "Yr " & cl & "@" & HostName & " email acc is full. Plz delete some msg's. A msg has been bounced."
                            quickmail getaccountinfo(cl, "alt"), "account overflowing", "Your " & cl & "@" & HostName & " email account is overflowing. Please delete some msg's to free up space, as an email was just rejected that was addressed to you."
                        End If
                    Else
                        outbuffer = outbuffer & "502 " & extractemail(previewtopline) & " Account doesn't exist," & vbCrLf
                    End If
                Else
                    toaddr.Add extractemail(previewtopline)
                    outbuffer = outbuffer & "250 " & extractemail(previewtopline) & " recpient ok, will forward." & vbCrLf
                End If
            End If
        ElseIf LCase(Left(previewtopline, 4)) = "help" Then
            outbuffer = outbuffer & "250 specify email recipient(s), 1 per command (ie RCPT TO: me@myemail.com). send the 'DATA' command when done." & vbCrLf
        
        ElseIf LCase(Left(previewtopline, 4)) = "data" Then
            If toaddr.Count > 0 Then
                outbuffer = outbuffer & "354 " & toaddr.Count & " recipient" & IIf(toaddr.Count <> 1, "s", "") & " specified, enter mail data and end with . on a new line" & vbCrLf
                State = dGettingdata
            Else
                outbuffer = outbuffer & "503 And send it to who? huh? (expected RCPT)" & vbCrLf
            End If
        ElseIf extractemail(previewtopline) = "" Then
            outbuffer = outbuffer & "501 Bad Email Address on the line: " & previewtopline & vbCrLf
        Else
            outbuffer = outbuffer & "503 Expected 'RCPT' or 'DATA', see help for example" & vbCrLf
        End If
        pulltopline
    Case smtpinstate.dGettingdata
        If Right(inbuffer, 5) = vbCrLf & "." & vbCrLf Then
            body = Mid(inbuffer, 1, InStr(1, inbuffer, vbCrLf & "." & vbCrLf) - 1)
            inbuffer = Mid(inbuffer, InStr(1, inbuffer, vbCrLf & "." & vbCrLf) + 3)
            State = eDone
            outbuffer = outbuffer & "250 Ok, body completed" & vbCrLf
        End If
    Case smtpinstate.eDone
        If Left(previewtopline, 4) = "QUIT" Then
        
            Dim ts As TextStream
            For a = 1 To toaddr.Count
                fn = fso.BuildPath(subfolder("out"), Timer * 100 & Int(Rnd() * 10000) & ".txt")
                
                If InStr(1, toaddr(a), "@" & HostName) Then
                    'ok, why dont we just save it in our contacts maildrop, instead of our outbox
                    cl = Mid(toaddr(a), 1, InStr(1, toaddr(a), "@") - 1)
                    fn = fso.BuildPath(subfolder(CStr(cl)), Timer * 100 & Int(Rnd() * 10000) & ".txt")
                    If getaccountinfo(cl, "sms") <> "" Then
                        'they are using the cell phone notification system,
                        'so, lets notify their cell phone
                        Dim sendit As New inmail
                        sendit.moreincomming "HELO " & HostName & " webmail" & vbCrLf
                        sendit.parsebuffer
                        sendit.moreincomming "MAIL FROM: " & "mailsubsystem@" & HostName & vbCrLf
                        sendit.parsebuffer
                        sendit.moreincomming "RCPT TO: " & getaccountinfo(cl, "sms") & vbCrLf
                        sendit.parsebuffer
                        sendit.moreincomming "DATA" & vbCrLf
                        sendit.parsebuffer
                        sendit.moreincomming Left(body, 2000) & vbCrLf & "." & vbCrLf
                        sendit.parsebuffer
                        sendit.moreincomming "QUIT" & vbCrLf
                        sendit.parsebuffer
                    End If
                End If
                
                Set ts = fso.OpenTextFile(fn, ForWriting, True)
                ts.WriteLine fromaddr
                ts.WriteLine toaddr(a)
                ts.WriteLine "0"
                ts.WriteLine Now
                ts.Write body
                ts.Close
                DoEvents
                
            Next a
            outbuffer = outbuffer & "250 Ok, message added to que, will send shortly." & vbCrLf
        ElseIf Left(previewtopline, 4) = "RSET" Then
            outbuffer = "250 Ok, message canceled" & vbCrLf
            fromaddr = ""
            For a = 1 To toaddr.Count
                toaddr.Remove 1
            Next a
            body = ""
            inbuffer = ""
            State = aWanthelo
        Else
            outbuffer = outbuffer & "502 expected either QUIT or RSET" & vbCrLf
        End If
        pulltopline
    End Select
End Sub

Private Function previewtopline() As String
    a = InStr(1, inbuffer, vbCrLf)
    If a > 0 Then
        previewtopline = Mid(inbuffer, 1, a - 1)
    End If
End Function

Private Function pulltopline() As String
    a = InStr(1, inbuffer, vbCrLf)
    If a > 0 Then
        pulltop = Mid(inbuffer, 1, a - 1)
        inbuffer = Mid(inbuffer, a + 2)
    End If
End Function




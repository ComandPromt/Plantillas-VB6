Attribute VB_Name = "webmailhtml"
Private Const htmlopen As String = "Content-type: text/html" & vbCrLf & vbCrLf & "<HTML><HEAD><TITLE>"
Private Const bodystart As String = "</TITLE></HEAD><BODY><TABLE width=100% height=100% cols=2 rows=1>" & _
"<tr><td valign=top width=1><IMG border=0 src=""/img/logo.png"" width=128 height=128 alt=""Ashleys webmail:"">" & _
"<P><A HREF=""inbox.webmail""><Img border=0 src=""/img/inbox.png"">Inbox</A><BR><A HREF=""compose.webmail""><Img " & _
"border=0 src=""/img/compose.png"">Compose</A><BR><A HREF=""settings.webmail""><Img border=0 src=""/img/settings.png"">Settings</A><BR><A HREF=""logout.webmail""><Img border=0 src=""/img/logout.png"">Logout</A><BR><A HREF=""signup.webmail""><Img border=0 src=""/img/signup.png"">Signup</A></td><td valign=top>"
Private Const bodyend As String = "</td></tr></table>"

Public Function dowebsite(FileName, vars As Dictionary, headers As Dictionary, ByRef pageheader) As String
    If getaccountinfo(vars("un"), "pw") = vars("pw") And Len(vars("pw")) > 0 Then
        pageheader = pageheader & "Set-Cookie: un=" & vars("un") & "; expires=Fri 28-Jun-2012 13:25:03 GMT;  path=/; domain=" & HostName & ";" & vbCrLf
        pageheader = pageheader & "Set-Cookie: pw=" & vars("pw") & "; expires=Fri 28-Jun-2012 13:25:03 GMT;  path=/; domain=" & HostName & ";" & vbCrLf
    Else
        If Not (FileName = "login.webmail" Or FileName = "signup.webmail") Then
            pageheader = "HTTP/1.0 302 FOUND" & vbNewLine & "Server: Ashleys opensource mailserver" & vbNewLine & "Host: " & _
            HostName & vbNewLine & "Url: /mail/login.webmail" & vbNewLine & "Location: /mail/login.webmail" & vbNewLine & "Connection: close" & vbNewLine & _
            vbNewLine
        End If
    End If
    If FileName = "login.webmail" Then
        dowebsite = showloginpage
        Exit Function
    End If
    If FileName = "inbox.webmail" Then
        If vars("Delete") = "Delete" Then
            'the user wishes to delete some messages before the mail list is shown
            For Each f In vars.Keys
                If vars(f) = "dm" Then
                    On Error Resume Next
                    fso.DeleteFile fso.BuildPath(subfolder(vars("un")), f & ".txt"), True
                    On Error GoTo 0
                End If
            Next
        End If
        
        accmsgcount = getmsgcount(vars("un"))
        accsize = getmailboxsize(vars("un"))
        dowebsite = htmlopen & "Inbox for " & vars("un") & bodystart & "There " & IIf(accmsgcount = 1, " is ", " are ") & accmsgcount & _
        " message" & IIf(accmsgcount = 1, "", "s") & " in your inbox, taking up a total of " & _
        accsize & " bytes.<P>"
        dowebsite = dowebsite & "<form action=""inbox.webmail"" method=post><TABLE><tr bgcolor=00FFFF><td>&nbsp;</td><td width=50><B>From:</B></td><td width=300><B>Subject:</B></TD><td><b>Size:</B></td></tr>"
        For a = 1 To Form1.File2.ListCount - 1
            eid = Left(Form1.File2.List(a), Len(Form1.File2.List(a)) - 4)
            fn = fso.BuildPath(Form1.File2.Path, Form1.File2.List(a))
            dowebsite = dowebsite & "<tr><td><input type=checkbox name=" & eid & " value=dm></td><td>" & getmailheader(CStr(fn), "from") & "</td><td><A href=""showmsg.webmail?msg=" & eid & """>" & getmailheader(CStr(fn), "subject") & "</a></td><td>" & getmailsize(CStr(fn)) & "</td></tr>" & vbCrLf
        Next a
        dowebsite = dowebsite & "</table><Input type=submit name=Delete value=Delete></Form>" & bodyend
    End If
    
    If FileName = "showmsg.webmail" Then
        dowebsite = htmlopen & getmailheader(fso.BuildPath(subfolder(vars("un")), vars("msg") & ".txt"), "subject") & " - " & getmailheader(fso.BuildPath(subfolder(vars("un")), vars("msg") & ".txt"), "from") & bodystart
        Source = getmail(fso.BuildPath(subfolder(vars("un")), vars("msg") & ".txt"))
        mh = Left(Source, InStr(1, Source, vbCrLf & vbCrLf) - 1)
        body = Mid(Source, InStr(1, Source, vbCrLf & vbCrLf) + 3)
        Set hlist = parseheaders(CStr(mh))
        dowebsite = dowebsite & "<TABLE border=1 bordercolor=000000 width=550><tr><td>"
        For Each a In Array("From", "To", "Subject", "Date")
            dowebsite = dowebsite & "<B>" & a & ":</B> " & Replace(Replace(hlist(a), ">,<", ">, <"), "<", "&lt;") & "<BR>" & vbCrLf
        Next
        dowebsite = dowebsite & "<P>" & Replace(Replace(body, vbCrLf & vbCrLf, "<P>"), vbCrLf, "<BR>")
        dowebsite = dowebsite & "</Table><P> &nbsp; <P> &nbsp; <P><B>Message source follows:</B><P><PRE>" & Replace(Source, "<", "&lt;") & "</PRE>" & bodyend
    End If
    
    If FileName = "logout.webmail" Then
        pageheader = pageheader & "Set-Cookie: un=" & vars("un") & "; expires=Thu 28-Jun-2000 13:25:03 GMT;  path=/; domain=" & HostName & ";" & vbCrLf
        pageheader = pageheader & "Set-Cookie: pw=" & vars("pw") & "; expires=Thu 28-Jun-2000 13:25:03 GMT;  path=/; domain=" & HostName & ";" & vbCrLf
        dowebsite = htmlopen & "You have been logged out" & bodystart & "<A HREF="" / "">Sign In</A>" & bodyend
    End If
    
    If FileName = "compose.webmail" Then
        dowebsite = htmlopen & "Compose a new email message" & bodystart & "<B>Compose a new email</B><P><FORM action=""send.webmail"" method=post>" & _
        "<CENTER><TABLE><tr><td align=right><B>To:</B></td><td><INPUT size=60 name=to value=" & vars("to") & "></td></tr>" & _
        "<tr><td align=right><B>From:</B></td><td><INPUT size=60 name=from value=" & vars("un") & "@" & HostName & "></td></tr>" & _
        "<tr><td align=right><B>Subject:</B></td><td><INPUT size=60 name=subject value=" & vars("subject") & "></td></tr>" & _
        "</TABLE><BR>"
        
        dowebsite = dowebsite & "<TEXTAREA rows=16 cols=70 name=body></TEXTAREA><P><INPUT type=submit name=""Send"" value=""Send""></CENTER>"
        
        dowebsite = dowebsite & bodyend
    End If
    
    If FileName = "send.webmail" Then
        from = vars("from")
        ato = vars("to")
        subject = vars("subject")
        body = vars("body")
        
        data = "Recieved: " & HostName & " webmail, user=" & vars("un") & vbCrLf & _
        "From: " & from & vbCrLf & _
        "To: " & ato & vbCrLf & _
        "Date: " & Now & vbCrLf & _
        "Subject: " & subject & vbCrLf & vbCrLf & body
    
        'creates a nwe instance of the class that is used to control new mail arivals
        'and creates a new instance of it, pluging in the data from the website as oposed
        'to from a socket
        
        Dim sendit As New inmail
        sendit.moreincomming "HELO " & HostName & " webmail" & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming "MAIL FROM: " & from & vbCrLf
        sendit.parsebuffer
        For Each r In Split(ato, ",")
            sendit.moreincomming "RCPT TO: " & r & vbCrLf
            sendit.parsebuffer
        Next
        sendit.moreincomming "DATA" & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming data & vbCrLf & "." & vbCrLf
        sendit.parsebuffer
        sendit.moreincomming "QUIT" & vbCrLf
        sendit.parsebuffer
        
        dowebsite = htmlopen & "Mail has been sent!" & bodystart & "<B>Your message has been sent</B>, follows is the log from the server:<P><PRE>" & sendit.outbuffer & "</PRE>" & bodyend
    End If
    
    If FileName = "signup.webmail" Then
        If vars("Signup") = "Signup" Then
            If Len(vars("pw1")) < 3 Then
                er = "Password is too short"
                GoTo nope
            End If
            If Len(vars("unp")) < 3 Then
                er = "username is too short"
                GoTo nope
            End If
            If fso.FileExists(fso.BuildPath(subfolder(vars("unp")), "!account.txt")) Then
                er = "Account is allready taken"
                GoTo nope
            End If
            If vars("pw1") <> vars("pw2") Then
                er = "Passwords didn't match"
                GoTo nope
            End If
            
            'everything is fine, create their account
            Dim ts As TextStream
            Set ts = fso.OpenTextFile(fso.BuildPath(subfolder(vars("unp")), "!account.txt"), ForWriting, True)
            ts.WriteLine "pw: " & vars("pw1")
            ts.WriteLine "alt: " & vars("alt")
            ts.WriteLine "sms: " & vars("sms")
            ts.Close
            
            data = "From: Welcomebot@" & HostName & vbCrLf & _
            "To: New " & HostName & " user" & vbCrLf & _
            "Subject: Welcome to " & HostName & vbCrLf & _
            "Date: " & Now & vbCrLf & vbCrLf & _
            "Welcome to the " & HostName & " mail service!" & vbCrLf & vbCrLf & _
            "Username: " & vars("unp") & vbCrLf & _
            "Password: " & vars("pw1") & vbCrLf & _
            "POP3 server: " & HostName & vbCrLf & _
            "SMTP server: " & HostName & vbCrLf & vbCrLf & _
            "Additionally, you can use webmail by going to http://" & HostName & vbCrLf & vbCrLf & _
            "Thank you for signing up with " & HostName & vbCrLf & vbCrLf
            
            'send them an email welcomming them
            Set sendit = New inmail
            sendit.moreincomming "HELO " & HostName & " webmail" & vbCrLf
            sendit.parsebuffer
            sendit.moreincomming "MAIL FROM: " & "Welcome! <Welcomebot@" & HostName & ">" & vbCrLf
            sendit.parsebuffer
            sendit.moreincomming "RCPT TO: " & vars("unp") & "@" & HostName & vbCrLf
            sendit.parsebuffer
            sendit.moreincomming "RCPT TO: " & vars("alt") & vbCrLf
            sendit.parsebuffer
            sendit.moreincomming "DATA" & vbCrLf
            sendit.parsebuffer
            sendit.moreincomming data & vbCrLf & "." & vbCrLf
            sendit.parsebuffer
            sendit.moreincomming "QUIT" & vbCrLf
            sendit.parsebuffer
            
            dowebsite = htmlopen & "Signup successful" & bodystart & "Signup sucessful, <A HREF=""login.webmail"">Login</A> to continue." & bodyend

        Else
nope:
            dowebsite = htmlopen & "Sign up for " & HostName & " mail service" & bodystart & _
            "Sign up for the " & HostName & " mail server and get:<P><UL>"
            For Each x In Array("Full POP3 access", "Webmail access", "SMS notification", "Full 1mb storage space", "Autoresponder")
                dowebsite = dowebsite & "<LI> " & x & "<BR>" & vbCrLf
            Next
            dowebsite = dowebsite & "</UL><P>To get all that, and more, simply sign up on the form below:<P>" & _
            "<FORM action=""signup.webmail"" method=post>" & _
            "<FONT SIZE=6 COLOR=FF0000>" & er & "</FONT>" & _
            "<TABLE>" & _
            "<tr><td>Username:</td><td><INPUT name=unp></td></tr>" & _
            "<tr><td>Password:</td><td><INPUT name=pw1 type=password></td></tr>" & _
            "<tr><td>Confirm:</td><td><INPUT name=pw2 type=password></td></tr>" & _
            "<tr><td colspan=2> &nbsp </td></tr>" & _
            "<tr><td>Alternate email:</td><td><INPUT name=alt></td></tr>" & _
            "<tr><td>SMS email:</td><td><INPUT name=sms></td></tr>" & _
            "</TABLE><INPUT type=submit name=Signup value=Signup></FORM>" & bodyend
        End If
    End If
    
    If FileName = "settings.webmail" Then
        dowebsite = htmlopen & "Modify your settings" & bodystart & "Modify your settings<P>" & vbCrLf
        
        If vars("Save") = "Save" Then
            If vars("mod") = "pw" Then
                If vars("pw1") <> vars("pw") Then
                    dowebsite = dowebsite & "<FONT size=5 color=FF0000>Enter your current password</FONT>"
                    GoTo no
                End If
                If vars("pw2") <> vars("pw3") Then
                    dowebsite = dowebsite & "<FONT size=5 color=FF0000>Those passwords dont match</FONT>"
                    GoTo no
                End If
                If Len(vars("pw2")) < 3 Then
                    dowebsite = dowebsite & "<FONT size=5 color=FF0000>Password must be 3 letters long</FONT>"
                    GoTo no
                End If
                Set ts = fso.OpenTextFile(fso.BuildPath(subfolder(vars("un")), "!account.txt"), ForAppending, True)
                ts.WriteLine "pw: " & vars("pw3")
                dowebsite = dowebsite & "<FONT size=5 color=FF0000>Password changed, please log back in</FONT>"
                ts.Close
            End If
            If vars("mod") = "alt" Then
                Set ts = fso.OpenTextFile(fso.BuildPath(subfolder(vars("un")), "!account.txt"), ForAppending, True)
                ts.WriteLine "alt: " & vars("alt")
                dowebsite = dowebsite & "<FONT size=5 color=FF0000>Alternate email changed</FONT>"
                ts.Close
            End If
            If vars("mod") = "sms" Then
                Set ts = fso.OpenTextFile(fso.BuildPath(subfolder(vars("un")), "!account.txt"), ForAppending, True)
                ts.WriteLine "sms: " & vars("sms")
                dowebsite = dowebsite & "<FONT size=5 color=FF0000>SMS email changed</FONT>"
                ts.Close
            End If
            vars("mod") = ""
        End If
no:
        dowebsite = dowebsite & "<FORM action=""settings.webmail"" method=POST><TABLE>"
        If vars("mod") = "pw" Then
            dowebsite = dowebsite & "<tr><td align=right><b>Old password</b>:</td><td><INPUT name=pw1 type=password></td></tr>" & vbCrLf
            dowebsite = dowebsite & "<tr><td align=right><b>New password</b>:</td><td><INPUT name=pw2 type=password></td></tr>" & vbCrLf
            dowebsite = dowebsite & "<tr><td align=right><b>Confirm</b>:</td><td><INPUT name=pw3 type=password></td></tr>" & vbCrLf
        Else
            dowebsite = dowebsite & "<tr><td align=right><b>Password</b>:</td><td>" & String(Len(vars("pw")), "*") & " <small><A HREF=""settings.webmail?mod=pw"">Change</A></small></td></tr>" & vbCrLf
        End If
        
        If vars("mod") = "alt" Then
            dowebsite = dowebsite & "<tr><td align=right><b>Alternate email</b>:</td><td><INPUT name=alt value=" & getaccountinfo(vars("un"), "alt") & "></td></tr>" & vbCrLf
        Else
            dowebsite = dowebsite & "<tr><td align=right><b>Alternate email</b>:</td><td>" & getaccountinfo(vars("un"), "alt") & " <small><A HREF=""settings.webmail?mod=alt"">Change</A></small></td></tr>" & vbCrLf
        End If
        
        If vars("mod") = "sms" Then
            dowebsite = dowebsite & "<tr><td align=right><b>SMS email</b>:</td><td><INPUT name=sms value=" & getaccountinfo(vars("un"), "sms") & "></td></tr>" & vbCrLf
        Else
            dowebsite = dowebsite & "<tr><td align=right><b>SMS email</b>:</td><td>" & getaccountinfo(vars("un"), "sms") & " <small><A HREF=""settings.webmail?mod=sms"">Change</A></small></td></tr>" & vbCrLf
        End If
        dowebsite = dowebsite & "</TABLE>"
        If vars("mod") <> "" Then dowebsite = dowebsite & "<INPUT type=hidden name=mod value=" & vars("mod") & "><INPUT type=submit name=Save value=Save></FORM>"
        dowebsite = dowebsite & bodyend
    End If
End Function

Public Function showloginpage() As String
    t = htmlopen & "Check your " & HostName & " mail from the web!!" & bodystart
    m = "No matter where you are on the planet, you can check your email from this page!" & "<P>" & _
    "Just enter your username and password to continue.<P>" & _
    "<FORM action=""/mail/inbox.webmail"" method=post>" & _
    "<INPUT type=text name=un> Username<BR>" & _
    "<INPUT type=password name=pw> Password<BR>" & _
    "<INPUT type=submit name=""login"" value=""login"">" & _
    "</FORM>" & _
    vbCrLf & printjavascript
    
    b = bodyend
    
    showloginpage = t & m & b
End Function

Public Function printjavascript() As String
p = "<SCRIPT>" & _
"if (document.location.host != '" & HostName & "')" & _
"{" & _
 "document.location='http://" & HostName & "';" & _
"}" & _
"</SCRIPT>"
printjavascript = p
End Function

Attribute VB_Name = "functions"
Public fso As New FileSystemObject

Private mxtable As New Dictionary

Public Const HostName As String = "192.168.0.20"

Public Const accountsize = 5242880 ' (5mb)

Private md5 As New md5

Public Function checksum(InString) As String
    checksum = md5.DigestStrToHexStr(CStr(InString))
End Function

Public Function mxlookup(dommain As String) As String
    If extractip(dommain) <> "" Then mxtable(dommain) = dommain
    mxlookup = mxtable(dommain)
    If mxlookup = "" Then
        Form1.mx.Domain = dommain
        mxtable(dommain) = Form1.mx.GetMX
        mxlookup = mxtable(dommain)
    End If
End Function

Public Function extractemail(emin As String) As String
    'extracts the email address from: (using regexps)
    'Ashley Harris <ashley___harris@hotmail.com>
    'ashley___harris@hotmail.com
    '<ashley___harris@hotmail.com>
    'RCPT TO: Ashley H@rris <Ashley___Harris@hotmail.com> <-- He made this
    'or any other circumstance
    Dim re As New RegExp
    re.IgnoreCase = True
    
    re.Pattern = "[abcdefghijklmnopqrstuvwxyz_.-0123456789]{1,64}@[abcdefghijklmnopqrstuvwxyz_.-0123456789]{1,64}\.[abcdefghijklmnopqrstuvwxyz0123456789]{1,6}"
    On Error Resume Next
    extractemail = re.Execute(emin)(0)
End Function

Public Function subfolder(subfoldername As String) As String
    subfolder = fso.BuildPath(fso.BuildPath(App.Path, "email"), subfoldername)
    If Not fso.FolderExists(subfolder) Then
        On Error Resume Next
        fso.CreateFolder (fso.GetParentFolderName(subfolder))
        fso.CreateFolder subfolder
    End If
End Function

Public Function extractip(emin As String) As String
    'extracts the ip address from: (using regexps)
    Dim re As New RegExp
    re.IgnoreCase = True
    
    re.Pattern = "[0123456789]{1,3}\.[0123456789]{1,3}\.[0123456789]{1,3}\.[0123456789]{1,3}"
    On Error Resume Next
    extractip = re.Execute(emin)(0)
End Function

Public Function getaccountinfo(accountname, key) As String
    If accountname = "" Then Exit Function
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(fso.BuildPath(subfolder(CStr(accountname)), "!account.txt"), ForReading, True)
    While Not ts.AtEndOfStream
        a = ts.ReadLine
        If Left(a, Len(key)) = key Then
            getaccountinfo = Mid(a, InStr(1, a, ":") + 1)
            'GoTo out
        End If
    Wend
out:
    ts.Close
    If Left(getaccountinfo, 1) = " " Then getaccountinfo = Mid(getaccountinfo, 2)
End Function

Public Function getmailsize(FileName As String) As Long
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(FileName)
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    content = ts.ReadAll
    getmailsize = Len(content)
    ts.Close
End Function

Public Function getmail(FileName As String) As String
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(FileName)
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    content = ts.ReadAll
    getmail = content
    ts.Close
End Function

Public Function getmailheader(FileName As String, headername As String) As String
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(FileName)
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    content = ts.ReadAll
    Dim b As Dictionary
    Set b = parseheaders(CStr(Mid(content, 1, InStr(1, content, vbCrLf & vbCrLf) - 1)))
    ts.Close
    getmailheader = b(headername)
End Function

Public Function getmailboxsize(mbox As String) As Long
    Form1.File2.Path = subfolder(mbox)
    Form1.File2.Refresh
    For a = 1 To Form1.File2.ListCount - 1
        getmailboxsize = getmailboxsize + getmailsize(fso.BuildPath(Form1.File2.Path, Form1.File2.List(a)))
    Next a
End Function

Public Function getmsgcount(mbox As String) As Long
    Form1.File2.Path = subfolder(mbox)
    Form1.File2.Refresh
    getmsgcount = Form1.File2.ListCount - 1
End Function

Public Sub quickmail(toaddr, subject, data)
    data = "To: " & toaddr & vbCrLf & _
    "From: mailsubsystem@" & HostName & vbCrLf & _
    "Subject: " & subject & vbCrLf & _
    "To: " & toaddr & vbCrLf & _
    "Date: " & Now & vbCrLf & _
    vbCrLf & data

    Dim sendit As New inmail
    sendit.moreincomming "HELO " & HostName & " webmail" & vbCrLf
    sendit.parsebuffer
    sendit.moreincomming "MAIL FROM: " & "mailsubsystem@" & HostName & vbCrLf
    sendit.parsebuffer
    sendit.moreincomming "RCPT TO: " & toaddr & vbCrLf
    sendit.parsebuffer
    sendit.moreincomming "DATA" & vbCrLf
    sendit.parsebuffer
    sendit.moreincomming data & vbCrLf & "." & vbCrLf
    sendit.parsebuffer
    sendit.moreincomming "QUIT" & vbCrLf
    sendit.parsebuffer
End Sub


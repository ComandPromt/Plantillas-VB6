Attribute VB_Name = "Module1"
'if you want to change the location of the INI file
'this is where you do it.  If you dont want to change
'the location of the ini file, then save it to the same
'directory as the application

Public Const con_INI_File As String = "text.INI"

Public strPath As String
Public strFileTypes As String

'Thanks to planetsourcecode for this snippet of code that
'reads and writes to an ini file
'---> http://www.planetsourcecode.com <----



'**************************************
' Name: INI File Read/Write
'
' Description:This code allows the user
'     to read and write INI file values withou
'     t using *any* API calls, therefore elimi
'     nating any incompatibility issues with
'     Win31/95/98/nt
'
' Inputs:  GetValue("category", "variable", "filename342.ini")
'          PutValue "category", "variable", "value", "filename.ini"
'
' Returns:GetValue - Value of Fieldname
'
' Assumes:None
'
'**************************************

Public Function GetValue(getcat, getfield, getfile) As String
    'example usage:
    'username = GetValue("UserInfo", "Username", "myprog.ini")
    
    If Dir(getfile) = "" Then Exit Function
    getcat = LCase(getcat)
    getfield = LCase(getfield)
    fnum = FreeFile
    Open getfile For Input As fnum

    Do While Not EOF(fnum)
        Line Input #fnum, l1
        l1 = Trim(l1)
        l1 = LCase(l1)


        If InStr(l1, "[") <> 0 Then
            If LCase(Mid(l1, (InStr(l1, "[") + 1), (Len(l1) - 2))) = getcat Then
                Do Until EOF(fnum) Or l2 = "["
                    Line Input #fnum, l2
                    l2 = Trim(l2)
                    If InStr(l2, "]") <> 0 Then
                        Close fnum
                        Exit Function
                    End If
                    If InStr(l2, "=") <> 0 Then
                        If LCase(Left(l2, (InStr(l2, "=") - 1))) = getfield Then
                            GetValue = Trim(Mid(l2, InStr(l2, "=") + 1, Len(l2)))
                            Close fnum
                            Exit Function
                        End If
                    End If
                Loop
            End If
        End If
    Loop
    Close fnum
End Function


Public Sub PutValue(putcat, putvar, putval, putfile)
    Dim fileCol(1 To 9000) As String
    Dim foundCat As Boolean
    Dim foundVar As Boolean
    Dim catPos As Integer
    Dim varPos As Integer
    fnum = FreeFile
    putcat = Trim(putcat)
    putcat = LCase(putcat)
    putfile = Trim(putfile)
    putfile = LCase(putfile)
    putvar = LCase(putvar)
    putvar = Trim(putvar)
    putval = LCase(putval)
    putval = Trim(putval)


    If Dir(putfile) = "" Then
        Open putfile For Append As #fnum
        Close #fnum
    End If
    Open putfile For Input As #fnum


    Do While Not EOF(fnum)
        DoEvents
            Counter = Counter + 1
            Line Input #fnum, l1
            fileCol(Counter) = l1
        Loop
        Close #fnum
        For i = 1 To Counter
            DoEvents
                If InStr(LCase(fileCol(i)), "[" & putcat & "]") <> 0 Then
                    foundCat = True
                    catPos = i
                    For x = i To Counter
                        DoEvents
                            If InStr(fileCol(x), "[") <> 0 And LCase(fileCol(x)) <> "[" & putcat & "]" Then Exit For
                            If InStr(LCase(fileCol(x)), putvar & "=") <> 0 Then
                                foundVar = True
                                varPos = x
                            End If
                        Next x
                    End If
                Next i
                If foundCat = True And foundVar = True Then
                    fileCol(varPos) = putvar & "=" & putval
                    Kill putfile
                    Open putfile For Append As #fnum
                    For i = 1 To Counter
                        Print #fnum, fileCol(i)
                        DoEvents
                        Next i
                        Close #fnum
                        Exit Sub
                    End If
                    If foundCat = True And foundVar = False Then
                        Kill putfile
                        Open putfile For Append As #fnum
                        For i = 1 To Counter
                            Print #fnum, fileCol(i)
                            If i = catPos Then Print #fnum, putvar & "=" & putval
                        Next i
                        Close #fnum
                        Exit Sub
                    End If
                    If foundCat = False And foundVar = False Then
                        Kill putfile
                        Open putfile For Append As #fnum
                        For i = 1 To Counter
                            Print #fnum, fileCol(i)
                        Next i
                        Print #fnum, "[" & putcat & "]"
                        Print #fnum, putvar & "=" & putval
                        Close #fnum
                    End If
                End Sub

'**************************************
'Encrypt/Decrypt
'**************************************


'Obviously it wouldn't be that hard to break this encryption
'
'This program is made that if you wanted more security
'It could be added to this mondule and thats it.  The program
'writes to an ini file (from planetsourcecode - thanks!) and
'password then gets encrypted witha a simple algorythm
'I had some problems using all the chars...due to either
'spaces or returns, but this works and a lot of the time thats
'all that is needed.  Simple password protection with encrytion.
'


'Modified to shift characters in password by 5 characters
Function Encrypt(strPW As String)
    
    Dim intLetterCntr As Integer
    Dim strLetter As String
    Dim intLetter As Integer
    Dim strEncPW As String
    
    strEncPW = ""
    
    For intLetterCntr = 1 To Len(strPW)
        strLetter = Mid(strPW, intLetterCntr, 1)
        intLetter = (Asc(strLetter) + 5)
        ' if you want to change it from 5 chars to
        ' whatever...do it in the decrypt too
        
        If intLetter > 122 Then
            intLetter = intLetter - 26
        End If
        
        strEncPW = strEncPW & Chr(intLetter)
        
    Next intLetterCntr
    Encrypt = strEncPW
    
End Function


Function Decrypt(strEncPW As String)
    Dim intLetterCntr As Integer
    Dim strLetter As String
    Dim intLetter As Integer
    Dim strDecPW As String
    
    strDecPW = ""
    
    For intLetterCntr = 1 To Len(strEncPW)
        strLetter = Mid(strEncPW, intLetterCntr, 1)
        intLetter = (Asc(strLetter) - 5)
        ' right here
        
        If intLetter < 97 Then
            intLetter = intLetter + 26
        End If
        
        
        strDecPW = strDecPW & Chr(intLetter)
        
    Next intLetterCntr
    Decrypt = strDecPW
    
End Function



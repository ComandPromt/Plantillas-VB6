Attribute VB_Name = "ModuleAgent"
Option Explicit

Public Char As IAgentCtlCharacterEx     'declare agent
Public arrSayings() As String           'array for sayings
Public arrActions() As String           'array for actions
Public arrMe() As String                'array for user info

'api used to get main operating system directory
Private Declare Function GetWindowsDirectory Lib "kernel32" _
    Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
   
Public Function GetWindowsDir() As String
    'gets windows directory whether windows or winnt
    Dim Temp As String
    Dim Ret As Long
    Const MAX_LENGTH = 145

    Temp = String$(MAX_LENGTH, 0)
    Ret = GetWindowsDirectory(Temp, MAX_LENGTH)
    Temp = Left$(Temp, Ret)
    If Temp <> "" And Right$(Temp, 1) <> "\" Then
        GetWindowsDir = Temp & "\"
    Else
        GetWindowsDir = Temp
    End If

End Function

Public Sub IdleOn()
    'set char to idle
    Char.Play "Restpose"
    Char.IdleOn = True
End Sub

Public Sub IdleOff()
    'turn off idle - really dont need it
    Char.IdleOn = False
End Sub

Public Sub RestPose()
    'set character to rest
    Char.Play "Restpose"
End Sub

Public Sub RandomSpeak()
    'randomly grab a sayinging the speak it
    Dim Upper As Integer
    
    If arrSayings(1) <> "" Then
        Upper = UBound(arrSayings)          'finds last item in array
        If Upper <> 1 Then
            Char.Speak arrSayings(Rnd * Upper)  'picks random from begining to end of array(upper)
        End If
    End If
    
    IdleOn
    
End Sub

Public Sub RandomMove()
    
    Dim Upper As Integer
    Dim Left As Integer
    Dim Top As Integer
    
    If arrActions(1) <> "" Then
         Upper = UBound(arrActions)         'finds last item in array
        
         Left = Rnd * 600                   'random left
         Top = Rnd * 800                    'random height
         Char.MoveTo Left, Top              'move char x,y
         Char.Play arrActions(Rnd * Upper)  'picks random from begining to end of array(upper)
    End If
    IdleOn
   
End Sub

Public Sub RandomQuote()
    'opens file, fills array with text strings
    'sets upper to last item in array then picks random saying
    'from begining to end
    Dim strFile As String
    Dim Upper As Integer
    ReDim arrQuotes(0)

    Open "C:\program files\DesktopAnnoyance\quotes.lst" For Input As #1
    
    Do While Not EOF(1)
        Input #1, strFile
        ReDim Preserve arrQuotes(UBound(arrQuotes) + 1)
        arrQuotes(UBound(arrQuotes)) = strFile
    Loop
    
    Close #1
    
    If arrQuotes(1) <> "" Then
        Upper = UBound(arrQuotes)
        Char.Speak arrQuotes(Rnd * Upper)
    End If
    
End Sub

Public Sub RandomFact()
    'opens file, fills array with text strings
    'sets upper to last item in array then picks random saying
    'from begining to end
    Dim arrFacts() As String
    Dim strFile As String
    Dim Upper As Integer
    ReDim arrFacts(0)

    Open "C:\program files\DesktopAnnoyance\facts.lst" For Input As #1
    
    Do While Not EOF(1)
        Input #1, strFile
        ReDim Preserve arrFacts(UBound(arrFacts) + 1)
        arrFacts(UBound(arrFacts)) = strFile
    Loop
    
    Close #1
    
    If arrFacts(1) <> "" Then
        Upper = UBound(arrFacts)
        Char.Speak arrFacts(Rnd * Upper)
    End If
    
End Sub

Public Sub RandomOxy()
    'opens file, fills array with text strings
    'sets upper to last item in array then picks random saying
    'from begining to end
    Dim arrOxy() As String
    Dim strFile As String
    Dim Upper As Integer
    ReDim arrOxy(0)

    Open "C:\program files\DesktopAnnoyance\oxymorons.lst" For Input As #1
    
    Do While Not EOF(1)
        Input #1, strFile
        ReDim Preserve arrOxy(UBound(arrOxy) + 1)
        arrOxy(UBound(arrOxy)) = strFile
    Loop
    
    Close #1
    
    If arrOxy(1) <> "" Then
        Upper = UBound(arrOxy)
        Char.Speak arrOxy(Rnd * Upper)
    End If
    
End Sub

Public Sub RandomPonder()
    'opens file, fills array with text strings
    'sets upper to last item in array then picks random saying
    'from begining to end
    Dim arrPonder() As String
    Dim strFile As String
    Dim Upper As Integer
    ReDim arrPonder(0)

    Open "C:\program files\DesktopAnnoyance\ponder.lst" For Input As #1
    
    Do While Not EOF(1)
        Input #1, strFile
        ReDim Preserve arrPonder(UBound(arrPonder) + 1)
        arrPonder(UBound(arrPonder)) = strFile
    Loop
    
    Close #1
    
    If arrPonder(1) <> "" Then
        Upper = UBound(arrPonder)
        Char.Speak arrPonder(Rnd * Upper)
    End If
    
End Sub

Public Sub RandomJokes()
    'opens file, fills array with text strings
    'sets upper to last item in array then picks random saying
    'from begining to end
    Dim arrJokes() As String
    Dim strFile As String
    Dim Upper As Integer
    ReDim arrJokes(0)

    Open "C:\program files\DesktopAnnoyance\lame.lst" For Input As #1
    
    Do While Not EOF(1)
        Input #1, strFile
        ReDim Preserve arrJokes(UBound(arrJokes) + 1)
        arrJokes(UBound(arrJokes)) = strFile
    Loop
    
    Close #1
    
    If arrJokes(1) <> "" Then
        Upper = UBound(arrJokes)
        Char.Speak arrJokes(Rnd * Upper)
    End If
    
End Sub

Public Sub RandomMurphy()
    'opens file, fills array with text strings
    'sets upper to last item in array then picks random saying
    'from begining to end
    Dim arrMurphy() As String
    Dim strFile As String
    Dim Upper As Integer
    ReDim arrMurphy(0)

    Open "C:\program files\DesktopAnnoyance\murphy.lst" For Input As #1
    
    Do While Not EOF(1)
        Input #1, strFile
        ReDim Preserve arrMurphy(UBound(arrMurphy) + 1)
        arrMurphy(UBound(arrMurphy)) = strFile
    Loop
    
    Close #1
    
    If arrMurphy(1) <> "" Then
        Upper = UBound(arrMurphy)
        Char.Speak arrMurphy(Rnd * Upper)
    End If
    
End Sub

Public Function KeyCheck(ByVal intKey As Integer) As Integer
 
    'only allow vbkeyescape, numeric
 
    Select Case intKey
        
        Case vbKeyEscape
            KeyCheck = intKey
        
        Case Asc("0") To Asc("9")
            KeyCheck = intKey
            
        Case vbKeyBack
            KeyCheck = intKey
        Case Else
            KeyCheck = 0
    
    End Select
 
End Function

Public Sub PlayIntro()
    'intro used for new user or daily opening
    
    Char.MoveTo 300, 0

    If frmAgent.txtName.Text = "" Then
        Char.Speak "Welcome"
        Char.Play "Greet"
        Char.Play "Restpose"
        Char.Play "Blink"
        Char.Speak "This is your Desktop Annoyance Window"
        Char.Play "Gesturedown"
        Char.Play "Restpose"
        Char.Play "blink"
        Char.Speak "Please help me know you better by filling out the" & _
                        " questions under the Me tab"
        Char.Play "Restpose"
        Char.Speak "Add some quotes you would like me to say during our time " & _
                        "together under the Sayings Tab."
        Char.Play "Gesturedown"
        Char.Speak "Edit me under the Options Tab."
        Char.Play "Restpose"
        Char.Play "Blink"
        Char.Speak "You can also Right-Click me to get a mini menu of choices."
        Char.Speak "Don't forget to email Alex with your credit card number " & _
                    "so he can clear out your bank account."
        Char.Speak "Did I ever tell you how great you look today!"
        Char.Play "Think"
        Char.Think "hehe, except for that spinach in your teeth."
        Char.Play "Blink"
        Char.Speak "Oops, I forgot you can see what I think, ha ha ha."
    Else
        Char.Play "Greet"
        If frmAgent.optBoy = True Then
            Char.Speak "It is so good to see you again master " & frmAgent.txtName.Text & "."
            Char.Speak "Today is a perfect day to be " & frmAgent.txtAge & "."
        Else
            Char.Speak "It is so good to see you again mistress " & frmAgent.txtName.Text
            Char.Speak "Today is a perfect day to be " & frmAgent.txtAge & "."
        End If
        
    End If
    
End Sub




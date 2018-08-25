Attribute VB_Name = "somehttpfunctions"
'These functions have been ripped strait from my webserver
' "(complete perl, plus php and asp support, resume downloads,
'   SSI) - Be your own geocities" on PSC.
'check it out. if you want.

Public Function tophpvariables(getdata As String, postdata As String, cookiedata As String) As Dictionary
    'takes the get data (?message=1), the post data (Feedback=Hi%2E+Im+Ashley etc), and cookies (session=1414)
    'and converts it all to one nice data dictonary
    Dim back As New Dictionary
    getdata = Replace(cookiedata, ";", "&") & "&" & getdata & "&" & postdata
    getdata = Replace(getdata, " ", "")
    getdata = Replace(getdata, vbCrLf, "")
    
    Keys = Split(getdata, "&")
    For a = 0 To UBound(Keys)
        If InStr(1, Keys(a), "=") = 0 Then Keys(a) = Keys(a) & "="
        K = Mid(Keys(a), 1, InStr(1, Keys(a), "=") - 1)
        v = Mid(Keys(a), InStr(1, Keys(a), "=") + 1)
        K = fromhttpstringtostring(CStr(K))
        v = fromhttpstringtostring(CStr(v))
        If K <> "" Then back(K) = v
    Next a
    Set tophpvariables = back
End Function

Public Function fromhttpstringtostring(httpstring As String) As String
    'turns 'This%20is%20cool' into 'This is cool'
    httpstring = Replace(httpstring, "+", " ")
    While InStr(1, httpstring, "%")
        fromhttpstringtostring = fromhttpstringtostring & Mid(httpstring, 1, InStr(1, httpstring, "%") - 1)
        httpstring = Mid(httpstring, InStr(1, httpstring, "%"))
        esc = Mid(httpstring, 1, 3)
        ch = Chr(hexdiget(Mid(esc, 2, 1)) * 16 + hexdiget(Mid(esc, 3, 1)))
        httpstring = Replace(httpstring, esc, ch)
    Wend
    fromhttpstringtostring = fromhttpstringtostring & httpstring
End Function

Public Function hexdiget(d) As Integer
    'converts a number from 0-15 into a hexeqiverlant (ie a=10)
    If d = Val(CStr(d)) Then hexdiget = d: Exit Function
    Select Case LCase(d)
    Case "a"
        hexdiget = 10
    Case "b"
        hexdiget = 11
    Case "c"
        hexdiget = 12
    Case "d"
        hexdiget = 13
    Case "e"
        hexdiget = 14
    Case "f"
        hexdiget = 15
    End Select
End Function


Public Function parseheaders(h As String) As Dictionary
    'turn
    'cookie: name=ashley
    'referer: www.pornrus.com
    'accept: all the stuff that goes here.
    'langauge: en-au
    'etc.
    'into a datadictonary.
    
    'would also be really usefull for email header parsing, but, I
    'only use it for http request parsing
    Dim K As String, v As String
    Set p = New Dictionary
    p.CompareMode = TextCompare
    h = h & vbNewLine
    h = Replace(h, ": ", ":")
    h = Replace(h, vbCrLf & " ", "")
    h = Replace(h, vbCrLf & vbTab, "")
    While h <> vbNewLine And h <> ""
        K = LCase(Mid(h, 1, InStr(1, h, ":") - 1))
        h = Mid(h, Len(K) + 2)
        v = Mid(h, 1, InStr(1, h, vbNewLine) - 1)

        h = Mid(h, Len(v) + 3)
        p(K) = v
    Wend
    Set parseheaders = p
End Function

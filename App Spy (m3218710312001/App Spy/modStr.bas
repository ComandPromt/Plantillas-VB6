Attribute VB_Name = "modStr"
Sub Searchfile(sFile As String, sSearch As String, ibyte _
    As Boolean, iUniCode As Boolean)
    'sFile - file name
    'sSearch - string to search for
    'ibyte - use byte array to search
    'iUniCode - look for UniCode strings
    Dim iHandle As Integer
    Dim sTemp As String
    Dim lSpot As Long
    Dim lFind As Long
    Dim sSearch1 As String
    Dim bTemp() As Byte
    'another advantage of using a byte array
    '
    'is that we can easily look for UniCode
    '     strings


    If iUniCode Or (Not ibyte) Then
        'this line will look for unicode strings
        '
        'when using byte arrays, regular
        'strings when using string variable
        sSearch1 = sSearch
    Else
        'this line will look for ANSII strings
        'when looking through a byte array
        sSearch1 = StrConv(sSearch, vbFromUnicode)
    End If
    iHandle = FreeFile
    Open sFile For Binary Access Read As iHandle


    If iHandle Then
        sTemp = Space$((LOF(iHandle) / 2) + 1)
        ReDim bTemp(LOF(iHandle)) As Byte


        If ibyte Then
            Get #iHandle, , bTemp
            sTemp = bTemp
        Else
            Get #iHandle, , sTemp
        End If
        Close iHandle
    End If


    Do


        If ibyte Then
            lFind = InStrB(lSpot + 1, sTemp, sSearch1, 1)
        Else
            lFind = InStr(lSpot + 1, sTemp, sSearch1, 1)
        End If
        lSpot = lFind
    Loop Until lFind = 0
End


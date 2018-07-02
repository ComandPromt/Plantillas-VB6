Attribute VB_Name = "modWavHeader"
Type WaveHeaderInfo
 channels As Integer
 freq As String
 bits As Integer
 kbps As Long
 wFilesize As String
 wPlaytime As String
End Type
Public wInfo As WaveHeaderInfo

Public Function wHeadInfo(wavFileName As String) As Boolean
Dim wMins, wSecs As String
Dim riff As String * 4
Dim freq1x As String
Dim freq1 As Byte

wInfo.wFilesize = FileLen(wavFileName)
Open wavFileName For Binary As #1   'open file #1 for read
    Get #1, 1, riff
    Get #1, 23, wInfo.channels
    Get #1, 35, wInfo.bits
    Get #1, 25, freq1
Close #1
freq1x = ConvertBase(Val(Str(freq1)), 10, 16)
If riff <> "RIFF" Then wHeadInfo = False: Exit Function
wHeadInfo = True

Select Case freq1x
Case 40
wInfo.freq = "8000"
Case 11
wInfo.freq = "11025"
Case 22
wInfo.freq = "22050"
Case 0
wInfo.freq = "32000"
Case 44
wInfo.freq = "44100"
Case 80
wInfo.freq = "48000"
End Select

wInfo.kbps = (CLng(wInfo.bits) * CLng(wInfo.channels) * CLng(wInfo.freq))
wInfo.wPlaytime = (((wInfo.wFilesize * 8) - 8000) / wInfo.kbps)
wMins = wInfo.wPlaytime \ 60
wSecs = wInfo.wPlaytime - (wMins * 60)
wInfo.wPlaytime = Format(wMins, "#0#") & ":" & Format(wSecs, "0#")
End Function


'This Function
Public Function ConvertBase(NumIn As String, BaseIn As Integer, _
    BaseOut As Integer) As String
    ' Converts a number from one base to another
    ' E.g. Binary = Base 2
    'Octal = Base 8
    'Decimal = Base 10
    'Hexadecimal = Base 16
    ' NumIn is the number which you wish to convert
        ' (A String including characters 0 - 9, A - Z)
    ' BaseIn is the base of NumIn (An integer value in
        ' Decimal between 1 & 36)
    ' BaseOut is the base of the number the function
        ' returns (An Integer value in Decimal between 1 & 36)
    ' Returns a string in the desired base containing the
        ' characters 0 - 9, A - Z)
    ' e.g. Debug.Print ConvertBase ("42", 8, 16) converts the octal n
    '     umber 42 into hexadecimal
    ' Returns the string "22"
    ' Returns the word "Error" if any of the input values
        ' are incorrect
    Dim i As Integer, CurrentCharacter As String, _
    CharacterValue As Integer, PlaceValue As Integer, _
    RunningTotal As Double, Remainder As Double, _
    BaseOutDouble As Double, NumInCaps As String
    ' Ensure input data is valid
    If NumIn = "" Or BaseIn < 2 Or BaseIn > 36 Or _
    BaseOut < 1 Or BaseOut > 36 Then
    ConvertBase = "Error"
    Exit Function
End If

' Ensure any letters in the input mumber are capitals
NumInCaps = UCase$(NumIn)
' Convert NumInCaps into Decimal
PlaceValue = Len(NumInCaps)

For i = 1 To Len(NumInCaps)
    PlaceValue = PlaceValue - 1
    CurrentCharacter = Mid$(NumInCaps, i, 1)
    CharacterValue = 0
    If Asc(CurrentCharacter) > 64 And _
    Asc(CurrentCharacter) < 91 Then _
    CharacterValue = Asc(CurrentCharacter) - 55


    If CharacterValue = 0 Then
        ' Ensure NumIn is correct
        If Asc(CurrentCharacter) < 48 Or _
        Asc(CurrentCharacter) > 57 Then
        ConvertBase = "Error"
        Exit Function
    Else
        CharacterValue = Val(CurrentCharacter)
    End If

End If

If CharacterValue < 0 Or CharacterValue > BaseIn - 1 Then
    ' Ensure NumIn is correct
    ConvertBase = "Error"
    Exit Function
End If

RunningTotal = RunningTotal + CharacterValue * _
(BaseIn ^ PlaceValue)
Next i

' Convert Decimal Number into the desired base using
    ' Repeated Division
Do
BaseOutDouble = CDbl(BaseOut)
Remainder = ModDouble(RunningTotal, BaseOutDouble)
RunningTotal = (RunningTotal - Remainder) / BaseOut

If Remainder >= 10 Then
    CurrentCharacter = Chr$(Remainder + 55)
Else
    CurrentCharacter = Right$(Str$(Remainder), _
    Len(Str$(Remainder)) - 1)
End If

ConvertBase = CurrentCharacter & ConvertBase
Loop While RunningTotal > 0

End Function

Public Function ModDouble(NumIn As Double, DivNum As Double) As Double
    ' Returns the Remainder when a number is divided by another
    ' (Works for double data-type)
    ModDouble = NumIn - (Int(NumIn / DivNum) * DivNum)
End Function


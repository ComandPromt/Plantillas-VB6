Attribute VB_Name = "Module1"

            
'**************************************
' Name: ConvertBase
' Description:Converts a string containi
'     ng a number in any base into another spe
'     cified base
' By: Aidan Crook
'
'
' Inputs:Input Number (as string), Base
'     of Input Number (as integer), Desired ou
'     tput base (as integer)
'
' Returns:A string containing the number
'     in the desired base
'
'Assumes:None
'
'Side Effects:None
'
'Warranty:
'Code provided by Planet Source Code(tm)
'     (http://www.Planet-Source-Code.com) 'as
'     is', without warranties as to performanc
'     e, fitness, merchantability,and any othe
'     r warranty (whether expressed or implied
'     ).
'Terms of Agreement:
'By using this source code, you agree to
'     the following terms...
' 1) You may use this source code in per
'     sonal projects and may compile it into a
'     n .exe/.dll/.ocx and distribute it in bi
'     nary format freely and with no charge.
' 2) You MAY NOT redistribute this sourc
'     e code (for example to a web site) witho
'     ut written permission from the original
'     author.Failure to do so is a violation o
'     f copyright laws.
' 3) You may link to this code from anot
'     her website, provided it is not wrapped
'     in a frame.
' 4) The author of this code may have re
'     tained certain additional copyright righ
'     ts.If so, this is indicated in the autho
'     r's description.
'**************************************

' Written by Aidan Crook
' Converts a number of any base into a d
'     ifferent base (e.g. Hexadecimal into Oct
'     al)
' This code should be placed in a module
'


Public Function ConvertBase(NumIn As String, BaseIn As Integer, BaseOut As Integer) As String
    ' NumIn is the number which you wish to
    '     convert (A string including characters 0
    '     - 9, A - Z)
    ' BaseIn is the base of NumIn (An intege
    '     r value in decimal between 1 & 36)
    ' BaseOut is the base of the number the


'     function returns (An integer value in de
    '     cimal between 1 & 36)
    ' Returns a string in the desired base c
    '     ontaining the characters 0 - 9, A - Z)
    ' e.g. ConvertBase ("42", 8, 16) convert
    '     s the octal number 42 into hexadecimal
    ' Returns the string "22"
    ' Returns the word "Error" if any of the
    '     input values are incorrect
    Dim i As Integer, CurrentCharacter As String, CharacterValue As Integer, PlaceValue As Integer, RunningTotal As Long, Remainder As Long
    ' Ensure input data is valid


    If NumIn = "" Or BaseIn < 1 Or BaseIn > 36 Or BaseOut < 1 Or BaseOut > 36 Then
        ConvertBase = "Error"
        Exit Function
    End If
    ' Convert NumIn into Decimal
    PlaceValue = Len(NumIn)


    For i = 1 To Len(NumIn)
        PlaceValue = PlaceValue - 1
        CurrentCharacter = Mid$(NumIn, i, 1)
        CharacterValue = 0
        If Asc(CurrentCharacter) > 64 And Asc(CurrentCharacter) < 91 Then CharacterValue = Asc(CurrentCharacter) - 55


        If CharacterValue = 0 Then


            If Asc(CurrentCharacter) < 48 Or Asc(CurrentCharacter) > 57 Then
                ' Ensure NumIn is correct
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
        RunningTotal = RunningTotal + CharacterValue * (BaseIn ^ PlaceValue)
    Next i
    ' Convert Decimal Number into the desire
    '     d base using Repeated Division


    Do
        Remainder = (RunningTotal Mod BaseOut)
        RunningTotal = (RunningTotal - Remainder) / BaseOut


        If Remainder >= 10 Then
            CurrentCharacter = Chr$(Remainder + 55)
        Else
            CurrentCharacter = Right$(Str$(Remainder), Len(Str$(Remainder)) - 1)
        End If
        ConvertBase = CurrentCharacter + ConvertBase
    Loop While RunningTotal > 0
End Function

 


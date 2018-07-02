Attribute VB_Name = "Module1"
'**************************************
' Name: Converting to/from Feet and Inch
'     es
' Description:These two functions conver
'     t numbers to and from standard architect
'     ural dimensions (Feet and Inches).
' By: Larry Serflaten
'
'
' Inputs:'
'The number is some value that represent
'     s inches.
'The string is in the form of FF' II" wh
'     ere
'FF is the number of feet and II is the
'     number of
'inches.
'
' Returns:'
'FeetAndInches returns a string
'TotalInches returns a single number
'
'Assumes:'
'The tick constant is used to denote the
'     smallest unit deired.
'Tick = 64' will allow measurement to 1/
'     64th of an inch:
'5 feet 3 inches and 7 - 64ths will read
'     as: 5' 3 7/64"
'Measurments are always rounded down to
'     the nearest tick value.
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



Function FeetAndInches(ByVal Num) As String
    Dim Recip As Single, Denom As Long
    Dim Out As String
    Const Tick = 64 '1/Tick = Smallest allowable increment
    Num = CSng(Num)


    If Num > 12 Then
        Denom = Num \ 12
        Num = Num - (Denom * 12)
        Out = CStr(Denom) & "'"
    End If


    If Num > 0 Then


        If Num >= 1 Then
            Out = Out & Space(1) & CStr(Int(Num))
            Num = Num - Int(Num)
        End If


        If Num > 0 Then
            Recip = Int(Num * Tick)
            Denom = Tick


            Do While (Recip And 1!) = 0
                Recip = Recip \ 2
                Denom = Denom \ 2
            Loop
            Out = Out & Space(1) & CStr(Recip) & "/" & CStr(Denom)
        End If
        Out = Out & """"
    End If
    FeetAndInches = Out
End Function


Function TotalInches(NewValue) As Single
    Dim Tmp&, Total As Single
    Dim Pos&, Measured$
    Measured = Trim$(CStr(NewValue))
    Pos = InStr(Measured, "'")


    If Pos Then
        Tmp = Val(Measured)
        Total = Tmp * 12
        Measured = Trim(Mid$(Measured, Pos + 2))
    End If
    Pos = InStr(Measured, " ")


    If Pos Then
        Tmp = Val(Left$(Measured, Pos))
        Total = Total + Tmp
        Measured = Mid$(Measured, Pos + 1)
    End If
    Tmp = Val(Measured)
    Pos = InStr(Measured, "/")


    If Pos Then
        Total = Total + CSng(Tmp / Val(Mid$(Measured, Pos + 1)))
    Else
        Total = Total + Tmp
    End If
    TotalInches = Total
End Function

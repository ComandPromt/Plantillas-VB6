Attribute VB_Name = "Converts"
'**************************************
' Name: Base 2-36 conversion
' Description:Converts a number from bas
'     e 2~36 to a number of base 2~36
' By: Joseph Wang
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:use all lower case if passed ba
'     se 10
'
'Side Effects:none
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

Option Explicit


Private Function dec2any(number As Long, convertb As Integer) As String
    On Error Resume Next
    Dim num As Long
    Dim sum As String
    Dim carry As Long
    sum = ""
    num = number


    If convertb > 1 And convertb < 37 Then


        Do
            carry = num Mod convertb


            If carry > 9 Then
                sum = Chr$(carry + 87) + sum
            Else
                sum = carry & sum
            End If
            num = Int(num / convertb)
        Loop Until num = 0
        dec2any = sum
    Else
        dec2any = -1
    End If
End Function


Private Function any2dec(num As String, Optional numbase As Integer = 10) As Long
    On Error Resume Next
    Dim sum As Long
    Dim length As Integer
    Dim count As Integer
    Dim digit As String * 1
    length = Len(num)


    If length > 0 And numbase > 0 And numbase < 37 Then


        For count = 1 To length
            digit = Mid$(num, count, 1)


            If digit <= "9" Then
                sum = sum + digit * numbase ^ (length - count)
            Else
                sum = sum + (Asc(digit) - 87) * numbase ^ (length - count)
            End If
        Next count
        any2dec = sum
    Else
        any2dec = -1
    End If
End Function


Private Function any2any(num1 As String, num1base As Integer, convertbase As Integer) As String
    Dim answer As Long
    If num1base <> convertbase And num1base > 0 And convertbase > 0 _
    And num1base < 37 And convertbase < 37 Then
    answer = any2dec(num1, num1base)
    any2any = dec2any(answer, convertbase)
Else
    any2any = -1
End If
End Function


Private Sub Form_Load()
    ' example: converts letter z of base 36
    '     to base 2 (binary)
    Me.Caption = any2any("z", 36, 2)
End Sub




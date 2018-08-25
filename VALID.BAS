Attribute VB_Name = "modValid"
Option Explicit

' This function checks the phone number passed in.
' The phone number must be in the format XXX-XXX-XXXX
Function CheckPhoneNumber(PhoneNumber)
    Dim Index ' used to scan phone number
    
    ' If the phone number is the wrong length,
    ' it's bad ... don't bother to scan
    If Len(PhoneNumber) <> 12 Then
        CheckPhoneNumber = False
        Exit Function
    Else
        For Index = 1 To 12
            Select Case Index
                Case 4, 8 ' Check for hypens
                    If Mid(PhoneNumber, Index, 1) <> "-" Then
                        CheckPhoneNumber = False
                        Exit Function
                        End If
                        
                Case Else ' Check for digits
                    If Not IsDigit(Mid(PhoneNumber, Index, 1)) Then
                        CheckPhoneNumber = False
                        Exit Function
                        End If
                End Select
            Next Index
        End If
        
    ' We made it through the whole phone number,
    ' must be good.
    CheckPhoneNumber = True
    End Function

Function CheckZipCode(ZipCode)
    Dim Index
    
    Select Case Len(ZipCode)
        Case 5, 9 ' normal five digit Zip Code or
                  ' nine digit Zip Code w/o hyphen
                  
            ' Check each characer in the Zip Code
            ' Each one must be a digit.
            For Index = 1 To Len(ZipCode)
                If Not IsDigit(Mid(ZipCode, Index, 1)) Then
                    CheckZipCode = False
                    Exit Function
                    End If
                Next Index
                
        Case 10 ' nine digit zip code w/ hyphen
            ' Check each characer in the Zip Code
            ' The sixth character must be a hyphen,
            ' the rest must be digits.
            For Index = 1 To Len(ZipCode)
                ' Check for the hyphen
                If (Index = 6) And (Mid(ZipCode, Index, 1) <> "-") Then
                    CheckZipCode = False
                    Exit Function
                ' Check for the digits.
                ElseIf (Index <> 6) And (Not IsDigit(Mid(ZipCode, Index, 1))) Then
                    CheckZipCode = False
                    Exit Function
                    End If
                Next Index
        
        Case Else ' error
            CheckZipCode = False
            Exit Function
        End Select
        
    ' We made it through every check.  This Zip
    ' Code must be good.
    CheckZipCode = True
    End Function


' This function returns True if the character passed in is
' a decimal digit.  False, otherwise.
Function IsDigit(Character)
    IsDigit = (InStr("0123456789", Character) <> 0)
    End Function



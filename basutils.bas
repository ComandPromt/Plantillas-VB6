Attribute VB_Name = "basUtils"
Option Explicit

Public Function TransitionText(intTransitionNumber As Integer) As String
    Select Case intTransitionNumber
        Case 1
            TransitionText = "Column"
        Case 2
            TransitionText = "Smash"
        Case 3
            TransitionText = "Rotate"
        Case 4
            TransitionText = "Unroll"
        Case 5
            TransitionText = "Tear"
        Case 6
            TransitionText = "Fade"
    End Select
End Function

Public Function TransitionNumber(strTransitionText As String) As Integer
    Select Case strTransitionText
        Case "Column"
            TransitionNumber = 1
        Case "Smash"
            TransitionNumber = 2
        Case "Rotate"
            TransitionNumber = 3
        Case "Unroll"
            TransitionNumber = 4
        Case "Tear"
            TransitionNumber = 5
        Case "Fade"
            TransitionNumber = 6
    End Select
End Function
Public Function StripNull(strTarget As String) As String
    If Len(strTarget) > 0 Then
        If Right(strTarget, 1) = Chr(0) Then
            strTarget = Left(strTarget, Len(strTarget) - 1)
        End If
    End If
    StripNull = strTarget
End Function



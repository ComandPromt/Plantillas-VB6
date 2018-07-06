Attribute VB_Name = "ExtendedRTFNotSoUsefulMod"
Option Explicit
'Copyright 2002 Roger Gilchrist
'these are attempts to solve the problem of identifying selection point format
' they sort of work so I left them in the Demo Program
'but the are slow and interfer with normal operations
'none of them is called anywhere in the Demo

Public Function IsCAPS(RTB As RichTextBox) As Boolean

  'Copyright 2002 Roger Gilchrist

  Dim t As String, oldlen As Long
  Static PrevTest As Long

    With RTB
        If PrevTest <> .SelStart Then
            PrevTest = .SelStart
            oldlen = .SelLength
                        If .SelLength = 0 Then
            .SelLength = 1
            End If


            t$ = .SelRTF
            IsCAPS = InStr(t$, "\caps")
            .SelLength = oldlen
            .SelStart = PrevTest
        End If
    End With 'RTB

End Function

Public Function IsExtendedRTF(RTB As RichTextBox, RTFCode$) As Boolean

  'Copyright 2002 Roger Gilchrist
  ' this is a generic version of the other routines in this module

  Dim t As String, oldlen As Long
  Static PrevTest As Long
  Static PrevCode As String
  Static hitting As Boolean

    If hitting Then
        Exit Function '>---> Bottom
    End If
    hitting = True
    With RTB
        If PrevTest <> .SelStart Or PrevCode$ <> RTFCode Then
            PrevTest = .SelStart
            PrevCode$ = RTFCode
            oldlen = .SelLength
                        If .SelLength = 0 Then
            .SelLength = 1
            End If


            t$ = .SelRTF
            IsExtendedRTF = InStr(t$, RTFCode$)
            .SelLength = oldlen
            .SelStart = PrevTest
        End If
    End With 'RTB
    hitting = False

End Function

Public Function IsHighLighted(RTB As RichTextBox) As Boolean

  'Copyright 2002 Roger Gilchrist

  Dim t As String, oldlen As Long
  Static PrevTest As Long

    With RTB
        If PrevTest <> .SelStart Then
            PrevTest = .SelStart
            oldlen = .SelLength
                        If .SelLength = 0 Then
            .SelLength = 1
            End If

            t$ = .SelRTF
            IsHighLighted = InStr(t$, "\highlight")
            .SelLength = oldlen
            .SelStart = PrevTest
        End If
    End With 'RTB

End Function

Public Function IsULWave(RTB As RichTextBox) As Boolean

  'Copyright 2002 Roger Gilchrist

  Dim Tpos As Long, oldlen As Long, t As String
  Static PrevStart As Long, PrevLen As Long, PrevTest As Long

    With RTB
        If PrevStart <> .SelStart And PrevLen <> .SelLength Then
            PrevTest = .SelStart
            PrevLen = .SelLength
            If .SelLength = 0 Then
            .SelLength = 1
            End If
            '    If .SelLength > 1 Then
            '    Stop
            '    End If
            t$ = .SelRTF
            IsULWave = InStr(t$, "\ulwave")
            .SelLength = PrevLen
            .SelStart = PrevTest
        End If
    End With 'RTB

End Function

':) Ulli's VB Code Formatter V2.13.6 (19/08/2002 11:28:29 AM) 6 + 98 = 104 Lines

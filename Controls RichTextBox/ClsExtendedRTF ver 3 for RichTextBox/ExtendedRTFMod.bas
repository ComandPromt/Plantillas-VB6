Attribute VB_Name = "ExtendedRTFSupportMod"
Option Explicit
Public MyRTB As New ClsExtendedRTF
Public RTBLooks As New ClsRTFFontPainter
Public RTBHigh As New ClsAPIHighlight
Public RTBZoom As New ClsAPIZoom
Public Manfst As New ClsManifestation


'This module exists simple to give support for the demo
'holds the Class initializations and a few support routines
'it is NOT part of the class

Public Function IsDebugMode(Optional bSetMode As Boolean = False) As Boolean

  'VB2MAX 'Tip of the Week: Check Whether VB Is in Debug Mode
  'Erik Perrohe (Seattle, WA)
  
  Static DebugMode As Boolean

    DebugMode = bSetMode
    If Not DebugMode Then
        Debug.Assert IsDebugMode(True)
    End If
    IsDebugMode = DebugMode

End Function

Public Sub PlaceControlOnToolBar(ctl As Control, Tb As Toolbar, Index As String, Ctl2But As Boolean)

  '*PURPOSE:'Place a control on a tool bar with
  '                   Ctl2But  = False control width set to toolbar button size
  '                              True toolbar button size set to control width,
  '                               if control is not ComboBox then set height to toolbar button
  '*CREATOR: Roger Gilchrist'*DATE:    AUG-1999

    With ctl
        .Move Tb.Buttons(Index).Left, Tb.Buttons(Index).Top
        Select Case Ctl2But
          Case True
            If .Width > 0 And Tb.Buttons(Index).Width > 0 Then
                .Width = Tb.Buttons(Index).Width
            End If
          Case False
            If .Width > 0 And Tb.Buttons(Index).Width Then
                Tb.Buttons(Index).Width = .Width
            End If
        End Select
       If Tb.Buttons(Index).Height > 0 And .Height > 0 And Not (TypeOf ctl Is ComboBox) Then
            .Height = Tb.Buttons(Index).Height
'ElseIf Tb.Buttons(Index).Height > 0 And .Height > 0 And TypeOf Ctl Is ComboBox Then
'            .Top = (Tb.Buttons(Index).Height - .Height) / 2
        End If
        
        .ZOrder
    End With 'CTL

End Sub

Public Sub DoSomething()

  'This routine exist to show that the Hot Key system in ClsExtendedRTF
  ' can call to other parts of your program. Use Ctrl+Alt+F4 to invoke it.

    MsgBox "Called from inside ClsExtender but not part of it." & vbNewLine & "This MsgBox is in ExtendedRTFSupportMod.bas", , "Ctrl+Alt+F4"

End Sub

':) Ulli's VB Code Formatter V2.13.6 (19/08/2002 11:28:30 AM) 12 + 11 = 23 Lines

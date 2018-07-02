Attribute VB_Name = "mdlGeneral"
Dim h_midiout As Long ' MIDIOUT Port Handle

Sub midi_ListOutdevs(C As Control)
Dim OutCaps As MIDIOUTCAPS
Dim Dev As Integer
For Dev = -1 To midiOutGetNumDevs() - 1
        If midiOutGetDevCaps(Dev, OutCaps, Len(OutCaps)) = 0 Then
        C.AddItem OutCaps.szPname
        C.ItemData(NewIndex) = Dev
        Else: MsgBox "Error Obtaining Midi Out Devices"
        End If
Next Dev
End Sub

Sub midi_outStatus(device As Integer, Status As Boolean)
'Status  : Open = True Close = False
If Status = True Then
Call midiOutClose(h_midiout)  ' To be safe
If midiOutOpen(h_midiout, device, 0, 0, CallBack_Null) <> 0 Then _
MsgBox "Error opening MIDI OUT Port"
Else
If midiOutClose(h_midiout) <> 0 Then _
MsgBox "Error closing MIDI OUT Port"
End If
End Sub

Sub midioutmsg(MEvent As Byte, Channel As Byte, Value1 As Byte, Value2 As Byte)
Dim midimsg As Long
Select Case MEvent
Case &H80
midimsg = &H80 + (Value1 * &H100) + (Value2 * &H10000) + Channel
Case &H90
midimsg = &H90 + (Value1 * &H100) + (Value2 * &H10000) + Channel
Case &HA0
midimsg = &HA0 + (Value1 * &H100) + (Value2 * &H10000) + Channel
Case &HB0
midimsg = &HB0 + (Value1 * &H100) + (Value2 * &H10000) + Channel
Case &HC0
midimsg = &HC0 + (Value1 * &H100) + Channel
Case &HD0
midimsg = &HD0 + (Value1 * &H100) + Channel
Case &HE0
midimsg = &HE0 + (Value1 * &H100) + (Value2 * &H10000) + Channel
End Select
If midiOutShortMsg(h_midiout, midimsg) <> 0 Then MsgBox "Error sending note"
End Sub

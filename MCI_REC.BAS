
Declare Function mciSendString Lib "mmsystem" (ByVal MCI_COMMAND As String, ByVal ReturnString As String, ByVal ReturnLength As Integer, ByVal Handle As Integer) As Long
Declare Function mciGetErrorString Lib "mmsystem" (ByVal MCI_ERROR As Long, ByVal ErororString As String, ByVal ReturnLenght As Integer) As Integer

Sub CloseSound ()
'close the sound file
    Dim Result&
    Dim errormsg%
    Dim ReturnString As String * 1024
    Dim ErrorString As String * 1024

    Result& = mciSendString("close mysound", ReturnString, 1024, 0)

End Sub

Sub PlayRecSound ()
'plays the recoreded sound aliased by mysound
    Dim Result&
    Dim errormsg%
    Dim ReturnString As String * 1024
    Dim ErrorString As String * 1024

    Result& = mciSendString("stop mysound", ReturnString, 1024, 0)
    If Not Result& = 0 Then
	errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
	MsgBox ErrorString, 0, "Error"
    End If
    
    Result& = mciSendString("play mysound from 1 wait", ReturnString, 1024, 0)
    If Not Result& = 0 Then
	errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
	MsgBox ErrorString, 0, "Error"
    End If
    
End Sub

Sub RecordSound ()
'records sound aliased as mysound to memory for six seconds
    Dim Result&
    Dim errormsg%
    Dim ReturnString As String * 1024
    Dim ErrorString As String * 1024

    CloseSound

    Result& = mciSendString("open new type waveaudio alias mysound", ReturnString, 1024, 0)
    If Not Result& = 0 Then
	errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
	MsgBox ErrorString, 0, "Error"
	Exit Sub
    End If

'The following command sets the time format to milliseconds
'and sets the waveform audio format to 8 bit, mono, 11 kHz:
    Result& = mciSendString("set mysound time format ms bitspersample 8 samplespersec 11025", ReturnString, 1024, 0)
    If Not Result& = 0 Then
	errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
	MsgBox ErrorString, 0, "Error"
	Exit Sub
    End If

'Record for 60000 milliseconds
    Result& = mciSendString("record mysound to 60000", ReturnString, 1024, 0)
    If Not Result& = 0 Then
	errormsg% = mciGetErrorString(Result&, ErrorString, 1024)
	MsgBox ErrorString, 0, "Error"
	Exit Sub
    End If
End Sub


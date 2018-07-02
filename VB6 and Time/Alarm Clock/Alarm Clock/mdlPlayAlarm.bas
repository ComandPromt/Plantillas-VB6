Attribute VB_Name = "mdlPlayAlarm"
Public filename1 As String ' Public global string representing the filename of the media file to be played
Public isPlaying As Boolean ' Public global boolean representing whether the file is playing or not

Public Function PlayMusic()

' Input: None
' Process: Calls the API functions that play the sounds
' Output: None

    Dim Openresult As String ' String representing the result of the OpenMPEG function
    Dim Playresult As String ' String representing the result of the PlayMPEG function

    If filename1 = "" Then ' The filename variable is empty
        MsgBox "Error!  File not found!", 16, "Error!"
        Exit Function ' Booting out of the function
    End If
    
    isPlaying = True ' Flagging whether the file is playing or not
    Openresult = OpenMPEG(frmMain.hwnd, filename1, "MPEGVideo") ' Calling the function to open the file
    Playresult = PlayMPEG(vbNullString, vbNullString) ' Calling the function to play the file
    
End Function

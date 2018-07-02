Attribute VB_Name = "mdlCallWinDir"
Public Function CallWinDir()

' Input: None
' Process: Calls the API function that returns the Windows directory
' Output: None

    Dim strWinDir As String ' String representing the WIndows directory
    Dim lngSize As Long ' Some long integer.  I don't know what.
    
    strWinDir = Space(255) ' Allocating 255 bytes for the string
    lngSize = 255
    
    Call GetWindowsDirectory(strWinDir, lngSize - 1) ' Calling the API function
    strWinDir = TrimNull(strWinDir) ' Trimming the null terminator
    frmSetAlarm.dlgFiles.InitDir = strWinDir ' Assigning the InitDir property of the Common Dialog Box
End Function
    

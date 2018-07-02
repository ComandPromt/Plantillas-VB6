Attribute VB_Name = "GameFunctions"
'*********************************
'* Game functions module
'*********************************
'* Description: Provides arbitrary functions for games
'*
'* Date: 08/09-1998 (European date system)
'*
'* Author: Søren Christensen - Rankan Software www.rankan.com
'*         commments etc. soren@rankan.com
'*
'*
'*
'* Types:
'*              Tile
'*              RECT
'*
'* Functions:
'*              ReadTileFile
'*              PlayBackGroundSound
'*              StopBackGroundSound
'*              ServiceBackgroundMusic
'*
'*              CreateMask
'*              ReleaseMask
'*              DetectCollision
'*
'*              GenerateDC
'*              CheckUpKey
'*              CheckDownKey
'*              CheckLeftKey
'*              CheckRightKey
'*
'*
'*********************************


Option Explicit
'API declarations
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, _
                                           ByVal nHeight As Long, _
                                           ByVal nPlanes As Long, _
                                           ByVal nBitCount As Long, _
                                           lpBits As Any) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                           ByVal hObject As Long) As Long

Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
                                         ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                            ByVal X As Long, _
                                            ByVal Y As Long, _
                                            ByVal nWidth As Long, _
                                            ByVal nHeight As Long, _
                                            ByVal hSrcDC As Long, _
                                            ByVal xSrc As Long, _
                                            ByVal ySrc As Long, _
                                            ByVal dwRop As Long) As Long

Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
                                         ByVal X As Long, _
                                         ByVal Y As Long, _
                                         ByVal nWidth As Long, _
                                         ByVal nHeight As Long, _
                                         ByVal hSrcDC As Long, _
                                         ByVal xSrc As Long, _
                                         ByVal ySrc As Long, _
                                         ByVal nSrcWidth As Long, _
                                         ByVal nSrcHeight As Long, _
                                         ByVal dwRop As Long) As Long

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
                                            (ByVal lpstrCommand As String, _
                                             ByVal lpstrReturnString As String, _
                                             ByVal uReturnLength As Long, _
                                             ByVal hwndCallback As Long) As Long

Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" _
                                            (ByVal dwError As Long, _
                                             ByVal lpstrBuffer As String, _
                                             ByVal uLength As Long) As Long

Declare Function LoadImage Lib "user32" Alias "LoadImageA" ( _
                                             ByVal hInst As Long, _
                                             ByVal lpsz As String, _
                                             ByVal un1 As Long, _
                                             ByVal n1 As Long, _
                                             ByVal n2 As Long, _
                                             ByVal un2 As Long) As Long

Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetTickCount Lib "kernel32" () As Long


'Constants

'**MCI Constants**
Public Const WAV As Long = 1
Public Const MIDI As Long = 2
'**LoadImage Constants**
Public Const IMAGE_BITMAP As Long = 0
Public Const LR_LOADFROMFILE As Long = 10
Public Const LR_CREATEDIBSECTION As Long = 2000
'**GetKeyState Constants**
Public Const VK_RIGHT As Long = &H27
Public Const VK_LEFT As Long = &H25
Public Const VK_DOWN As Long = &H28
Public Const VK_UP As Long = &H26
Public Const VK_ESCAPE = &H1B
Public Const VK_KEYDOWN As Long = -127
Public Const VK2_KEYDOWN As Long = -128
'***********



'Types

'Tile type
'One of the special sprite data could be removed, if only few effects are used
Type Tile
    SpecialData As Single 'Special Sprite data
    SpecialData1 As Single 'Another special sprite data
    SourceX As Integer 'The X position on the source
    SourceY As Integer 'The Y position on the source
    TileWidth As Integer
    TileHeight As Integer
End Type

'RECT type
Type RECT
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
End Type




'Function purpose: Translate a text file into a 2-dimensional array
'Note: This funciton is only usable for ordinary text files, composed of
'       numbers from 0 to 9
'IN: FileName: The file name of the string
'    TileArray: The String array to put the translated text file into
'OUT: True on completed success
'     False on failure
Public Function ReadTileFile(FileName As String, TileArray() As Long) As Boolean
On Error GoTo Err_Handler

Dim FreeFileNr As Integer
Dim TempString As String
Dim I As Integer, J As Integer, P As Integer

FreeFileNr = FreeFile

'We need the number of lines the file has, so open it and count them
Open FileName For Input As #FreeFileNr
    
    Do Until EOF(FreeFileNr)    'Count the number of lines
        Line Input #FreeFileNr, TempString
        J = J + 1
    Loop
    
Close #FreeFileNr

'Now we do the actual parsing of the file
'Looking for, and storing the numbers
Open FileName For Input As #FreeFileNr
    
    Do Until EOF(FreeFileNr)
        
        Line Input #FreeFileNr, TempString
        
        If TempString <> "" Then
            ReDim Preserve TileArray(J - 1, Len(TempString) - 1)
            For I = 0 To Len(TempString) - 1 'The number counter in each line
                TileArray(P, I) = CLng(Mid$(TempString, I + 1, 1))
            Next I
        End If
        
        P = P + 1 'This is the line counter - first index in the 2d-array
        
    Loop

Close #FreeFileNr

Err_Handler:
    
    Select Case Err
        
        Case 0
            Err.Clear
            ReadTileFile = True
            
        Case Else
                        
            ReadTileFile = False
            
    End Select

End Function


Public Function CreateMask(GraphicDC As Long, Width As Long, Height As Long) As Long
Dim hMemDC As Long
Dim hBitmap As Long
Dim TempBkColor As Long
Dim rt As Long


'Create the Memory device context
hMemDC = CreateCompatibleDC(GraphicDC)

If hMemDC < 1 Then  'Error
    CreateMask = 0
    Exit Function
End If


'Create a new monochrome bitmap with the size of the passed arguments
hBitmap = CreateBitmap(Width, Height, 1, 1, 0)

If hBitmap < 1 Then 'Error
    'Clean Up
    DeleteDC hMemDC
    CreateMask = 0
    Exit Function
End If

'Select the bitmap into the DC
SelectObject hMemDC, hBitmap


'Set the background color of the source dc to black, the current color will be stored
'in the TempBkColor variable
TempBkColor = SetBkColor(GraphicDC, 0)

'Blit the source Dc into the memory dc
rt = BitBlt(hMemDC, 0, 0, Width, Height, GraphicDC, 0, 0, vbSrcCopy)

'Restore background color
SetBkColor GraphicDC, TempBkColor

If rt < 1 Then  'Blit operation failed release sources
    DeleteDC hMemDC
    DeleteObject hBitmap
    CreateMask = 0
    Exit Function
Else
    DeleteObject hBitmap
    CreateMask = hMemDC
End If

End Function

'Releases a Mask created with the CreateMask function
'IN: The DC to release
'OUT: Over 0 when succesfull
Public Function ReleaseMask(MaskDC As Long) As Long

If MaskDC > 0 Then
    ReleaseMask = DeleteDC(MaskDC)
End If

End Function

'Initializes background sound and starts playing it
'IN: File name of the sound (MIDI or WAV)
'    The format of the sound (MIDI or WAV), defined in global constants
'OUT: The identifier of the sound
Public Function PlayBackGroundSound(FileName As String, Device As Long) As String
Dim rt As Long
Dim ErrorString As String
Static DeviceIdentifier As Integer

'Up the count (making it unique)
DeviceIdentifier = DeviceIdentifier + 1

ErrorString = Space(255)

Select Case Device

    Case WAV 'Wav files
        
        rt = mciSendString("OPEN " & FileName & " TYPE WAVEAUDIO ALIAS WAVNR" _
                            & CStr(DeviceIdentifier) & " BUFFER 4", "", 0, 0)
        
        If rt <> 0 Then 'Zero on success
                                    
            PlayBackGroundSound = Str(0)
                                    
        Else
            
            rt = mciSendString("PLAY WAVNR" & CStr(DeviceIdentifier) & " FROM 0", "", 0, 0)
            
            'Used for debugging..Fills the ErrorString variable witha description on the error
            'mciGetErrorString rt, ErrorString, Len(ErrorString)
            'Debug.Print ErrorString
            
            PlayBackGroundSound = "WAVNR" & DeviceIdentifier
                       
        End If
    
    Case MIDI
    
        rt = mciSendString("OPEN " & FileName & " TYPE SEQUENCER ALIAS MIDINR" _
                            & CStr(DeviceIdentifier), "", 0, 0)
        
        If rt <> 0 Then 'Zero on success
                        
            PlayBackGroundSound = Str(0)
            
        Else
            
            rt = mciSendString("PLAY MIDINR" & CStr(DeviceIdentifier) & " FROM 0", "", 0, 0)
            
            'Used for debugging..Fills the ErrorString variable witha description on the error
            'mciGetErrorString rt, ErrorString, Len(ErrorString)
            'Debug.Print ErrorString
            
            PlayBackGroundSound = "MIDINR" & DeviceIdentifier
                       
        End If
    
End Select


End Function
'Stops background music from playing
'IN: Identifier: Device identifier
'OUT: 1: No errors music stopped
'     0: Error, music not stopped (might not even be playing)
Public Function StopBackgroundMusic(Identifier As String) As Long
Dim rt As Long
Dim ErrorString As String

ErrorString = Space(255)

If Identifier <> "" Then
        
        rt = mciSendString("STOP " & Identifier, "", 0, 0)
        
        If rt <> 0 Then 'Error, Zero on success
            
            'Used for debugging..Fills the ErrorString variable witha description on the error
            'mciGetErrorString rt, ErrorString, Len(ErrorString)
            'Debug.Print ErrorString
            
            StopBackgroundMusic = 0
        
        Else
        
            StopBackgroundMusic = 1
        
        End If
Else
    
    StopBackgroundMusic = 0

End If

End Function

'Use this function to service background music, ie Start playing it if has stopped
'IN: Device identifier of the multimedia device to be serviced
'OUT: 1: No errors, but service not needed
'     2: No erros, service was performed
'     0: Error
Public Function ServiceBackgroundMusic(Identifier As String) As Long
Dim rt As Long
Dim Status As String

Status = Space(25)

rt = mciSendString("STATUS " & Identifier & " MODE", Status, Len(Status), 0)

If rt = 0 Then
    Status = Trim$(Status)
    
    If Left(UCase$(Status), Len("STOPPED")) = "STOPPED" Then 'Music has stopped play it again
        
        rt = mciSendString("PLAY " & Identifier & " FROM 0", "", 0, 0)
        
        If rt = 0 Then
            
            ServiceBackgroundMusic = 2
        
        Else
            
            ServiceBackgroundMusic = 0
        
        End If
    Else
        
        ServiceBackgroundMusic = 1
    End If

Else

    'Used for debugging..Fills the ErrorString variable witha description on the error
    'mciGetErrorString rt, ErrorString, Len(ErrorString)
    'Debug.Print ErrorString
    
    ServiceBackgroundMusic = 0

End If
           
    
End Function

'Reads a bitmap file and generates a Memory context for it
'IN: CompatibleDC: The context, which the generated DC should be compatible with
'    FileName: The file name of the graphics
'OUT: The Generated DC
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context
DC = CreateCompatibleDC(0)

If DC < 1 Then
    GenerateDC = 0
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

DeleteObject hBitmap

End Function

'Destroy a Device Context created with the GenerateDC function
Public Function DestroyDC(DC As Long) As Long

If DC > 0 Then
    DestroyDC = DeleteDC(DC)
End If

End Function

'Checks whether the LEFT Arrow key is pressed
'OUT: True if the key is pressed, else false
Public Function CheckLeftKey() As Boolean
Dim vkLeft As Long

vkLeft = GetKeyState(VK_LEFT)

If vkLeft = VK_KEYDOWN Or vkLeft = VK2_KEYDOWN Then
    CheckLeftKey = True
Else
    CheckLeftKey = False
End If

End Function

'Checks whether the RIGHT Arrow key is pressed
'OUT: True if the key is pressed, else false
Public Function CheckRightKey() As Boolean
Dim vkRight As Long

vkRight = GetKeyState(VK_RIGHT)

If vkRight = VK_KEYDOWN Or vkRight = VK2_KEYDOWN Then
    CheckRightKey = True
Else
    CheckRightKey = False
End If

End Function

'Checks whether the Down Arrow key is pressed
'OUT: True if the key is pressed, else false
Public Function CheckDownKey() As Boolean
Dim vkDown As Long

vkDown = GetKeyState(VK_DOWN)

If vkDown = VK_KEYDOWN Or vkDown = VK2_KEYDOWN Then
    CheckDownKey = True
Else
    CheckDownKey = False
End If

End Function

'Checks whether the UP Arrow key is pressed
'OUT: True if the key is pressed, else false
Public Function CheckUPKey() As Boolean
Dim vkUP As Long

vkUP = GetKeyState(VK_UP)

If vkUP = VK_KEYDOWN Or vkUP = VK2_KEYDOWN Then
    CheckUPKey = True
Else
    CheckUPKey = False
End If

End Function

'Checks if Escape key is pressed
Public Function IsEscapePressed() As Boolean
Dim vkEscape As Long

vkEscape = GetKeyState(VK_ESCAPE)

If vkEscape = VK_KEYDOWN Or vkEscape = VK2_KEYDOWN Then
    IsEscapePressed = True
Else
    IsEscapePressed = False
End If

End Function

'Collision detection.
'IN: Two RECTS Structures
Public Function DetectCollision(FirstRect As RECT, SecondRect As RECT) As Boolean
Dim Collision As Boolean

Collision = True

If ((FirstRect.X2 < SecondRect.X1) Or (FirstRect.X1 > SecondRect.X2) Or _
    (FirstRect.Y1 < SecondRect.Y1) Or (FirstRect.Y1 > SecondRect.Y2)) Then
        
        Collision = False
End If

DetectCollision = Collision

End Function


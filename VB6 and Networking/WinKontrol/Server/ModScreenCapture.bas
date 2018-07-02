Attribute VB_Name = "ModScreenCapture"
Option Explicit

Option Base 0

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    'Enough for 256 colors
    palPalEntry(255) As PALETTEENTRY
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type


Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PicBmp
   Size As Long
   bitMapType As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function GetWindowDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "USER32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
        
        Public Const MAX_CHUNK = 4096

Public Function CaptureForm(frmSrc As Form) As Picture
    On Error GoTo ErrorRoutineErr

    'Call CaptureWindow to capture the entire form
    'given it's window
    'handle and then return the resulting Picture object
    Set CaptureForm = CaptureWindow(frmSrc.hwnd, 0, 0, _
            frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), _
            frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))

ErrorRoutineResume:
    Exit Function
ErrorRoutineErr:
    MsgBox "Project1.Module1.CaptureForm" & Err & Error
    Resume Next
End Function

Public Function CreateBitmapPicture(ByVal hBmp As Long, _
    ByVal hPal As Long) As Picture

    On Error GoTo ErrorRoutineErr

    Dim r As Long
    Dim Pic As PicBmp
    'IPicture requires a reference to "Standard OLE Types"
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID

    'Fill in with IDispatch Interface ID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill Pic with necessary parts
    With Pic
    'Length of structure
        .Size = Len(Pic)
    'Type of Picture (bitmap)
        .bitMapType = vbPicTypeBitmap
    'Handle to bitmap
        .hBmp = hBmp
    'Handle to palette (may be null)
        .hPal = hPal
    End With

    'Create Picture object
    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

    'Return the new Picture object
    Set CreateBitmapPicture = IPic

ErrorRoutineResume:
    Exit Function
ErrorRoutineErr:
    MsgBox "Project1.Module1.CreateBitmapPicture" & Err & Error
    Resume Next
End Function

Public Function CaptureWindow(ByVal hWndSrc As Long, _
    ByVal LeftSrc As Long, _
    ByVal TopSrc As Long, ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long) As Picture

    On Error GoTo ErrorRoutineErr

    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim rc As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long

    Dim LogPal As LOGPALETTE

    'get device context for the window
    hDCSrc = GetWindowDC(hWndSrc)

    'Create a memory device context for the copy process
    hDCMemory = CreateCompatibleDC(hDCSrc)
    'Create a bitmap and place it in the memory DC
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    'get screen properties
    'Raster capabilities
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
    'Palette support
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    'Size of palette
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)

    'If the screen has a palette, make a copy
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    'Create a copy of the system palette
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        rc = GetSystemPaletteEntries(hDCSrc, 0, 256, _
                LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
    'Select the new palette into the memory
    'DC and realize it
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        rc = RealizePalette(hDCMemory)
    End If

    'Copy the image into the memory DC
    rc = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, _
            hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

    'Remove the new copy of the  on-screen image
    'hBmp = SelectObject(hDCMemory, hBmpPrev)

    'If the screen has a palette get back the palette that was
    'selected in previously
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    'Release the device context resources back to the system
    rc = DeleteDC(hDCMemory)
    rc = ReleaseDC(hWndSrc, hDCSrc)

    'Call CreateBitmapPicture to create a picture
    'object from the bitmap and palette handles.
    'then return the resulting picture object.
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)

ErrorRoutineResume:
    Exit Function
ErrorRoutineErr:
    MsgBox "CaptureWindow" & Err & Error
    Resume Next
End Function

Sub SendScreenShot(Fname As String)
    Dim DataChunk As String
    Dim passes As Long

    
    SendData "SCREENSHOT," 'Tell the CLIENT that the SCREENSHOT is about to be sent...
    Pause 200 'Pause. Let the Client get ready..
    
    Open Fname$ For Binary As #1
        Do While Not EOF(1)
            'passes& = passes& + 1 'Update pass Variable
            DataChunk$ = Input(MAX_CHUNK, #1) 'Get some of the data chunk
            SendData DataChunk$ 'Send Chunk to client
            Pause 200 'Pause again...
            DoEvents
        Loop 'Loop until Screenshot is completly sent
        
        SendData "ENDSCREENSHOT,"
        'passes& = 0
    Close #1
End Sub


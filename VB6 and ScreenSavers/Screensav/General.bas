Attribute VB_Name = "modGeneral"
Option Explicit
'
'----------------------------------------------------------------------
' Public Variables.
'----------------------------------------------------------------------
'

Public glDisplayHwnd As Long          ' Handle of Preview window.
Public glRunMode     As Long          ' Screen saver running mode (run, preview, setup)
Public glDeskDC      As Long          ' Desktop device context handle.
Public glSpriteCount As Long          ' Active sprites.
Public glRefreshRate As Long          ' Sprite animation frame movement rate.
Public glSpriteSize  As Long          ' Relative sprite size option.
Public glSpriteSpeed As Long          ' Active sprite velocity.
Public glBmpYUnits   As Long          ' # sprite frames on the y axis
Public glBitMap_ID   As Long          ' Res File bitmap image ID
Public gsSpriteImage As String        ' Image to Display.
Public gsPassword    As String        ' Screen Saver password.
Public gbClearScreen As Boolean       ' Clear Screen first.
Public gbUseTracers  As Boolean       ' Tracers option (sprite doesn't clean up trails).
Public gbRefreshRND  As Boolean       ' Random refresh rate option.
Public gbSizeRND     As Boolean       ' Randomize sprite size.
Public gbSpeedRND    As Boolean       ' Randomize sprite speed.
Public gbUsePassword As Boolean       ' Password enabled flag.
Public gaSprite()    As Sprite        ' Array of active sprites.
Public gSprite       As ResBitmap     ' Bitmap resource loading bucket.
Public gDispRec      As RECT          ' Rectangle values of Preview window.
Public gDeskBmp      As BITMAP        ' Bitmap copy of the desktop.
'
'----------------------------------------------------------------------
' Application Specific Constants.
'----------------------------------------------------------------------
'
Public Const cIMAGE0 = "Jordan"       ' Image to display.
Public Const cIMAGE1 = "TheScarms"    ' Image to display.
Public Const cIMAGE2 = "Vertex"       ' Image to display.
Public Const cBMPXUNITS = 1           ' # sprite frames on the x axis
Public Const cDEF_SPRITECOUNT = 8     ' Default sprite counts.
Public Const cMIN_SPRITECOUNT = 1     ' Minimum number of sprites.
Public Const cMAX_SPRITECOUNT = 30    ' Maximum number of sprites.
Public Const cMIN_REFRESHRATE = 1     ' 1 / 1000 second.
Public Const cMAX_REFRESHRATE = 100   ' 1 / 10   second.
Public Const cMIN_SPRITESIZE = 25     ' 25%  of normal size.
Public Const cMAX_SPRITESIZE = 150    ' 150% of normal size.
Public Const cMIN_SPRITESPEED = 1     ' Move in 1  pixel increments.
Public Const cMAX_SPRITESPEED = 50    ' Move in 50 pixel increments.
Public Const cBASE_MASS = 100         ' Relative base mass for sprite size.
Public Const cPREVIEW_WINDOW = "Display Properties"
Public Const cREGKEY = "Software\TheScarms\ScreenSaver"
'
' ScreenSaver Running Modes.
'
Public Const RM_NORMAL = 1
Public Const RM_CONFIGURE = 2
Public Const RM_PREVIEW = 4
'
'----------------------------------------------------------------------
'Public API Declares.
'----------------------------------------------------------------------
'
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal fShow As Integer) As Integer
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As Any) As Long
Public Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As String) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'
'----------------------------------------------------------------------
'Public Constants.
'----------------------------------------------------------------------
'
Public Const WS_CHILD = &H40000000
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)

Public Const SPI_SCREENSAVERRUNNING = 97

Public Const HWND_TOPMOST = -1&
Public Const HWND_TOP = 0&
Public Const HWND_BOTTOM = 1&
Public Const HWND_NOTOPMOST = -2

Public Const SWP_NOSIZE = &H1&
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
'
' Windows messages.
'
Public Const WM_PAINT = &HF&
Public Const WM_ACTIVATEAPP = &H1C&
Public Const SW_SHOWNOACTIVATE = 4&
'
' Get Windows Long Constants.
'
Public Const GWL_USERDATA = (-21&)
Public Const GWL_WNDPROC = (-4&)
'
'----------------------------------------------------------------------
'Public Type Defs.
'----------------------------------------------------------------------
'
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Public Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength              As Long
    lpSecurityDescriptor As Long
    bInheritHandle       As Boolean
End Type

Public Type ResBitmap
    ResID  As Long
    Sprite As StdPicture
End Type

Sub Main()
Dim l        As Long
Dim rc       As Long
Dim lStyle   As Long
Dim lLen     As Long
Dim lTemp    As Long
Dim sStr     As String
Dim sCommand As String
Dim sOption  As String
'
' Get the command line parameters.
'
sCommand = LCase$(Trim$(Command()))
sOption = Left$(sCommand, 2)
'
' Only allow a single instance of the screen saver
' under normal operation.  When the PC is idle for
' the specified period of time, Windows will launch
' your screen saver continually with a "/s" parameter.
' When you click the Preview button on the Display
' Properties dialog the screen saver is also started
' with the "/s" switch. In this case you want a second
' instance to run since the first instance will be
' running in the small preview window.  To distinguish
' between the two scenarios, use FindWindow to see if
' the "Display Properties" dialog is open.
'
lTemp = FindWindow(vbNullString, cPREVIEW_WINDOW)
If App.PrevInstance And sOption = "/s" And lTemp = 0 Then End
'
' Process the command line parameters.
'
Select Case sOption
    Case "", "/s" '/s
        '
        ' Start the Screen Saver.
        '
        ' Store screen saver's run mode.
        ' Get the Desktop window's dimensions.
        ' Load the main form.
        '
        glRunMode = RM_NORMAL
        Call GetWindowRect(GetDesktopWindow(), gDispRec)
        Load frmMain
        '
        ' Maximize the main form and make it the top-most window.
        '
        Call SetWindowPos(frmMain.hwnd, _
             HWND_TOPMOST, 0&, 0&, gDispRec.Right, gDispRec.Bottom, _
             SWP_SHOWWINDOW)
        '
        ' Prevent the user from using ALT+TAB to switch
        ' to another application or CTRL+ALT+DELETE to
        ' kill the Screen Saver.
        '
        Call SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, lTemp, 0)
    Case "/p" '/p <hwnd>
        '
        ' Preview Mode.  Run inside of the Screen
        ' Saver Configuration Viewer.
        '
        ' When the screen saver is called in Preview mode
        ' it is passed "/p <hwnd>" where <hwnd> is the handle
        ' of the Preview window.
        '
        ' Get the handle and client area dimensions
        ' of the Preview DeskTop window.
        '
        glRunMode = RM_PREVIEW
        sStr = sCommand
        lLen = Len(sStr)

        For l = lLen To 1 Step -1
            sStr = Right$(sStr, l)
            If IsNumeric(sStr) Then
                glDisplayHwnd = Val(sStr)
                Exit For
            End If
        Next

        Call GetClientRect(glDisplayHwnd, gDispRec)
        '
        ' Load the Screen Saver form.
        '
        Load frmMain
        '
        ' Set its caption consistant with Windows screen savers.
        ' Get the form's current window style.
        ' Convert it to a child window.
        ' Set its parent window to be the Preview window.
        ' Save the Preview window's handle in the form's window structure.
        '
        With frmMain
            .Caption = "Preview"
            lStyle = GetWindowLong(.hwnd, GWL_STYLE)
            lStyle = lStyle Or WS_CHILD
            Call SetWindowLong(.hwnd, GWL_STYLE, lStyle)
            Call SetParent(.hwnd, glDisplayHwnd)
            Call SetWindowLong(.hwnd, GWL_HWNDPARENT, glDisplayHwnd)
        End With
        '
        ' Show the screen saver in the Preview window.
        '
        Call SetWindowPos(frmMain.hwnd, HWND_TOP, 0&, 0&, _
            gDispRec.Right, gDispRec.Bottom, SWP_NOZORDER Or _
            SWP_NOACTIVATE Or SWP_SHOWWINDOW)
    Case "/c"  '/c:<hwnd>
        '
        ' Display the screen saver configuration dialog.
        '
        frmSetup.Show vbModal
    Case "/a"  '/a <hwnd>
        '
        ' Display the password change dialog.
        '
        frmChgPswd.Show vbModal
    Case Else
End Select
End Sub
Public Function fShrinkBmp(dispHdc As Long, hBmp As Long, RatioX As Single, RatioY As Single) As Long
'
' Scale a bitmap by an X and Y percentage and
' return a handle to the new bitmap.
'
Dim hBmpOut As Long   ' Output bitmap handle.
Dim hdcMem1 As Long   ' Temporary memory bitmap handles.
Dim hdcMem2 As Long
Dim bm1     As BITMAP ' Temporary bitmap structures.
Dim bm2     As BITMAP
'
' Create memory DCs compatible to the display DC.
'
hdcMem1 = CreateCompatibleDC(dispHdc)
hdcMem2 = CreateCompatibleDC(dispHdc)
'
' Get the bitmap information and save it in bm1.
'
Call GetObject(hBmp, LenB(bm1), bm1)
'
' Copy bitmap 1 to bitmap 2.
'
LSet bm2 = bm1
'
' Scale output bitmap width and height.
' Calculate bitmap width bytes.
'
With bm2
    .bmWidth = CLng(.bmWidth * RatioX)
    .bmHeight = CLng(.bmHeight * RatioY)
    .bmWidthBytes = ((((.bmWidth * .bmBitsPixel) + 15) \ 16) * 2)
End With
'
' Create a handle to output bitmap indirectly from new bm2.
'
hBmpOut = CreateBitmapIndirect(bm2)
'
' Select original bitmap into the memory dc.
' Select new bitmap into the memory dc.
'
Call SelectObject(hdcMem1, hBmp)
Call SelectObject(hdcMem2, hBmpOut)
'
' Stretch old bitmap into new bitmap.
'
Call StretchBlt(hdcMem2, 0, 0, bm2.bmWidth, bm2.bmHeight, _
        hdcMem1, 0, 0, bm1.bmWidth, bm1.bmHeight, vbSrcCopy)
'
' Delete memory DCs
'
Call DeleteDC(hdcMem1)
Call DeleteDC(hdcMem2)
'
' Return handle to new bitmap
'
fShrinkBmp = hBmpOut
End Function

Public Sub pInitDeskDC(OutHdc As Long, OutBmp As BITMAP, gDispRec As RECT)
'
' Create and return a bitmap that looks like the current desktop
' but that is stretched or compressed to a pre-defined width and
' height as specified by gDispRec.
'
Dim DskHwnd As Long  ' Handle of desktop window.
Dim DskHdc  As Long  ' DC handle of desktop window.
Dim hOutBmp As Long  ' Handle to output bitmap.
Dim rc      As Long  ' Function return code.
Dim DskRect As RECT  ' Rect size of desktop.
'
' Get the handle of the desktop window.
'
DskHwnd = GetDesktopWindow()
'
' Get the device context (DC) for the entire window,
' including title bar, menus, and scroll bars. A window DC
' permits painting anywhere in a window, because the origin
' of the device context is the upper-left corner of the
' window instead of the client area.
'
DskHdc = GetWindowDC(DskHwnd)
'
' Get the dimensions of the desktop window.
'
rc = GetWindowRect(DskHwnd, DskRect)

With gDispRec
    '
    ' Create a bitmap compatible with the desktop
    ' window and return its handle. The dimensions
    ' are 1 pixel wider and taller than the desktop
    ' window.
    '
    hOutBmp = CreateCompatibleBitmap(DskHdc, (.Right - .Left + 1), (.Bottom - .Top + 1))
    '
    ' Fill the output bitmap's structure with
    ' the width, height and color information
    ' of the newly created bitmap.
    '
    rc = GetObject(hOutBmp, Len(OutBmp), OutBmp)
    '
    ' Create a memory DC compatible with the
    ' desktop window DC. The new memory DC's
    ' display surface is one monochrome pixel
    ' wide and one monochrome pixel high.
    '
    OutHdc = CreateCompatibleDC(DskHdc)
    '
    ' Copy the desktop compatible bitmap (hOutBmp)
    ' into the output/memory DC (OutHdc). rc is the
    ' handle of the replaced the existing DC.
    '
    rc = SelectObject(OutHdc, hOutBmp)
    '
    ' Copy the desktop compatible bitmap to the
    ' output DC streching or compressing it as
    ' required by the specified dimensions.
    '
    rc = StretchBlt(OutHdc, 0, 0, (.Right - .Left + 1), _
            (.Bottom - .Top + 1), DskHdc, 0, 0, _
            (DskRect.Right - DskRect.Left + 1), _
            (DskRect.Bottom - DskRect.Top + 1), _
             vbSrcCopy)
    '
    ' If the clear screen option was selected, set
    ' the output DC to black.
    '
    If gbClearScreen Then
       Call BitBlt(OutHdc, 0, 0, (.Right - .Left + 1), _
       (.Bottom - .Top + 1), DskHdc, 0, 0, vbBlackness)
    End If

End With
'
' Delete the output bitmap.
' Release the desktop DC.
'
rc = DeleteObject(hOutBmp)
rc = ReleaseDC(DskHwnd, DskHdc)
End Sub

Public Sub pPaintDeskDC(InHdc As Long, InBmp As BITMAP, OutHwnd As Long)
'
' Paint the picture, specified by InBmp, to the
' output window streching or compressing it as required.
'
Dim OutHdc  As Long   ' Output window DC handle.
Dim rc      As Long   ' Function return code
Dim OutRect As RECT   ' Rectangular size of output window.
'
' Get the dimensions of the client area of
' the destination window. Also get the
' destinations window's DC handle.
'
rc = GetClientRect(OutHwnd, OutRect)
OutHdc = GetWindowDC(OutHwnd)
'
' Paint the desktop picture to the output window
' streching or compressing it as required.
'
With OutRect
    rc = StretchBlt(OutHdc, 0, 0, (.Right - .Left + 1), _
        (.Bottom - .Top + 1), InHdc, 0, 0, _
         InBmp.bmWidth, InBmp.bmHeight, vbSrcCopy)
End With
'
' Release the source DC.
'
rc = ReleaseDC(OutHwnd, OutHdc)
End Sub

Public Sub pDrawTransparentBitmap(lHDCDest As Long, lBmSource As Long, _
        lMaskColor As Long, Optional lDestStartX As Long, _
        Optional lDestStartY As Long, Optional lDestWidth As Long, _
        Optional lDestHeight As Long, Optional lSrcStartX As Long, _
        Optional lSrcStartY As Long, Optional BkGrndHdc As Long)
'
' Draw the sprite onto the Form.  The background of
' the sprite is made transparent so the form's image
' shows through.
'
Dim lColorRef    As Long 'COLORREF
Dim lBmAndBack   As Long 'HBITMAP
Dim lBmAndObject As Long
Dim lBmAndMem    As Long
Dim lBmSave      As Long
Dim lBmBackOld   As Long
Dim lBmObjectOld As Long
Dim lBmMemOld    As Long
Dim lBmSaveOld   As Long
Dim lHDCMem      As Long 'HDC
Dim lHDCBack     As Long
Dim lHDCObject   As Long
Dim lHDCTemp     As Long
Dim lHDCSave     As Long
Dim x            As Long
Dim y            As Long
Dim udtBitMap    As BITMAP
Dim udtSize      As POINTAPI 'POINT
'
' Create a temporary DC compatible with the
' Destination DC (main form's DC).
'
lHDCTemp = CreateCompatibleDC(lHDCDest)
'
' Select the sprite's bitmap into the temporary DC.
' Store the sprite bitmap's characteristics in
' the udtBitMap.
'
Call SelectObject(lHDCTemp, lBmSource)
Call GetObject(lBmSource, Len(udtBitMap), udtBitMap)
'
' Set the size of the temporary bitmap.
'
With udtSize
    .x = udtBitMap.bmWidth
    .y = udtBitMap.bmHeight
    '
    ' Use the optionally passed in width and height values.
    '
    If lDestWidth <> 0 Then .x = lDestWidth
    If lDestHeight <> 0 Then .y = lDestHeight
    x = .x
    y = .y
End With
'
' Create some DCs compatible with
' the main form to hold temporary data.
'
lHDCBack = CreateCompatibleDC(lHDCDest)
lHDCObject = CreateCompatibleDC(lHDCDest)
lHDCMem = CreateCompatibleDC(lHDCDest)
lHDCSave = CreateCompatibleDC(lHDCDest)
'
' Create a bitmap for each DC.  DCs are required
' for a number of GDI functions.
'
' Monochrome bitmaps.
'
'02/025/2002
'lBmAndBack = CreateBitmap(x, y, 1, 1, 0&)
'lBmAndObject = CreateBitmap(x, y, 1, 1, 0&)
lBmAndBack = CreateBitmap(x, y, 1, 1, ByVal 0&)
lBmAndObject = CreateBitmap(x, y, 1, 1, ByVal 0&)
'
' Color Compatible bitmaps.
'
lBmAndMem = CreateCompatibleBitmap(lHDCDest, x, y)
lBmSave = CreateCompatibleBitmap(lHDCDest, x, y)
'
' Each DC must select a bitmap object to store pixel data.
'
' Monochrome.
'
lBmBackOld = SelectObject(lHDCBack, lBmAndBack)
lBmObjectOld = SelectObject(lHDCObject, lBmAndObject)
'
' Color.
'
lBmMemOld = SelectObject(lHDCMem, lBmAndMem)
lBmSaveOld = SelectObject(lHDCSave, lBmSave)
'
' Set the mapping mode of the temporary (sprite)
' DC to that of the form's DC.  The mapping mode
' defines the unit of measure used to transform
' page-space units into device-space units, and
' also defines the orientation of the device's
' x and y axes.
'
Call SetMapMode(lHDCTemp, GetMapMode(lHDCDest))
'
' Save the sprite bitmap that was passed in
' because it will be overwritten.
'
Call BitBlt(lHDCSave, 0&, 0&, x, y, lHDCTemp, lSrcStartX, lSrcStartY, vbSrcCopy)
'
' Set the background color of the sprite's DC to
' the color in the sprite that should be transparent.
'
' Example:
' The background of our sprite is black. Set the
' background color of the DC to black.
'
lColorRef = SetBkColor(lHDCTemp, lMaskColor)
'
' Create a mask for the sprite by performing a BitBlt from
' the sprite's bitmap to a monochrome bitmap.  The result is
' a matrix of 1's and 0's where 0 represents the foreground
' color and 1 represents the the background color.
'
' Suppose our sprite is a red "X" on a black blackground.
' The mask will contain 0's where the "X" is and 1's
' everywhere else.
'
Call BitBlt(lHDCObject, 0&, 0&, x, y, lHDCTemp, lSrcStartX, lSrcStartY, vbSrcCopy)
'
' Set the background color of the sprite's
' DC back to its original color.
'
Call SetBkColor(lHDCTemp, lColorRef)
'
' Create the inverse of the mask.
'
' In Our Example:
' The mask will have 1's where the "X" is and 0's
' everywhere else representing the black background.
'
Call BitBlt(lHDCBack, 0&, 0&, x, y, lHDCObject, 0&, 0&, vbNotSrcCopy)
'
' Copy the background of the main DC to the destination.
' The lHDCMem is a color version of the desktop image.
'
If (BkGrndHdc = 0) Then
    Call BitBlt(lHDCMem, 0&, 0&, x, y, lHDCDest, lDestStartX, lDestStartY, vbSrcCopy)
Else
    Call BitBlt(lHDCMem, 0&, 0&, x, y, BkGrndHdc, lDestStartX, lDestStartY, vbSrcCopy)
End If
'
' Mask out the places where the bitmap will be placed by AND-ing
' the memory DC with the mask with the 0's where the "X" is.
'
Call BitBlt(lHDCMem, 0&, 0&, x, y, lHDCObject, 0&, 0&, vbSrcAnd)
'
' Mask out the transparent colored pixels on the bitmap. The
' background of the sprite is masked out so only the "X" remains.
' lHDCTemp is the colored sprite.  This is AND-ed with the mask
' which has the 1's where the "X" is and 0's where the background is.
'
Call BitBlt(lHDCTemp, lSrcStartX, lSrcStartY, x, y, lHDCBack, 0&, 0&, vbSrcAnd)
'
' Combine the colored foreground of the sprite with the colored
' desktop image by OR-ing the bitmap in the prior step with
' that from two steps back.
'
Call BitBlt(lHDCMem, 0&, 0&, x, y, lHDCTemp, lSrcStartX, lSrcStartY, vbSrcPaint)
'
' Copy the final combination of desktop bitmap
' and the sprite's foreground to the form.
'
Call BitBlt(lHDCDest, lDestStartX, lDestStartY, x, y, lHDCMem, 0&, 0&, vbSrcCopy)
'
' Place the original sprite bitmap back into the bitmap sent here.
'
Call BitBlt(lHDCTemp, lSrcStartX, lSrcStartY, x, y, lHDCSave, 0&, 0&, vbSrcCopy)
'
' Delete memory bitmaps.
'
DeleteObject SelectObject(lHDCBack, lBmBackOld)
DeleteObject SelectObject(lHDCObject, lBmObjectOld)
DeleteObject SelectObject(lHDCMem, lBmMemOld)
DeleteObject SelectObject(lHDCSave, lBmSaveOld)
'
' Delete memory DC's
'
DeleteDC lHDCMem
DeleteDC lHDCBack
DeleteDC lHDCObject
DeleteDC lHDCSave
DeleteDC lHDCTemp
End Sub

Public Sub pSaveSettings()
'
' Save the current options to the registry.
'
Call fWriteValue("HKCU", cREGKEY, "Clear Screen", "B", gbClearScreen)
Call fWriteValue("HKCU", cREGKEY, "Use Tracers", "B", gbUseTracers)
Call fWriteValue("HKCU", cREGKEY, "Random Rate", "B", gbRefreshRND)
Call fWriteValue("HKCU", cREGKEY, "Random Size", "B", gbSizeRND)
Call fWriteValue("HKCU", cREGKEY, "Random Speed", "B", gbSpeedRND)
Call fWriteValue("HKCU", cREGKEY, "Sprite Count", "D", glSpriteCount)
Call fWriteValue("HKCU", cREGKEY, "Refresh Rate", "D", glRefreshRate)
Call fWriteValue("HKCU", cREGKEY, "Sprite Size", "D", glSpriteSize)
Call fWriteValue("HKCU", cREGKEY, "Sprite Speed", "D", glSpriteSpeed)
Call fWriteValue("HKCU", cREGKEY, "Sprite Image", "S", gsSpriteImage)
End Sub

Public Sub pLoadSettings()
'
' Read the current options from the registry.
'
Call fReadValue("HKCU", cREGKEY, "Clear Screen", "B", False, gbClearScreen)
Call fReadValue("HKCU", cREGKEY, "Use Tracers", "B", False, gbUseTracers)
Call fReadValue("HKCU", cREGKEY, "Random Rate", "B", True, gbRefreshRND)
Call fReadValue("HKCU", cREGKEY, "Random Size", "B", True, gbSizeRND)
Call fReadValue("HKCU", cREGKEY, "Random Speed", "B", True, gbSpeedRND)
Call fReadValue("HKCU", cREGKEY, "Password", "S", "", gsPassword)
Call fReadValue("HKCU", cREGKEY, "Sprite Image", "S", cIMAGE1, gsSpriteImage)
Select Case gsSpriteImage
    Case cIMAGE0
        glBmpYUnits = 10
        glBitMap_ID = 101
    Case cIMAGE2
        glBmpYUnits = 4
        glBitMap_ID = 103
    Case Else
        glBmpYUnits = 9
        glBitMap_ID = 102
End Select

Call fReadValue("HKCU", cREGKEY, "Sprite Count", "D", cDEF_SPRITECOUNT, glSpriteCount)
If (glSpriteCount < cMIN_SPRITECOUNT) Then glSpriteCount = cDEF_SPRITECOUNT
If (glSpriteCount > cMAX_SPRITECOUNT) Then glSpriteCount = cMAX_SPRITECOUNT

Call fReadValue("HKCU", cREGKEY, "Refresh Rate", "D", cMIN_REFRESHRATE, glRefreshRate)
If (glRefreshRate < cMIN_REFRESHRATE) Then glRefreshRate = cMIN_REFRESHRATE
If (glRefreshRate > cMAX_REFRESHRATE) Then glRefreshRate = cMIN_REFRESHRATE

Call fReadValue("HKCU", cREGKEY, "Sprite Size", "D", cMIN_SPRITESIZE, glSpriteSize)
If (glSpriteSize < cMIN_SPRITESIZE) Then glSpriteSize = cMIN_SPRITESIZE
If (glSpriteSize > cMAX_SPRITESIZE) Then glSpriteSize = cMAX_SPRITESIZE

Call fReadValue("HKCU", cREGKEY, "Sprite Speed", "D", cMIN_SPRITESPEED, glSpriteSpeed)
If (glSpriteSpeed < cMIN_SPRITESPEED) Then glSpriteSpeed = cMIN_SPRITESPEED
If (glSpriteSpeed > cMAX_SPRITESPEED) Then glSpriteSpeed = cMAX_SPRITESPEED
End Sub


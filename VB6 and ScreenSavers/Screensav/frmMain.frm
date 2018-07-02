VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   2790
   ClientLeft      =   5640
   ClientTop       =   2205
   ClientWidth     =   4440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrSprite 
      Interval        =   50
      Left            =   3930
      Top             =   2250
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oSprite As Sprite ' Sprite builder engine.


Private Sub Form_Load()
Dim l         As Long
Dim ScaleSize As Single
'
' Initialize the desktop image information.
'
tmrSprite.Enabled = False
Call fReadValue("HKCU", cREGKEY, "Clear Screen", "B", False, gbClearScreen)
Call pInitDeskDC(glDeskDC, gDeskBmp, gDispRec)
'
' Get the screen saver's settings from the registry.
' See if a screen saver password is used.
'
Call pLoadSettings
Call fReadValue("HKCU", "Control Panel\Desktop", _
        "ScreenSaveUsePassword", "D", False, gbUsePassword)
'
' Create new a sprite and
' resize the active sprite array.
'
Set oSprite = New Sprite
ReDim gaSprite(glSpriteCount - 1) As Sprite
'
' Initialize each sprite.
'
For l = LBound(gaSprite) To UBound(gaSprite)
    '
    ' Size the sprite either randomly or
    ' based on the registry value.
    '
    If gbSizeRND Then
        ScaleSize = (((cMAX_SPRITESIZE - cMIN_SPRITESIZE) * Rnd) + cMIN_SPRITESIZE) / 100
    Else
        ScaleSize = glSpriteSize / 100
    End If
    '
    ' Create a new active sprite.
    '
    Set gaSprite(l) = oSprite.CreateSprite(Me, glDeskDC, glBitMap_ID, vbBlack, _
        cBMPXUNITS * glBmpYUnits, cBMPXUNITS, glBmpYUnits, ScaleSize, ScaleSize, l)
    '
    ' Initialize the sprite.
    '
    With gaSprite(l)
        ' Calculate width and height of the display.
        '
        .BdrX = gDispRec.Right - CLng(.uWidth * 0.8)
        .BdrY = gDispRec.Bottom - CLng(.uHeight * 0.8)
        '
        ' Set the horizontal and vertical speed of the
        ' sprite either randomly or based on the user's
        ' registry value.
        '
        If gbSpeedRND Then
            .Dx = CLng(((20 * Rnd) + 1) * ScaleSize)
            .Dy = CLng(((20 * Rnd) + 1) * ScaleSize)
        Else
            .Dx = CLng(glSpriteSpeed * ScaleSize) + 1
            .Dy = .Dx
        End If
        '
        ' Randomly place sprite on x and y axes.
        ' Turn tracers on or off.
        '
        .x = CLng(.BdrX * Rnd) + 1
        .y = CLng(.BdrY * Rnd) + 1
        .DDx = 1      ' (Sprite acceleration) Not currently used.
        .DDy = 1      ' (Sprite acceleration) Not currently used.
        .Tracers = gbUseTracers
    End With
Next
'
' Set the animation timer interval to
' either random or the value read from
' the registry.
'
If gbRefreshRND Then
    tmrSprite.Interval = CLng((cMAX_REFRESHRATE - cMIN_REFRESHRATE + 1) * Rnd) + cMIN_REFRESHRATE
Else
    tmrSprite.Interval = (cMAX_REFRESHRATE - cMIN_REFRESHRATE) + 2 - glRefreshRate
End If
'
' Start the timer to animate the active sprites.
'
tmrSprite.Enabled = True
Set oSprite = Nothing
End Sub
Private Sub Form_Click()
If glRunMode = RM_NORMAL Then Call pRespond
End Sub
Private Sub Form_DblClick()
If glRunMode = RM_NORMAL Then Call pRespond
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If glRunMode = RM_NORMAL Then Call pRespond
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If glRunMode = RM_NORMAL Then Call pRespond
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If glRunMode = RM_NORMAL Then Call pRespond
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iCount As Long
'
' Besides firing when the mouse moves, this event is
' also fired when the form is first loaded and sized.
' This code prevents the screen saver from unloading
' under the wrong circumstances.
'
If glRunMode = RM_NORMAL Then
    If iCount > 2 Then
        Call pRespond
    Else
        iCount = iCount + 1
    End If
End If
End Sub
Private Sub pRespond()
Dim lPrev As Long
'
' Stop the screen saver from being the top most window.
'
Call ShowCursor(True)
Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0&, 0&, 0, 0, SWP_NOSIZE)
'
' If a password is required, prompt for it.
' Otherwise, end the screen saver.
'
If gbUsePassword Then
    frmPassword.Show vbModal
Else
    '
    ' Enable Ctrl-Alt-Delete and Alt-Tab.
    '
    Call SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, lPrev, 0)
    Unload Me
    End
End If
End Sub

Private Sub Form_Paint()
'
' Repaint the desktop bitmap to the form.
'
Call pPaintDeskDC(glDeskDC, gDeskBmp, hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim l As Integer
'
' Turn off the timer before destroying the
' sprite object otherwise you may deadlock.
'
tmrSprite.Enabled = False
'
' Destroy each active sprite.
'
For l = LBound(gaSprite) To UBound(gaSprite)
    Set gaSprite(l) = Nothing
Next
'
' Clean up the Desktop device context
' to prevent memory leaks.
'
Call DeleteDC(glDeskDC)
'
' Show the MousePointer
'
If (glRunMode = RM_NORMAL) Then Call ShowCursor(True)
Screen.MousePointer = vbDefault
Erase gaSprite
Set oSprite = Nothing
Set frmMain = Nothing
End Sub

Private Sub tmrSprite_Timer()
Dim l As Long
'
' Automatically move each active sprite.
'
For l = LBound(gaSprite) To UBound(gaSprite)
    gaSprite(l).AutoMove
Next
End Sub


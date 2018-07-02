VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'-----------------------------------------------------------------
' Public Variables
'-----------------------------------------------------------------
'
Public Index        As Long      ' Global array index value.
Public idxMin       As Long      ' Minimum sprite frame index value.
Public idxMax       As Long      ' Maximum sprite frame index value.
Public xUnits       As Long      ' # of horizontal sprite frames.
Public yUnits       As Long      ' # of virtical sprite frames.
Public uWidth       As Long      ' Sprite frame width in pixels.
Public uHeight      As Long      ' Sprite frame height in pixels.
Public DestHDC      As Long      ' Destination window hdc.
Public hBitmap      As Long      ' Handle to animation bitmap.
Public hDisplayBack As Long      ' Handle to background bitmap.
Public MASKCOLOR    As Long      ' Transparency blt color mask.
Public SprtH        As Long      ' Animation bitmap height in pixels.
Public SprtW        As Long      ' Animation bitmap width in pixels.
Public Mass         As Long      ' Sprite mass (virtual).
Public Tracers      As Boolean   ' Use tracers flag.
'
'-----------------------------------------------------------------
' AutoMove Programmable Variables
'-----------------------------------------------------------------
'
Public x            As Long      ' Sprite's current screen x coordinate.
Public y            As Long      ' Sprite's current screen y coordinate.
Public BdrX         As Long      ' Border width.
Public BdrY         As Long      ' Border height.
Public Dx           As Long      ' Current x veloCity.
Public Dy           As Long      ' Current y veloCity.
Public DDx          As Long      ' Current x acceleration (= 1 not currently used).
Public DDy          As Long      ' Current Y acceleration (= 1 not currently used).
Public ScreenW      As Long      ' Width of screen.
Public ScreenH      As Long      ' Height of screen.
Public bmpIdx       As Long      ' Current animated bitmap frame index.
'
'-----------------------------------------------------------------
' Private Variables
'-----------------------------------------------------------------
'
Private LastX       As Long      ' Previous x coordinate.
Private LastY       As Long      ' Previous y coordinate.

Public Function CollisionTest(Sprite As Sprite) As Boolean
'
' See if two sprites overlap.
'
Dim l1 As Long  ' Left, Right, Top, Bottom positions of sprite.
Dim r1 As Long
Dim t1 As Long
Dim b1 As Long
Dim l2 As Long  ' Left, Right, Top, Bottom positions of sprite.
Dim r2 As Long
Dim t2 As Long
Dim b2 As Long
'
' Don't compare a sprite with itself.
'
If (Sprite Is Me) Then Exit Function
'
' Current sprite. Get its coordinates.
'
With Me
    l1 = .x
    t1 = .y
    r1 = l1 + .uWidth
    b1 = t1 + .uHeight
End With
'
' Other sprite.  Get its coordinates.
'
With Sprite
    l2 = .x
    t2 = .y
    r2 = l2 + .uWidth
    b2 = t2 + .uHeight
End With
'
' Test for sprite collision
'
'02/05/2002
'CollisionTest = (((l2 <= l1) And (l1 <= r2)) Or _
'                 ((l2 <= r1) And (r1 <= r2))) And _
'                (((t2 <= t1) And (t1 <= b2)) Or _
'                 ((t2 <= b1) And (b1 <= b2)))
CollisionTest = (((l2 <= l1) And (l1 <= r2)) Or _
                ((l2 <= r1) And (r1 <= r2)) Or _
                ((l1 <= l2) And (l2 <= r1)) Or _
                 ((l1 <= r2) And (r2 <= r1))) And _
                (((t2 <= t1) And (t1 <= b2)) Or _
                 ((t2 <= b1) And (b1 <= b2)) Or _
                 ((t1 <= t2) And (t2 <= b1)) Or _
                 ((t1 <= b2) And (b2 <= b1)))

End Function
Public Function CreateSprite(Display As Form, hDisplay As Long, _
        ResID As Long, MaskCol As Long, picCount As Long, _
        xUnits As Long, yUnits As Long, Optional ScaleW As Single, _
        Optional ScaleH As Single, Optional Idx As Long) As Sprite
'
' Create a new sprite initializing its properties.
'
' The sprite is a "framed" or "phased" image meaning
' its bitmap consists of many images which are displayed
' in a sequence (like a cartoon) resulting in the
' appearance of motion.  The first frame displayed is
' picked ramdomly.
'
Dim cSprite As Sprite ' ScreenSaver sprite.
Dim RatioX  As Single ' X and Y shrinkage factors.
Dim RatioY  As Single
'
' Set the randomization seed, x and y
' ratio base and the form's scalemode to pixels.
'
Randomize Timer
RatioX = 1
RatioY = 1
Display.ScaleMode = vbPixels
'
' Create a new screen saver sprite.
'
Set cSprite = New Sprite
With cSprite
    '
    ' Set the display hdc.
    ' Calculate the screen dimensions in pixels.
    '
    .DestHDC = Display.hdc
    .ScreenW = (Screen.Width \ Screen.TwipsPerPixelX)
    .ScreenH = (Screen.Height \ Screen.TwipsPerPixelY)
    '
    ' Load the sprite from the resource file if not already loaded.
    '
    If (gSprite.ResID <> ResID) Then
        gSprite.ResID = ResID
        Set gSprite.Sprite = LoadResPicture(ResID, vbResBitmap)
    End If
    '
    ' Scale x ratio based on display size.
    ' Scale y ratio based on display size.
    '
    If (.ScreenW <> 0) Then RatioX = (gDispRec.Right - gDispRec.Left) / .ScreenW
    If (.ScreenH <> 0) Then RatioY = (gDispRec.Bottom - gDispRec.Top) / .ScreenH
    If (ScaleW = 0) Then ScaleW = 1
    If (ScaleH = 0) Then ScaleH = 1
    '
    ' Adjust ratios based on sprite scale dimensions.
    ' Scale dimension are based on the user selecting
    ' random sized sprites or picking a sprite size.
    '
    RatioX = RatioX * ScaleW
    RatioY = RatioY * ScaleH
    '
    ' Shrink the animated sprite frames.
    '
    .hBitmap = fShrinkBmp(.DestHDC, gSprite.Sprite.Handle, RatioX, RatioY)
    '
    ' Calculate the dimensions of the sprite image in terms
    ' of scale units in use on the main screen saver form.
    '
    .SprtW = CLng(Display.ScaleX(gSprite.Sprite.Width, vbHimetric, vbPixels) * RatioX)
    .SprtH = CLng(Display.ScaleY(gSprite.Sprite.Height, vbHimetric, vbPixels) * RatioY)
    
    .xUnits = xUnits                              ' Save # of horizontal frames.
    .yUnits = yUnits                              ' Save # of verticle frames.
    .uWidth = .SprtW \ xUnits                     ' Calculate single frame width.
    .uHeight = .SprtH \ yUnits                    ' Calculate single frame height.
    .idxMin = 0                                   ' Initialize frame index = 0.
    .idxMax = picCount - 1                        ' Init max frame index = # frames - 1.
    .bmpIdx = (.idxMax - .idxMin) * Rnd + .idxMin ' Randomize the first picture frame to be displayed.
    .hDisplayBack = hDisplay                      ' Save handle to the display bitmap.
    .MASKCOLOR = MaskCol                          ' Save the bitmap mask color.
    .Mass = CLng(cBASE_MASS * ScaleW * ScaleH)    ' Calculate mass based on scaled surface area.
    .Index = Idx                                  ' Save index possition in global array.
End With
'
' Return sprite reference.
' Destroy local sprite reference.
'
Set CreateSprite = cSprite
Set cSprite = Nothing
End Function

Private Function Atn2(y As Double, x As Double) As Double
'
' Calculate the ATan2(x,y) function.
'
If (x <> 0) Then           ' Prevent divide by zero
    Atn2 = Atn(y / x)      ' Atan2(y,x) = Atn(y/x) when x <> 0
Else                       ' Handle special case
    Atn2 = 2 * Atn(Sgn(y)) ' as N ~> infinity Atn(N) ~> (sign(N)*PI/2) = 2 * Atn(Sgn(y))
End If
End Function

Public Function ResolveCollision() As Boolean
'
' When 2 sprites collide, they are moved until
' the collision is resloved.  This relies on
' a bit of math and physics.
'
Dim x1    As Long
Dim y1    As Long
Dim x2    As Long
Dim y2    As Long
Dim a     As Double
Dim cos_a As Double
Dim sin_a As Double
Dim vn1   As Double
Dim vn2   As Double
Dim vp1   As Double
Dim vp2   As Double
Dim vx1   As Long
Dim vx2   As Long
Dim vy1   As Long
Dim vy2   As Long
Dim m1    As Double
Dim m2    As Double
Dim k     As Double
Dim e     As Double
Dim vn2p1 As Double
Dim vn2p2 As Double
Dim temp1 As Double
Dim Idx   As Integer
Dim First As Integer
Dim Last  As Integer

First = Me.Index + 1
Last = UBound(gaSprite)

For Idx = First To Last
    '
    ' Check for collisions between sprites.
    '
    If CollisionTest(gaSprite(Idx)) Then
        With gaSprite(Idx)
            '
            ' Compute the coordinates of the centers of the sprites.
            '
            x1 = Me.x + (Me.uWidth \ 2)
            y1 = Me.y + (Me.uHeight \ 2)
            x2 = .x + (.uWidth \ 2)
            y2 = .y + (.uHeight \ 2)
            '
            ' Compute the angle of the line joining their centers.
            '
            a = Atn2((y2 - y1), (x2 - x1))
            cos_a = Cos(a)
            sin_a = Sin(a)
            '
            ' Compute the velocities normal and
            ' perpendicular to the center line.
            '
            vx1 = Me.Dx
            vy1 = Me.Dy
            vx2 = .Dx
            vy2 = .Dy
            vn1 = (vx1 * cos_a) + (vy1 * sin_a)
            vp1 = (vy1 * cos_a) - (vx1 * sin_a)
            vn2 = (vx2 * cos_a) + (vy2 * sin_a)
            vp2 = (vy2 * cos_a) - (vx2 * sin_a)
            '
            ' Compute the momentum along the center line.
            '
            m1 = Me.Mass
            m2 = .Mass
            k = (m1 * vn1) + (m2 * vn2)
            '
            ' Compute the energy.
            '
            e = 0.5 * ((m1 * vn1 ^ 2) + (m2 * vn2 ^ 2))
            '
            ' There are two possible solutions to the equations.
            ' Compute both and choose.
            ' (Convert to long to fix Floating Point Error Bug.)
            '
            temp1 = Sqr(Fix(k ^ 2 - ((m1 / m2) + 1) * (-2 * e * m1 + k ^ 2)))
            vn2p1 = (k + temp1) / (m1 + m2)
            vn2p2 = (k - temp1) / (m1 + m2)
            '
            ' Choose the solution that is not the current state.
            '
            If (vn2p1 = vn2) Then
                vn2 = vn2p2
            Else
                vn2 = vn2p1
            End If
            '
            ' Compute the new vn1 value.
            '
            vn1 = (k - m2 * vn2) / m1
            '
            ' Compute the new x and y velocities.
            '
            vx1 = (vn1 * cos_a) - (vp1 * sin_a)
            vy1 = (vn1 * sin_a) + (vp1 * cos_a)
            vx2 = (vn2 * cos_a) - (vp2 * sin_a)
            vy2 = (vn2 * sin_a) + (vp2 * cos_a)
            '
            ' Save new change in velocity.
            '
            Me.Dx = vx1
            Me.Dy = vy1
            .Dx = vx2
            .Dy = vy2
            '
            ' Move the sprites until they are no longer in collision.
            '
            If ((vx1 <> 0) Or (vy1 <> 0) Or (vx2 <> 0) Or (vy2 <> 0)) Then
                Do While CollisionTest(gaSprite(Idx))
                    If ((Dx <> 0) Or (Dy <> 0)) Then        ' If 0 then update wont matter.
                        UpdatePosition                      ' Move sprite out of the way.
                    ElseIf ((.Dx <> 0) Or (.Dy <> 0)) Then  ' If 0 then update wont matter.
                        .UpdatePosition                     ' Move sprite out of the way.
                    Else
                        Exit Do                             ' Exit to avoid dead lock (infinite loop).
                    End If
                Loop
            End If
            ResolveCollision = True
        End With
    End If
Next
End Function

Public Sub UpdatePosition()
'
' Update the x and y positions of a sprite.
'
x = x + Dx
y = y + Dy
'
' Reverse direction when hitting a border.
'
If (x < 0) Then Dx = Abs(Dx)
If (x > BdrX) Then Dx = -1 * Abs(Dx)
If (y < 0) Then Dy = Abs(Dy)
If (y > BdrY) Then Dy = -1 * Abs(Dy)
'
' Acceleration sprite x and y velocities.
'
Dx = Dx * DDx
Dy = Dy * DDy
End Sub

Public Sub AutoMove()
'
' Move a sprite to next the coordinate.
'
Call DrawNext(x, y)
'
' Check for collision or update current position.
'
If Not ResolveCollision Then Call UpdatePosition
End Sub

Public Sub DrawNext(PosX As Long, PosY As Long)
'
' Paint a sprite in a new position.
'
Dim rc As Long
Dim x  As Long
Dim y  As Long  ' Source indexed bmp coordinates.
Dim x1 As Long
Dim y1 As Long
Dim w1 As Long
Dim h1 As Long  ' Repaint Rectangle # 1 screen coordinates.
Dim x2 As Long
Dim y2 As Long
Dim w2 As Long
Dim h2 As Long  ' Repaint Rectangle # 2 screen coordinates.
'
' Get the x and y coordinates of the next bitmap frame.
'
x = CLng((bmpIdx Mod xUnits) * (SprtW / xUnits))
y = CLng((bmpIdx \ xUnits) * (SprtH / yUnits))
'
' Paint the sprite in the new position.
'
If (Tracers) Then
    '
    ' When using Tracers don't clean up previous image.
    ' Leave off the last parameter.
    '
    Call pDrawTransparentBitmap(DestHDC, hBitmap, MASKCOLOR, PosX, PosY, uWidth, uHeight, x, y)
Else
    '
    ' Clean up & calculate unused sprite space.
    '
    Select Case PosX
        Case Is < LastX         ' PosX <=== LastX
            x1 = PosX + uWidth
            w1 = LastX - PosX
            x2 = LastX
            w2 = uWidth
        Case LastX              ' PosX ==== LastX
            x2 = LastX
            w2 = uWidth
        Case Is > LastX         ' PosX ===> LastX
            x1 = LastX
            w1 = PosX - LastX
            x2 = LastX
            w2 = uWidth
    End Select
    
    Select Case PosY
        Case Is < LastY             ' PosY <=== LastY
            y1 = LastY
            h1 = uHeight - (LastY - PosY)
            y2 = PosY + uHeight
            h2 = uHeight - h1
        Case LastY                  ' PosY ==== LastY
            y1 = LastY
            h1 = uHeight
        Case Is > LastY             ' PosY ===> LastY
            y1 = PosY
            h1 = uHeight - (PosY - LastY)
            y2 = LastY
            h2 = uHeight - h1
    End Select
    '
    ' Paint the sprite in the new position.
    '
    Call pDrawTransparentBitmap(DestHDC, hBitmap, MASKCOLOR, PosX, PosY, uWidth, uHeight, x, y, hDisplayBack)
    '
    ' If sprite has moved repaint the previous
    ' unoccupied positions.
    '
    If ((LastX <> PosX) Or (LastY <> PosY)) Then
        If ((w1 > 0) And (h1 > 0)) Then
            Call BitBlt(DestHDC, x1, y1, w1, h1, hDisplayBack, x1, y1, vbSrcCopy)
        End If
        If ((w2 > 0) And (h2 > 0)) Then
            Call BitBlt(DestHDC, x2, y2, w2, h2, hDisplayBack, x2, y2, vbSrcCopy)
        End If
    End If
End If
'
' Save previous x and y positions.
'
LastX = PosX
LastY = PosY
'
' Increment bitmap frame index or
' reset it to the beginning.
'
If (bmpIdx < idxMax) Then
    bmpIdx = bmpIdx + 1
Else
    bmpIdx = idxMin
End If
End Sub

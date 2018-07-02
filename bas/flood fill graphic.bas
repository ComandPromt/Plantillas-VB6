Attribute VB_Name = "FldFill"
Option Explicit
'Author:        Mike Kenton - mkenton@ix.netcom.com
'Date:          December 3, 1997
'Functions:     Fill,Twip2PixelX(),Twip2PixelY()
'API:           ExtFloodFill
'16 Bit:        True
'32 Bit:        True
'Description:   The module perfoms flood fill graphic functions

#If Win32 Then
Private Declare Function ExtFloodFill Lib "gdi32" _
                         (ByVal hDC As Long, _
                          ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal crColor As Long, _
                          ByVal wFillType As Long) As Long
#Else
Declare Function ExtFloodFill Lib "GDI" _
                         (ByVal hDC As Integer, _
                          ByVal X As Integer, _
                          ByVal Y As Integer, _
                          ByVal crColor As Long, _
                          ByVal wFillType As Integer) As Integer
#End If
 
Const FloodFillBorder = 0  ' "wFillType" Fill until crColor color encountered.
Const FloodFillSurface = 1 ' "wFillType" Fill surface until crColor color not encountered.
Public Sub Fill(Picture As Object, X As Long, Y As Long, Color As Long, Boarder As Long)
'Author:        Mike Kenton - mkenton@ix.netcom.com
'Date:          December 4, 1997
'Independant:   Twip2PixelX(),Twip2PixelY()
'API:           ExtFloodFill
'16 Bit:        True
'32 Bit:        True
'Description:   This routine Flood or fills an area on the screen using an api.
'               Picture is any object that has a FillStyle and FillColor properties.
'               X and Y are the coordinants of ware to begin the process in twips.
'               Color is an RGB value indicateing the color to fill with.
'               Boarder is the color at which the operation will halt or confine itself to.
'               Tje routine first sets the Picture Object FillStyle to 0 or solid.  Then
'               the FillColor value is set to the value passed.
'               Now routine first takes the x,y coordinants passed and converts them
'               into pixel coordinants for use by the api.

Dim iPixelX As Long
Dim iPixelY As Long
Dim ReturnValue As Long

Picture.FillStyle = 0
Picture.FillColor = Color

iPixelX = X \ Twip2PixelX(Picture)
iPixelY = Y \ Twip2PixelY(Picture)

ReturnValue = ExtFloodFill(Picture.hDC, iPixelX, iPixelY, Boarder, FloodFillBorder)
End Sub
Private Function Twip2PixelX(Picture As Object) As Long
'Author:        Mike Kenton - mkenton@ix.netcom.com
'Date:          December 4, 1997
'Independant:   True
'API:           None
'16 Bit:        True
'32 Bit:        True
'Description:   This function converts twip coordinants to Pixal coordinants and
'               returns a long value.
'               Picture is any object that has a ScaleMode property.
'               The function first stores the ScaleMode of the Picture object and then
'               sets the scalemode to 1 or Twip.  TwipX stores the ScaleWidth.  Now the
'               scale mode is set to 3 or Pixel.  PixelX stroes another ScaleWidth value.
'               The Picture object's scale mode is now returned to its origianl value.
'               The ratio of Twips to Pixels or TwipX/PixelX returns a conversion from
'               Twips to Pixels.

Dim PixelX As Integer
Dim TwipX As Long
Dim iScaleMode As Integer

iScaleMode = Picture.ScaleMode

Picture.ScaleMode = 1
TwipX = Picture.ScaleWidth

Picture.ScaleMode = 3
PixelX = Picture.ScaleWidth

Picture.ScaleMode = iScaleMode
Twip2PixelX = TwipX / PixelX

End Function
Private Function Twip2PixelY(Picture As Object) As Long
'Author:        Mike Kenton - mkenton@ix.netcom.com
'Date:          December 4, 1997
'Independant:   True
'API:           None
'16 Bit:        True
'32 Bit:        True
'Description:   This function converts twip coordinants to Pixal coordinants and
'               returns a long value.
'               Picture is any object that has a ScaleMode property.
'               The function first stores the ScaleMode of the Picture object and then
'               sets the scalemode to 1 or Twip.  TwipY stores the ScaleWidth.  Now the
'               scale mode is set to 3 or Pixel.  PixelY stroes another ScaleWidth value.
'               The Picture object's scale mode is now returned to its origianl value.
'               The ratio of Twips to Pixels or TwipY/PixelY returns a conversion from
'               Twips to Pixels.

Dim PixelY As Integer
Dim TwipY As Long
Dim iScaleMode As Integer

iScaleMode = Picture.ScaleMode

Picture.ScaleMode = 1
TwipY = Picture.ScaleHeight

Picture.ScaleMode = 3
PixelY = Picture.ScaleHeight

Picture.ScaleMode = iScaleMode
Twip2PixelY = TwipY / PixelY

End Function

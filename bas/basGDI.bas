Attribute VB_Name = "basGDI"
Option Explicit

Private Const LF_FACESIZE = 32

Private Type LogFont
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * LF_FACESIZE
 End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias _
   "CreateFontIndirectA" (lpLogFont As LogFont) As Long
   
Private Declare Function DeleteObject Lib "gdi32" _
   (ByVal hObject As Long) As Long
   
Private Declare Function SelectObject Lib "gdi32" _
   (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" _
   (ByVal hdc As Long, ByVal nBkMode As Long) As Long
   
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2

Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
   ByVal nIndex As Long) As Long
   
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
   ByVal hdc As Long) As Long
   
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Type TEXTMETRIC
   tmHeight As Integer
   tmAscent As Integer
   tmDescent As Integer
   tmInternalLeading As Integer
   tmExternalLeading As Integer
   tmAveCharWidth As Integer
   tmMaxCharWidth As Integer
   tmWeight As Integer
   tmItalic As String * 1
   tmUnderlined As String * 1
   tmStruckOut As String * 1
   tmFirstChar As String * 1
   tmLastChar As String * 1
   tmDefaultChar As String * 1
   tmBreakChar As String * 1
   tmPitchAndFamily As String * 1
   tmCharSet As String * 1
   tmOverhang As Integer
   tmDigitizedAspectX As Integer
   tmDigitizedAspectY As Integer
End Type

Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" _
  (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
  
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, _
   ByVal nMapMode As Long) As Long

Private Const MM_TEXT = 1

' Constants for get device caps
Private Const PHYSICALOFFSETX = 112
Private Const PHYSICALOFFSETY = 113
Private Const PLANES = 14
Private Const BITSPIXEL = 12
   
Public Const MARGIN_TOP = 1
Public Const MARGIN_BOTTOM = 2
Public Const MARGIN_LEFT = 3
Public Const MARGIN_RIGHT = 4


Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

'
' Gets the minumum margins for the printer.
' All returned values are in twips.
' It should also be noted the physical location 0,0
' of the printer object falls at the minimum top and left
' margins.
'
Public Function GetPrinterMinMargin(ByVal t As Integer) As Long
   Select Case t
    Case MARGIN_TOP:
       GetPrinterMinMargin = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY) _
           * Printer.TwipsPerPixelY
    Case MARGIN_BOTTOM:
       GetPrinterMinMargin = _
          Printer.Height - Printer.ScaleHeight - _
          (GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY) * Printer.TwipsPerPixelY)
    Case MARGIN_LEFT:
       GetPrinterMinMargin = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX) _
           * Printer.TwipsPerPixelX
       
    Case MARGIN_RIGHT:
       GetPrinterMinMargin = _
          Printer.Width - Printer.ScaleWidth - _
          GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX) * Printer.TwipsPerPixelX
   
    Case Else
       ' There's an error
       GetPrinterMinMargin = -1
    End Select
End Function

'
' Shades the form in a similar manner to many
' install programs.
'
' Optional Arguments:
' StartColor is what color to start with.
'   (Default = vbBlue)
' Fstep is the number of steps to use to fill the form.
'   (Default = 64)
' Cstep is the color step (change in color per step).
'   (Default = 4)
'
' Note: the effect can be reversed by calling ShadeForm with
'    a StartColor near black (but not completely 0) and by
'    setting a negative color step.
'
Public Sub ShadeForm(f As Form, Optional StartColor As Variant, Optional Fstep As Variant, Optional Cstep As Variant)
   Dim FillStep As Single  ' Not an integer because sometimes
                           ' rounding leaves a large bottom region
   Dim c As Long
   Dim FillArea As RECT
   Dim i As Integer
   Dim oldm As Integer
   Dim hBrush As Long
   Dim c2(1 To 3) As Long
   Dim cs2(1 To 3) As Long
   Dim fs As Long
   Dim cs As Integer
      
   ' Set defaults
   fs = IIf(IsMissing(Fstep), 64, CLng(Fstep))
   cs = IIf(IsMissing(Cstep), 4, CInt(Cstep))
   c = IIf(IsMissing(StartColor), vbBlue, CLng(StartColor))
   
   
   oldm = f.ScaleMode
   f.ScaleMode = vbPixels
   FillStep = f.ScaleHeight / fs
   FillArea.Left = 0
   FillArea.Right = f.ScaleWidth
   FillArea.Top = 0

   ' Break down the color and set individual
   ' color steps
   c2(1) = c And 255#
   cs2(1) = IIf(c2(1) > 0, cs, 0)
   c2(2) = (c \ 256#) And 255#
   cs2(2) = IIf(c2(2) > 0, cs, 0)
   c2(3) = (c \ 65536#) And 255#
   cs2(3) = IIf(c2(3) > 0, cs, 0)
   
   
   For i = 1 To fs
      FillArea.Bottom = FillStep * i

      hBrush = CreateSolidBrush(RGB(c2(1), c2(2), c2(3)))
      FillRect f.hdc, FillArea, hBrush
      DeleteObject hBrush
      
      ' Could do this in a loop, but it's simple
      ' and may be faster.
      c2(1) = (c2(1) - cs2(1)) And 255#
      c2(2) = (c2(2) - cs2(2)) And 255#
      c2(3) = (c2(3) - cs2(3)) And 255#
      
      FillArea.Top = FillArea.Bottom
   Next i
   
   f.ScaleMode = oldm
End Sub

'
'  Returns true if the system is using small fonts,
'  false if using large fonts
'
'  Source: the MS knowlege base article Q152136.
'
Public Function SmallFonts() As Boolean
   Dim hdc As Long
   Dim hwnd As Long
   Dim PrevMapMode As Long
   Dim tm As TEXTMETRIC

   ' Set the default return value to small fonts
   SmallFonts = True
   
   ' Get the handle of the desktop window
   hwnd = GetDesktopWindow()

   ' Get the device context for the desktop
   hdc = GetWindowDC(hwnd)
   If hdc Then
      ' Set the mapping mode to pixels
      PrevMapMode = SetMapMode(hdc, MM_TEXT)
      
      ' Get the size of the system font
      GetTextMetrics hdc, tm

      ' Set the mapping mode back to what it was
      PrevMapMode = SetMapMode(hdc, PrevMapMode)

      ' Release the device context
      ReleaseDC hwnd, hdc
     
      ' If the system font is more than 16 pixels high,
      ' then large fonts are being used
      If tm.tmHeight > 16 Then SmallFonts = False
   End If

End Function
'
' Returns the number of colors in the display.
'
Public Function GetNColors() As Long
  Dim hSrcDC As Integer

  hSrcDC = GetDC(GetDesktopWindow())
  GetNColors = GetDeviceCaps(hSrcDC, PLANES) * 2 ^ GetDeviceCaps(hSrcDC, BITSPIXEL)
  Call ReleaseDC(GetDesktopWindow(), hSrcDC)
End Function
'
' ob is a form, printer, or picturbox object
' You MUST call RestoreText with the handles (array)
' It should be called immediately after printing
' the rotated text and before changing any fonts, etc.
' or a leak in GDI resourses may occur.
'
' Note:  When printing rotated fonts to the printer
'        the .Transparent property is apparently ignored.
'        Use the SetTransparent() function to fix this.
'
' Bug: This doesn't work yet on forms or imageboxes :(
Public Function RotateText(ob As Object, ByVal angle As Single) As Variant
   Dim t As LogFont
   Dim i As Long
   Dim v(1 To 2) As Variant
   
   If ob Is Printer Then
      t.lfHeight = ob.FontSize * -20 / Printer.TwipsPerPixelY
   Else
      t.lfHeight = ob.FontSize * -20 / Screen.TwipsPerPixelY
   End If
   
   t.lfWidth = 0
   t.lfEscapement = CLng(angle * 10#)
   t.lfOrientation = t.lfEscapement
   t.lfWeight = ob.Font.Weight
   t.lfItalic = IIf(ob.FontItalic, 255, 0)
   t.lfUnderline = IIf(ob.FontUnderline, 255, 0)
   t.lfStrikeOut = IIf(ob.FontStrikethru, 255, 0)
   t.lfCharSet = 0
   t.lfOutPrecision = 0
   t.lfClipPrecision = 0
   t.lfQuality = 0
   t.lfPitchAndFamily = 0
   t.lfFaceName = ob.FontName & Chr$(0)

   i = CreateFontIndirect(t)
      
   v(1) = SelectObject(ob.hdc, i)
   v(2) = i
   
   RotateText = v
End Function
'
' Usually the same as ob.Transparent = t except that
' rotated fonts apparently ignore this object with
' the printer object.
'
Public Sub SetTransparent(ob As Object, ByVal t As Boolean)
   Call SetBkMode(ob.hdc, IIf(t, TRANSPARENT, OPAQUE))
End Sub

Public Sub RestoreText(ob As Object, handles As Variant)
   SelectObject ob.hdc, CLng(handles(1))
   DeleteObject CLng(handles(2))
End Sub



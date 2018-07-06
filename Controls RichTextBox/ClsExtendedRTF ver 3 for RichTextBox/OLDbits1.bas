Attribute VB_Name = "OLDbits1"
 Option Explicit

'Public Sub RainBow(Optional RedMagenta As Boolean = True, Optional ForeBack As Boolean = True)
'
'  'Copyright 2002 Roger Gilchrist
'  '*---PROGRAMMER MODIFICATION POINT---*
'  'delete if not needed
'
'  Dim i As Long, SPos As Long, EPos As Long
'  Dim RbowMod As Long, Rsection As Integer, Rstep As Long
'  Dim Cycler As Long, RainBowArray() As Long, drv As Long
'
'    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
'        Exit Sub                            'sets start and end and Exits if no Selection
'    End If
'    RbowMod = (EPos - SPos) \ 6            'this is used to change spectrum sections
'    If RbowMod < 1 Then
'        RbowMod = 1
'    End If
'    ReDim RainBowArray(SPos To EPos) As Long ' initialise array
'    For i = SPos To EPos
'        Rstep = (255 * Cycler / RbowMod)
'
'        If Cycler Mod RbowMod = 0 And Cycler <> 0 Then
'            Cycler = 0
'            Rsection = Rsection + 1
'            Rstep = 0
'        End If
'
'        Cycler = Cycler + 1 '    count through the spectrum section
'        RainBowArray(i) = RainbowColor(IIf(Rsection > 5, 5, Rsection), Rstep)
'    Next i
'If RedMagenta = False Then ' invert spectrum
'    RainBowArray = ArrayInvert(RainBowArray)
'End If
'    ColourApplicator RainBowArray, ForeBack
'
'End Sub

'
'Private Sub RainBowSubSet(Range As Spectrum, Optional InOut As Boolean = False, Optional LeftRight As Boolean = True, Optional ForeBack As Boolean = True)
'
'  'Copyright 2002 Roger Gilchrist
'  '*---PROGRAMMER MODIFICATION POINT---*
'  'delete if not needed
'
'  Dim i As Long
'  Dim SPos As Long, EPos As Long
'  Dim RbowMod As Long, Rstep As Long
'  Dim Cycler As Long, Rev As Boolean, ClrArray() As Long, RainBowArrayFlip() As Long, drv As Long
'  Dim DIST As Long
'
'    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
'        Exit Sub                            'sets start and end and Exits if no Selection
'    End If
'
'    RbowMod = IIf(InOut, (EPos - SPos) / 2, EPos - SPos) 'this is used to change spectrum sections
'    If RbowMod < 1 Then
'        RbowMod = 1
'    End If
'    ReDim ClrArray(SPos To EPos) As Long
'    ReDim RainBowArrayFlip(SPos To EPos) As Long
'    For i = SPos To EPos
'        Rstep = 255 * Cycler / RbowMod
'        If Rstep = 255 Then
'            Rev = True
'        End If
'        Cycler = Cycler + IIf(Rev, -1, 1)
'        drv = i
'        If LeftRight = False Then ' invert spectrum
'            drv = SPos + EPos - i
'        End If
'
'        ClrArray(drv) = RainbowColor(CInt(Range), Rstep)
'
'    Next i
'
'    If InOut And LeftRight = False Then ' offsets the array to allow style 4
'        DIST = EPos - SPos
'        For i = LBound(ClrArray) To LBound(ClrArray) + DIST \ 2
'            RainBowArrayFlip(i) = ClrArray(i + (DIST - 1) / 2)
'        Next i
'        For i = LBound(ClrArray) + DIST \ 2 To UBound(ClrArray)
'            RainBowArrayFlip(i) = RainbowColor(i - DIST \ 2, 0)
'            RainBowArrayFlip(i) = RainbowColor(i - DIST \ 2, 1)
'        Next i
'        ClrArray = RainBowArrayFlip
'    End If
'
'
'If InOut Then
'    ClrArray = ArrayInOut(ClrArray)
'
'End If
'
'If LeftRight = False Then ' invert spectrum
'    ClrArray = ArrayInvert(ClrArray)
'End If
'    ColourApplicator ClrArray, ForeBack
'
'End Sub

'Public Sub SpectrumSector(col As Spectrum, Mode As Integer, Optional ForeBack As Boolean)
'
'  'Copyright 2002 Roger Gilchrist
'  'public interface for RainBowSubset
'  '*---PROGRAMMER MODIFICATION POINT---*
'  'delete if unwanted
'
'
'    Select Case Mode
'      Case LeftRight
'        RainBowSubSet col, False, True, ForeBack
'      Case RightLeft
'        RainBowSubSet col, False, False, ForeBack
'      Case InOutLeftRight
'        RainBowSubSet col, True, True, ForeBack
'      Case InOutRightLeft
'        RainBowSubSet col, True, False, ForeBack
'    End Select
'
'End Sub

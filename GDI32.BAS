Attribute VB_Name = "basGDI32"
' Module Name               basGDI32
' Module Description        Module which contains ALL Declarations
'                           for API calls in the GDI32 Library
'
Option Explicit

Declare Function AbortDoc _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function AbortPath _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function AddFontResource _
    Lib "gdi32" _
    Alias "AddFontResourceA" ( _
        ByVal lpFileName As String) _
    As Long

Declare Function AngleArc _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal dwRadius As Long, _
        ByVal eStartAngle As Double, ByVal eSweepAngle As Double) _
    As Long

Declare Function AnimatePalette _
    Lib "gdi32" _
    Alias "AnimatePaletteA" ( _
        ByVal hPalette As Long, _
        ByVal wStartIndex As Long, _
        ByVal wNumEntries As Long, _
        lpPaletteColors As PALETTEENTRY) _
    As Long

Declare Function Arc _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) _
    As Long

Declare Function ArcTo _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) _
    As Long

Declare Function BeginPath _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function BitBlt _
    Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
    As Long

Declare Function CheckColorsInGamut _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpv As Any, _
        lpv2 As Any, _
        ByVal dw As Long) _
    As Long

Declare Function ChoosePixelFormat _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) _
    As Long

Declare Function Chord _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) _
    As Long

Declare Function CloseEnhMetaFile _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function CloseFigure _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function CloseMetaFile _
    Lib "gdi32" ( _
        ByVal hMF As Long) _
    As Long

Declare Function ColorMatchToTarget _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal hdc2 As Long, _
        ByVal dw As Long) _
    As Long

Declare Function CombineRgn _
    Lib "gdi32" ( _
        ByVal hDestRgn As Long, _
        ByVal hSrcRgn1 As Long, _
        ByVal hSrcRgn2 As Long, _
        ByVal nCombineMode As Long) _
    As Long

Declare Function CombineTransform _
    Lib "gdi32" ( _
        lpxformResult As XFORM, _
        lpxform1 As XFORM, _
        lpxform2 As XFORM) _
    As Long

Declare Function CopyEnhMetaFile _
    Lib "gdi32" _
    Alias "CopyEnhMetaFileA" ( _
        ByVal hemfSrc As Long, _
        ByVal lpszFile As String) _
    As Long

Declare Function CopyMetaFile _
    Lib "gdi32" _
    Alias "CopyMetaFileA" ( _
        ByVal hMF As Long, _
        ByVal lpFileName As String) _
    As Long

Declare Function CreateBitmap _
    Lib "gdi32" ( _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal nPlanes As Long, _
        ByVal nBitCount As Long, _
        lpBits As Any) _
    As Long

Declare Function CreateBitmapIndirect _
    Lib "gdi32" ( _
        lpBitmap As BITMAP) _
    As Long

Declare Function CreateBrushIndirect _
    Lib "gdi32" ( _
        lpLogBrush As LOGBRUSH) _
    As Long

Declare Function CreateColorSpace _
    Lib "gdi32" _
    Alias "CreateColorSpaceA" ( _
        lplogcolorspace As LOGCOLORSPACE) _
    As Long

Declare Function CreateCompatibleBitmap _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long) _
    As Long

Declare Function CreateCompatibleDC _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function CreateDiscardableBitmap _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long) _
    As Long

Declare Function CreateEllipticRgn _
    Lib "gdi32" ( _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal X2 As Long, _
        ByVal Y2 As Long) _
    As Long

Declare Function CreateEllipticRgnIndirect _
    Lib "gdi32" ( _
        lpRect As RECT) _
    As Long

Declare Function CreateEnhMetaFile _
    Lib "gdi32" _
    Alias "CreateEnhMetaFileA" ( _
        ByVal hdcRef As Long, _
        ByVal lpFileName As String, _
        lpRect As RECT, _
        ByVal lpDescription As String) _
    As Long

Declare Function CreateFont _
    Lib "gdi32" _
    Alias "CreateFontA" ( _
        ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, _
        ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, _
        ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) _
    As Long

Declare Function CreateFontIndirect _
    Lib "gdi32" _
    Alias "CreateFontIndirectA" ( _
        lpLogFont As LOGFONT) _
    As Long

Declare Function CreateHalftonePalette _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function CreateHatchBrush _
    Lib "gdi32" ( _
        ByVal nIndex As Long, _
        ByVal crColor As Long) _
    As Long

Declare Function CreateIC _
    Lib "gdi32" _
    Alias "CreateICA" ( _
        ByVal lpDriverName As String, _
        ByVal lpDeviceName As String, _
        ByVal lpOutput As String, _
        lpInitData As DEVMODE) _
    As Long

Declare Function CreateMetaFile _
    Lib "gdi32" _
    Alias "CreateMetaFileA" ( _
        ByVal lpString As String) _
    As Long

Declare Function CreatePalette _
    Lib "gdi32" ( _
        lpLogPalette As LOGPALETTE) _
    As Long

Declare Function CreatePatternBrush _
    Lib "gdi32" ( _
        ByVal hBitmap As Long) _
    As Long

Declare Function CreatePen _
    Lib "gdi32" ( _
        ByVal nPenStyle As Long, _
        ByVal nWidth As Long, _
        ByVal crColor As Long) _
    As Long

Declare Function CreatePenIndirect _
    Lib "gdi32" ( _
        lpLogPen As LOGPEN) _
    As Long

Declare Function CreatePolygonRgn _
    Lib "gdi32" ( _
        lpPoint As POINTAPI, _
        ByVal nCount As Long, _
        ByVal nPolyFillMode As Long) _
    As Long

Declare Function CreatePolyPolygonRgn _
    Lib "gdi32" ( _
        lpPoint As POINTAPI, _
        lpPolyCounts As Long, _
        ByVal nCount As Long, _
        ByVal nPolyFillMode As Long) _
    As Long

Declare Function CreateRectRgn _
    Lib "gdi32" ( _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal X2 As Long, _
        ByVal Y2 As Long) _
    As Long

Declare Function CreateRectRgnIndirect _
    Lib "gdi32" ( _
        lpRect As RECT) _
    As Long

Declare Function CreateRoundRectRgn _
    Lib "gdi32" ( _
        ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long) _
    As Long

Declare Function CreateScalableFontResource _
    Lib "gdi32" _
    Alias "CreateScalableFontResourceA" ( _
        ByVal fHidden As Long, ByVal lpszResourceFile As String, _
        ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) _
    As Long

Declare Function CreateSolidBrush _
    Lib "gdi32" ( _
        ByVal crColor As Long) _
    As Long

Declare Function DeleteColorSpace _
    Lib "gdi32" ( _
        ByVal hcolorspace As Long) _
    As Long

Declare Function DeleteDC _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function DeleteEnhMetaFile _
    Lib "gdi32" ( _
        ByVal hemf As Long) _
    As Long

Declare Function DeleteMetaFile _
    Lib "gdi32" ( _
        ByVal hMF As Long) _
    As Long

Declare Function DeleteObject _
    Lib "gdi32" ( _
        ByVal hObject As Long) _
    As Long

Declare Function DescribePixelFormat _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal n As Long, _
        ByVal un As Long, _
        lpPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) _
    As Long

Declare Function DPtoLP _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpPoint As POINTAPI, _
        ByVal nCount As Long) _
    As Long

Declare Function DrawEscape _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nEscape As Long, _
        ByVal cbInput As Long, _
        ByVal lpszInData As String) _
    As Long

Declare Function Ellipse _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long) _
    As Long

Declare Function EndDoc _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function EndPage _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function EndPath _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function EqualRgn _
    Lib "gdi32" ( _
        ByVal hSrcRgn1 As Long, _
        ByVal hSrcRgn2 As Long) _
    As Long

Declare Function Escape _
    Lib "gdi32" ( _
        ByVal hDC As Long, ByVal nEscape As Long, _
        ByVal nCount As Long, ByVal lpInData As String, _
        lpOutData As Any) _
    As Long

Declare Function ExcludeClipRect _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long) _
    As Long

Declare Function ExtCreatePen _
    Lib "gdi32" ( _
        ByVal dwPenStyle As Long, ByVal dwWidth As Long, _
        lplb As LOGBRUSH, _
        ByVal dwStyleCount As Long, lpStyle As Long) _
    As Long

Declare Function ExtCreateRegion _
    Lib "gdi32" ( _
        lpXform As XFORM, _
        ByVal nCount As Long, _
        lpRgnData As RGNDATA) _
    As Long

Declare Function ExtEscape _
    Lib "gdi32" ( _
        ByVal hDC As Long, ByVal nEscape As Long, _
        ByVal cbInput As Long, ByVal lpszInData As String, _
        ByVal cbOutput As Long, ByVal lpszOutData As String) _
    As Long

Declare Function ExtFloodFill _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal crColor As Long, ByVal wFillType As Long) _
    As Long

Declare Function ExtSelectClipRgn _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal hRgn As Long, _
        ByVal fnMode As Long) _
    As Long

Declare Function ExtTextOut _
    Lib "gdi32" _
    Alias "ExtTextOutA" ( _
        ByVal hDC As Long, ByVal x As Long, ByVal y As Long, _
        ByVal wOptions As Long, lpRect As RECT, _
        ByVal lpString As String, ByVal nCount As Long, lpDx As Long) _
    As Long

Declare Function FillPath _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function FillRgn _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal hRgn As Long, _
        ByVal hBrush As Long) _
    As Long

Declare Function FrameRgn _
    Lib "gdi32" ( _
        ByVal hDC As Long, ByVal hRgn As Long, _
        ByVal hBrush As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long) _
    As Long

Declare Function GdiComment _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal cbSize As Long, _
        lpData As Byte) _
    As Long

Declare Function GdiFlush _
    Lib "gdi32" () _
    As Long

Declare Function GdiGetBatchLimit _
    Lib "gdi32" () _
    As Long

Declare Function GdiSetBatchLimit _
    Lib "gdi32" ( _
        ByVal dwLimit As Long) _
    As Long

Declare Function GetArcDirection _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetAspectRatioFilterEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpAspectRatio As Size) _
    As Long

Declare Function GetBitmapBits _
    Lib "gdi32" ( _
        ByVal hBitmap As Long, _
        ByVal dwCount As Long, _
        lpBits As Any) _
    As Long

Declare Function GetBitmapDimensionEx _
    Lib "gdi32" ( _
        ByVal hBitmap As Long, _
        lpDimension As Size) _
    As Long

Declare Function GetBkColor _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetBkMode _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetBoundsRect _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lprcBounds As RECT, _
        ByVal flags As Long) _
    As Long

Declare Function GetBrushOrgEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpPoint As POINTAPI) _
    As Boolean

Declare Function GetCharABCWidths _
    Lib "gdi32" _
    Alias "GetCharABCWidthsA" ( _
        ByVal hDC As Long, _
        ByVal uFirstChar As Long, ByVal uLastChar As Long, _
        lpabc As ABC) _
    As Long

Declare Function GetCharABCWidthsFloat _
    Lib "gdi32" _
    Alias "GetCharABCWidthsFloatA" ( _
        ByVal hDC As Long, _
        ByVal iFirstChar As Long, ByVal iLastChar As Long, _
        lpABCF As ABCFLOAT) _
    As Long

Declare Function GetCharacterPlacement _
    Lib "gdi32" _
    Alias " GetCharacterPlacementA" ( _
        ByVal hDC As Long, ByVal lpsz As String, _
        ByVal n1 As Long, ByVal n2 As Long, _
        lpGcpResults As GCP_RESULTS, ByVal dw As Long) _
    As Long

Declare Function GetCharWidth _
    Lib "gdi32" _
    Alias "GetCharWidthA" ( _
        ByVal hDC As Long, _
        ByVal wFirstChar As Long, ByVal wLastChar As Long, _
        lpBuffer As Long) _
    As Long

Declare Function GetCharWidth32 _
    Lib "gdi32" _
    Alias "GetCharWidth32A" ( _
        ByVal hDC As Long, _
        ByVal iFirstChar As Long, ByVal iLastChar As Long, _
        lpBuffer As Long) _
    As Long

Declare Function GetCharWidthFloat _
    Lib "gdi32" _
    Alias "GetCharWidthFloatA" ( _
        ByVal hDC As Long, _
        ByVal iFirstChar As Long, ByVal iLastChar As Long, _
        pxBuffer As Double) _
    As Long

Declare Function GetClipBox _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpRect As RECT) _
    As Long

Declare Function GetClipRgn _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal hRgn As Long) _
    As Long

Declare Function GetColorAdjustment _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpca As ColorAdjustment) _
    As Long

Declare Function GetColorSpace _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetDCOrgEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpPoint As POINTAPI) _
    As Long

Declare Function GetDeviceCaps _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nIndex As Long) _
    As Long

Declare Function GetDeviceGammaRamp _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpv As Any) _
    As Long

Declare Function GetDIBColorTable _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal un1 As Long, _
        ByVal un2 As Long, _
        pRGBQuad As RGBQUAD) _
    As Long

Declare Function GetDIBits _
    Lib "gdi32" ( _
        ByVal aHDC As Long, ByVal hBitmap As Long, _
        ByVal nStartScan As Long, ByVal nNumScans As Long, _
        lpBits As Any, lpBI As BITMAPINFO, _
        ByVal wUsage As Long) _
    As Long

Declare Function GetEnhMetaFile _
    Lib "gdi32" _
    Alias "GetEnhMetaFileA" ( _
        ByVal lpszMetaFile As String) _
    As Long

Declare Function GetEnhMetaFileBits _
    Lib "gdi32" ( _
        ByVal hemf As Long, _
        ByVal cbBuffer As Long, _
        lpbBuffer As Byte) _
    As Long

Declare Function GetEnhMetaFileDescription _
    Lib "gdi32" _
    Alias "GetEnhMetaFileDescriptionA" ( _
        ByVal hemf As Long, _
        ByVal cchBuffer As Long, _
        ByVal lpszDescription As String) _
    As Long

Declare Function GetEnhMetaFileHeader _
    Lib "gdi32" ( _
        ByVal hemf As Long, _
        ByVal cbBuffer As Long, _
        lpemh As ENHMETAHEADER) _
    As Long

Declare Function GetEnhMetaFilePaletteEntries _
    Lib "gdi32" ( _
        ByVal hemf As Long, _
        ByVal cEntries As Long, _
        lppe As PALETTEENTRY) _
    As Long

Declare Function GetFontData _
    Lib "gdi32" _
    Alias "GetFontDataA" ( _
        ByVal hDC As Long, _
        ByVal dwTable As Long, ByVal dwOffset As Long, _
        lpvBuffer As Any, ByVal cbData As Long) _
    As Long

Declare Function GetFontLanguageInfo _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetGlyphOutline _
    Lib "gdi32" _
    Alias "GetGlyphOutlineA" ( _
        ByVal hDC As Long, ByVal uChar As Long, _
        ByVal fuFormat As Long, lpgm As GLYPHMETRICS, _
        ByVal cbBuffer As Long, lpBuffer As Any, lpmat2 As MAT2) _
    As Long

Declare Function GetGraphicsMode _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetICMProfile _
    Lib "gdi32" _
    Alias "GetICMProfileA" ( _
        ByVal hDC As Long, _
        ByVal dw As Long, _
        ByVal lpStr As String) _
    As Long

Declare Function GetKerningPairs _
    Lib "gdi32" _
    Alias "GetKerningPairsA" ( _
        ByVal hDC As Long, _
        ByVal cPairs As Long, _
        lpkrnpair As KERNINGPAIR) _
    As Long

Declare Function GetLogColorSpace _
    Lib "gdi32" _
    Alias "GetLogColorSpaceA" ( _
        ByVal hcolorspace As Long, _
        ByVal lplogcolorspace As LOGCOLORSPACE, _
        ByVal dw As Long) _
    As Long

Declare Function GetMapMode _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetMetaFile _
    Lib "gdi32" _
    Alias "GetMetaFileA" ( _
        ByVal lpFileName As String) _
    As Long

Declare Function GetMetaFileBitsEx _
    Lib "gdi32" ( _
        ByVal hMF As Long, _
        ByVal nSize As Long, _
        lpvData As Any) _
    As Long

Declare Function GetMetaRgn _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal hRgn As Long) _
    As Long

Declare Function GetMiterLimit _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        peLimit As Double) _
    As Long

Declare Function GetObject _
    Lib "gdi32" _
    Alias "GetObjectA" ( _
        ByVal hObject As Long, _
        ByVal nCount As Long, _
        lpObject As Any) _
    As Long

Declare Function GetObjectType _
    Lib "gdi32" ( _
        ByVal hgdiobj As Long) _
    As Long

'Declare Function GetOutlineTextMetrics _
'    Lib "gdi32" _
'    Alias "GetOutlineTextMetricsA" ( _
'        ByVal hDC As Long, _
'        ByVal cbData As Long, _
'        lpotm As OUTLINETEXTMETRIC) _
'    As Long

Declare Function GetPaletteEntries _
    Lib "gdi32" ( _
        ByVal hPalette As Long, _
        ByVal wStartIndex As Long, _
        ByVal wNumEntries As Long, _
        lpPaletteEntries As PALETTEENTRY) _
    As Long

Declare Function GetPath _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpPoint As POINTAPI, _
        lpTypes As Byte, _
        ByVal nSize As Long) _
    As Long

Declare Function GetPixel _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long) _
    As Long

Declare Function GetPixelFormat _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetPolyFillMode _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetRasterizerCaps _
    Lib "gdi32" ( _
        lpraststat As RASTERIZER_STATUS, _
        ByVal cb As Long) _
    As Long

Declare Function GetRegionData _
    Lib "gdi32" _
    Alias "GetRegionDataA" ( _
        ByVal hRgn As Long, _
        ByVal dwCount As Long, _
        lpRgnData As RGNDATA) _
    As Long

Declare Function GetRgnBox _
    Lib "gdi32" ( _
        ByVal hRgn As Long, _
        lpRect As RECT) _
    As Long

Declare Function GetROP2 _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetStockObject _
    Lib "gdi32" ( _
        ByVal nIndex As Long) _
    As Long

Declare Function GetStretchBltMode _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetTextAlign _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetTextCharacterExtra _
    Lib "gdi32" _
    Alias "GetTextCharacterExtraA" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetTextCharset _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetTextCharsetInfo _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpSig As FONTSIGNATURE, _
        ByVal dwFlags As Long) _
    As Long

Declare Function GetTextColor _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function GetTextExtentExPoint _
    Lib "gdi32" _
    Alias "GetTextExtentExPointA" ( _
        ByVal hDC As Long, ByVal lpszStr As String, _
        ByVal cchString As Long, ByVal nMaxExtent As Long, _
        lpnFit As Long, alpDx As Long, lpSize As Size) _
    As Long

Declare Function GetTextExtentPoint _
    Lib "gdi32" _
    Alias "GetTextExtentPointA" ( _
        ByVal hDC As Long, ByVal lpszString As String, _
        ByVal cbString As Long, lpSize As Size) _
    As Long

Declare Function GetTextExtentPoint32 _
    Lib "gdi32" _
    Alias "GetTextExtentPoint32A" ( _
        ByVal hDC As Long, ByVal lpsz As String, _
        ByVal cbString As Long, lpSize As Size) _
    As Long

Declare Function GetTextFace _
    Lib "gdi32" _
    Alias "GetTextFaceA" ( _
        ByVal hDC As Long, _
        ByVal nCount As Long, _
        ByVal lpFacename As String) _
    As Long

Declare Function GetTextMetrics _
    Lib "gdi32" _
    Alias "GetTextMetricsA" ( _
        ByVal hDC As Long, _
        lpMetrics As TEXTMETRIC) _
    As Long

Declare Function GetViewportExtEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpSize As Size) _
    As Long

Declare Function GetViewportOrgEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpPoint As POINTAPI) _
    As Long

Declare Function GetWindowExtEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpSize As Size) _
    As Long

Declare Function GetWindowOrgEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpPoint As POINTAPI) _
    As Long

Declare Function GetWinMetaFileBits _
    Lib "gdi32" ( _
        ByVal hemf As Long, ByVal cbBuffer As Long, _
        lpbBuffer As Byte, ByVal fnMapMode As Long, _
        ByVal hdcRef As Long) _
    As Long

Declare Function GetWorldTransform _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpXform As XFORM) _
    As Long

Declare Function IntersectClipRect _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long) _
    As Long

Declare Function InvertRgn _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal hRgn As Long) _
    As Long

Declare Function LineDDA _
    Lib "gdi32" ( _
        ByVal n1 As Long, ByVal n2 As Long, _
        ByVal n3 As Long, ByVal n4 As Long, _
        ByVal lpLineDDAProc As Long, ByVal lParam As Long) _
    As Boolean

Declare Function LineTo _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long) _
    As Long

Declare Function LPtoDP _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpPoint As POINTAPI, _
        ByVal nCount As Long) _
    As Long

Declare Function MaskBlt _
    Lib "gdi32" ( _
        ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, _
        ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal hbmMask As Long, _
        ByVal xMask As Long, ByVal yMask As Long, ByVal dwRop As Long) _
    As Long

Declare Function ModifyWorldTransform _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpXform As XFORM, _
        ByVal iMode As Long) _
    As Long

Declare Function MoveToEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        lpPoint As POINTAPI) _
    As Long

Declare Function OffsetClipRgn _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long) _
    As Long

Declare Function OffsetRgn _
    Lib "gdi32" ( _
        ByVal hRgn As Long, _
        ByVal x As Long, _
        ByVal y As Long) _
    As Long

Declare Function OffsetViewportOrgEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nX As Long, _
        ByVal nY As Long, _
        lpPoint As POINTAPI) _
    As Long

Declare Function OffsetWindowOrgEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nX As Long, _
        ByVal nY As Long, _
        lpPoint As POINTAPI) _
    As Long

Declare Function PaintRgn _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal hRgn As Long) _
    As Long

Declare Function PatBlt _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal dwRop As Long) _
    As Long

Declare Function PathToRegion _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function Pie _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) _
    As Long

Declare Function PlayEnhMetaFile _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal hemf As Long, _
        lpRect As RECT) _
    As Long

Declare Function PlayEnhMetaFileRecord _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpHandletable As HANDLETABLE, _
        lpEnhMetaRecord As ENHMETARECORD, _
        ByVal nHandles As Long) _
    As Long

Declare Function PlayMetaFile _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal hMF As Long) _
    As Long

Declare Function PlayMetaFileRecord _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpHandletable As HANDLETABLE, _
        lpMetaRecord As METARECORD, _
        ByVal nHandles As Long) _
    As Long

Declare Function PlgBlt _
    Lib "gdi32" ( _
        ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, _
        ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, _
        ByVal yMask As Long) _
    As Long

Declare Function PolyBezier _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lppt As POINTAPI, _
        ByVal cPoints As Long) _
    As Long

Declare Function PolyBezierTo _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lppt As POINTAPI, _
        ByVal cCount As Long) _
    As Long

Declare Function PolyDraw _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lppt As POINTAPI, _
        lpbTypes As Byte, _
        ByVal cCount As Long) _
    As Long

Declare Function Polygon _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpPoint As POINTAPI, _
        ByVal nCount As Long) _
    As Long

Declare Function Polyline _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpPoint As POINTAPI, _
        ByVal nCount As Long) _
    As Long

Declare Function PolylineTo _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lppt As POINTAPI, _
        ByVal cCount As Long) _
    As Long

Declare Function PolyPolygon _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpPoint As POINTAPI, _
        lpPolyCounts As Long, _
        ByVal nCount As Long) _
    As Long

Declare Function PolyPolyline _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lppt As POINTAPI, _
        lpdwPolyPoints As Long, _
        ByVal cCount As Long) _
    As Long

Declare Function PolyTextOut _
    Lib "gdi32" _
    Alias "PolyTextOutA" ( _
        ByVal hDC As Long, _
        pptxt As POLYTEXT, _
        cStrings As Long) _
    As Long

Declare Function PtInRegion _
    Lib "gdi32" ( _
        ByVal hRgn As Long, _
        ByVal x As Long, _
        ByVal y As Long) _
    As Long

Declare Function PtVisible _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long) _
    As Long

Declare Function RealizePalette _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function Rectangle _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long) _
    As Long

Declare Function RectInRegion _
    Lib "gdi32" ( _
        ByVal hRgn As Long, _
        lpRect As RECT) _
    As Long

Declare Function RectVisible _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpRect As RECT) _
    As Long

Declare Function RemoveFontResource _
    Lib "gdi32" _
    Alias "RemoveFontResourceA" ( _
        ByVal lpFileName As String) _
    As Long

Declare Function ResetDC _
    Lib "gdi32" _
    Alias "ResetDCA" ( _
        ByVal hDC As Long, _
        lpInitData As DEVMODE) _
    As Long

Declare Function ResizePalette _
    Lib "gdi32" ( _
        ByVal hPalette As Long, _
        ByVal nNumEntries As Long) _
    As Long

Declare Function RestoreDC _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nSavedDC As Long) _
    As Long

Declare Function RoundRect _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long) _
    As Long

Declare Function SaveDC _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function ScaleViewportExtEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nXnum As Long, ByVal nXdenom As Long, _
        ByVal nYnum As Long, ByVal nYdenom As Long, _
        lpSize As Size) _
    As Long

Declare Function ScaleWindowExtEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nXnum As Long, ByVal nXdenom As Long, _
        ByVal nYnum As Long, ByVal nYdenom As Long, _
        lpSize As Size) _
    As Long

Declare Function SelectPalette _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal hPalette As Long, _
        ByVal bForceBackground As Long) _
    As Long

Declare Function SetAbortProc _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal lpAbortProc As Long) _
    As Long

Declare Function SetArcDirection _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal ArcDirection As Long) _
    As Long

Declare Function SetBitmapBits _
    Lib "gdi32" ( _
        ByVal hBitmap As Long, _
        ByVal dwCount As Long, _
        lpBits As Any) _
    As Long

Declare Function SetBitmapDimensionEx _
    Lib "gdi32" ( _
        ByVal hbm As Long, _
        ByVal nX As Long, _
        ByVal nY As Long, _
        lpSize As Size) _
    As Long

Declare Function SetBkColor _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal crColor As Long) _
    As Long

Declare Function SetBkMode _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nBkMode As Long) _
    As Long

Declare Function SetBoundsRect _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lprcBounds As RECT, _
        ByVal flags As Long) _
    As Long

Declare Function SetBrushOrgEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nXOrg As Long, _
        ByVal nYOrg As Long, _
        lppt As POINTAPI) _
    As Long

Declare Function SetColorAdjustment _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpca As ColorAdjustment) _
    As Long

Declare Function SetColorSpace _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal hcolorspace As Long) _
    As Long

Declare Function SetDeviceGammaRamp _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpv As Any) _
    As Long

Declare Function SetDIBColorTable _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal un1 As Long, _
        ByVal un2 As Long, _
        pcRGBQuad As RGBQUAD) _
    As Long

Declare Function SetDIBits _
    Lib "gdi32" ( _
        ByVal hDC As Long, ByVal hBitmap As Long, _
        ByVal nStartScan As Long, ByVal nNumScans As Long, _
        lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) _
    As Long

Declare Function SetDIBitsToDevice _
    Lib "gdi32" ( _
        ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, _
        ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, _
        ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) _
    As Long

Declare Function SetEnhMetaFileBits _
    Lib "gdi32" ( _
        ByVal cbBuffer As Long, _
        lpData As Byte) _
    As Long

Declare Function SetMapMode _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nMapMode As Long) _
    As Long

Declare Function SetMapperFlags _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal dwFlag As Long) _
    As Long

Declare Function SetMetaFileBitsEx _
    Lib "gdi32" ( _
        ByVal nSize As Long, _
        lpData As Byte) _
    As Long

Declare Function SetMetaRgn _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function SetMiterLimit _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal eNewLimit As Double, _
        peOldLimit As Double) _
    As Long

Declare Function SetPaletteEntries _
    Lib "gdi32" ( _
        ByVal hPalette As Long, _
        ByVal wStartIndex As Long, _
        ByVal wNumEntries As Long, _
        lpPaletteEntries As PALETTEENTRY) _
    As Long

Declare Function SetPixel _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal crColor As Long) _
    As Long

Declare Function SetPixelFormat _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal n As Long, _
        pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) _
    As Boolean

Declare Function SetPixelV _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal crColor As Long) _
    As Long

Declare Function SetPolyFillMode _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nPolyFillMode As Long) _
    As Long

Declare Function SetRectRgn _
    Lib "gdi32" ( _
        ByVal hRgn As Long, _
        ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long) _
    As Long

Declare Function SetROP2 _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nDrawMode As Long) _
    As Long

Declare Function SetStretchBltMode _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nStretchMode As Long) _
    As Long

Declare Function SetSystemPaletteUse _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal wUsage As Long) _
    As Long

Declare Function SetTextAlign _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal wFlags As Long) _
    As Long

Declare Function SetTextCharacterExtra _
    Lib "gdi32" _
    Alias "SetTextCharacterExtraA" ( _
        ByVal hDC As Long, _
        ByVal nCharExtra As Long) _
    As Long

Declare Function SetTextColor _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal crColor As Long) _
    As Long

Declare Function SetTextJustification _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nBreakExtra As Long, _
        ByVal nBreakCount As Long) _
    As Long

Declare Function SetViewportExtEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nX As Long, _
        ByVal nY As Long, _
        lpSize As Size) _
    As Long

Declare Function SetViewportOrgEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nX As Long, _
        ByVal nY As Long, _
        lpPoint As POINTAPI) _
    As Long

Declare Function SetWindowExtEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nX As Long, _
        ByVal nY As Long, _
        lpSize As Size) _
    As Long

Declare Function SetWindowOrgEx _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nX As Long, _
        ByVal nY As Long, _
        lpPoint As POINTAPI) _
    As Long

Declare Function SetWinMetaFileBits _
    Lib "gdi32" ( _
        ByVal cbBuffer As Long, _
        lpbBuffer As Byte, _
        ByVal hdcRef As Long, _
        lpmfp As METAFILEPICT) _
    As Long

Declare Function SetWorldTransform _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        lpXform As XFORM) _
    As Long

Declare Function StartDoc _
    Lib "gdi32" _
    Alias "StartDocA" ( _
        ByVal hDC As Long, _
        lpdi As DOCINFO) _
    As Long

Declare Function StartPage _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function StretchBlt _
    Lib "gdi32" ( _
        ByVal hDC As Long, ByVal x As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, _
        ByVal nSrcHeight As Long, ByVal dwRop As Long) _
    As Long

Declare Function StretchDIBits _
    Lib "gdi32" ( _
        ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, _
        ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, _
        ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, _
        ByVal dwRop As Long) _
    As Long

Declare Function StrokeAndFillPath _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function StrokePath _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function SwapBuffers _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Boolean

Declare Function TextOut _
    Lib "gdi32" _
    Alias "TextOutA" ( _
        ByVal hDC As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal lpString As String, ByVal nCount As Long) _
    As Long

Declare Function TranslateCharsetInfo _
    Lib "gdi32" ( _
        lpSrc As Long, _
        lpcs As CHARSETINFO, _
        ByVal dwFlags As Long) _
    As Long

Declare Function UnrealizeObject _
    Lib "gdi32" ( _
        ByVal hObject As Long) _
    As Long

Declare Function UpdateColors _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long

Declare Function WidenPath _
    Lib "gdi32" ( _
        ByVal hDC As Long) _
    As Long



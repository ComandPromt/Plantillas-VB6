VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lib_SplitterBar"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
        
'-------------------------------------------------------------------------------------------
'-- Class:    Lib_SplitterBar
'-- Filename: cSplit.cls
'-- Author:   Gary Noble (gwnoble@msn.com)
'-- Date:     15 Jan 2002
'-------------------------------------------------------------------------------------------
'-- History
'-------------------------------------------------------------------------------------------
'--         15-1-2002  Initial Version  Gary Noble
'-------------------------------------------------------------------------------------------
'-- Notes:
'--         Thanks To Steve At VBAccelerator For His cSplitDC Class
'--
'-------------------------------------------------------------------------------------------



'-------------------------------------------------------------------------------------------
'-- some global declarations
'-------------------------------------------------------------------------------------------

Private bDraw       As Boolean
Private rcCurrent   As RECT
Private rcNew       As RECT
Private rcWindow    As RECT
Private rcWindows   As RECT

Private LastX       As Long
Private LastY       As Long

Dim hBrush          As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
   left             As Long
   top              As Long
   Right            As Long
   Bottom           As Long
End Type

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Const R2_BLACK = 1
Private Const R2_COPYPEN = 13
Private Const R2_LAST = 16
Private Const R2_MASKNOTPEN = 3
Private Const R2_MASKPEN = 9
Private Const R2_MASKPENNOT = 5
Private Const R2_MERGENOTPEN = 12
Private Const R2_MERGEPEN = 15
Private Const R2_MERGEPENNOT = 14
Private Const R2_NOP = 11
Private Const R2_NOT = 6
Private Const R2_NOTCOPYPEN = 4
Private Const R2_NOTMASKPEN = 8
Private Const R2_NOTMERGEPEN = 2
Private Const R2_NOTXORPEN = 7 '10
Private Const R2_WHITE = 16
Private Const R2_XORPEN = 7
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Sub ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As RECT)
Private Declare Sub ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CYCAPTION = 4
Private Const SM_CYMENU = 15


'-------------------------------------------------------------------------------------------
'-- COMBINEREGION
'-------------------------------------------------------------------------------------------
Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5
Private Const WM_NCACTIVATE = &H86
Private Const WM_MOVE = &H3
Private Const WM_CLOSE = &H10
Private Const HTCAPTION = 2
Private Const DFC_CAPTION = 1
Private Const DFCS_CAPTIONCLOSE = &H0
Private Const DFCS_FLAT = &H4000
Private Const DFCS_INACTIVE = &H100
Private Const DFCS_PUSHED = &H200
Private Const SPI_GETNONCLIENTMETRICS = 41

'-------------------------------------------------------------------------------------------
'-- Sizer constants
'-------------------------------------------------------------------------------------------
Private Const HTRIGHT = 11
Private Const HTLEFT = 10
Private Const HTTOP = 12
Private Const HTBOTTOM = 15
Private Const HTBOTTOMRIGHT = 17


'-------------------------------------------------------------------------------------------
'-- global declarations
'-------------------------------------------------------------------------------------------
Private Declare Function apiCombineRgn Lib "gdi32" Alias "CombineRgn" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function apiCreateBitmap Lib "gdi32" Alias "CreateBitmap" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function apiCreatePatternBrush Lib "gdi32" Alias "CreatePatternBrush" (ByVal hBitmap As Long) As Long
Private Declare Function apiCreateRectRgn Lib "gdi32" Alias "CreateRectRgn" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function apiCreateRectRgnIndirect Lib "gdi32" Alias "CreateRectRgnIndirect" (lpRect As RECT) As Long
Private Declare Function apiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function apiGetClipBox Lib "gdi32" Alias "GetClipBox" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function apiPatBlt Lib "gdi32" Alias "PatBlt" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function apiSelectClipRgn Lib "gdi32" Alias "SelectClipRgn" (ByVal hdc As Long, ByVal hRGN As Long) As Long
Private Declare Function apiSelectObject Lib "gdi32" Alias "SelectObject" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function apiSetRectRgn Lib "gdi32" Alias "SetRectRgn" (ByVal hRGN As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


'-------------------------------------------------------------------------------------------
'-- USER32
'-------------------------------------------------------------------------------------------
Private Declare Function apiCopyRect Lib "user32" Alias "CopyRect" (lpDstRect As RECT, lpSrcRect As RECT) As Long
Private Declare Function apiFillRect Lib "user32" Alias "FillRect" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function apiInflateRect Lib "user32" Alias "InflateRect" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function apiIntersectRect Lib "user32" Alias "IntersectRect" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function apiIsRectEmpty Lib "user32" Alias "IsRectEmpty" (lpRect As RECT) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetWindowPosAPI Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private m_hDC           As Long
Private m_hBmp          As Long
Private m_hBmpOld       As Long

'-------------------------------------------------------------------------------------------
'-- private members
'-------------------------------------------------------------------------------------------
Private m_lngPower2(0 To 31) As Long

'-------------------------------------------------------------------------------------------
'-- client metrics (for caption height)
'-------------------------------------------------------------------------------------------
Private m_bButtonDown       As Boolean
Private m_bButtonOver       As Boolean
Private m_bMouseDown        As Boolean
Private m_bRunTime          As Boolean
Private m_sCaption          As String
Private m_bDragging         As Boolean
Private m_lOldSize          As Long

Private m_hWnd              As Long
Private m_eOrientation      As eOrientationConstants
Private m_lBorder(1 To 4)   As Long
Private m_lSpecial(1 To 4)  As Long
Private m_oSplit            As Object

Private m_bUseInternalAjustments As Boolean
Private m_lSPBottom         As Long
Private m_lSPTop            As Long

Private m_bIsMDI            As Boolean
Private m_bSplitting        As Boolean

Public Enum eOrientationConstants
    espVertical = 1
    espHorizontal = 2
End Enum

Public Enum ESplitBorderTypes
   espbLeft = 1
   espbtop = 2
   espbright = 3
   espbBottom = 4
End Enum

Public Enum ESPValuesTypes
   espvleft = 1
   espvtop = 2
   espvright = 3
   espvBottom = 4
End Enum

Public Property Let UseInternalAjustments(ByVal vData As Boolean)
    m_bUseInternalAjustments = vData
End Property


Public Property Get UseInternalAjustments() As Boolean
    UseInternalAjustments = m_bUseInternalAjustments
End Property

Public Property Get SplitObject() As Object
    Set SplitObject = m_oSplit
End Property
Public Property Let SplitObject(ByRef oThis As Object)
    Set m_oSplit = oThis
    On Error Resume Next
    oThis.BorderStyle = 0
    If (m_eOrientation = espHorizontal) Then
        oThis.MousePointer = vbSizeNS
    Else
        oThis.MousePointer = vbSizeWE
    End If
End Property
Public Property Let Border(ByVal eBorderType As ESplitBorderTypes, ByVal lSize As Long)
   m_lBorder(eBorderType) = lSize
End Property
Public Property Get Border(ByVal eBorderType As ESplitBorderTypes) As Long
   Border = m_lBorder(eBorderType)
End Property

Public Property Let SPValues(ByVal eSPV As ESPValuesTypes, ByVal lSize As Long)
   m_lSpecial(eSPV) = lSize
End Property
Public Property Get SPValues(ByVal eSPV As ESPValuesTypes) As Long
   SPValues = m_lSpecial(eSPV)
End Property


Public Property Get Orientation() As eOrientationConstants
    Orientation = m_eOrientation
End Property
Public Property Let Orientation(ByVal eOrientation As eOrientationConstants)
    
    m_eOrientation = eOrientation
    
    If Not (m_oSplit Is Nothing) Then
        
        If (m_eOrientation = espHorizontal) Then
            
            m_oSplit.MousePointer = vbSizeNS
            
            m_lBorder(espbtop) = 64
            
            m_lBorder(espbBottom) = 64
            
            m_lBorder(espbLeft) = 0
            
            m_lBorder(espbright) = 0
        
        Else
            
            m_oSplit.MousePointer = vbSizeWE
            
            m_lBorder(espbtop) = 0
            
            m_lBorder(espbBottom) = 0
            
            m_lBorder(espbLeft) = 64
            
            m_lBorder(espbright) = 64
        
        End If
    
    End If

End Property

Public Sub SplitterMouseDown( _
        ByVal hWnd As Long, _
        ByVal x As Long, _
        ByVal y As Long _
    )
Dim tp As POINTAPI

    m_hWnd = hWnd

    
    Dim cd As Control
   
    
    '-------------------------------------------------------------------------------------------
    '-- Send subsequent mouse messages to the owner window
    '-------------------------------------------------------------------------------------------
    SetCapture m_hWnd
    
    '-------------------------------------------------------------------------------------------
    '-- Get the window rectangle on the desktop of the owner window
    '-------------------------------------------------------------------------------------------
    GetWindowRect m_hWnd, rcWindow
    
    '-------------------------------------------------------------------------------------------
    '-- Set The Splitter Limits
    '-------------------------------------------------------------------------------------------
        If (m_eOrientation = espVertical) Then
            
            rcWindow.left = rcWindow.left + m_lBorder(espbLeft) + 10
            
            rcWindow.Right = rcWindow.Right - m_lBorder(espbright)
        
        Else
            
            rcWindow.left = rcWindow.left + m_lBorder(espbLeft)
            
            rcWindow.Right = rcWindow.Right - m_lBorder(espbright)
            
            rcWindow.top = rcWindow.top + m_lBorder(espbtop)
            
            rcWindow.Bottom = rcWindow.Bottom - m_lBorder(espbBottom)
        
        End If
        
    
    
    '-------------------------------------------------------------------------------------------
    '-- Clip the cursor so it can't move outside the window
    '-------------------------------------------------------------------------------------------
    ClipCursorRect rcWindow
    
    '-------------------------------------------------------------------------------------------
    '-- Check if this is an MDI form
    '-------------------------------------------------------------------------------------------
        If (ClassName(m_hWnd) = "ThunderMDIForm") Then
        '-------------------------------------------------------------------------------------------
        ' Get the inside portion of the MDI form:
        ' Assuming we have a caption,menu and border in your MDI
        '-------------------------------------------------------------------------------------------
        rcWindow.left = rcWindow.left + GetSystemMetrics(SM_CXBORDER)
        
        rcWindow.Right = rcWindow.Right - GetSystemMetrics(SM_CXBORDER)
        
        rcWindow.Bottom = rcWindow.Bottom - GetSystemMetrics(SM_CYBORDER)
        
        rcWindow.top = rcWindow.top + GetSystemMetrics(SM_CYBORDER) * 3 + GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYMENU)
        
        m_bIsMDI = True
        
        tp.x = rcWindow.Right
            
        tp.y = rcWindow.Bottom
        
                
        If m_eOrientation = espVertical Then
            
            If SPValues(espvtop) > 0 Then
                rcWindow.top = rcWindow.top + SPValues(espvtop) / Screen.TwipsPerPixelY
            End If
            
            If SPValues(espvBottom) > 0 Then
                rcWindow.Bottom = rcWindow.Bottom - SPValues(espvBottom) / Screen.TwipsPerPixelY
            End If
            
        End If
      

    Else
    
    
        '-------------------------------------------------------------------------------------------
        '-- Get the client rectangle of the window in screen coordinates
        '-------------------------------------------------------------------------------------------
        GetClientRect m_hWnd, rcWindow
        
        If m_eOrientation = espVertical Then
            
            tp.x = rcWindow.left + m_lBorder(espbLeft)
            
            tp.y = rcWindow.top
        
        End If
        
        ClientToScreen m_hWnd, tp
        rcWindow.left = tp.x
        rcWindow.top = tp.y
        
        If m_eOrientation = espHorizontal Then
            
            tp.x = rcWindow.left
            tp.x = rcWindow.Right
            
            If SPValues(espvleft) > 0 Then
                rcWindow.left = rcWindow.left + SPValues(espvleft) / Screen.TwipsPerPixelX
            End If
          
            If SPValues(espvright) > 0 Then
                rcWindow.Right = SPValues(espvright) / Screen.TwipsPerPixelX ' rcWindow.Right - ((SPValues(espvright) / Screen.TwipsPerPixelX))  'SPValues(espvright) * Screen.TwipsPerPixelX))
                tp.x = rcWindow.Right
            End If
        
        Else
            
            tp.x = rcWindow.Right
            
            tp.y = rcWindow.Bottom
                        
               
        '-------------------------------------------------------------------------------------------
        '-- Adjust Rectangle Accordingly
        '-------------------------------------------------------------------------------------------
                If SPValues(espvBottom) > 0 Then
                    tp.y = rcWindow.Bottom - SPValues(espvBottom) / Screen.TwipsPerPixelY
                End If
            
                If SPValues(espvtop) > 0 Then
                    rcWindow.top = rcWindow.top + SPValues(espvtop) / Screen.TwipsPerPixelY
                End If
            
        End If
        
        ClientToScreen m_hWnd, tp
        
        '-------------------------------------------------------------------------------------------
        '-- Set Parameters
        '-------------------------------------------------------------------------------------------
        rcWindow.Right = tp.x
        rcWindow.Bottom = tp.y
        
        If m_eOrientation = espHorizontal Then
            
            tp.x = rcWindow.Right
            
            tp.y = rcWindow.Bottom
        
        Else
            
            tp.x = rcWindow.Right
            
            tp.y = rcWindow.Bottom
        
        End If

        
        m_bIsMDI = False
    
    End If
    
    bDraw = True  '-- start actual drawing from next move message
    
    rcCurrent.left = 0: rcCurrent.top = 0: rcCurrent.Right = 0: rcCurrent.Bottom = 0
    
    x = (m_oSplit.left + x) \ Screen.TwipsPerPixelX
    y = (m_oSplit.top + y) \ Screen.TwipsPerPixelY
    
    SplitterContainer_MouseMove x, y
    
End Sub

Public Function SplitterContainer_MouseMove( _
      ByVal x As Long, _
      ByVal y As Long)


    Dim hdc     As Long
    Dim tp      As POINTAPI
    Dim hBrush  As Long
    
    If (bDraw) Then
        '-----------------------------------------------------------------------
        '-- Draw two rectangles in the screen DC to cause splitting:
        '-- First get the Desktop DC:
        '-----------------------------------------------------------------------
        hdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        
        '-----------------------------------------------------------------------
        '-- Set the draw mode to XOR:
        '-----------------------------------------------------------------------
       ' SetROP2 hDC, R2_NOTXORPEN
        
        '-----------------------------------------------------------------------
        '-- Get The Cursor Position
        '-----------------------------------------------------------------------
        GetCursorPos tp
        
        '-----------------------------------------------------------------------
        '-- Determine where to draw the splitter:
        '-----------------------------------------------------------------------
        If m_bIsMDI Then
            If (m_eOrientation = espHorizontal) Then
                    
                    rcNew.left = rcWindow.left
                    
                    rcNew.Right = rcWindow.Right
                    
                    If (tp.y >= rcWindow.top) And (tp.y < rcWindow.Bottom) Then
                        
                        rcNew.top = tp.y - 4
                        
                        rcNew.Bottom = tp.y + 2
                    
                    Else
                        If (tp.y < rcWindow.top) Then
                            
                            rcNew.top = rcWindow.top - 10
                            
                            rcNew.Bottom = rcNew.top + 20
                        
                        Else
                            
                            rcNew.top = rcWindow.Bottom - 10
                            
                            rcNew.Bottom = rcNew.top + 20
                        
                        End If
                    
                    End If
                
                Else
                    
                    
                '-----------------------------------------------------------------------
                '-- Vertical
                '-----------------------------------------------------------------------
                    rcNew.top = rcWindow.top
                    rcNew.Bottom = rcWindow.Bottom
                    
                        
                    
                    If (tp.x >= rcWindow.left) And (tp.x <= rcWindow.Right) Then
                        
                        rcNew.left = tp.x '- 1
                        
                        rcNew.Right = tp.x + 5
                    
                    Else
                        
                        If (tp.x < rcWindow.left) Then
                            
                            rcNew.left = rcWindow.left '+ 5 ' 40
                            
                            rcNew.Right = rcNew.left + 5 ' + 5
                        
                        Else
                            
                            rcNew.left = rcWindow.Right '+ 5 ' 50
                            
                            rcNew.Right = rcNew.left '+ 5
                        
                        End If
                    
                    End If
                
                End If
        
        Else
        

        If (m_eOrientation = espHorizontal) Then
            
            rcNew.left = rcWindow.left
            
            rcNew.Right = rcWindow.Right
            
            If (tp.y >= rcWindow.top) And (tp.y < rcWindow.Bottom) Then
                
                rcNew.top = tp.y - 4
                
                rcNew.Bottom = tp.y + 2
            
            Else '(m_eOrientation = espVertical)
                
                If (tp.y < rcWindow.top) Then
                    
                    rcNew.top = rcWindow.top - 10
                    
                    rcNew.Bottom = rcNew.top + 20
                
                Else
                    
                    rcNew.top = rcWindow.Bottom - 10
                    
                    rcNew.Bottom = rcNew.top + 20
                
                End If
            
            End If
        
        Else
            
            
        '-----------------------------------------------------------------------
        '-- Vertical
        '-----------------------------------------------------------------------
            rcNew.top = rcWindow.top
            rcNew.Bottom = rcWindow.Bottom
            
                
            
            If (tp.x >= rcWindow.left) And (tp.x <= rcWindow.Right) Then
                
                rcNew.left = tp.x - 1
                
                rcNew.Right = tp.x + 5
            
            Else
                
                If (tp.x < rcWindow.left) Then
                    
                    rcNew.left = rcWindow.left + 5 ' 40
                    
                    rcNew.Right = rcNew.left + 5
                
                Else
                    
                    rcNew.left = rcWindow.Right + 5 ' 50
                    
                    rcNew.Right = rcNew.left + 5
                
                End If
            
            End If
        
        End If
        
        End If
   
DrawSplitter:

        '-----------------------------------------------------------------------
        '-- Draw the new rectangle
        '-----------------------------------------------------------------------
        hBrush = gdiCreateHalftoneBrush()
        gdiDrawDragRectangle hdc, rcNew.left, rcNew.top, rcNew.Right - rcNew.left, rcNew.Bottom - rcNew.top, rcCurrent.left, rcCurrent.top, rcCurrent.Right - rcCurrent.left, rcCurrent.Bottom - rcCurrent.top, 15, 15, hBrush, 0
    
        '-----------------------------------------------------------------------
        '-- Store this position so we can erase it next time:
        '-----------------------------------------------------------------------
        LSet rcCurrent = rcNew
        
        '-----------------------------------------------------------------------
        '-- Free the reference to the Desktop DC we got (make sure you do this!)
        '-----------------------------------------------------------------------
        DeleteDC hdc
        
    
    End If


End Function

Public Function SplitterContainer_MouseUp(ByVal x As Long, ByVal y As Long) As Boolean


Dim hdc         As Long
Dim tp          As POINTAPI
Dim hWndClient  As Long
Dim hBrush      As Long

    '-------------------------------------------------------------------------------------------
    '-- Don't leave orphaned rectangle on desktop; erase last rectangle.
    '-------------------------------------------------------------------------------------------
    If (bDraw) Then
       bDraw = False
       
       '-------------------------------------------------------------------------------------------
       '-- Release mouse capture
       '-------------------------------------------------------------------------------------------
       ReleaseCapture
       
       '-------------------------------------------------------------------------------------------
       '-- Release the cursor clipping region (must do this!)
       '-------------------------------------------------------------------------------------------
       ClipCursorClear 0&
       
       '-------------------------------------------------------------------------------------------
       '-- Get the Desktop DC
       '-------------------------------------------------------------------------------------------
       hdc = CreateDCAsNull("DISPLAY", 0, 0, 0)
       
       '-------------------------------------------------------------------------------------------
       '-- Set to XOR drawing mode
       '-- SetROP2 hDC, R2_NOTXORPEN
       '-- Erase the last rectangle
       '-------------------------------------------------------------------------------------------
       
       '-----------------------------------------------------------------------
       '-- Modification: New Splitter Implementation (01-05-2002 G Noble)
       '-- Delete The Last Splitter Bar
       '-----------------------------------------------------------------------
       hBrush = gdiCreateHalftoneBrush()
       gdiDrawDragRectangle hdc, rcCurrent.left, rcCurrent.top, rcCurrent.Right - rcCurrent.left, rcCurrent.Bottom - rcCurrent.top, 0, 0, 0, 0, 5, 5, hBrush, 0
       
       '-------------------------------------------------------------------------------------------
       '-- Clear up the desktop DC
       '-------------------------------------------------------------------------------------------
       DeleteDC hdc
       
       '-------------------------------------------------------------------------------------------
       '-- Here we ensure the splitter is within bounds before releasing
       '-------------------------------------------------------------------------------------------
       GetCursorPos tp

       ScreenToClient m_hWnd, tp
       
       '-------------------------------------------------------------------------------------------
       '-- Move the splitter to the validated final position
       '-------------------------------------------------------------------------------------------
       If (m_eOrientation = espHorizontal) Then
           
           m_oSplit.top = (tp.y - 4) * Screen.TwipsPerPixelY
           
           Else
           
           m_oSplit.left = (tp.x) * Screen.TwipsPerPixelX
       
       End If
       
       '-------------------------------------------------------------------------------------------
       '-- Return true to tell the owner we have completed splitting
       '-------------------------------------------------------------------------------------------
       SplitterContainer_MouseUp = True
   
   End If
 
  RefreshALLControls
    
    
End Function

Private Function RefreshALLControls()
On Error Resume Next

    Dim ctrl As Control
        
        For Each ctrl In Me.SplitObject.Parent.Controls
            ctrl.Refresh
        Next
        
        
        
End Function

Private Sub Class_Initialize()
   
   m_eOrientation = espVertical
   
   m_lBorder(espbLeft) = 64
   
   m_lBorder(espbright) = 64
           
End Sub
Private Function ClassName(ByVal lHwnd As Long) As String

Dim lLen As Long
Dim sBuf As String
    
    lLen = 260
    
    sBuf = String$(lLen, 0)
    
    lLen = GetClassName(lHwnd, sBuf, lLen)
    
    If (lLen <> 0) Then
        ClassName = left$(sBuf, lLen)
    End If

End Function





Private Sub gdiDrawDragRectangle(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal DY As Long, ByVal LastX As Long, ByVal LastY As Long, ByVal LastDX As Long, ByVal LastDY As Long, ByVal Size As Long, ByVal LastSize As Long, ByVal hBrush As Long, ByVal hBrushLast As Long, Optional ByVal hClipRgn As Long)

'-----------------------------------------------------------------------
'-- draw drag rectangle, erasing old as needed
'-----------------------------------------------------------------------

    Dim tRect       As RECT, tRectLast  As RECT
    Dim hOutsideRgn As Long, hInsideRgn As Long
    Dim hLastRgn    As Long, hUpdateRgn As Long
    Dim hNewRgn     As Long
    Dim tNewRect    As RECT
    Dim hOldBrush   As Long
    
'    Debug.Assert hBrush <> 0
    
    '-----------------------------------------------------------------------
    '-- init vars
    '-----------------------------------------------------------------------
    With tRect
        .left = x
        .top = y
        .Right = x + DX
        .Bottom = y + DY
        
    End With
    With tRectLast
        .left = LastX
        .top = LastY
        .Right = LastX + LastDX
        .Bottom = LastY + LastDY
    
    End With
    
    '-----------------------------------------------------------------------
    '-- determine update region and select it
    '-----------------------------------------------------------------------
    hOutsideRgn = apiCreateRectRgnIndirect(tRect)
    apiCopyRect tNewRect, tRect
    apiInflateRect tNewRect, -Size, -Size
    apiIntersectRect tNewRect, tNewRect, tRect
    hInsideRgn = apiCreateRectRgnIndirect(tNewRect)
    hNewRgn = apiCreateRectRgn(0, 0, 0, 0)
    apiCombineRgn hNewRgn, hOutsideRgn, hInsideRgn, RGN_XOR
    If hBrushLast = 0 Then
        hBrushLast = hBrush
        
    End If
    If apiIsRectEmpty(tRectLast) = 0 Then
        
        '-----------------------------------------------------------------------
        '-- find difference between new region and old region
        '-----------------------------------------------------------------------
        hLastRgn = apiCreateRectRgn(0, 0, 0, 0)
        apiSetRectRgn hOutsideRgn, tRectLast.left, tRectLast.top, tRectLast.Right, tRectLast.Bottom
        apiCopyRect tNewRect, tRectLast
        apiInflateRect tNewRect, -LastSize, -LastSize
        apiIntersectRect tNewRect, tNewRect, tRectLast
        apiSetRectRgn hInsideRgn, tNewRect.left, tNewRect.top, tNewRect.Right, tNewRect.Bottom
        apiCombineRgn hLastRgn, hOutsideRgn, hInsideRgn, RGN_XOR
        
        '-----------------------------------------------------------------------
        '-- only diff them if brushes are the same
        '-----------------------------------------------------------------------
        If (hBrush = hBrushLast) Then
            hUpdateRgn = apiCreateRectRgn(0, 0, 0, 0)
            apiCombineRgn hUpdateRgn, hLastRgn, hNewRgn, RGN_XOR
            
        End If
    
    End If
    If (hBrush <> hBrushLast) And (apiIsRectEmpty(tRectLast) = 0) Then
        '-----------------------------------------------------------------------
        '-- brushes are different -- erase old region first
        '-----------------------------------------------------------------------
        apiSelectClipRgn hdc, hLastRgn
        apiGetClipBox hdc, tNewRect
        hOldBrush = apiSelectObject(hdc, hBrushLast)
        apiPatBlt hdc, tNewRect.left, tNewRect.top, tNewRect.Right - tNewRect.left, tNewRect.Bottom - tNewRect.top, vbPatInvert
        apiSelectObject hdc, hOldBrush
        apiDeleteObject hOldBrush
        
    End If
    
    '-----------------------------------------------------------------------
    '-- draw into the update/new region
    '-----------------------------------------------------------------------
    If hUpdateRgn <> 0 Then
        If hClipRgn <> 0 Then
            apiCombineRgn hUpdateRgn, hClipRgn, hUpdateRgn, RGN_AND
        
        End If
        apiSelectClipRgn hdc, hUpdateRgn
        
    Else
        If hClipRgn <> 0 Then
            apiCombineRgn hNewRgn, hClipRgn, hNewRgn, RGN_AND
        
        End If
        apiSelectClipRgn hdc, hNewRgn
    
    End If
    apiGetClipBox hdc, tNewRect
    hOldBrush = apiSelectObject(hdc, hBrush)
    apiPatBlt hdc, tNewRect.left, tNewRect.top, tNewRect.Right - tNewRect.left, tNewRect.Bottom - tNewRect.top, vbPatInvert

    '-----------------------------------------------------------------------
    '-- cleanup DC
    '-----------------------------------------------------------------------
    If (hOldBrush <> 0) Then
        apiSelectObject hdc, hOldBrush
    End If
    apiSelectClipRgn hdc, 0&
    
    '-----------------------------------------------------------------------
    '-- delete objects (A MUST)
    '-----------------------------------------------------------------------
    If hUpdateRgn <> 0 Then
        apiDeleteObject hUpdateRgn
    End If
    
    If hLastRgn <> 0 Then
        apiDeleteObject hLastRgn
        
    End If
    If hNewRgn <> 0 Then
        apiDeleteObject hNewRgn
        
    End If
    If hInsideRgn <> 0 Then
        apiDeleteObject hInsideRgn
        
    End If
    If hOutsideRgn <> 0 Then
        apiDeleteObject hOutsideRgn
    End If

    '-- >>
    If hBrush <> 0 Then
        apiDeleteObject hBrush
    End If
    
    If hOldBrush <> 0 Then
        apiDeleteObject hOldBrush
    End If
    
    
End Sub
Private Sub gdiDrawSolidRectangleBrush(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal DY As Long, ByVal hBrush As Long)

'-----------------------------------------------------------------------
'-- draw filled rectangle using supplied brush
'-----------------------------------------------------------------------
    Dim tRect As RECT
    
    '-----------------------------------------------------------------------
    '-- build RECT
    '-----------------------------------------------------------------------
    gdiSetRect tRect, x, y, DX, DY
    
    '-----------------------------------------------------------------------
    '-- draw filled rect
    '-----------------------------------------------------------------------
    apiFillRect hdc, tRect, hBrush
    
End Sub
Private Sub gdiSetRect(tRect As RECT, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal DY As Long)
    
    '-----------------------------------------------------------------------
    '-- build RECT
    '-----------------------------------------------------------------------
    With tRect
        .left = x
        .top = y
        .Right = x + DX
        .Bottom = y + DY
        
    End With
End Sub
Public Sub gdiDeleteBrush(ByVal hBrush As Long)

'-----------------------------------------------------------------------
'-- delete brush object
'-----------------------------------------------------------------------
    apiDeleteObject hBrush

End Sub
Private Function gdiCreateHalftoneBrush() As Long

'-----------------------------------------------------------------------
'-- Create a halftoned brush object
'-----------------------------------------------------------------------

    Dim nGreyPattern(8) As Integer
    Dim nBitCount       As Integer
    Dim hBitmap         As Long
    Dim lngResult       As Long
    
    '-----------------------------------------------------------------------
    '-- create greyed bitmap
    '-----------------------------------------------------------------------
    For nBitCount = 0 To 7
        nGreyPattern(nBitCount) = LShiftWord(&H5555, (nBitCount And 1))
        
    Next nBitCount
    hBitmap = apiCreateBitmap(8, 8, 1, 1, nGreyPattern(0))
    
    '-----------------------------------------------------------------------
    '-- create halftone pattern brush
    '-----------------------------------------------------------------------
    gdiCreateHalftoneBrush = apiCreatePatternBrush(hBitmap)
    
    '-----------------------------------------------------------------------
    '-- delete bitmap
    '-----------------------------------------------------------------------
    apiDeleteObject hBitmap
    '->> apiDeleteObject gdiCreateHalftoneBrush
    
    
End Function
Private Function LShiftWord(ByVal nWord As Integer, ByVal nBitCount As Integer) As Integer

'-----------------------------------------------------------------------
'-- left shift dword value lngDWord by nBitCount
'-----------------------------------------------------------------------
    Dim lngDWord As Long

    Debug.Assert (nBitCount >= 0 And nBitCount <= 15)   'design-time only
    If Not (nBitCount >= 0 And nBitCount <= 15) Then Exit Function
    
    lngDWord = nWord * Power2(nBitCount)
    If lngDWord And &H8000& Then
        LShiftWord = CInt(lngDWord And &H7FFF&) Or &H8000
        
    Else
        LShiftWord = lngDWord And &HFFFF&
        
    End If
End Function
Private Function Power2(ByVal nIndex As Integer) As Long

'-----------------------------------------------------------------------
'-- build lookup table for bitshifting operations
'-----------------------------------------------------------------------
    If Not (nIndex >= 0 And nIndex <= 31) Then Exit Function
    
    If m_lngPower2(0) = 0 Then 'only set array if not yet been used
        m_lngPower2(0) = &H1&
        m_lngPower2(1) = &H2&
        m_lngPower2(2) = &H4&
        m_lngPower2(3) = &H8&
        m_lngPower2(4) = &H10&
        m_lngPower2(5) = &H20&
        m_lngPower2(6) = &H40&
        m_lngPower2(7) = &H80&
        m_lngPower2(8) = &H100&
        m_lngPower2(9) = &H200&
        m_lngPower2(10) = &H400&
        m_lngPower2(11) = &H800&
        m_lngPower2(12) = &H1000&
        m_lngPower2(13) = &H2000&
        m_lngPower2(14) = &H4000&
        m_lngPower2(15) = &H8000&
        m_lngPower2(16) = &H10000
        m_lngPower2(17) = &H20000
        m_lngPower2(18) = &H40000
        m_lngPower2(19) = &H80000
        m_lngPower2(20) = &H100000
        m_lngPower2(21) = &H200000
        m_lngPower2(22) = &H400000
        m_lngPower2(23) = &H800000
        m_lngPower2(24) = &H1000000
        m_lngPower2(25) = &H2000000
        m_lngPower2(26) = &H4000000
        m_lngPower2(27) = &H8000000
        m_lngPower2(28) = &H10000000
        m_lngPower2(29) = &H20000000
        m_lngPower2(30) = &H40000000
        m_lngPower2(31) = &H80000000
        
    End If
    Power2 = m_lngPower2(nIndex)
End Function




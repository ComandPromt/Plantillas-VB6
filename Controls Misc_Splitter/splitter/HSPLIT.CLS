VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Internal variables for forms and controls
Private ctlTop As Control
Private ctlBottom As Control
Private objContainer As Object

' Sizes of borders and pixels
Private ySplit As Single
Private dySplit As Single
Private xPixel As Single
Private yPixel As Single
Private dxBorder As Single
Private dyBorder As Single

' Flags
Private fResize As Boolean
Private fAutoBorder As Boolean
Private fDragging As Boolean
Private fDragIcon As Boolean
Private fCreated As Boolean

' Old mouse pointer, draw style, and draw mode
Private mpOld As Integer
Private dsOld As Integer
Private dmOld As Integer

' AutoRedraw
Private arOld As Boolean

' Create a splitter window
Function Create(vTopControl As Control, vBottomControl As Control, _
                Optional vBorderPixels As Variant, _
                Optional vAutoBorder As Variant, _
                Optional vResizeable As Variant) As Boolean
                
    Create = True
    fCreated = False
    On Error GoTo CreateError
    ' Set internal controls
    Set ctlTop = vTopControl
    Set ctlBottom = vBottomControl
    Set objContainer = ctlTop.Container
    objContainer.AutoRedraw = True
    If objContainer.ClipControls Then GoTo CreateError
    
    ' Save resizable and AutoBorder flags
    If IsMissing(vResizeable) Then vResizeable = True
    fResize = vResizeable
    If IsMissing(vAutoBorder) Then vAutoBorder = True
    fAutoBorder = vAutoBorder
    
    ' Size of one in pixel in current scale
    xPixel = objContainer.ScaleX(1, vbPixels, objContainer.ScaleMode)
    yPixel = objContainer.ScaleY(1, vbPixels, objContainer.ScaleMode)
    ' Set splitter size
    If IsMissing(vBorderPixels) Then
        fAutoBorder = True
        vBorderPixels = 4
    End If
    dySplit = vBorderPixels * yPixel
    ' Set border size
    If fAutoBorder Then
        dxBorder = ctlTop.Left
        dyBorder = ctlTop.Top
    Else
        dxBorder = vBorderPixels * xPixel
        dyBorder = vBorderPixels * yPixel
    End If

    ' Size the controls
    If ctlBottom.Top < ctlTop.Top Then GoTo CreateError
    If yBottom(ctlBottom) < yBottom(ctlTop) Then GoTo CreateError
    Resize
    fCreated = True
    Exit Function
    
CreateError:
    Create = False
End Function

Sub Resize()

    ' Move everything in border size from the edge
    ctlTop.Left = dxBorder
    ctlTop.Top = objContainer.ScaleTop + dyBorder
    ctlTop.Width = objContainer.ScaleWidth - (2 * dxBorder)
    ' ctlTop.Height ' Unchanged
        
    ctlBottom.Left = dxBorder
    ctlBottom.Top = yBottom(ctlTop) + dySplit
    ctlBottom.Width = ctlTop.Width
    ctlBottom.Height = objContainer.ScaleHeight - ctlBottom.Top - dyBorder

End Sub

Sub HSplitter_MouseMove(Button As Integer, Shift As Integer, _
                        X As Single, Y As Single)
With objContainer
    If Not fCreated Then Exit Sub
    Dim yPos As Single
    ' Change the cursor to splitter or back
    If Y <= ctlBottom.Top And Y >= yBottom(ctlTop) Then
        If .MousePointer <> 99 And .MousePointer <> vbSizeNS Then
            mpOld = .MousePointer
            If .MouseIcon.Type <> vbPicTypeIcon Then
                .MousePointer = vbSizeNS
            Else
                .MousePointer = 99
            End If
        End If
    Else
        If (.MousePointer = 99 Or .MousePointer = vbSizeNS) _
           And Button <> vbLeftButton Then
            .MousePointer = mpOld
        End If
    End If
    
    ' Move the splitter line if within range
    If fDragging And (ySplit <> Y) And _
       (Y > (yPixel * 20)) And (Y < (.ScaleHeight - (yPixel * 40))) Then
        .DrawStyle = vbInsideSolid
        .DrawMode = vbInvert
        yPos = ySplit
        ' Erase old line
        objContainer.Line (ctlTop.Left, yPos - yPixel)-(ctlTop.Width, yPos + yPixel), , B
        ' Draw new line
        yPos = Y
        objContainer.Line (ctlTop.Left, yPos - yPixel)-(ctlTop.Width, yPos + yPixel), , B
        ySplit = yPos
    End If
End With
End Sub

' Put in MouseMove of the contained controls
Sub HSplitter_MouseOff()
With objContainer
    If Not fCreated Then Exit Sub
    If .MousePointer = 99 Or .MousePointer = vbSizeWE Then .MousePointer = mpOld
End With
End Sub

Sub HSplitter_MouseDown(Button As Integer, Shift As Integer, _
                        X As Single, Y As Single)
With objContainer
    If Not fCreated Then Exit Sub
    Dim yPos As Single
    yPos = yBottom(ctlTop)
    ' If over splitter start a drag
    If (yPos < Y) And (Y < ctlBottom.Top) Then
        If (Button = vbLeftButton) And (yPos < Y) And (Y < ctlBottom.Top) Then
            ' Save and restore state
            fDragging = True
            dsOld = .DrawStyle
            dmOld = .DrawMode
            arOld = .AutoRedraw
            .DrawStyle = vbInsideSolid
            .DrawMode = vbInvert
            .AutoRedraw = False
            ' Draw the splitter line and save position
            yPos = yPos + (dxBorder / 3)
            objContainer.Line (ctlTop.Left, yPos - yPixel)-(ctlTop.Width, yPos + yPixel), , B
            ySplit = yPos
        End If
    Else
        If .MousePointer = 99 Or .MousePointer = vbSizeWE Then .MousePointer = mpOld
    End If
End With
End Sub

Sub HSplitter_MouseUp(Button As Integer, Shift As Integer, _
                      X As Single, Y As Single)
With objContainer
    Dim yPos As Single
    If Not fCreated Then Exit Sub
    If fDragging Then
        ' Erase old line
        .DrawStyle = vbInsideSolid
        .DrawMode = vbInvert
        yPos = ySplit
        objContainer.Line (ctlTop.Left, yPos - yPixel)-(ctlTop.Width, yPos + yPixel), , B
        .DrawStyle = dsOld
        .DrawMode = dmOld
        fDragging = False
        ' Resize the panes if in range
        If Y > (yPixel * 20) And Y < (.ScaleHeight - (yPixel * 20)) Then
            ctlTop.Height = Y - ctlTop.Top - (dySplit / 2)
            ctlBottom.Top = yBottom(ctlTop) + dySplit
            ctlBottom.Height = .ScaleHeight - ctlBottom.Top - dyBorder
        End If
        .DrawStyle = dsOld
        .DrawMode = dmOld
        .AutoRedraw = arOld
    End If
End With
End Sub

Sub HSplitter_Resize()
    If objContainer Is Nothing Then Exit Sub
    If Not fCreated Then Exit Sub
    ' Only forms have WindowState
    On Error Resume Next
    If objContainer.WindowState <> vbMinimized And fResize Then Resize
    ' Must not be form
    If Err And fResize Then Resize
End Sub

Private Function xRight(obj As Object) As Single
    xRight = obj.Left + obj.Width
End Function

Private Function yBottom(obj As Object) As Single
    yBottom = obj.Top + obj.Height
End Function

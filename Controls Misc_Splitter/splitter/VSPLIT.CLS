VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Internal variables for forms and controls
Private ctlLeft As Control
Private ctlRight As Control
Private objContainer As Object

' Sizes of borders and pixels
Private xSplit As Single
Private dxSplit As Single
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
Function Create(vLeftControl As Control, vRightControl As Control, _
                Optional vBorderPixels As Variant, _
                Optional vAutoBorder As Variant, _
                Optional vResizeable As Variant) As Boolean
                
    Create = True
    fCreated = False
    On Error GoTo CreateError
    ' Set internal controls
    Set ctlLeft = vLeftControl
    Set ctlRight = vRightControl
    Set objContainer = ctlLeft.Container
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
    dxSplit = vBorderPixels * xPixel
    ' Set border size
    If fAutoBorder Then
        dxBorder = ctlLeft.Left
        dyBorder = ctlLeft.Top
    Else
        dxBorder = vBorderPixels * xPixel
        dyBorder = vBorderPixels * yPixel
    End If

    ' Size the controls
    If ctlRight.Left < ctlLeft.Left Then GoTo CreateError
    If xRight(ctlRight) < xRight(ctlLeft) Then GoTo CreateError
    Resize
    fCreated = True
    Exit Function
    
CreateError:
    Create = False
End Function

Sub Resize()

    ' Move everything in border size from the edge
    ctlLeft.Left = dxBorder
    ctlLeft.Top = objContainer.ScaleTop + dyBorder
    ' ctlLeft.Width ' Unchanged
    ctlLeft.Height = objContainer.ScaleHeight - (2 * dyBorder)
    
    ctlRight.Left = xRight(ctlLeft) + dxSplit
    ctlRight.Top = dyBorder
    ctlRight.Width = objContainer.ScaleWidth - ctlRight.Left - dxBorder
    ctlRight.Height = ctlLeft.Height

End Sub

Sub VSplitter_MouseMove(Button As Integer, Shift As Integer, _
                        X As Single, Y As Single)
With objContainer
    If Not fCreated Then Exit Sub
    Dim xPos As Single
    ' Change the cursor to splitter or back
    If X <= ctlRight.Left And X >= xRight(ctlLeft) Then
        If .MousePointer <> 99 And .MousePointer <> vbSizeWE Then
            mpOld = .MousePointer
            If .MouseIcon.Type <> vbPicTypeIcon Then
                .MousePointer = vbSizeWE
            Else
                .MousePointer = 99
            End If
        End If
    Else
        If (.MousePointer = 99 Or .MousePointer = vbSizeWE) _
           And Button <> vbLeftButton Then
            .MousePointer = mpOld
        End If
    End If
    
    ' Move the splitter line if within range
    If fDragging And (xSplit <> X) And _
       (X > (xPixel * 20)) And (X < (.ScaleWidth - (xPixel * 40))) Then
        .DrawStyle = vbInsideSolid
        .DrawMode = vbInvert
        xPos = xSplit
        ' Erase old line
        objContainer.Line (xPos - xPixel, ctlLeft.Top)-(xPos + xPixel, yBottom(ctlLeft)), , B
        ' Draw new line
        xPos = X
        objContainer.Line (xPos - xPixel, ctlLeft.Top)-(xPos + xPixel, yBottom(ctlLeft)), , B
        xSplit = xPos
    End If
End With
End Sub

' Put in MouseMove of the contained controls
Sub VSplitter_MouseOff()
With objContainer
    If Not fCreated Then Exit Sub
    If .MousePointer = 99 Or .MousePointer = vbSizeWE Then .MousePointer = mpOld
End With
End Sub

Sub VSplitter_MouseDown(Button As Integer, Shift As Integer, _
                        X As Single, Y As Single)
With objContainer
    If Not fCreated Then Exit Sub
    Dim xPos As Single
    xPos = xRight(ctlLeft)
    ' If over splitter start a drag
    If (xPos < X) And (X < ctlRight.Left) Then
        If Button = vbLeftButton Then
            ' Save and restore state
            fDragging = True
            dsOld = .DrawStyle
            dmOld = .DrawMode
            arOld = .AutoRedraw
            .DrawStyle = vbInsideSolid
            .DrawMode = vbInvert
            .AutoRedraw = False
            ' Draw the splitter line and save position
            xPos = xPos + (dxBorder / 3)
            objContainer.Line (xPos - xPixel, ctlLeft.Top)-(xPos + xPixel, yBottom(ctlLeft)), , B
            xSplit = xPos
        End If
    Else
        If .MousePointer = 99 Or .MousePointer = vbSizeWE Then .MousePointer = mpOld
    End If
End With
End Sub

Sub VSplitter_MouseUp(Button As Integer, Shift As Integer, _
                     X As Single, Y As Single)
With objContainer
    If Not fCreated Then Exit Sub
    Dim xPos As Single
    If fDragging Then
        ' Erase old line
        .DrawStyle = vbInsideSolid
        .DrawMode = vbInvert
        xPos = xSplit
        objContainer.Line (xPos - xPixel, ctlLeft.Top)-(xPos + xPixel, yBottom(ctlLeft)), , B
        .DrawStyle = dsOld
        .DrawMode = dmOld
        fDragging = False
        ' Resize the panes if in range
        If X > (xPixel * 20) And X < (.ScaleWidth - (xPixel * 20)) Then
            ctlLeft.Width = X - ctlLeft.Left - (dxSplit / 2)
            ctlRight.Left = xRight(ctlLeft) + dxSplit
            ctlRight.Width = .ScaleWidth - ctlRight.Left - dxBorder
        End If
        .DrawStyle = dsOld
        .DrawMode = dmOld
        .AutoRedraw = arOld
    End If
End With
End Sub

Sub VSplitter_Resize()
    If objContainer Is Nothing Then Exit Sub
    If Not fCreated Then Exit Sub
    On Error Resume Next
    ' Only forms have WindowState
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



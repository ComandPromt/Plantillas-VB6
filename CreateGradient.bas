Attribute VB_Name = "Module1"
Type COLORREF
    lRed As Single
    lGreen As Single
    lBlue As Single
End Type

Public Function CreateGradient(ctrl As Object, startcolor As COLORREF, endcolor As COLORREF)
Dim ct As Long ' standard counter
Dim lDiff As Single ' The difference between the top and bottom.
Dim lClr As Long ' Placeholder for color
Dim lTop As Long
Dim lBottom As Long

On Error GoTo errhandler
Dim h As Long
h = ctrl.hWnd

lDiff = ctrl.Height / 63
lTop = 0
lBottom = lTop + lDiff

For ct = 0 To ctrl.Height Step lDiff
    lClr = RGB(endcolor.lRed, endcolor.lGreen, endcolor.lBlue)
    ctrl.Line (0, lTop)-(ctrl.Width, lBottom), lClr, BF
    endcolor.lRed = iif(endcolor.lRed - 4 < 0, 0, endcolor.lRed - 4)
    endcolor.lGreen = iif(endcolor.lGreen - 4 < 0, 0, endcolor.lGreen - 4)
    endcolor.lBlue = iif(endcolor.lBlue - 4 < 0, 0, endcolor.lBlue - 4)
    lTop = lTop + lDiff
    lBottom = lBottom + lDiff
Next ct
On Error GoTo 0
Exit Function
errhandler:
On Error GoTo 0
End Function

Public Function iif(boolCondition As Boolean, varTrue As Variant, varFalse As Variant) As Variant
    'Replaces IIF, to eliminate the need for the DLL
    If boolCondition Then
        iif = varTrue
    Else
        iif = varFalse
    End If
End Function

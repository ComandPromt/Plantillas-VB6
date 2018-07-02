Attribute VB_Name = "Module2"
'*********************************************************************
' TABS.BAS - Creates a tabbed dialog effect for a form.
'---------------------------------------------------------------------
' USAGE:    1. Set AutoRedraw = True on the destination form.
'           2. Create a control array of Frames named fratabs.
'           3. Label each frame in Tabs with an appropriate caption.
'           4. From your forms Form_Load() event, call SetupTabs.
'           5. Paste "fratabs(DrawTabs(Me, X, Y) - 1).ZOrder" into the
'              forms Form_MouseUp event.
'*********************************************************************
Option Explicit
Dim strTabLabels() As String

'*********************************************************************
' DrawTabs - Draws fratabs on a form that look like Word & Excel's
'---------------------------------------------------------------------
' FormName      (Form)      Name of the form to draw the tabs
' Tabs()        (String)    Array of names for the tabs
' sngXPos, sngYPos    (Single)    Point clicked on the form
'*********************************************************************
Function DrawTabs(FormName As Form, ByVal sngXPos As Single, ByVal sngYPos As Single) As Integer
 
Dim intNumTabs As Integer
Dim intTabWidth As Integer
Dim intLpctr As Integer
Dim intActiveTab As Integer
Dim sngX As Single
Dim sngX1 As Single

Const TABHEIGHT = 18
Const OFFSET = 4
    '*****************************************************************
    ' The form's ScaleMode MUST be in pixels, or else...
    '*****************************************************************
    FormName.PicTabContainer.ScaleMode = 3
    '*****************************************************************
    ' Only respond to clicks within a tab
    '*****************************************************************
    If sngYPos < OFFSET Or sngYPos > OFFSET + TABHEIGHT Then Exit Function
    '*****************************************************************
    ' Cache the upper index of Tabs
    '*****************************************************************
    intNumTabs = UBound(strTabLabels)
    '*****************************************************************
    ' Setup the width of the tabs
    '*****************************************************************
    intTabWidth = (FormName.PicTabContainer.ScaleWidth - 2) / intNumTabs
    '*****************************************************************
    ' Draw a block over any existing tabs
    '*****************************************************************
    FormName.PicTabContainer.Line (0, 0)-(Screen.Width - 2, TABHEIGHT + OFFSET + 1), FormName.PicTabContainer.BackColor, BF
    '*****************************************************************
    ' Draw a black border around the tabs
    '*****************************************************************
    For intLpctr = 1 To intNumTabs
        FormName.PicTabContainer.Line (sngX, TABHEIGHT + OFFSET)-(sngX, 4 + OFFSET), 0
        FormName.PicTabContainer.Line (sngX, 4 + OFFSET)-(sngX + 4, 0 + OFFSET), 0
        FormName.PicTabContainer.Line (sngX + 4, 0 + OFFSET)-(sngX + intTabWidth - 4, 0 + OFFSET), 0
        FormName.PicTabContainer.Line (sngX + intTabWidth - 4, 0 + OFFSET)-(sngX + intTabWidth, 4 + OFFSET), 0
        FormName.PicTabContainer.Line (sngX + intTabWidth, 4 + OFFSET)-(sngX + intTabWidth, TABHEIGHT + OFFSET + 2), 0
        sngX = sngX + intTabWidth
    Next intLpctr
    '*****************************************************************
    ' Draw a black border around the form
    '*****************************************************************
    FormName.PicTabContainer.Line (0, TABHEIGHT + OFFSET)-(0, FormName.ScaleHeight - 1), 0
    FormName.PicTabContainer.Line (0, FormName.PicTabContainer.ScaleHeight - 1)-((intTabWidth * intNumTabs), FormName.PicTabContainer.ScaleHeight - 1), 0
    FormName.PicTabContainer.Line ((intTabWidth * intNumTabs), FormName.PicTabContainer.ScaleHeight - 1)-((intTabWidth * intNumTabs), TABHEIGHT + OFFSET), 0
    '*****************************************************************
    ' Draw the 3D effect for the form
    '*****************************************************************
    FormName.PicTabContainer.Line (1, TABHEIGHT + OFFSET)-(1, FormName.PicTabContainer.ScaleHeight - 1), QBColor(15)
    FormName.PicTabContainer.Line (2, TABHEIGHT + OFFSET)-(2, FormName.PicTabContainer.ScaleHeight - 1), QBColor(15)
    FormName.PicTabContainer.Line (2, FormName.PicTabContainer.ScaleHeight - 2)-((intTabWidth * intNumTabs) - 1, FormName.PicTabContainer.ScaleHeight - 2), QBColor(8)
    FormName.PicTabContainer.Line (3, FormName.PicTabContainer.ScaleHeight - 3)-((intTabWidth * intNumTabs) - 2, FormName.PicTabContainer.ScaleHeight - 3), QBColor(8)
    FormName.PicTabContainer.Line ((intTabWidth * intNumTabs) - 1, FormName.PicTabContainer.ScaleHeight - 2)-((intTabWidth * intNumTabs) - 1, TABHEIGHT + OFFSET), QBColor(8)
    FormName.PicTabContainer.Line ((intTabWidth * intNumTabs) - 2, FormName.PicTabContainer.ScaleHeight - 2)-((intTabWidth * intNumTabs) - 2, TABHEIGHT + OFFSET), QBColor(8)
    '*****************************************************************
    ' Determine which tab was clicked
    '*****************************************************************
    If sngXPos <> 0 Then intActiveTab = Int(sngXPos / intTabWidth) + 1
    '*****************************************************************
    ' Make suret that intActiveTab >= 1
    '*****************************************************************
    If intActiveTab < 1 Then intActiveTab = 1
    '*****************************************************************
    ' Draw the 3D effect around the active tab
    '*****************************************************************
    sngX = (intActiveTab - 1) * intTabWidth
    FormName.PicTabContainer.Line (sngX + 1, TABHEIGHT + OFFSET)-(sngX + 1, 4 + OFFSET), QBColor(15)
    FormName.PicTabContainer.Line (sngX + 1, 4 + OFFSET)-(sngX + 4, 1 + 0 + OFFSET), QBColor(15)
    FormName.PicTabContainer.Line (sngX + 2, TABHEIGHT + OFFSET)-(sngX + 2, 4 + OFFSET), QBColor(15)
    FormName.PicTabContainer.Line (sngX + 2, 4 + OFFSET)-(sngX + 5, 1 + 0 + OFFSET), QBColor(15)
    FormName.PicTabContainer.Line (sngX + 4, 1 + 0 + OFFSET)-(sngX + intTabWidth - 4, 1 + 0 + OFFSET), QBColor(15)
    FormName.PicTabContainer.Line (sngX + intTabWidth - 4, 1 + 0 + OFFSET)-(sngX + intTabWidth - 1, 4 + OFFSET), QBColor(8)
    FormName.PicTabContainer.Line (sngX + intTabWidth - 1, 4 + OFFSET)-(sngX + intTabWidth - 1, TABHEIGHT + OFFSET + 2), QBColor(8)
    FormName.PicTabContainer.Line (sngX + intTabWidth - 5, 1 + 0 + OFFSET)-(sngX + intTabWidth - 2, 4 + OFFSET), QBColor(8)
    FormName.PicTabContainer.Line (sngX + intTabWidth - 2, 4 + OFFSET)-(sngX + intTabWidth - 2, TABHEIGHT + OFFSET + 2), QBColor(8)
    '*****************************************************************
    ' Draw a horizontal 3D line to the left of the active tab
    '*****************************************************************
    sngX = 2
    sngX1 = ((intActiveTab - 1) * intTabWidth) + 1
    If sngX <> sngX1 + 1 Then
        FormName.PicTabContainer.Line (sngX - 1, TABHEIGHT + OFFSET)-(sngX1, TABHEIGHT + OFFSET), 0
        FormName.PicTabContainer.Line (sngX, TABHEIGHT + OFFSET + 1)-(sngX1 + 1, TABHEIGHT + OFFSET + 1), QBColor(15)
    End If
    '*****************************************************************
    ' Draw a horizontal 3D line to the right of the active tab
    '*****************************************************************
    sngX = intActiveTab * intTabWidth
    sngX1 = (intTabWidth * intNumTabs) - 2
    If sngX <> sngX1 + 2 Then
        FormName.PicTabContainer.Line (sngX, TABHEIGHT + OFFSET)-(sngX1 + 1, TABHEIGHT + OFFSET), 0
        FormName.PicTabContainer.Line (sngX - 1, TABHEIGHT + OFFSET + 1)-(sngX1, TABHEIGHT + OFFSET + 1), QBColor(15)
    End If
    '*****************************************************************
    ' Print the text on the tabs
    '*****************************************************************
    sngX = 0
    FormName.PicTabContainer.CurrentY = OFFSET + ((TABHEIGHT / 2) - (FormName.PicTabContainer.TextHeight("sngX") / 2))
    For intLpctr = 1 To intNumTabs
        FormName.PicTabContainer.FontBold = IIf(intLpctr = intActiveTab, True, False)
        FormName.PicTabContainer.CurrentX = sngX + (intTabWidth / 2) - (FormName.PicTabContainer.TextWidth(Trim(strTabLabels(intLpctr))) / 2)
        '*************************************************************
        ' A semi-colon is required to prevent changing CurrentY
        '*************************************************************
        FormName.PicTabContainer.Print Trim(strTabLabels(intLpctr));
        sngX = sngX + intTabWidth
    Next intLpctr
    '*****************************************************************
    ' Return the active tab index
    '*****************************************************************
    
    
    
    Realize_GraphSettings intActiveTab
       
    DrawTabs = intActiveTab
End Function

'*********************************************************************
' Setup Tabs - Prepares a form to be a tabbed dialog
'---------------------------------------------------------------------
' FormName      (Form)      Name of the form to draw the tabs
' NumTabs       (String)    The number of Tabs() frames on FormName
'*********************************************************************
Sub SetupTabs(FormName As Form, intNumTabs As Integer)

Dim intLpctr As Integer

    '*****************************************************************
    ' Set the backcolor of the form
    '*****************************************************************
    FormName.PicTabContainer.BackColor = QBColor(7)
    '*****************************************************************
    ' Build the array that holds the tab labels
    '*****************************************************************
    ReDim Preserve strTabLabels(1 To intNumTabs)
    '*****************************************************************
    ' Fill the array with the values provided by Labels
    '*****************************************************************
    For intLpctr = 1 To intNumTabs
        If (FormName.fratabs(intLpctr - 1) <> "") Then
          strTabLabels(intLpctr) = FormName.fratabs(intLpctr - 1)
          FormName.fratabs(intLpctr - 1) = ""
          FormName.fratabs(intLpctr - 1).BackColor = QBColor(7)
        End If
    Next intLpctr
    '*****************************************************************
    ' Initialize the tabs
    '*****************************************************************
    FormName.fratabs(DrawTabs(FormName, 10, 10) - 1).ZOrder
    '*****************************************************************
    ' Put the frames on top of each other
    '*****************************************************************
    For intLpctr = 0 To intNumTabs - 1
        FormName.fratabs(intLpctr).Move 8, 24, FormName.PicTabContainer.ScaleWidth - 17, FormName.PicTabContainer.ScaleHeight - 32
    Next intLpctr
End Sub


VERSION 5.00
Begin VB.UserControl Splitter 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   150
   ControlContainer=   -1  'True
   ScaleHeight     =   2145
   ScaleWidth      =   150
   ToolboxBitmap   =   "Splitter.ctx":0000
   Begin VB.PictureBox SplitterBar 
      Height          =   2130
      Left            =   0
      ScaleHeight     =   2070
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "Splitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************************************************************
' Orientation (Boolean, r/w) True: the splitter bar will be horizontal
' SplitPercent (Byte, r/w)   10-90 Percentage of the width of the control for first pane
' SplitterWidth (Byte, r/w) 0-*, width of the splitterbar
' SplitterColor (Long, r/w) Color value or constant that you want the splitterbar to be
' Child1 (Object, r/w) The control to act as pane1, the upper or left pane
' Child2 (Object, r/w) The control to act as pane2, the lower or right pane
' Be sure to set the Child1 and Child2 properties in the form load event to controls
' contained within the Splitterbar control.
'****************************************************************************************

'********************************
' Constants for properties
'********************************
Private Const nChild1  As String = "Child1"
Private Const nChild2  As String = "Child2"

'********************************
' Variables for properties
'********************************
Private mOrientation As Boolean
Private mChild1        As Object
Private mChild2        As Object
Private mSplitPercent    As Single
Private mSplitterWidth      As Byte
Private mSplitterColor  As Long

Public Event Resize()

Public Enum splitterOrientation
    splitVertical = False
    splitHorizontal = True
End Enum

Public Property Get Orientation() As splitterOrientation
    Orientation = mOrientation
End Property

Public Property Let Orientation(val As splitterOrientation)
    mOrientation = val
    SplitterBar.MousePointer = IIf(Orientation, vbSizeNS, vbSizeWE)
    PropertyChanged "Orientation"
    UserControl_Resize
End Property

Public Property Get Child1() As Object
    Set Child1 = mChild1
End Property

Public Property Set Child1(ctl As Object)
    Set mChild1 = ctl
    PropertyChanged nChild1
    UserControl_Resize
End Property

Public Property Get Child2() As Object
    Set Child2 = mChild2
End Property

Public Property Set Child2(ctl As Object)
    Set mChild2 = ctl
    PropertyChanged nChild2
    UserControl_Resize
End Property

Public Property Get SplitPercent() As Byte
    SplitPercent = mSplitPercent * 100
End Property

Public Property Let SplitPercent(val As Byte)
    mSplitPercent = val / 100
    PropertyChanged "SplitPercent"
    UserControl_Resize
End Property

Public Property Get SplitterWidth() As Byte
    SplitterWidth = mSplitterWidth
End Property

Public Property Let SplitterWidth(val As Byte)
    mSplitterWidth = val
    PropertyChanged "SplitterWidth"
End Property

Public Property Get SplitterColor() As Long
    SplitterColor = mSplitterColor
End Property

Public Property Let SplitterColor(val As Long)
    mSplitterColor = val
    PropertyChanged "SplitterColor"
End Property

Private Sub UserControl_Initialize()
    SplitterBar.BorderStyle = vbBSNone
End Sub

'********************************
' Set up the defaults
'********************************
Private Sub UserControl_InitProperties()
    Orientation = False
    SplitPercent = 50
    SplitterWidth = 50
    SplitterBar.Width = SplitterWidth
End Sub

Private Sub UserControl_Paint()
    If SplitterColor <> vbButtonFace Then
        If SplitterColor = 0 Then
            SplitterBar.BackColor = vbButtonFace
        Else
            SplitterBar.BackColor = SplitterColor
        End If
    Else
        SplitterBar.BackColor = vbButtonFace
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Orientation = PropBag.ReadProperty("Orientation", False)
    SplitPercent = PropBag.ReadProperty("SplitPercent", 50)
    SplitterWidth = PropBag.ReadProperty("SplitterWidth", 80)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Orientation", Orientation, False
    PropBag.WriteProperty "SplitPercent", SplitPercent, 50
    PropBag.WriteProperty "SplitterWidth", SplitterWidth, 80
End Sub

Private Sub SplitterBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With SplitterBar
        .BackColor = SplitterColor     ' Make the splitter visible
        .ZOrder
    End With
End Sub

Private Sub SplitterBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If mOrientation Then        ' horizontal figures
            Y = SplitterBar.Top - (SplitterWidth - Y)
            mSplitPercent = Y / UserControl.Height
            SplitterBar.Move 0, Y
        Else                                    ' vertical
            X = SplitterBar.Left - (SplitterWidth - X)
            mSplitPercent = X / UserControl.Width
            SplitterBar.Move X
        End If
        If mSplitPercent < 0.001 Then mSplitPercent = 0.001     ' Check if in range
        If mSplitPercent > 0.999 Then mSplitPercent = 0.999
    End If
End Sub

Private Sub SplitterBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SplitterBar.BackColor = SplitterColor  ' change the color back to normal
    UserControl_Resize                  ' update the panes
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    If UserControl.Ambient.UserMode Then    ' get rid of border in run mode
        UserControl.BorderStyle = vbBSNone
    End If
    
    Dim pane1 As Single
    Dim pane2 As Single
    Dim totwidth As Single
    Dim totheight As Single
    totwidth = UserControl.Width
    totheight = UserControl.Height
    If mOrientation Then
        pane1 = (totheight - SplitterWidth) * mSplitPercent
        pane2 = (totheight - SplitterWidth) * (1 - mSplitPercent)
        mChild1.Move 0, 0, totwidth, pane1
        mChild2.Move 0, pane1 + SplitterWidth, totwidth, pane2
        SplitterBar.Move 0, pane1, totwidth, SplitterWidth
    Else
        pane1 = (totwidth - SplitterWidth) * mSplitPercent
        pane2 = (totwidth - SplitterWidth) * (1 - mSplitPercent)
        mChild1.Move 0, 0, pane1, totheight
        mChild2.Move pane1 + SplitterWidth, 0, pane2, totheight
        SplitterBar.Move pane1, 0, SplitterWidth, totheight
    End If
    mChild1.Refresh
    mChild2.Refresh
    RaiseEvent Resize
End Sub




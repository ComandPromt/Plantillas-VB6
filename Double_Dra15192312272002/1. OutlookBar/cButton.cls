VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
'   You are free to use this source as long as this copyright message
'     appears on your program's "About" dialog:
'
'   Outlook Bar Project
'   Copyright (c) 2002 Vlad Vissoultchev (wqweto@myrealbox.com)
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "cButton"

'Purpose: Raised when a reference to <b>Outlook Bar</b> control is needed. Implemented internally.
Event GetControl(oValue As ctxOutlookBar)
Attribute GetControl.VB_HelpID = 1005

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsButtonClassType
    ucsBtnClassItem = 0         ' The button is representing an item in a group.
    ucsBtnClassGroup = 1        ' The button is a group of the control.
    ucsBtnClassControl = 2      ' The button is the top most element the hierarchy.
End Enum

Public Enum UcsIconStyle
    ucsIcsSmallIcons = 0        ' (Default) All containing items are decorated with their small icons.
    ucsIcsLargeIcons = 1        ' All containing items are decorated with their large icons.
End Enum

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'=========================================================================
' Private constants and variables
'=========================================================================

Private m_sCaption              As String
Private m_imgSmallIcon          As StdPicture
Private m_imgLargeIcon          As StdPicture
Private m_sKey                  As String
Private m_vTag                  As Variant
Private m_sTooltipText          As String
Private m_oParent               As cButton
Private m_cItems                As cButtons
Private m_eClass                As UcsButtonClassType
Private m_eIconsType            As UcsIconStyle
Private m_bVisible              As Boolean
Private m_bEnabled              As Boolean
#If DebugMode Then
    Private m_sDebugID          As String
#End If

'=========================================================================
' Properties
'=========================================================================

'Purpose: Returns or sets the text to be displayed in a cButton object.
Property Get Caption() As String
Attribute Caption.VB_HelpID = 1001
Attribute Caption.VB_UserMemId = 0
    Caption = m_sCaption
End Property

Property Let Caption(sValue As String)
    m_sCaption = sValue
    pvPaintControl
End Property

'Purpose: Returns or sets the picture to be displayed in a cButton object when parent group is in small icons mode.
Property Get SmallIcon() As StdPicture
Attribute SmallIcon.VB_HelpID = 1014
    Set SmallIcon = m_imgSmallIcon
End Property

Property Set SmallIcon(ByVal oValue As StdPicture)
    Set m_imgSmallIcon = oValue
    pvPaintControl
End Property

'Purpose: Returns or sets the picture to be displayed in a cButton object when parent group is in large icons mode.
Property Get LargeIcon() As StdPicture
Attribute LargeIcon.VB_HelpID = 1010
    Set LargeIcon = m_imgLargeIcon
End Property

Property Set LargeIcon(ByVal oValue As StdPicture)
    Set m_imgLargeIcon = oValue
    pvPaintControl
End Property

'Purpose: Returns a value that represents the index of an object in a collection.
Property Get Index() As Long
Attribute Index.VB_HelpID = 1008
    If Not m_oParent Is Nothing Then
        For Index = 1 To m_oParent.GroupItems.Count
            If m_oParent.GroupItems(Index) Is Me Then
                Exit Property
            End If
        Next
    End If
End Property

'Purpose: Returns a value that represents the position of an item in its parent visible items.
Property Get Position() As Long
Attribute Position.VB_HelpID = 1012
    Dim lIdx            As Long
    
    If Not m_oParent Is Nothing Then
        Position = 1
        For lIdx = 1 To m_oParent.GroupItems.Count
            If m_oParent.GroupItems(lIdx) Is Me Then
                Exit Property
            End If
            Position = Position + Abs(m_oParent.GroupItems(lIdx).Visible)
        Next
    End If
End Property

'Purpose: Returns or sets a string that uniquely identifies a member in a <b>cButtons</b> collection.
Property Get Key() As String
Attribute Key.VB_HelpID = 1009
    Key = m_sKey
End Property

Property Let Key(sValue As String)
    Dim lIdx            As Long
    
    m_sKey = sValue
    If Not Parent Is Nothing Then
        '--- fix key in Parent collection
        lIdx = Index
        With Parent.GroupItems
            .frRemove lIdx
            .frAdd Me, lIdx
        End With
    End If
End Property

Friend Property Let frKey(sValue As String)
    m_sKey = sValue
End Property

'Purpose: Returns or sets a value that determines whether the cButton is currently selected in the <b>Outlook Bar</b> control.
Property Get Selected() As Boolean
Attribute Selected.VB_HelpID = 1013
    Dim oCtl            As ctxOutlookBar
    
    Set oCtl = Control
    If Not oCtl Is Nothing Then
        Select Case m_eClass
        Case ucsBtnClassItem
            Selected = oCtl.SelectedItem Is Me
        Case ucsBtnClassGroup
            Selected = oCtl.SelectedGroup Is Me
        Case ucsBtnClassControl
            '--- always
            Selected = True
        End Select
    End If
End Property

Property Let Selected(ByVal bValue As Boolean)
    Dim oCtl            As ctxOutlookBar
    
    Set oCtl = Control
    If Not oCtl Is Nothing Then
    If bValue = True Then
        If Not Parent Is Nothing Then
            Parent.Selected = True
        End If
        Select Case m_eClass
        Case ucsBtnClassItem
            Set oCtl.SelectedItem = Me
        Case ucsBtnClassGroup
            Set oCtl.SelectedGroup = Me
        Case ucsBtnClassControl
            '--- nothing to do
        End Select
    Else
        '--- if desected select first in parent's items
        If Not Parent Is Nothing Then
            If Parent.GroupItems.Count > 0 Then
                Parent.GroupItems(1).Selected = True
            End If
        End If
    End If
    End If
End Property

'Purpose: Returns or sets an expression that stores any extra data needed for your program.
Property Get Tag() As Variant
Attribute Tag.VB_HelpID = 1015
    If IsObject(m_vTag) Then
        Set Tag = m_vTag
    Else
        Tag = m_vTag
    End If
End Property

Property Let Tag(oValue As Variant)
    m_vTag = oValue
End Property

Property Set Tag(oValue As Variant)
    Set m_vTag = oValue
End Property

'Purpose: Returns or sets a string that is to be displayed as a tooltip for the cButton object.
Property Get TooltipText() As String
Attribute TooltipText.VB_HelpID = 1016
    TooltipText = m_sTooltipText
End Property

Property Let TooltipText(sValue As String)
    m_sTooltipText = sValue
End Property

Friend Property Get Items() As cButtons
    Set Items = m_cItems
End Property

Friend Property Set Items(ByVal oValue As cButtons)
    Set m_cItems = Nothing
End Property

'Purpose: Returns the <b>cButtons</b> collection in a <b>cButton</b> object that represents a group in the <b>Outlook Bar</b> control.
Property Get GroupItems() As cButtons
Attribute GroupItems.VB_HelpID = 1006
    If m_cItems Is Nothing Then
        Set m_cItems = New cButtons
        If Class > 0 Then
            m_cItems.Class = Class - 1
        End If
        Set m_cItems.Parent = Me
    End If
    Set GroupItems = m_cItems
End Property

'Purpose: Returns or sets a value indicating class's RTTI.
Public Property Get Class() As UcsButtonClassType
Attribute Class.VB_HelpID = 1002
    Class = m_eClass
End Property

Property Let Class(ByVal eValue As UcsButtonClassType)
    m_eClass = eValue
End Property

'Purpose: Returns or sets a reference to the parent <b>cButton</b>.
Public Property Get Parent() As cButton
Attribute Parent.VB_HelpID = 1011
    Set Parent = m_oParent
End Property

Friend Property Set Parent(ByVal oValue As cButton)
    '--- ToDo: make it a weak reference to Parent
    Set m_oParent = oValue
End Property

'Purpose: Returns a reference to parent <b>Outlook Bar</b> control.
Property Get Control() As ctxOutlookBar
Attribute Control.VB_HelpID = 1003
    Dim oCtl        As ctxOutlookBar
    
    If Not Parent Is Nothing Then
        Set Control = Parent.Control
    Else
        RaiseEvent GetControl(oCtl)
        Set Control = oCtl
    End If
End Property

'Purpose: Returns or sets a value that represents the type icons displayed in <b>cButton</b> objects.
Property Get IconsType() As UcsIconStyle
Attribute IconsType.VB_HelpID = 1007
    IconsType = m_eIconsType
End Property

Property Let IconsType(ByVal eValue As UcsIconStyle)
    On Error Resume Next
    m_eIconsType = eValue
    Control.frGetMeasures '--- control's m_lGroupHeight depends on IconsType
    pvPaintControl
End Property

'Purpose: Returns or sets a value indicating whether a <b>cButton</b> object is visible or hidden.
Property Get Visible() As Boolean
Attribute Visible.VB_HelpID = 1017
    Visible = m_bVisible
End Property

Property Let Visible(ByVal bValue As Boolean)
    m_bVisible = bValue
    pvPaintControl
End Property

'Purpose: Returns or sets a value indicating whether a <b>cButton</b> object is enabled or disabled.
Property Get Enabled() As Boolean
Attribute Enabled.VB_HelpID = 1004
    Enabled = m_bEnabled
End Property

Property Let Enabled(ByVal bValue As Boolean)
    m_bEnabled = bValue
    pvPaintControl
End Property

Private Sub pvPaintControl()
    On Error Resume Next
    Control.RefreshControl
End Sub

#If DebugMode Then
    Private Sub Class_Initialize()
        DebugInit m_sDebugID, MODULE_NAME
    End Sub
    
    Private Sub Class_Terminate()
        DebugTerm m_sDebugID
    End Sub
#End If

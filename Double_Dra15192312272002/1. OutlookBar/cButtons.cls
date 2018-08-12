VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
Private Const MODULE_NAME As String = "cButtons"

'=========================================================================
' Private constants and variables
'=========================================================================

Private Const DEF_CAPTION           As String = ""
Private Const DEF_KEY               As String = ""
Private Const DEF_TOOLTIPTEXT       As String = ""
Private Const DEF_CLASS             As Long = ucsBtnClassItem
Private Const DEF_ICONSTYPE         As Long = ucsIcsSmallIcons
Private Const DEF_VISIBLE           As Boolean = True
Private Const DEF_ENABLED           As Boolean = True

Private m_cItems                As Collection
Private m_hWnd                  As Long
Private m_oParent               As cButton
Private m_eClass                As UcsButtonClassType
Private m_bInSet                As Boolean
#If DebugMode Then
    Private m_sDebugID          As String
#End If

'=========================================================================
' Error handling
'=========================================================================

Private Sub RaiseError(sFunc As String)
    PushError sFunc, MODULE_NAME
    PopRaiseError
End Sub

'Private Function ShowError(sFunc As String) As VbMsgBoxResult
'    PushError sFunc, MODULE_NAME
'    ShowError = PopShowError(CAP_MSG)
'End Function

'=========================================================================
' Propeties
'=========================================================================

'Purpose: Returns a specific <b>cButton</b> object of a <b>cButtons</b> collection, either by index of by key.
Property Get Item(IdxOrKey) As cButton
Attribute Item.VB_HelpID = 2007
Attribute Item.VB_UserMemId = 0
    Set Item = m_cItems(IdxOrKey)
End Property

'Purpose: Returns a specific <b>cButton</b> object of a <b>cButtons</b> collection, according by its visible position.
Property Get ItemByPosition(ByVal Pos As Long) As cButton
Attribute ItemByPosition.VB_HelpID = 2008
    Dim lIdx        As Long
    
    For lIdx = 1 To m_cItems.Count
        If Item(lIdx).Visible Then
            Pos = Pos - 1
            If Pos = 0 Then
                Set ItemByPosition = Item(lIdx)
                Exit Property
            End If
        End If
    Next
End Property

Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_cItems.[_NewEnum]
End Property

'Purpose: Returns the number of objects in a collection.
Property Get Count() As Long
Attribute Count.VB_HelpID = 2005
    Count = m_cItems.Count
End Property

Friend Property Get Parent() As cButton
    Set Parent = m_oParent
End Property

Friend Property Set Parent(ByVal oValue As cButton)
    Set m_oParent = oValue
End Property

Friend Property Get hwnd() As Long
    hwnd = m_hWnd
End Property

Friend Property Let hwnd(ByVal lValue As Long)
    m_hWnd = lValue
End Property

'Purpose: Returns or sets the contents of the collection as a serialized byte array.
Property Get Contents() As Byte()
Attribute Contents.VB_HelpID = 2004
    Const FUNC_NAME     As String = "Contents [get]"
    Dim oBag            As New PropertyBag
    Dim lIdx            As Long

    On Error GoTo EH
    With oBag
        .WriteProperty "Cn", Count, 0
        For lIdx = 1 To Count
            With Item(lIdx)
                oBag.WriteProperty "C" & lIdx, .Caption, DEF_CAPTION
                WritePictureProperty oBag, "LI" & lIdx, .LargeIcon, Nothing
                WritePictureProperty oBag, "SI" & lIdx, .SmallIcon, Nothing
                oBag.WriteProperty "K" & lIdx, .Key, DEF_KEY
                oBag.WriteProperty "T" & lIdx, .TooltipText, DEF_TOOLTIPTEXT
                oBag.WriteProperty "CL" & lIdx, .Class, DEF_CLASS
                oBag.WriteProperty "TG" & lIdx, .Tag, Empty
                If Not .Items Is Nothing Then
                    oBag.WriteProperty "G" & lIdx, .Items.Contents
                End If
                oBag.WriteProperty "I" & lIdx, .IconsType, DEF_ICONSTYPE
                oBag.WriteProperty "V" & lIdx, .Visible, DEF_VISIBLE
                oBag.WriteProperty "E" & lIdx, .Enabled, DEF_ENABLED
            End With
        Next
        Contents = .Contents
    End With
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Property Let Contents(aValue() As Byte)
    Const FUNC_NAME     As String = "Contents [let]"
    Dim oBag            As New PropertyBag
    Dim lIdx            As Long
    
    On Error GoTo EH
    With oBag
        .Contents = aValue
        If Err.Number = 0 Then
            Clear
            m_bInSet = True
            For lIdx = 1 To .ReadProperty("Cn", 0)
                With Add(.ReadProperty("C" & lIdx, DEF_CAPTION), _
                            ReadPictureProperty(oBag, "LI" & lIdx, Nothing), _
                            ReadPictureProperty(oBag, "SI" & lIdx, Nothing), _
                            .ReadProperty("K" & lIdx, DEF_KEY))
                    Set .Parent = Parent
                    .TooltipText = oBag.ReadProperty("T" & lIdx, DEF_TOOLTIPTEXT)
                    .Tag = oBag.ReadProperty("TG" & lIdx, Empty)
                    .Class = oBag.ReadProperty("CL" & lIdx, DEF_CLASS)
                    If IsArray(oBag.ReadProperty("G" & lIdx, "")) Then
                        .GroupItems.Contents = oBag.ReadProperty("G" & lIdx)
                    End If
                    .IconsType = oBag.ReadProperty("I" & lIdx, DEF_ICONSTYPE)
                    .Visible = oBag.ReadProperty("V" & lIdx, DEF_VISIBLE)
                    .Enabled = oBag.ReadProperty("E" & lIdx, DEF_ENABLED)
                End With
            Next
            m_bInSet = False
            pvRefreshControl
        End If
    End With
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Friend Property Get Class() As UcsButtonClassType
Attribute Class.VB_HelpID = 2002
    Class = m_eClass
End Property

Friend Property Let Class(ByVal eValue As UcsButtonClassType)
    m_eClass = eValue
End Property

'=========================================================================
' Methods
'=========================================================================

'Purpose: Adds a <b>cButton</b> object to the collection and returns a reference to the newly created object.
Public Function Add( _
            Caption As String, _
            Optional ByVal LargeIcon As StdPicture, _
            Optional ByVal SmallIcon As StdPicture, _
            Optional Key, _
            Optional Before, _
            Optional After) As cButton
Attribute Add.VB_HelpID = 2001
    Const FUNC_NAME     As String = "Add"
    
    On Error GoTo EH
    Set Add = New cButton
    Set Add.Parent = Parent
    Add.Class = Class
    Add.Caption = Caption
    Add.TooltipText = Caption
    Set Add.LargeIcon = LargeIcon
    Set Add.SmallIcon = SmallIcon
    If Not IsMissing(Key) Then
        Add.frKey = C2Str(Key)
    End If
    Add.Enabled = True
    Add.Visible = True
    If Add.Key <> "" Then
        m_cItems.Add Add, Key, Before, After
    Else
        m_cItems.Add Add, , Before, After
    End If
    If m_cItems.Count = 1 Then
        If Class = ucsBtnClassGroup Then
            Add.Selected = True
        End If
    End If
    pvRefreshControl
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Friend Sub frAdd( _
            oBtn As cButton, _
            Optional ByVal lIndex As Long)
    Const FUNC_NAME     As String = "frAdd"
    
    On Error GoTo EH
    Set oBtn.Parent = Parent
    If lIndex > 0 And lIndex <= m_cItems.Count Then
        '--- insert in between
        If Len(oBtn.Key) > 0 Then
            m_cItems.Add oBtn, oBtn.Key, lIndex
        Else
            m_cItems.Add oBtn, , lIndex
        End If
    Else
        '--- append to back
        If Len(oBtn.Key) > 0 Then
            m_cItems.Add oBtn, oBtn.Key
        Else
            m_cItems.Add oBtn
        End If
    End If
    pvRefreshControl
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'Purpose: Removes a specific member from the <b>cButtons</b> collection.
Public Sub Remove(IdxOrKey As Variant)
Attribute Remove.VB_HelpID = 2009
    Const FUNC_NAME     As String = "Remove"
    Dim oBtn            As cButton
    Dim lIdx            As Long
    Dim oCtl            As ctxOutlookBar
    
    On Error GoTo EH
    '--- init local vars
    Set oBtn = Item(IdxOrKey)
    lIdx = oBtn.Index
    Set oCtl = oBtn.Control
    '--- clear children
    If Not oBtn.Items Is Nothing Then
        Set oBtn.Items.Parent = Nothing
        oBtn.Items.Clear
        Set oBtn.Items = Nothing
    End If
    '--- disconnect parent
    Set oBtn.Parent = Nothing
    '--- remove from current collection
    m_cItems.Remove IdxOrKey
    '--- fix removed control's selected item/group
    If Not oCtl Is Nothing Then
        If oCtl.SelectedItem Is oBtn Then
            Set oCtl.SelectedItem = Nothing
        End If
        If oCtl.SelectedGroup Is oBtn Then
            If oCtl.Groups.Count >= lIdx Then
                Set oCtl.SelectedGroup = oCtl.Groups(lIdx)
            Else
                If oCtl.Groups.Count > 0 Then
                    Set oCtl.SelectedGroup = oCtl.Groups(lIdx - 1)
                Else
                    Set oCtl.SelectedGroup = Nothing
                End If
            End If
        End If
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Friend Sub frRemove(IdxOrKey As Variant)
    Const FUNC_NAME     As String = "frRemove"
    
    On Error GoTo EH
    '--- remove from current collection
    m_cItems.Remove IdxOrKey
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'Purpose: Removes all objects in a collection.
Public Sub Clear()
Attribute Clear.VB_HelpID = 2003
    Const FUNC_NAME As String = "Clear"
    
    On Error GoTo EH
    Do While m_cItems.Count > 0
        '--- recusevely cleanup
        If Not Item(1).Items Is Nothing Then
            Item(1).Items.Clear
        End If
        Remove 1
    Loop
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'Purpose: Returns a value that indicates whether a <b>cButton</b> objects exists in a collection.
Public Function Exists(IdxKeyOrObject) As Boolean
Attribute Exists.VB_HelpID = 2006
    Dim oItem           As cButton
    
    On Error Resume Next
    If IsObject(IdxKeyOrObject) Then
        For Each oItem In Me
            If oItem Is IdxKeyOrObject Then
                Exit For
            End If
        Next
    Else
        Set oItem = Item(IdxKeyOrObject)
    End If
    '--- success (or failure)
    Exists = Not oItem Is Nothing
End Function

Private Sub pvRefreshControl()
    Const FUNC_NAME     As String = "pvRefreshControl"

    On Error GoTo EH
    If Not m_bInSet Then
        If Not Parent.Control Is Nothing Then
            Parent.Control.RefreshControl
        End If
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'=========================================================================
' Class events
'=========================================================================

Private Sub Class_Initialize()
    Set m_cItems = New Collection
    #If DebugMode Then
        DebugInit m_sDebugID, MODULE_NAME
    #End If
End Sub

#If DebugMode Then
    Private Sub Class_Terminate()
        DebugTerm m_sDebugID
    End Sub
#End If


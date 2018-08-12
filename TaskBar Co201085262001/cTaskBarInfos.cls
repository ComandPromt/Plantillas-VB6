VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTaskBarInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"cTaskBarInfo"
Attribute VB_Ext_KEY = "Member0" ,"cTaskBarInfo"
Option Explicit

'=========================================================================
' Constants and variables
'=========================================================================

Private m_oCol          As Collection
Private m_hWnd          As Long

'=========================================================================
' Properties
'=========================================================================

Public Property Get Item(IdxOrKey As Variant) As cTaskBarInfo
Attribute Item.VB_UserMemId = 0
    Set Item = m_oCol(IdxOrKey)
End Property

Public Property Get Count() As Long
    Count = m_oCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_oCol.[_NewEnum]
End Property

Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Friend Property Let hWnd(ByVal lValue As Long)
    m_hWnd = lValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Add( _
        Optional Caption As String, _
        Optional TooltipText As String, _
        Optional Icon As StdPicture, _
        Optional sKey As String) As cTaskBarInfo
    On Error Resume Next
    Set Add = New cTaskBarInfo
    Add.Key = sKey
    Add.hWnd = hWnd
    Set Add.Icon = Icon
    Add.TooltipText = TooltipText
    Add.Caption = Caption
    If Len(sKey) = 0 Then
        m_oCol.Add Add
    Else
        m_oCol.Add Add, sKey
    End If
    RefreshDC hWnd
End Function

Public Sub Remove(IdxOrKey As Variant)
    m_oCol.Remove IdxOrKey
    RefreshDC hWnd
End Sub

'=========================================================================
' Controls events
'=========================================================================

Private Sub Class_Initialize()
    Set m_oCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_oCol = Nothing
End Sub


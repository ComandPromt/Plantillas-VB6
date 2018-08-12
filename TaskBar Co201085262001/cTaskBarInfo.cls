VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTaskBarInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cTaskBarInfos"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'=========================================================================
' Constants and variables
'=========================================================================

Private m_sCaption      As String
Private m_sTooltipText  As String
Private m_oIcon         As StdPicture
Private m_hWnd          As Long
Private m_sKey          As String

'=========================================================================
' Properties
'=========================================================================

Property Get Caption() As String
    Caption = m_sCaption
End Property

Property Let Caption(sValue As String)
    m_sCaption = sValue
    Refresh
End Property

Property Get TooltipText() As String
    TooltipText = m_sTooltipText
End Property

Property Let TooltipText(sValue As String)
    m_sTooltipText = sValue
    Refresh
End Property

Property Get Icon() As StdPicture
    Set Icon = m_oIcon
End Property

Property Set Icon(ByVal oValue As StdPicture)
    Set m_oIcon = oValue
    Refresh
End Property

Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Friend Property Let hWnd(ByVal lValue As Long)
    m_hWnd = lValue
End Property

Property Get Key() As String
    Key = m_sKey
End Property

Friend Property Let Key(sValue As String)
    m_sKey = sValue
End Property

'=========================================================================
' Methods
'=========================================================================

Private Function Refresh()
    RefreshDC hWnd
End Function

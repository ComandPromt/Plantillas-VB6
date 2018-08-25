VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrowsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'This Collection Class is used to Expose the Browser Collection from the User control
Private InternalCollection As Collection
Public Function NewEnum() As IUnknown 'Allow Enumeration of our internal collection
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = InternalCollection.[_NewEnum]
End Function
Friend Function AddItem(Item As SHDocVw.InternetExplorer)
    InternalCollection.Add Item
End Function
Friend Function Clear()
    Set InternalCollection = Nothing
    Set InternalCollection = New Collection
End Function
Private Sub Class_Initialize()
    Clear
End Sub
Function Item(Index) As SHDocVw.InternetExplorer
Attribute Item.VB_UserMemId = 0
    Set Item = InternalCollection.Item(Index)
End Function
Public Property Get Count() As Integer
    Count = InternalCollection.Count
End Property


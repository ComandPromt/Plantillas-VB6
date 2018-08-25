VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'IBrowser - Interface Class
'This class defines the interface for calling classes.
Public Sub SetBrowserInst(Browser As SHDocVw.InternetExplorer)
    '
End Sub
Public Sub BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'
End Sub

Public Sub DocumentComplete(ByVal pDisp As Object, URL As Variant)
'
End Sub

Public Sub DownLoadBegin()
'
End Sub

Public Sub DownLoadComplete()
'
End Sub

Public Sub FileDownload(Cancel As Boolean)
'
End Sub

Public Sub NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'
End Sub

Public Sub NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
'
End Sub

Public Sub NewWindow2(ppDisp As Object, Cancel As Boolean)
'
End Sub

Public Sub OnFullScreen(ByVal FullScreen As Boolean)
'
End Sub

Public Sub ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'
End Sub

Public Sub TitleChange(ByVal Text As String)
'
End Sub

Public Sub WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
'
End Sub

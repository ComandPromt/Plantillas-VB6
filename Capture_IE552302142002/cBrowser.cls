VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Browser As SHDocVw.InternetExplorer
Attribute Browser.VB_VarHelpID = -1
Public InterFace As IBrowser
Private Sub Browser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    InterFace.SetBrowserInst Browser
    InterFace.BeforeNavigate2 pDisp, URL, Flags, TargetFrameName, PostData, Headers, Cancel
End Sub

Private Sub Browser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'
    InterFace.SetBrowserInst Browser
    InterFace.DocumentComplete pDisp, URL
End Sub

Private Sub Browser_DownloadBegin()
'
    InterFace.SetBrowserInst Browser
    InterFace.DownLoadBegin
End Sub

Private Sub Browser_DownloadComplete()
'
    InterFace.SetBrowserInst Browser
    InterFace.DownLoadComplete
End Sub

Private Sub Browser_FileDownload(Cancel As Boolean)
'
    InterFace.SetBrowserInst Browser
    InterFace.FileDownload Cancel
End Sub

Private Sub Browser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'
    InterFace.SetBrowserInst Browser
    InterFace.NavigateComplete2 pDisp, URL
End Sub

Private Sub Browser_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
'
    InterFace.SetBrowserInst Browser
    InterFace.NavigateError pDisp, URL, Frame, StatusCode, Cancel
End Sub

Private Sub Browser_NewWindow2(ppDisp As Object, Cancel As Boolean)
'
    InterFace.SetBrowserInst Browser
    InterFace.NewWindow2 ppDisp, Cancel
End Sub

Private Sub Browser_OnFullScreen(ByVal FullScreen As Boolean)
'
    InterFace.SetBrowserInst Browser
    InterFace.OnFullScreen FullScreen
End Sub

Private Sub Browser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'
    InterFace.SetBrowserInst Browser
    InterFace.ProgressChange Progress, ProgressMax
End Sub

Private Sub Browser_TitleChange(ByVal Text As String)
'
    InterFace.SetBrowserInst Browser
    InterFace.TitleChange Text
End Sub

Private Sub Browser_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
'
    InterFace.SetBrowserInst Browser
    InterFace.WindowClosing IsChildWindow, Cancel
End Sub

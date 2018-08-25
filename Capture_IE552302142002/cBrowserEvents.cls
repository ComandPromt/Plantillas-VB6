VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrowserEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event BrowserNavigating(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Event DocumentComplete(Browser As SHDocVw.InternetExplorer, pDisp As Object, URL As Variant)
Event DownLoadBegin(Browser As SHDocVw.InternetExplorer)
Event DownLoadComplete(Browser As SHDocVw.InternetExplorer)
Event FileDownload(Browser As SHDocVw.InternetExplorer, Cancel As Boolean)
Event NavigateComplete(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant)
Event NavigateError(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
Event NewWindow(Browser As SHDocVw.InternetExplorer, ppDisp As Object, Cancel As Boolean)
Event OnFullScreen(Browser As SHDocVw.InternetExplorer, ByVal FullScreen As Boolean)
Event ProgressChange(Browser As SHDocVw.InternetExplorer, ByVal Progress As Long, ByVal ProgressMax As Long)
Event TitleChange(Browser As SHDocVw.InternetExplorer, ByVal Text As String)
Event WindowClosing(Browser As SHDocVw.InternetExplorer, ByVal IsChildWindow As Boolean, Cancel As Boolean)
Event BrowserCreated(Browser As SHDocVw.InternetExplorer)
Event BrowserDestroyed()

Implements IBrowser ' Implement the IBrowser Interface Class
Private m_OwnerBrCollClass As cBrowsers ' Internal ref to cBrowsers collection class
Private m_oBrowser As SHDocVw.InternetExplorer
Private WithEvents m_oShell As SHDocVw.ShellWindows
Attribute m_oShell.VB_VarHelpID = -1
Dim coll As New Collection ' Hold all of our instances in internal collection
Private Sub Class_Initialize()
    SyncCollection
End Sub
'#####################################################################
'#Author    : Richard Friend,                                        #
'#Date      : Wed Sep 2001 10:09:55                                  #
'#Comments  :                                                        #
'#####################################################################
Private Function KeyInCollection(col As Collection, strKey As String)
    On Error Resume Next
    col.Item strKey
    KeyInCollection = Err.Number = 0
End Function
'Sycronise Both the Internal Events Collection
'And our public collection exposed
Friend Sub SyncCollection()
    
    Dim oTemp As SHDocVw.InternetExplorer
    Dim oo As cBrowser
    Dim sTemp As String
    Set m_oShell = Nothing 'Destroy ShellWindow object
    Set coll = Nothing 'Destroy Collection
    Set m_oShell = New SHDocVw.ShellWindows
    If Not m_OwnerBrCollClass Is Nothing Then
        m_OwnerBrCollClass.Clear
    End If
    For Each oTemp In m_oShell
        Set oo = New cBrowser
        Set oo.InterFace = Me
        Set oo.Browser = oTemp
        sTemp = ""
        While KeyInCollection(coll, "_" & oo.Browser.hWnd & "_" & sTemp)
            'This window has Child windows using the same HWND
            sTemp = CStr(CLng(Val(sTemp)) + 1)
            'Append Our Child Instane Number to the key
        Wend
        coll.Add oo, "_" & oo.Browser.hWnd & "_" & sTemp
        If Not m_OwnerBrCollClass Is Nothing Then
            m_OwnerBrCollClass.AddItem oTemp
        End If
    Next oTemp
    
End Sub
Friend Sub SetOwnerBrowserCollection(pBrColl As cBrowsers)
    Set m_OwnerBrCollClass = pBrColl
End Sub
Private Function GetNewestInstance() As SHDocVw.InternetExplorer
    Dim oTempBr As SHDocVw.InternetExplorer
    Dim oTempBr2 As cBrowser
    Dim blnFound As Boolean
    Dim lCount As Long
    For Each oTempBr In m_oShell
        blnFound = False
        Debug.Print oTempBr.hWnd
        For Each oTempBr2 In coll
            If oTempBr2.Browser.hWnd = oTempBr.hWnd Then
                blnFound = True
            End If
        Next oTempBr2
        If Not blnFound Then
            Set GetNewestInstance = oTempBr 'Newest instance
            Exit For
        End If
    Next oTempBr
    Set oTempBr2 = Nothing
End Function

Private Sub Class_Terminate()
    Set coll = Nothing
End Sub

Private Sub IBrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    RaiseEvent BrowserNavigating(m_oBrowser, pDisp, URL, Flags, TargetFrameName, PostData, Headers, Cancel)
End Sub

Private Sub IBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    RaiseEvent DocumentComplete(m_oBrowser, pDisp, URL)
End Sub

Private Sub IBrowser_DownloadBegin()
'
    RaiseEvent DownLoadBegin(m_oBrowser)
End Sub

Private Sub IBrowser_DownloadComplete()
'
    RaiseEvent DownLoadComplete(m_oBrowser)
End Sub

Private Sub IBrowser_FileDownload(Cancel As Boolean)
'
    RaiseEvent FileDownload(m_oBrowser, Cancel)
End Sub

Private Sub IBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'
    RaiseEvent NavigateComplete(m_oBrowser, pDisp, URL)
End Sub

Private Sub IBrowser_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
'
    RaiseEvent NavigateError(m_oBrowser, pDisp, URL, Frame, StatusCode, Cancel)
End Sub

Private Sub IBrowser_NewWindow2(ppDisp As Object, Cancel As Boolean)
'
    RaiseEvent NewWindow(m_oBrowser, ppDisp, Cancel)
End Sub

Private Sub IBrowser_OnFullScreen(ByVal FullScreen As Boolean)
'
    RaiseEvent OnFullScreen(m_oBrowser, FullScreen)
End Sub

Private Sub IBrowser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'
    RaiseEvent ProgressChange(m_oBrowser, Progress, ProgressMax)
End Sub

Private Sub IBrowser_SetBrowserInst(Browser As SHDocVw.IWebBrowser2)
    Set m_oBrowser = Browser
End Sub

Private Sub IBrowser_TitleChange(ByVal Text As String)
'
    RaiseEvent TitleChange(m_oBrowser, Text)
End Sub

Private Sub IBrowser_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
'
    RaiseEvent WindowClosing(m_oBrowser, IsChildWindow, Cancel)
End Sub

Private Sub m_oShell_WindowRegistered(ByVal lCookie As Long)
    Dim oTempBrowser As SHDocVw.InternetExplorer
    Set oTempBrowser = GetNewestInstance
    'Lets get the new browser in the ShellWindows collection
    'By comparing it against our Internal collection
    'We must do this before we update our internal collection
    If oTempBrowser Is Nothing Then
        'This occurs when you click to open a new window
        'The windows have the same HWND!!!
        'dont raise an event since the NewWindow2 event will raise!
    Else
        SyncCollection 'Update our collection before raising the event!
        RaiseEvent BrowserCreated(oTempBrowser)
    End If
    
End Sub

Private Sub m_oShell_WindowRevoked(ByVal lCookie As Long)
    SyncCollection 'Update our collection
    RaiseEvent BrowserDestroyed
End Sub

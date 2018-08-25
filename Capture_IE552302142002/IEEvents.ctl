VERSION 5.00
Begin VB.UserControl IEEvents 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3105
   ScaleHeight     =   735
   ScaleWidth      =   3105
   ToolboxBitmap   =   "IEEvents.ctx":0000
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "IEEvents.ctx":0312
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "IE Event Control"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   3135
      End
   End
End
Attribute VB_Name = "IEEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents oBrowserEvents As cBrowserEvents
Attribute oBrowserEvents.VB_VarHelpID = -1
Private m_Browsers As cBrowsers
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
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDc As Long, lpRect As RECT) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Property Get Enabled() As Boolean
    Enabled = Not oBrowserEvents Is Nothing
    PropertyChanged "Enabled"
End Property
'Destroy our browser collection and get new one
Public Sub Refresh()
    Set oBrowserEvents = Nothing
    Set oBrowserEvents = New cBrowserEvents
    oBrowserEvents.SetOwnerBrowserCollection m_Browsers
    oBrowserEvents.SyncCollection
End Sub
'Must set enabled to get events
Public Property Let Enabled(ByVal blnNewValue As Boolean)
    If blnNewValue Then
        If oBrowserEvents Is Nothing Then
            'Setting enabled when already enabled does nothing
            Set oBrowserEvents = New cBrowserEvents
            oBrowserEvents.SetOwnerBrowserCollection m_Browsers
            oBrowserEvents.SyncCollection
        End If
    Else
        Set oBrowserEvents = Nothing
    End If
    PropertyChanged "Enabled"
End Property

Private Sub oBrowserEvents_BrowserCreated(Browser As SHDocVw.InternetExplorer)
    RaiseEvent BrowserCreated(Browser)
End Sub

Private Sub oBrowserEvents_BrowserDestroyed()
    RaiseEvent BrowserDestroyed
End Sub

Private Sub oBrowserEvents_BrowserNavigating(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    RaiseEvent BrowserNavigating(Browser, pDisp, URL, Flags, TargetFrameName, PostData, Headers, Cancel)
End Sub

Private Sub oBrowserEvents_DocumentComplete(Browser As SHDocVw.InternetExplorer, pDisp As Object, URL As Variant)
    RaiseEvent DocumentComplete(Browser, pDisp, URL)
End Sub

Private Sub oBrowserEvents_DownLoadBegin(Browser As SHDocVw.InternetExplorer)
    RaiseEvent DownLoadBegin(Browser)
End Sub

Private Sub oBrowserEvents_DownLoadComplete(Browser As SHDocVw.InternetExplorer)
    RaiseEvent DownLoadComplete(Browser)
End Sub

Private Sub oBrowserEvents_FileDownload(Browser As SHDocVw.InternetExplorer, Cancel As Boolean)
    RaiseEvent FileDownload(Browser, Cancel)
End Sub

Private Sub oBrowserEvents_NavigateComplete(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant)
    RaiseEvent NavigateComplete(Browser, pDisp, URL)
End Sub

Private Sub oBrowserEvents_NavigateError(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
    RaiseEvent NavigateError(Browser, pDisp, URL, Frame, StatusCode, Cancel)
End Sub

Private Sub oBrowserEvents_NewWindow(Browser As SHDocVw.InternetExplorer, ppDisp As Object, Cancel As Boolean)
    RaiseEvent NewWindow(Browser, ppDisp, Cancel)
End Sub

Private Sub oBrowserEvents_OnFullScreen(Browser As SHDocVw.InternetExplorer, ByVal FullScreen As Boolean)
    RaiseEvent OnFullScreen(Browser, FullScreen)
End Sub

Private Sub oBrowserEvents_ProgressChange(Browser As SHDocVw.InternetExplorer, ByVal Progress As Long, ByVal ProgressMax As Long)
    RaiseEvent ProgressChange(Browser, Progress, ProgressMax)
End Sub

Private Sub oBrowserEvents_TitleChange(Browser As SHDocVw.InternetExplorer, ByVal Text As String)
    RaiseEvent TitleChange(Browser, Text)
End Sub

Private Sub oBrowserEvents_WindowClosing(Browser As SHDocVw.InternetExplorer, ByVal IsChildWindow As Boolean, Cancel As Boolean)
    RaiseEvent WindowClosing(Browser, IsChildWindow, Cancel)
End Sub


Private Sub UserControl_Initialize()
    Grad Picture2, 3
    SavePicture Image1.Picture, "c:\a.bmp"
    Label2.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
    'Grad Picture1, 2
    DrawRectangle 0, 0, Picture2.Width / 15, Picture2.Height / 15, vbYellow, True
    Set m_Browsers = New cBrowsers
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 3135
    UserControl.Height = 735
End Sub

Private Sub UserControl_Terminate()
    Set oBrowserEvents = Nothing
    Set m_Browsers = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", oBrowserEvents Is Nothing, False
End Sub

Public Property Get Browsers() As cBrowsers
    Set Browsers = m_Browsers
End Property
Private Sub Grad(f As Object, Factor As Long)
Dim Amount As Long
Dim i As Single
Dim r As Long
Dim g As Long
Dim b As Long

On Error Resume Next
    f.AutoRedraw = True
    Amount = (255 / f.ScaleHeight)
    If Amount = 0 Then Amount = 1
    For i = 0 To f.ScaleHeight
        r = CInt(Amount * i / Factor) - 100
        If r < 0 Then r = 0
        g = CInt(Amount * i / Factor) - 50
        If g < 0 Then g = 0
        b = CInt(Amount * i / Factor) - 20
        If b < 0 Then b = 0
        f.Line (0, i)-(f.Width, i), RGB(r, g, b)
    Next i
End Sub
Private Sub DrawRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)
Dim bRect As RECT
Dim hBrush As Long
Dim Ret As Long

bRect.Left = x
bRect.Top = y
bRect.Right = x + Width
bRect.Bottom = y + Height

hBrush = CreateSolidBrush(Color)

If OnlyBorder = False Then
    Ret = FillRect(Picture2.hDc, bRect, hBrush)
Else
    Ret = FrameRect(Picture2.hDc, bRect, hBrush)
End If

Ret = DeleteObject(hBrush)
End Sub










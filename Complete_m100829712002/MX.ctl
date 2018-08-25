VERSION 5.00
Begin VB.UserControl MX 
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   585
   ScaleWidth      =   615
   Begin VB.Label Title 
      Caption         =   "MX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "MX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'someone elses code, came with a mailer I got of PSC.

Private Sub UserControl_Initialize()
    SetWinVersion
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width <> 32 Then
        UserControl.Width = 400
    End If
    If UserControl.Height <> 32 Then
        UserControl.Height = 250
    End If
    Title.Left = 0
End Sub

Public Function GetMX() As String
    'THE LINES I (ASHLEY HARRIS) commented out are a PAIN IN THE ARSE!
    '(besides, whod run a webserver behind a proxy anyway?)
    '(my computer is behind an 'adkiller', which this thinks is a proxy, but isn't really.)

    If IsNetConnectOnline = True Then
        'If Not IsNetConnectViaProxy Then
            GetMX = MX_Query
        'Else
        '    Err.Raise 0, "GetMX", "This computer is connected via a proxy server." & vbCrLf & "At this time, the wMX control does not support proxy servers."
        '    Exit Function
        'End If
    Else
    '    Err.Raise 0, "GetMX", "This computer is not currently connected to the internet."
        Debug.Print "How the hell can email be sent while offline?"
        'If GetMX = "" Then GetMX = Domain
    End If
End Function

Public Property Get DNSCount() As Integer
    DNSCount = mi_DNSCount
End Property

Public Property Get MXCount() As Integer
    MXCount = mi_MXCount
End Property

Public Property Get PrefCount() As Integer
    PrefCount = mi_MXCount
End Property

Public Property Get Domain() As String
    Domain = ms_Domain
End Property

Public Property Let Domain(ByVal New_Domain As String)
    If Len(New_Domain) > 4 Then 'its a good host
        ms_Domain = New_Domain
    End If
End Property

Public Function DNS(ByVal Index As String) As String
    DNS = sDNS(Index)
End Function

Public Function mx(ByVal Index As String) As String
    mx = sMX(Index)
End Function

Public Function Pref(ByVal Index As String) As String
    Pref = sPref(Index)
End Function


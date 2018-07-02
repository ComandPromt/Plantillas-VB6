VERSION 5.00
Begin VB.Form FormBlaireau 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crack80"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CheckPrint 
      BackColor       =   &H00000000&
      Caption         =   "Imprimantes ou ressources"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      ToolTipText     =   "Attention... impossible de se connecteraux imprimantes"
      Top             =   600
      Width           =   1935
   End
   Begin VB.ListBox ListPing 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1860
      ItemData        =   "FormBlaireau.frx":0000
      Left            =   6360
      List            =   "FormBlaireau.frx":0002
      TabIndex        =   21
      ToolTipText     =   "Toutes machines présentes"
      Top             =   4560
      Width           =   3255
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Height          =   315
      Left            =   9360
      TabIndex        =   12
      Top             =   4200
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Height          =   315
      Left            =   3600
      TabIndex        =   11
      Top             =   4200
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.ListBox ListAutre 
      BackColor       =   &H00C0C0C0&
      DataField       =   "IP"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1860
      ItemData        =   "FormBlaireau.frx":0004
      Left            =   120
      List            =   "FormBlaireau.frx":0006
      TabIndex        =   20
      ToolTipText     =   "Machines présentes mais non accessibles"
      Top             =   4560
      Width           =   6135
   End
   Begin VB.TextBox TimeOut 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Text            =   "500"
      ToolTipText     =   "Cette valeur permet de fixer la valeur de TimeOut  lors du contrôle d'éxistence de la cible"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton CmdStop 
      BackColor       =   &H00808080&
      Caption         =   "&Stop"
      Height          =   375
      Left            =   6600
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Retour au menu"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox IPd1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox IPd2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox IPd3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox IPd4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox IPa1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Text            =   "254"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton CmdGo 
      BackColor       =   &H00808080&
      Caption         =   "Rechercher les &partages"
      Default         =   -1  'True
      Height          =   495
      Left            =   6600
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Recherche des répertoires et imprimantes partagées"
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox ListConnect 
      BackColor       =   &H00C0C0C0&
      DataField       =   "IP"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      ItemData        =   "FormBlaireau.frx":0008
      Left            =   120
      List            =   "FormBlaireau.frx":000A
      TabIndex        =   13
      ToolTipText     =   "Liste des connexions possibles !"
      Top             =   1440
      Width           =   9495
   End
   Begin VB.CommandButton CmdBye 
      BackColor       =   &H00808080&
      Caption         =   "&Quit"
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Retour au menu"
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton CmdMoi 
      BackColor       =   &H00808080&
      Caption         =   "Ressources déjà &connectées"
      Height          =   495
      Left            =   8160
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Affiche ce a quoi vous etes connecté"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Cibles répondant au ping"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   6360
      TabIndex        =   23
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Cibles windows non partagées"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Cibles windows partagées"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "TimeOut (ms)"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label statut2 
      BackColor       =   &H00404040&
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   7920
      TabIndex        =   17
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "IP de départ :"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "IP de fin :"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label statut 
      BackColor       =   &H00404040&
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6480
      Width           =   7815
   End
End
Attribute VB_Name = "FormBlaireau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdBye_Click()
    FormBlaireau.Visible = False
    Unload Me
    FormMenu.Visible = True
End Sub

Private Sub CmdMoi_Click()
Dim NomMachine As String
Dim nr1 As NETRES2
Dim strDisque As String

    nr1.dwScope = RESOURCE_CONNECTED
    nr1.dwDisplayType = 0
    nr1.dwUsage = 0
    nr1.lpLocalName = ""
    nr1.lpRemoteName = ""
    nr1.lpComment = ""
    nr1.lpProvider = ""
    If CheckPrint.Value = 1 Then
        nr1.dwType = RESOURCETYPE_ANY
    Else
        nr1.dwType = RESOURCETYPE_DISK
    End If
    
    cbBuff = 2048
    cCount = &HFFFFFFFF
    lpBuff = 0
    
    CmdGo.Enabled = False
    bNique = False
    Screen.MousePointer = vbHourglass
    ListConnect.Clear
    ListAutre.Clear
    
    statut.Caption = "Recherche en cours sur " + IPname
    Me.Refresh
    ListConnect.ToolTipText = "Liste de vos connexions actives : cliquer sur un lecteur pour le déconnecter"
        
    res = WNetOpenEnum(nr1.dwScope, nr1.dwType, 0, nr1, hEnum)
    DoEvents
        If res = 0 Then
            lpBuff = GlobalAlloc(GPTR, cbBuff)
            res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
            DoEvents
            If res = 0 Then
                p = lpBuff
                For i = 1 To cCount
                    CopyMemory nr, ByVal p, LenB(nr)
                    p = p + LenB(nr)
                    txtShare = PointerToString(nr.lpRemoteName)
                    strDisque = PointerToString(nr.lpLocalName)
                    ListConnect.AddItem (strDisque & " " & txtShare)
                Next i
            End If
            If lpBuff <> 0 Then
                GlobalFree (lpBuff)
            End If
        End If
    DoEvents
    res = WNetCloseEnum(hEnum)
    ListConnect.Refresh
    statut.Caption = "Recherche terminée !"
    Screen.MousePointer = vbNormal
    CmdGo.Enabled = True
End Sub

Private Sub cmdGo_Click()
Dim NomMachine As String
Dim nr2 As NETRES2
Dim i As Integer
Dim ECHO As ICMP_ECHO_REPLY

    timeout_ping = Val(TimeOut.Text)
    
    CmdMoi.Enabled = False
    CmdStop.Enabled = True
    bNique = True
    bStop = True
    
    Screen.MousePointer = vbHourglass
    ListConnect.Clear
    ListAutre.Clear
    ListPing.Clear
    ListConnect.ToolTipText = "Liste des ressources partagées : cliquer pour se connecter"
    statut.Caption = ""
    
    nr2.dwScope = RESOURCE_GLOBALNET
    nr2.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
    nr2.dwUsage = RESOURCEUSAGE_CONNECTABLE
    nr2.lpLocalName = ""
    nr2.lpRemoteName = ""
    nr2.lpComment = ""
    nr2.lpProvider = ""
    If CheckPrint.Value = 1 Then
        nr2.dwType = RESOURCETYPE_ANY
    Else
        nr2.dwType = RESOURCETYPE_DISK
    End If
    
    For IP = Val(IPd4) To Val(IPa1)
        If bStop = False Then
            Exit For
        End If
        Compose_IP
        ' On regarde si la machine est presente sur le reseau
        statut.Caption = "Recherche en cours... " + IP_dep
        statut.Refresh
        If Val(TimeOut.Text) <> 0 Then
            Call Ping(IP_dep, ECHO)
            statut2.Caption = Str(ECHO.status)
        Else
            ECHO.status = 0
            statut2.Caption = "Pas de ping..."
        End If
        ' Si la machine est presente on recherche les partages
        If ECHO.status = 0 Then
            ' On affiche l adresse pingee si necessaire dans la liste
            If Check2.Value = 1 Then
                ListPing.AddItem (IP_dep & " Ping OK en " & ECHO.RoundTripTime & " ms")
                ListPing.Refresh
            End If
        
        cbBuff = 16384
        cCount = &HFFFFFFFF
        lpBuff = 0
        nr2.lpRemoteName = "\\" & IP_dep
        ' Recherche des connexions
        res = WNetOpenEnum(nr2.dwScope, nr2.dwType, 0, nr2, hEnum)
        'DoEvents
        If res = 0 Then
            lpBuff = GlobalAlloc(GPTR, cbBuff)
            res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
            DoEvents
            If res = 0 Then
                NomMachine = DottedIPToDNS(IP_dep)
                p = lpBuff
                For i = 1 To cCount
                    CopyMemory nr, ByVal p, LenB(nr)
                    p = p + LenB(nr)
                    txtShare = PointerToString(nr.lpRemoteName)
                    ListConnect.AddItem (NomMachine & "     " & txtShare)
                    ListConnect.Refresh
                Next i
            Else
                If Check1.Value = 1 Then
                    NomMachine = DottedIPToDNS(IP_dep)
                    ListAutre.AddItem (NomMachine & " " & IP_dep)
                    ListAutre.Refresh
                End If
            End If
            If lpBuff <> 0 Then
                GlobalFree (lpBuff)
            End If
        Else
            If Check1.Value = 1 Then
                    ListAutre.AddItem (IP_dep & " OpenEnum err " & Str(res))
                    ListAutre.Refresh
                End If
        End If
        DoEvents
        res = WNetCloseEnum(hEnum)
        End If
    Next IP
    statut.Caption = "Recherche terminée !"
    statut2.Caption = ""
    Screen.MousePointer = vbNormal
    CmdMoi.Enabled = True
    CmdStop.Enabled = False
End Sub

Private Sub Compose_IP()
    IP_dep = IPd1 + "." + IPd2 + "." + IPd3 + "." + Mid(Str(IP), 2, Len(Str(IP)))
End Sub
Private Function PointerToString(p As Long) As String
Dim s As String
    s = String(255, Chr$(0))
    CopyPointer2String s, p
    PointerToString = Left(s, InStr(s, Chr$(0)) - 1)
End Function

Private Sub CmdStop_Click()
    bStop = False
End Sub

Private Sub Form_Load()
    ListConnect.Clear
    CmdStop.Enabled = False
    IPd1.Text = IP1
    IPd2.Text = IP2
    IPd3.Text = IP3
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WSACleanup
End Sub

Private Sub IPd1_GotFocus()
    IPd1.SelStart = 0
    IPd1.SelLength = Len(IPd1)
End Sub
Private Sub IPd2_GotFocus()
    IPd2.SelStart = 0
    IPd2.SelLength = Len(IPd2)
End Sub
Private Sub IPd3_GotFocus()
    IPd3.SelStart = 0
    IPd3.SelLength = Len(IPd3)
End Sub
Private Sub IPd4_GotFocus()
    IPd4.SelStart = 0
    IPd4.SelLength = Len(IPd4)
End Sub
Private Sub IPa1_GotFocus()
    IPa1.SelStart = 0
    IPa1.SelLength = Len(IPa1)
End Sub

Private Sub ListAutre_Click()
    ListAutre.ToolTipText = "Non connectable  :-(("
    statut.Caption = ListAutre.Text
End Sub

Private Sub TimeOut_GotFocus()
    TimeOut.SelStart = 0
    TimeOut.SelLength = Len(TimeOut)
End Sub

Private Sub ListConnect_Click()
Dim pos As Integer
Dim IPSelect As String
Dim drive_letter, temp_drive As String * 2
Dim share_name As String
Dim password As String
Dim retour_action As Long
Dim nDisks As Long
Dim compteur As Integer
   
If bNique = True Then           ' On peut connecter un lecteur reseau
    ' Recuperation des Drives disponibles
    nDisks = GetLogicalDrives
    For compteur = 0 To 25
        If (nDisks And 2 ^ compteur) <> 0 Then
            temp_drive = Chr$(65 + compteur) + ":"
        Else
            drive_letter = Chr$(65 + compteur) + ":"
                If drive_letter <> "B:" Then
                    Exit For
                End If
        End If
    Next compteur
    
    pos = InStrRev(ListConnect.Text, "\\", -1)
    share_name = Mid(ListConnect.Text, pos, Len(ListConnect.Text))
    'pos = InStr(1, ListConnect.Text, " ", 1)
    'share_name = "\\" & Mid(ListConnect.Text, 1, pos - 1) & share_name
    password = ""
    
    statut.Caption = "Connexion en cours..."
    Screen.MousePointer = vbHourglass
    DoEvents

    retour_action = WNetAddConnection(share_name, password, _
        drive_letter)
        
    Select Case retour_action
        Case 0
            statut.Caption = "Disque connecté !"
        Case 1326
            statut.Caption = "Désolé... il faut un mot de passe !"
        Case Else
            statut.Caption = "Erreur de connexion : " & Str(retour_action)
    End Select
    
    Screen.MousePointer = vbDefault
Else                        ' On peut deconnecter un lecteur reseau
    If Mid(ListConnect.Text, 2, 1) = ":" Then
        share_name = Mid(ListConnect.Text, 1, 2)
        retour_action = WNetCancelConnection(share_name, True)
        If retour_action > 0 Then
            statut.Caption = "Erreur déconnexion disque : " + Str(retour_action)
        Else
            statut.Caption = "Disque déconnecté !"
            Call CmdMoi_Click
        End If
    Else
        statut.Caption = "Impossible à déconnecter !"
    End If
End If

End Sub

Private Sub ListConnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    statut.Caption = "Crack80 est un programme pour tester VOS sécurités et non celles des autres !"
End If
End Sub
Private Sub Form_Activate()
    Call couleur(Me)
End Sub

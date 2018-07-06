VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAgent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desktop Annoyance"
   ClientHeight    =   4065
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5475
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "agent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5475
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerSpeak 
      Interval        =   60000
      Left            =   4500
      Top             =   3720
   End
   Begin VB.Timer TimerAction 
      Interval        =   60000
      Left            =   4920
      Top             =   3720
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   420
      Top             =   3660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      TabIndex        =   4
      Top             =   3660
      Width           =   1215
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2820
      TabIndex        =   5
      Top             =   3660
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3555
      Left            =   120
      TabIndex        =   21
      Top             =   60
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   6271
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Me"
      TabPicture(0)   =   "agent.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(3)=   "Label15"
      Tab(0).Control(4)=   "txtName"
      Tab(0).Control(5)=   "txtAge"
      Tab(0).Control(6)=   "optGirl"
      Tab(0).Control(7)=   "optBoy"
      Tab(0).Control(8)=   "Picture2"
      Tab(0).Control(9)=   "Picture3"
      Tab(0).Control(10)=   "Picture1"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Sayings"
      TabPicture(1)   =   "agent.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label17"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lstSayings"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdDelete"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdAddSaying"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtAddSaying"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdSpeak2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdClear2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Options"
      TabPicture(2)   =   "agent.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(2)=   "cmdOpen"
      Tab(2).Control(3)=   "lstAnimations"
      Tab(2).Control(4)=   "cmdPlay"
      Tab(2).Control(5)=   "txtSpeech"
      Tab(2).Control(6)=   "cmdSpeak"
      Tab(2).Control(7)=   "cmdClear"
      Tab(2).Control(8)=   "cmdStop"
      Tab(2).Control(9)=   "Frame2"
      Tab(2).Control(10)=   "Frame1"
      Tab(2).Control(11)=   "chkMore"
      Tab(2).ControlCount=   12
      Begin VB.PictureBox Picture1 
         Height          =   1215
         Left            =   -74820
         Picture         =   "agent.frx":0496
         ScaleHeight     =   1155
         ScaleWidth      =   4935
         TabIndex        =   41
         Top             =   2220
         Width           =   4995
      End
      Begin VB.CheckBox chkMore 
         Caption         =   "Add all ""Tell  Me's"" to timer"
         Height          =   300
         Left            =   -72690
         TabIndex        =   40
         Top             =   1380
         Width           =   2895
      End
      Begin VB.PictureBox Picture3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -71820
         Picture         =   "agent.frx":1D068
         ScaleHeight     =   495
         ScaleWidth      =   435
         TabIndex        =   34
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -70560
         Picture         =   "agent.frx":1D4AA
         ScaleHeight     =   495
         ScaleWidth      =   435
         TabIndex        =   33
         Top             =   1560
         Width           =   495
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -72885
         TabIndex        =   28
         Top             =   345
         Width           =   3015
         Begin VB.TextBox txtActionTime 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   14
            Text            =   "5"
            Top             =   150
            Width           =   495
         End
         Begin VB.TextBox txtSpeakTime 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   15
            Text            =   "2"
            Top             =   570
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "minutes"
            Height          =   255
            Left            =   2340
            TabIndex        =   32
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "minutes"
            Height          =   255
            Left            =   2340
            TabIndex        =   31
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Perform an Action every"
            Height          =   255
            Left            =   60
            TabIndex        =   30
            Top             =   195
            Width           =   1815
         End
         Begin VB.Label Label6 
            Caption         =   "Say Something every"
            Height          =   255
            Left            =   60
            TabIndex        =   29
            Top             =   600
            Width           =   1635
         End
      End
      Begin VB.OptionButton optBoy 
         Caption         =   "Boy"
         Height          =   495
         Left            =   -72420
         TabIndex        =   2
         Top             =   1590
         Width           =   615
      End
      Begin VB.OptionButton optGirl 
         Caption         =   "Girl"
         Height          =   435
         Left            =   -71280
         TabIndex        =   3
         Top             =   1620
         Width           =   675
      End
      Begin VB.TextBox txtAge 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74220
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1620
         Width           =   555
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74220
         TabIndex        =   0
         Top             =   960
         Width           =   4155
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear"
         Height          =   315
         Left            =   4320
         TabIndex        =   10
         Top             =   3060
         Width           =   795
      End
      Begin VB.CommandButton cmdSpeak2 
         Caption         =   "Speak"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   7
         Top             =   2220
         Width           =   795
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -72900
         TabIndex        =   24
         Top             =   1605
         Width           =   3015
         Begin VB.CommandButton cmdMore 
            Caption         =   "Advanced"
            Height          =   405
            Left            =   960
            TabIndex        =   39
            Top             =   300
            Width           =   1305
         End
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   -73920
         TabIndex        =   18
         Top             =   3120
         Width           =   795
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   315
         Left            =   -72120
         TabIndex        =   19
         Top             =   3120
         Width           =   795
      End
      Begin VB.TextBox txtAddSaying 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   4995
      End
      Begin VB.CommandButton cmdAddSaying 
         Caption         =   "Add"
         Height          =   315
         Left            =   3480
         TabIndex        =   11
         Top             =   3060
         Width           =   795
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   315
         Left            =   4320
         TabIndex        =   8
         Top             =   2220
         Width           =   795
      End
      Begin VB.ListBox lstSayings 
         Height          =   1185
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   4995
      End
      Begin VB.CommandButton cmdSpeak 
         Caption         =   "Speak"
         Height          =   315
         Left            =   -71280
         TabIndex        =   20
         Top             =   3120
         Width           =   795
      End
      Begin VB.TextBox txtSpeech 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72900
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   2700
         Width           =   3015
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   315
         Left            =   -74760
         TabIndex        =   17
         Top             =   3120
         Width           =   795
      End
      Begin VB.ListBox lstAnimations 
         Height          =   1860
         Left            =   -74820
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open Character"
         Height          =   435
         Left            =   -74640
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Add Saying:"
         Height          =   255
         Left            =   180
         TabIndex        =   38
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Manage the sayings of your character."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   300
         TabIndex        =   37
         Top             =   480
         Width           =   4635
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Help me interact with you by answering a few questions."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   36
         Top             =   540
         Width           =   4935
      End
      Begin VB.Label Label10 
         Caption         =   "Add your own sayings:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   35
         Top             =   1920
         Width           =   2115
      End
      Begin VB.Label Label9 
         Caption         =   "I am a"
         Height          =   255
         Left            =   -73020
         TabIndex        =   27
         Top             =   1710
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Age:"
         Height          =   315
         Left            =   -74760
         TabIndex        =   26
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Name:"
         Height          =   315
         Left            =   -74760
         TabIndex        =   25
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Say:"
         Height          =   255
         Left            =   -72900
         TabIndex        =   23
         Top             =   2460
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Character Animations"
         Height          =   255
         Left            =   -74820
         TabIndex        =   22
         Top             =   960
         Width           =   1875
      End
   End
   Begin AgentObjectsCtl.Agent AgentControl 
      Left            =   60
      Top             =   3600
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences"
      End
      Begin VB.Menu FileSept1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpeak 
         Caption         =   "Tell Me"
         Begin VB.Menu mnuMySayings 
            Caption         =   "my Sayings"
         End
         Begin VB.Menu mnuJoke 
            Caption         =   "a Joke"
         End
         Begin VB.Menu mnuFact 
            Caption         =   "a Fact"
         End
         Begin VB.Menu mnuQuote 
            Caption         =   "a Quote"
         End
         Begin VB.Menu mnuMurphy 
            Caption         =   "a Murphy Law"
         End
         Begin VB.Menu mnuOxy 
            Caption         =   "an OxyMoron"
         End
         Begin VB.Menu mnuPondering 
            Caption         =   "a Pondering"
         End
         Begin VB.Menu mnuTime 
            Caption         =   "the Time"
         End
      End
      Begin VB.Menu mnuRead 
         Caption         =   "Read From"
         Begin VB.Menu mnuClipboard 
            Caption         =   "Clipboard"
         End
      End
      Begin VB.Menu mnuPerform 
         Caption         =   "Perform"
         Begin VB.Menu mnuRandom 
            Caption         =   "Random"
         End
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu FileSept2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelpText 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Desktop Annoyance"
      End
   End
End
Attribute VB_Name = "frmAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**used for timing actions
Dim intActionTime As Integer
Dim aMinutes As Integer

'**used for timing speaking
Dim intSpeakTime As Integer
Dim sMinutes As Integer

Private Sub AgentControl_Click(ByVal CharacterID As String, ByVal Button As Integer, _
                        ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    
    '**when the agent is clicked (right opens menu)
    If Button = vbRightButton Then frmAgent.PopupMenu mnuFile
    If Button = vbLeftButton Then Char.Play "Restpose"

End Sub

Private Sub cmdAddSaying_Click()
    If txtAddSaying.Text <> "" Then
        lstSayings.AddItem txtAddSaying.Text
        txtAddSaying.Text = ""
    End If
End Sub

Private Sub cmdDelete_Click()
    '** remove saying from list
    Dim i As Integer
    For i = lstSayings.ListCount - 1 To 0 Step -1
        If lstSayings.Selected(i) Then
            lstSayings.RemoveItem (i)
        End If
    Next i
End Sub

Private Sub cmdMore_Click()
    '**opens agents property sheet
    AgentControl.PropertySheet.Visible = True
End Sub

Private Sub cmdOpen_Click()

    Dim DirName As String
    Dim varAnimation As Variant
    DirName = GetWindowsDir()
    '**gets windows directory and adds magent\char to it for looking up characters
    With CommonDialog
        .InitDir = DirName + "msagent\chars"
        .Filter = "*.acs"
        .FilterIndex = 1
        .ShowOpen
    End With
    
    '**for open dialog window
    If CommonDialog.FileName = "" Then
       Exit Sub
    End If
    
    AgentControl.Characters.Unload ("CharacterID")                      '**unload previous char
    AgentControl.Characters.Load "CharacterID", CommonDialog.FileName   '**load char
    Set Char = AgentControl.Characters("CharacterID")                   '**set char id to "char"
    Char.LanguageID = &H409                                             '**english
    Char.Show                                                           '**show char
    
    lstAnimations.Clear
    For Each varAnimation In Char.AnimationNames                        '**fill list with
        lstAnimations.AddItem varAnimation                              '**animations
    Next
    
End Sub

Private Sub cmdPlay_Click()
    '**plays animation from list
    If lstAnimations.SelCount >= 1 Then
        With Char
            .Play lstAnimations.List(lstAnimations.ListIndex)
            .Play "RestPose"
        End With
    End If
End Sub

Private Sub Form_Activate()
    '**show char upon activation of form
    Char.Show
End Sub

Private Sub Form_Load()
      
    '**set timer minutes to 1
    sMinutes = 1
    aMinutes = 1
    
    Dim varAnimation As Variant
    Dim fileHandle As Integer
    Dim strFile As String
    ReDim arrSayings(0)
    ReDim arrActions(0)
    ReDim arrMe(0)
   
    '**used in conjuction with rnd function
    Randomize
     
    Open App.Path & "\msagent.lst" For Input As #1  ' Open file for input.
    Do While Not EOF(1)                             '** Loop until end of file.
        Input #1, strFile                           '** Read data into two variables.
    Loop
    Close #1                                        '** Close file.
    
    '**load char saved in msagent.lst
    AgentControl.Characters.Load "CharacterID", strFile & ".acs"
    
    '**set char language and id <- id is given
    Set Char = AgentControl.Characters("CharacterID")
    Char.LanguageID = &H409
    Char.AutoPopupMenu = False
    
    'list animations
    For Each varAnimation In Char.AnimationNames
        lstAnimations.AddItem varAnimation
        ReDim Preserve arrActions(UBound(arrActions) + 1)
        arrActions(UBound(arrActions)) = varAnimation
    Next
    
    '**gets user sayings and fills them into array
    '** use ubound to find count of last item in array.  will
    '**     use it later for 'pick random number 1 to upper bound count'
    Open "C:\program files\DesktopAnnoyance\sayings.lst" For Input As #1
    
    Do While Not EOF(1)
        Input #1, strFile
        lstSayings.AddItem (strFile)
        ReDim Preserve arrSayings(UBound(arrSayings) + 1)
        arrSayings(UBound(arrSayings)) = strFile
    Loop
    Close #1
    
    
    'gets personal info and fills into array
    Open "C:\program files\DesktopAnnoyance\me.lst" For Input As #1
    
    Do While Not EOF(1)
        Input #1, strFile
        ReDim Preserve arrMe(UBound(arrMe) + 1)
        arrMe(UBound(arrMe)) = strFile
    Loop
    Close #1
    
    
    '**loads user info
    txtName.Text = arrMe(1)
    txtAge.Text = arrMe(2)
    If arrMe(3) = "boy" Then
        optBoy = True
    Else
        optGirl = True
    End If
        
    intActionTime = txtActionTime.Text
    intSpeakTime = txtSpeakTime.Text
        
    Char.Show
    PlayIntro
    
End Sub

Private Sub cmdExit_Click()
    'save user sayings and me info
    Dim i As Integer
    Dim strSex As String
    ReDim arrMe(3)          'only three items - orginally more but didnt feel like adding them
    
    If optBoy = True Then
        strSex = "boy"
    Else
        strSex = "girl"
    End If
    
    arrMe(1) = txtName.Text
    arrMe(2) = txtAge.Text
    arrMe(3) = strSex
        
    'save my settings
    Open "C:\program files\DesktopAnnoyance\me.lst" For Output As #1
    For i = 1 To 3
        Write #1, arrMe(i)
    Next i
    Close #1
    
    'save my sayings
    Open "C:\program files\DesktopAnnoyance\sayings.lst" For Output As #1
        For i = lstSayings.ListCount - 1 To 0 Step -1
            Write #1, lstSayings.List(i)
        Next i
    Close #1
    
    Open "C:\program files\DesktopAnnoyance\msagent.lst" For Output As #1
        Write #1, Char.Name
    Close #1
    
    AgentControl.Characters.Unload "CharacterID"
    Unload Me
    End

End Sub

Private Sub lstSayings_Click()
    cmdSpeak2.Enabled = True
End Sub

Private Sub mnuAbout_Click()
    frmMore.Show
End Sub

Private Sub mnuClipboard_Click()
    Char.Speak Clipboard.GetText
End Sub

Private Sub mnuFact_Click()
    RandomFact
End Sub

Private Sub mnuJoke_Click()
    RandomJokes
End Sub

Private Sub mnuMurphy_Click()
    RandomMurphy
End Sub

Private Sub mnuMySayings_Click()
    RandomSpeak
End Sub

Private Sub mnuOxy_Click()
    RandomOxy
End Sub

Private Sub mnuPondering_Click()
    RandomPonder
End Sub

Private Sub mnuQuote_Click()
    RandomQuote
End Sub

Private Sub mnuRandom_Click()
    RandomMove
End Sub

Private Sub mnuStop_Click()
    Char.Stop
End Sub

Private Sub TimerAction_Timer()
    'play when users time is fired
    If aMinutes = intActionTime Then
        RandomMove
        aMinutes = 0
    Else
        aMinutes = aMinutes + 1
    End If

End Sub

Private Sub TimerSpeak_Timer()
    'play when users time is fired
    '**checks if check box for all menu items clicked
    '** it then randomly picks what function to perform
    If sMinutes = intSpeakTime Then
        If chkMore = 0 Then
            RandomSpeak
        Else
            Dim intX As Integer
            intX = Rnd * 7
            Select Case intX
            Case 1
                RandomSpeak
            Case 2
                RandomQuote
            Case 3
                RandomMurphy
            Case 4
                RandomFact
            Case 5
                RandomOxy
            Case 6
                RandomPonder
            Case 7
                RandomJokes
            Case Else
                RandomSpeak
            End Select
        End If
    Else
        '**if has not reached set time add 1 minute to counter
        '** will continue until it reaches final count
        sMinutes = sMinutes + 1
    End If

End Sub

Private Sub txtActionTime_Change()
    '**sets timer increment
    If txtActionTime.Text <> "" Then
        intActionTime = txtActionTime.Text
    End If
End Sub

Private Sub txtActionTime_KeyPress(KeyAscii As Integer)
   KeyAscii = KeyCheck(KeyAscii)
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyCheck(KeyAscii)
End Sub

Private Sub txtSpeakTime_Change()
    '**sets timer increment
    If txtSpeakTime.Text <> "" Then
        intSpeakTime = txtSpeakTime.Text
    End If
End Sub

Private Sub mnuTime_Click()
    Char.Speak "The current time is " & Time
    IdleOn
End Sub

Private Sub cmdClear_Click()
    txtSpeech.Text = ""
End Sub

Private Sub cmdClear2_Click()
    txtAddSaying.Text = ""
End Sub

Private Sub cmdHide_Click()
    frmAgent.Hide
End Sub

Private Sub cmdSpeak_Click()
    '**say whatever is in textbox
    If txtSpeech.Text <> "" Then
        Char.Speak txtSpeech.Text
        IdleOn
    End If
End Sub

Private Sub cmdSpeak2_Click()
    '**say whatever is in textbox
    Char.Speak lstSayings.Text
    IdleOn
    cmdSpeak2.Enabled = False
End Sub

Private Sub cmdStop_Click()
    Char.Stop
End Sub

Private Sub lstAnimations_DblClick()
    Char.Play lstAnimations.Text
    IdleOn
End Sub

Private Sub lstSayings_DblClick()
    Char.Speak lstSayings.Text
    IdleOn
End Sub

Private Sub mnuExit_Click()
    cmdExit_Click
End Sub

Private Sub mnuPreferences_Click()
    frmAgent.Show
    Char.Show
End Sub

Private Sub txtSpeakTime_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyCheck(KeyAscii)
End Sub

Attribute VB_Name = "Module1"
Option Explicit

'Declare structure types: ships, players, planets

Type StarShip
    Troops As Integer
    AssaultTroops As Integer
    '
    HaveShields As Boolean
    
    HaveShield1 As Boolean
    HaveShield2 As Boolean
    
    HaveWeapons As Boolean
    HaveBigShip As Boolean
    HaveCloakingDevice As Boolean
    Hidden As Boolean
    Sabotage As Boolean
    '
    CombatStrength As Integer
    WarpPosition As Integer     '0 if not in use, 1-8 if on warp path
    Coordinate As String        'coord of originating planet
    CenterX As Integer          'used to draw line from starting
    CenterY As Integer          'planet to possible destinations
    Launched As Boolean         'if launched, disable button to launch again
    ShipNumber As Integer       '0 or 1 for array of ship(2) in player structure
End Type

Type PlanetaryBody
    Name As String
    Owner As Integer            '0,1,2 for neutral, player 1, player 2
    Troops As Integer
    AssaultTroops As Integer
    '
    HaveShields As Boolean
    HaveMissiles As Boolean     'CS depends on whether player has
                                'missile1 or 2 researched...

    HaveScanner As Boolean
    HaveJammer As Boolean
    InScannerRange As Boolean   'if in range of scanner - replaces sharing of "inrange" with landing
    '
    BioRocketETA As Integer
    BioDistance As Variant
    Contaminated As Boolean     'toxic to humans
    NukedResources As Boolean   'resources wiped out by biorocket
    InBioRange As Boolean       'if in range for biorocket
    LaunchSite As Boolean       'which planet rocket firing from - used
                                'to erase the lines!
    BioFailed As Boolean        'if enemy tried to use biorocket but it was shot down
    
    '
    CenterX As Integer
    CenterY As Integer
    '
    ImprovedResources As Boolean
    CombatStrength As Integer
    Coordinate As String
    Resources As Integer        'random # resources produced per turn
    InRange As Boolean          'if in range for landing
    BackGround As Integer       '1-5 for landscape picture
    Picture As Integer          '1-5 for galaxy picture - used to restore after contamination
    JustLanded As Boolean       'don't let ship take off after landing on a new planet
    '
    Sabotaged As Boolean            'in frmGameScreen, SabotageLanding, to give other player
    SabotageReduction As Integer    'a message the next turn
    SabotagedFactory As Boolean     'if factory destroyed in sabotage mission
    Damaged As Boolean              'enables cmdRepairIndustry button
                                    'will add 2-4 (up to 8), set to false once resources up to 4
    Captured As Boolean         'for giving a message if other player took over planet on his/her turn
    '
    FailedInvasion As Boolean               'to notify player that opponent tried and failed to attack planet
    FailedInvasionTroopLosses As Integer    'show losses suffered in defense of planet
    FailedInvasionMechLosses As Integer
    '
End Type

Type PlayerType
    Name As String
    NumTroops As Integer            'total # of troops
    NumAssaultTroops As Integer
    NumPlanets As Integer           'total # of planets owned
    NumResources As Integer         'total of all resources
    Ship(2) As StarShip
    HomePlanet As Integer           'index number of home planet
    WasBig As Boolean               'announce if empire shrinking
    
    '***Research***
    Missile1ResearchDone As Integer
    Missile1Researched As Boolean
    
    Missile2ResearchDone As Integer
    Missile2Researched As Boolean

    ShieldResearchDone As Integer    'time to research planetary shields
    ShieldResearched As Boolean      'whether this tech researched

    LaserResearchDone As Integer
    LaserResearched As Boolean
    
    PlasmaResearchDone As Integer
    PlasmaResearched As Boolean
    
    MechResearchDone As Integer     'time to research assault tech
    MechResearched As Boolean       'whether or not player has researched this tech
    
    BioRocketResearchDone As Integer
    BioRocketResearched As Boolean
    
    LongBioResearchDone As Integer
    LongBioResearched As Boolean
    
    ShipShield1ResearchDone As Integer
    ShipShield1Researched As Boolean
    
    ShipShield2ResearchDone As Integer
    ShipShield2Researched As Boolean
    
    BigShipResearchDone As Integer
    BigShipResearched As Boolean
    
    UltraWarpResearchDone As Integer
    UltraWarpResearched As Boolean
    
    CloakingResearchDone As Integer
    CloakingResearched As Boolean
     
    ResourceResearchDone As Integer     'time to research resource improvement
    ResourcesResearched As Boolean      'whether player has researched this tech
    
    BioCleanupResearchDone As Integer
    BioCleanupResearched As Boolean
    
    RegenerationResearchDone As Integer
    RegenerationResearched As Boolean
   
    ScannerResearchDone As Integer
    ScannerResearched As Boolean
    
    DeepScannerResearchDone As Integer
    DeepScannerResearched As Boolean
    
    JammerResearchDone As Integer
    JammerResearched As Boolean
    
    WarpScannerResearchDone As Integer
    WarpScannerResearched As Boolean
    
    '*****
    Message1Given As Boolean        'warning when enemy has 10 planets
    Message2Given As Boolean        'warning when enemy has 20 planets
End Type

'Declare variables of defined types
Public Player(2) As PlayerType
Public Planet(50) As PlanetaryBody

'Declare Global Variables
Public Const Modal = 1              'used to display forms as modal/nonmodal
Public Const NonModal = 0

Public gFileNum As Integer          'for reading/writing files
Public TurnNumber As Integer

'*****************************
Public Current As Integer           '0 or 1, use to see whose turn it is
Public Other As Integer             '0 or 1
Public Const Neutral = 2            'As Integer    '2 for now...
Public Const Alien = 3              'As Integer    '3 for now
'*****************************
Public ActivePlanet                 'used for buying stuff
Public UnitCost                     '
Public Quantity                     '
Public PurchasePrice                '
'*****************************
Public IncomingMessage As String    'for the messages players
Public OutgoingMessage As String    'send to each other

Public activeship As Integer        'either 0 or 1 for ship number

'************
'Combat:
Public AttackStrength As Integer
Public DefenceStrength As Integer
Public Losses As Integer
Public TroopLosses As Integer
Public AssaultLosses As Integer

Public NumPlanetsCaptured As Integer    'counter to see how many planets of other player taken over this turn

Public NumFailedInvasions As Integer    'counter for failed invasions by other player
'************

Public TechLevel As Integer         'for setting up techdone form
                                    'sets which title, image and description to show
                             
Public ReadyToLand1 As Boolean      'toggle for land1 button
Public ReadyToLand2 As Boolean

Public ScannerOn As Boolean         'toggle for mousemove, how much info shown
                                    'in txtstatus.text for planets
                                    
Public BioRocketOn As Boolean       'toggle for biorocket launch button
Public RegenerateOn As Boolean      'toggle
Public DetoxifyOn As Boolean        'toggle

'Sound API
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Sound As String
Public SoundOn As Boolean           'for sound toggle

'************************
Public GalaxySize As Integer        'for different number of planets E/M/H
'************************
Public MessageType As String        'to set picture and text for Message Form
                                    'ie. biorocket explosion, planet overrun
Public Announceline1 As String      'to set text for beeping text announcement
Public Announceline2 As String
Public Announceline3 As String

Public LoadCancelled As Boolean     'if player aborts loading a game
                                    'this goes back to frmcover
                                  
Public GameNumber As Integer        'to set the proper default game number when saving
                                    'so people don't save the wrong game # by mistake
            
Public GameName As String           'to show the user after compressing

'to disable the X button on the control box
Public Declare Function GetSystemMenu Lib "User32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Long) As Long
Public Declare Function RemoveMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "User32" (ByVal hWnd As Long) As Long

Public Const MF_BYPOSITION = &H400&
Public Const MF_DISABLED = &H2&
'**************************************

'NOT IMPLEMENTED - To check if Install.log exists - if not, have to register zlibtool.ocx
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public sysDir As String

'***NEW for March 99***config file
Public DefaultGameSize As Integer       'to set size of playing screen: 0 is normal, 2 is maximized
Public DefaultGameSound As Boolean      'set default sound setting (on/off)


Sub main()

Randomize

'I added a zlib dll control to frmCover, to check for registration errors right
'at the start - if it can't load frmCover, that's because of the zlib control
'on the form.  As noted elsewhere, for some reason the install program is not
'registering the custom control zlibOCX2.dll, which compresses the saved game info.

'If there's an error, the user gets a messagebox explaining that
'they have to run the program 'regfile', which is included in the install package.
'The regfile program checks to make sure the user has regsvr32.exe and zlibOCX2.dll in the
'windows\system directory - if not, another messagebox tells them this.
'Alternatively, the user could run regsvr32 manually - that's still covered
'in the readme file...


On Error GoTo ZipError

frmCover.Show
On Error GoTo 0

Exit Sub


ZipError:
    Dim ErrorMessage As String
    ErrorMessage = "The install program did not properly register a required file." + Chr(13)
    ErrorMessage = ErrorMessage + "Please run 'regfile.exe', located in the game directory." + Chr(13) + Chr(13)
    ErrorMessage = ErrorMessage + "If you're still having trouble, please refer to the" + Chr(13)
    ErrorMessage = ErrorMessage + "Troubleshooting section of the Readme file."
    
    MsgBox ErrorMessage, vbOKOnly + vbInformation, "A Slight Installation Problem..."
    End

End Sub



Public Sub SetCombatStrength(i As Integer)
'run after: launch, buy troops/assault troops, buy shields/missiles, battles
Dim X, Y

Planet(i).CombatStrength = 0
X = Planet(i).Troops
Y = Planet(i).AssaultTroops

'laser and plasma for troops?
If Player(Current).LaserResearched Then
    X = Int(X * 1.15)
End If

If Player(Current).PlasmaResearched Then
    X = Int(X * 1.3)
End If

'****Main Formula:
Planet(i).CombatStrength = X + (Y * 5)
'****

If Planet(i).HaveMissiles Then
    If Player(Current).Missile1Researched Then
        Planet(i).CombatStrength = Planet(i).CombatStrength + 5
    End If

    If Player(Current).Missile2Researched Then
        Planet(i).CombatStrength = Planet(i).CombatStrength + 7
    End If
End If

If Planet(i).HaveShields Then
    'add 25 to CS
    Planet(i).CombatStrength = Planet(i).CombatStrength + 25
End If

End Sub

Public Sub playBeep()
On Error Resume Next
If SoundOn Then
    Sound = App.Path + "\Button1.wav"
    sndPlaySound Sound, 3
    On Error GoTo 0
End If

End Sub


Public Sub PlaySoundEffect(Sound As String)
'play soundeffects as called throughout program
'uses the sndPlaySound API, declared in Declarations section of this module

On Error GoTo ErrHandler

If SoundOn And Sound <> "" Then

    Select Case Sound
        Case "Abort"
            Sound = App.Path + "\badnews.wav"
            sndPlaySound Sound, 0

        Case "Access"
            Sound = App.Path + "\access.wav"
            sndPlaySound Sound, 3
  
        Case "Ambient1"
            Sound = App.Path + "\ambient1.wav"
            sndPlaySound Sound, 3
    
        Case "Ambient3"
            Sound = App.Path + "\ambient3.wav"
            sndPlaySound Sound, 3

        Case "Attack"
            Sound = App.Path + "\attack.wav"
            sndPlaySound Sound, 3

        Case "BioFail"
            Sound = App.Path + "\biofail.wav"
            sndPlaySound Sound, 3

        Case "Button1"
            Sound = App.Path + "\button1.wav"
            sndPlaySound Sound, 3

        Case "Button2"
            Sound = App.Path + "\button2.wav"
            sndPlaySound Sound, 3

        Case "Button3"
            Sound = App.Path + "\button3.wav"
            sndPlaySound Sound, 3

        Case "Button4"
            Sound = App.Path + "\button4.wav"
            sndPlaySound Sound, 3

        Case "Button5"
            Sound = App.Path + "\button5.wav"
            sndPlaySound Sound, 3

        Case "Detonate"
            Sound = App.Path + "\detonate.wav"
            sndPlaySound Sound, 3

        Case "Disintegrate"
            Sound = App.Path + "\disintegrate.wav"
            sndPlaySound Sound, 3

        Case "Explosion"
            Sound = App.Path + "\explode.wav"
            sndPlaySound Sound, 3

        Case "Intro"
            Sound = App.Path + "\intro.wav"
            sndPlaySound Sound, 3

        Case "Launch"
            Sound = App.Path + "\launch3.wav"
            sndPlaySound Sound, 3
    
        Case "Overrun"
            Sound = App.Path + "\overrun.wav"
            sndPlaySound Sound, 3

        Case "Quiet"
            Sound = App.Path + "\quiet.wav"
            sndPlaySound Sound, 3
    
        Case "Research"
            Sound = App.Path + "\research.wav"
            sndPlaySound Sound, 3

        Case "Sabotage"
            Sound = App.Path + "\sabotage.wav"
            sndPlaySound Sound, 3
    
        Case "Warning"
            Sound = App.Path + "\warning.wav"
            sndPlaySound Sound, 3

    End Select

Else
    'sound is off, so don't let Windows use default button sounds
    Sound = App.Path + "\quiet.wav"
    sndPlaySound Sound, 3
    
End If


Exit Sub


ErrHandler:
    'error playing sound
    Exit Sub
    

End Sub

Public Sub PlayRandomSound()
'plays a random sound effect every 30 seconds
'activated by tmrRandomSounds in frmGameScreen

On Error GoTo SoundError

If SoundOn Then
  Randomize
  Dim X As Integer

  X = Int(Rnd * 6) + 1

  Select Case X
  Case 1
    Sound = App.Path + "\short1.wav"
    sndPlaySound Sound, 3
  Case 2
    Sound = App.Path + "\short2.wav"
    sndPlaySound Sound, 3
  Case 3
    Sound = App.Path + "\short3.wav"
    sndPlaySound Sound, 3
  Case 4
    Sound = App.Path + "\short4.wav"
    sndPlaySound Sound, 3
  Case 5
    Sound = App.Path + "\short5.wav"
    sndPlaySound Sound, 3
  Case 6
    Sound = App.Path + "\short6.wav"
    sndPlaySound Sound, 3
  End Select
Else
    Exit Sub
End If

Exit Sub


SoundError:
    'sound problem - exit sub
    Exit Sub

End Sub

Public Sub ShowQuickHelp()
'show quick help when user hits escape key
        
Dim message As String

message = "Function Key Commands:" & Chr(13) & Chr(13)
message = message + "Esc   This Quick Help Screen" + Chr(13)
message = message + "F1     4000 A.D. Help File" + Chr(13)
message = message + "F2     Save Game" + Chr(13)
message = message + "F3     Abort Game" + Chr(13)
message = message + "F4     Toggle Grid Lines On/Off" + Chr(13)
message = message + "F5     Toggle Sound On/Off" + Chr(13)
message = message + "F6     Configure Game Options" + Chr(13)
        
MsgBox message, vbOKOnly, "4000 A.D. Quick Help"


End Sub

Public Sub DisableX(Frm As Form)
    
Dim hMenu As Long, nCount As Long
    
'Get handle to system menu
hMenu = GetSystemMenu(Frm.hWnd, 0)

'Get number of items in menu
nCount = GetMenuItemCount(hMenu)
    
'Remove last item from system menu (last item is 'Close')
Call RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)

'Redraw menu
DrawMenuBar Frm.hWnd


End Sub

Public Function FileExists%(FullPathAndFile$)
'NOT IMPLEMENTED
'used to see if Install.log exists, showing that the
'zlibtool.ocx has been registered

On Error Resume Next

If FileLen(FullPathAndFile) > 0& Then
     If Err = 0 Then FileExists = True
End If

End Function

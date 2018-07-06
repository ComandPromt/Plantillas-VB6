VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsManifestation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*NAME:ClsManifestation version 2
'*CREATOR:Roger Gilchrist
'*DATE:    14-Jul-2002
'MODIFIED :  20-Jul-2002 added more control types

'HEY I know this is old hat but I only wnet XP a couple of weeks ago so it's new to me
'
'see the WARNINGS at end of this comment
'
'This code is ©2002 By Roger Gilchrist {rojagilkrist@hotmail.com}
'________________________________________________________________________________
'       NOTE this class is released in DEMO formation.
'       if you got it with my CLsExtendedRTF demo then the Demo stuff is still here
'       but has been turned off in Class Initialize
'       Search for "TO REMOVE DEMO MODE" and follow the notes
'       at each point to remove the demo aspects of the class
'       Search for "SPECIALISED CONTROL" to edit out controls you are not using
'________________________________________________________________________________

'*PURPOSE: Encapsulate routines for creating/removing and using manifest files
'          and allowing end users to decide whether or not to use the full XP theme look without
'          getting into the messy stuff.

' STRUCTURE:PUBLIC Routines
'           Manifest: If there is a manifest file then use it
'           Action: If there is a manifest delete it else create it
'           CheckBox:         \
'           CommandButton:     |
'           ListBox:           |
'           Menu1Item:         | Control(s) to fire Action.
'           Menu2Items:        | Control.Parent.Hwnd needed for the LaunchSelf routine
'           OptionButtons:     | For use in own program edit out the ones you dont want
'           ToolBarButton      |
'           ToolBarButtonMenu:/
'
'           PRIVATE Routines
'           ActionResponse: Offer to restart program immediately. If Yes launches new and close present instance.
'           Caption : set the caption for the control(also tests and disables if not WIndowsXP)
'           Content: the Contents string for the manifest file
'           CreateManifest: write Contents to a Manifest file
'           HasAManifest: test existence of Manifest file
'           IsWindowXP: test that operating system is XP
'           LaunchSelf: If ActionResponse is Yes launches a new instance of program.
'           ManifestFile: Creates a name for the manifest file
'           PairedCheckControls: Centralises setting pairs of controls
'           RefreshMe: updates the caption/image/check for control after taking action
'           RemoveManifest: Delete the manifest file
'           SingleCheckControl: Centralises setting a checkable control
'           SingleToggleControl: Centralises setting a toggling control
'
'           DEMO Routines
'           Search for "TO REMOVE DEMO MODE" and follow the notes
'           at each point to remove the demo aspects of the class
'           These routines can be deleted completely as you probably
'           never want to give then to an end-user
'           DemoDisplayManifest: Display the Private routine Content
'           DemoVB6Manifest:     Display and Create a VB6.exe.manifest file
'
'*USAGE: 1. Add the class to your program
'                   (create a new copy and delete DEMO parts and
'                       any of the public controls you don't want)
'
'        2. Place the following line to the top of any code module
'            (you can of course call it whatever you like)
'
'                 Public MyManifest As New ClsManifestation

'        3. Place the following in Sub Form_Initialize() of the first form that
'               your program loads.
'
'                MyManifest.Manifest
'       4. Place your Control where ever you want it in your program.
'           For all controls the class provides the captions.
'           You can caption the button whatever you like for identifying it in the IDE
'           For ListBox make sure you set Style to CheckBox. You have to supply the other items
'           for the ListBox and leave a blank or dummy value for the class to write over.
'           You should fill the Listbox before calling the Class.ListBox. Add a Boolean value
'           to stop the ListBox_Click from triggering while you are filling the list so that
'           if you fill the list a 2nd time you don't trigger Class.Action by accident.
'           (See Demo for details)
'           If you need different Captions for Localisation or you just don't like my captions
'           search for "LOCALIZATION" in code
'           and change the Constants WinXP, Classic, NotXP and ToolTip.
'
'       5. Place ONE of the following in Sub Form_Load() of the form containing the Control
'
'               MyManifest.CheckBox SomeChkBox
'               MyManifest.CommandButton SomeCmdButton
'               MyManifest.ListBox SomeLstBox,ItemNumber
'               MyManifest.Menu1item Menuname
'               MyManifest.Menu2items Menuname(0), Menuname(1)
'               MyManifest.Menu1item SubMenuname
'               MyManifest.Menu2items SubMenuname(0), SubMenuname(1)
'               MyManifest.OptionButtons  SomeOptButton(0),SomeOptButton(1)
'               MyManifest.ToolBarButton Toolbar, "ButtonKey", [True|False], [True|False]
'               MyManifest.ToolBarButtonMenu Toolbar, "ButtonName",, [True|False], [True|False], [True|False], "key1", "key2"
'
'       6. Place the following in SomeControl_Click  or SomeControl_Click(Index as Integer)
'           (and for Control Arrays place within the Index or Key test structure)
'
'               MyManifest.Action
'
'          All controls toggle their own appearance but
'          if you are using multiple controls you need to call each
'          by repeating a call to whatever you coded in step 5
'
'       7. That's it. Your program can now add/remove and (if your in XP) use XP theme styles
'
'               WARNING 1 remember to fix the Black Outline on Frames problem
'                        Place any optionbutton in a frame on a PictureBox in the frame
'               WARNING 2 CommandButtons with Style=Graphical are visible but remain Classic in appearance
'                         They just look ugly on an XP style form
'
' If you can think of another Standard or Professional VB control which might be used let me know
'
'This code is ©2002 By Roger Gilchrist {rojagilkrist@hotmail.com}
'and freely available to use with acknowledgment.
'Please leave copyright comments in code when editing out the demo stuff.
'
'Thanks to Ulli's VB Code Formatter V2.12.7 at Planet Source Code for tidy layout
'Barry Garvin 101 VBPJ Tech Tips 11 March 2001 p1 for Enforce Case for Enums
' *********************************************************************

Option Explicit

'I'm a bit compulsive about modularity so each module has its own
'Private calls to to the API. This means if you copy just one module
'you don't have to keep coming back to find that missing Public variable
'in another module. Once you have inculde a module in your program you can
'remove the Private Declare statements if you have a Public one in your project

Private Declare Sub InitCommonControls Lib "comctl32.dll" () 'load the XP Common controls
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                          ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
                          ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long ' used to launch a new instance of program
Private Const SW_SHOWNORMAL As Integer = 1

Private m_Control As Control   'The button for end-user access to the class
Private m_Control2 As Control   'Used by OptionButtons/Menu/ToolBar

Private m_Name As String        'You can set these 3 elements of the manifest content but see
Private m_Description As String 'Class_Initialize for automatic values
Private m_Version As String     'and why you may as well use them

Private m_DemoMode As Boolean   'TO REMOVE DEMO MODE 'comment out/delete this line

Private Const NL As String = vbNewLine  'These constants make complex strings easier to read
Private Const DQ As String = """"       'NL is short and DQ is much prettier than Chr$(34) isn't it?

Private SettingValue As Boolean ' prevent Action firing when you set the Check value of controls

'Refresh uses these to allow Action to reset the caption/image for a control
'SPECIALISED CONTROL remove the sections which will not be needed for your program
Private m_Checked As Boolean
Private m_Index1 As Variant
Private m_Index2 As Variant
Private m_Icn As Boolean
Private m_Txt As Boolean
Private m_ToolBarStyle As Integer
Private m_ButtonKey As String
Private m_SubToggle As Boolean
Private m_Inverse As Boolean
Private m_MenuStyle As Integer

'LOCALIZATION change these to fit your language/needs 'ALTERNATE CAPTIONS SETS
Private Const WinXP As String = "WindowsXP Themes"    '"Add Manifest"     '"Use XP Themes"
Private Const Classic As String = "Classic Windows"   '"Remove Manifest"  ' "Don't Use XP Themes",
Private Const NotXP As String = "WindowsXP only"      '"Unavailable"      '
Private Const ToolTip As String = "Click to change program style"

Private Enum CaptionStyles
    Toggle
    XP
    NoXP
    InVert
End Enum
Rem Mark Off
'stops Ulli's Code formatter from noticing these as Duplicated Name without Scope or TypeCasting
#If False Then '    'Enforce Case for Enums (does not compile)
Dim Toggle          'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Dim XP
Dim NoXP
Dim InVert
#End If
Rem Mark On
'Support for IsWindowsXP routine
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'Private Function DemoValidControls(C As Control, Optional C2 As Control, Optional Item As Integer) As Boolean
'   'Commented out this was a control validity test for DemoControl in Version 1 & 2
'  'TO REMOVE DEMO MODE
'  'comment out/delete this Sub routine
'  ' This routine is probably not needed as the IDE will do the checking for you
'
'    If TypeOf C Is CommandButton Then
'        DemoValidControls = True ' always
'      ElseIf TypeOf C Is CheckBox Then 'NOT TYPEOF...
'        DemoValidControls = True ' always
'      ElseIf TypeOf C Is OptionButton Then 'NOT TYPEOF...
'        DemoValidControls = (TypeOf C2 Is OptionButton)
'      ElseIf TypeOf C Is Menu Then 'NOT TYPEOF...
'        DemoValidControls = True
'        If TypeOf C2 Is Menu Then
'            DemoValidControls = True
'        End If
'      ElseIf TypeOf C Is Listbox Then 'NOT TYPEOF...
'        DemoValidControls = (C.Style = vbListBoxCheckbox) And (Item > 0 And Item < C.ListCount)
'    End If
'
'End Function
Public Sub About()

  Dim Msg As String

    Msg = "ClsManifestation" & vbNewLine & _
          "This class allows you to add The option for your XP users to choose between Classic and WindowsXP appearance" & vbNewLine & _
          "for your program. You need only add 4 lines to your code" & vbNewLine & _
          "1. Add 'Public Manfst As New ClsManifestation' to top of any module or form(that loads first)" & vbNewLine & _
          "2. Add 'ManFst.Manifest' to Sub Form_Initialize() of first form loaded by your program" & vbNewLine & _
          "3. Add 'ManFst.CommandButton CommandButton1' or any of the other controls possible (See Routines for additional paramaters)" & vbNewLine & _
          "4. Add 'Manfst.Action' to the click event." & vbNewLine & _
          "That's it. The control will take a caption or icon(if you provide one) indicating either present or alternative appearance." & vbNewLine & _
          "When clicked the user is offered the option to restart immediately to see the effect or waiting until next launch." & vbNewLine & vbNewLine & _
          "Please download 'ClsManifestation ver 3' at www.planet-source-code.com for a full description and Demo."
    MsgBox Msg, , "ClsManifestation"

End Sub

Public Sub Action()

  'Call from the Click event of m_Control

  Dim Success As Boolean ' Safety : let ActionResponse know that the manifest file has been created/removed

    If SettingValue = False Then ' allows OptionButton/CheckBox/ListBox to be set without taking action.
        If HasAManifest Then
            Success = RemoveManifest
          Else 'HASAMANIFEST = FALSE
            Success = CreateManifest
        End If
        'TO REMOVE DEMO MODE
        'comment out/delete the following 'If...Then...End If" structure
        'leaving only the line marked <*LEAVE ME*>
        If m_DemoMode Then
            DemoActionResponse Success
          Else 'M_DEMOMODE = FALSE
            ActionResponse Success '<*LEAVE ME*>
        End If
        If Success Then
            RefreshMe
        End If
    End If

End Sub

Private Sub ActionResponse(DidIt As Boolean)

    If DidIt Then  'you only see this Msgbox if .Action created or destroyed a manifest
        Select Case MsgBox( _
               "Changing a manifest file has no effect until you restart the program." & NL & _
               " Do you wish to restart now?", _
               vbDefaultButton1 + vbExclamation + vbYesNo)
          Case vbYes
            LaunchSelf  ' launch new instance of program
            End         ' close down current instance of program
          Case vbNo
            'Do nothing keep current instance of program running
        End Select
    End If

End Sub

Private Function Caption(Optional Mode As CaptionStyles = Toggle) As String

  '"SPECIALISED CONTROL" Edit out references to controls you are not using
  'if you are sure your users will have XP then edit out the IsWindowsXP test too

  'used to set a display string
  ' also disables the control if not in WindowsXP
  'Invert is used by OptionButtons to create the second caption
  'also sets the tooltiptext if you hav not.

    Select Case Mode
      Case Toggle
        If HasAManifest Then
            Caption = Classic
          Else 'HASAMANIFEST = FALSE
            Caption = WinXP
        End If
      Case XP
        Caption = WinXP
      Case NoXP
        Caption = Classic
      Case InVert
        If HasAManifest Then
            Caption = WinXP
          Else 'HASAMANIFEST = FALSE
            Caption = Classic
        End If
    End Select
    If Not (TypeOf m_Control Is Menu) Then ' menus don't support tooltip directly
        If Not (TypeOf m_Control Is Toolbar) Then ' toolbars don't support the next test properly so don't do it
            If Len(m_Control.ToolTipText) = 0 Then ' do not change if tooltip already exists
                m_Control.ToolTipText = ToolTip
            End If
        End If
    End If

    If Not IsWindowsXP Then
        If TypeOf m_Control Is CommandButton Then
            m_Control.Enabled = False
            Caption = NotXP
          ElseIf TypeOf m_Control Is CheckBox Then 'NOT TYPEOF...
            m_Control.Enabled = False
          ElseIf TypeOf m_Control Is OptionButton Or TypeOf m_Control Is Menu Then 'NOT TYPEOF...
            m_Control.Enabled = False
            If Not (m_Control2 Is Nothing) Then
                m_Control2.Enabled = False
            End If
          ElseIf TypeOf m_Control Is Listbox Then 'NOT TYPEOF...
            '   you don't want to disable the whole ListBox
            '   but this will prevent it from taking action
            Caption = NotXP
            SettingValue = True
        End If
    End If

End Function

Public Sub CheckBox(C As CheckBox)

  ' Call from Sub Form_Load() of the form which hold the control(s)
  ' you can actually call this from several different places in your program
  ' its main purpose is to keep the Caption in sync and to supply the Hwnd
  ' for ShellExecute in Sub LaunchSelf using m_Control.Parent.Hwnd.
  'SPECIALISED CONTROL 'comment out/delete this Sub routine if you are not using a CheckBox

    SingleCheckControl C

End Sub

Private Sub Class_Initialize()

    m_DemoMode = False 'True 'TO REMOVE DEMO MODE 'comment out/delete this line
    'most comments could also go for a working version
    'but please leave the copyright notices

    ' for Beginners: use Projects|Properties|Make tab to set the App.Properties
    'Quick and Dirty :) you could fill these properties from somewhere else in your program
    '                   using Public Property Let routines but they serve no real purpose (as yet)
    '                   so why not make them automatic?

    m_Description = App.Comments 'or if you like un-comment the next line instead
    'm_Description = App.FileDescription
    'm_Description = ""'if you aren't interested in what your manifest file says
    ' (and it doesn't seem to matter)

    m_Name = App.CompanyName & "." & App.ProductName & "." & App.EXEName
    ' this is the style used in Microsoft's FAQ article
    '"INFO: Visual Basic 6.0 Does Not Support Windows XP Themes or Visual Styles (Q309366)"
    ' but any value other than a zero length string will do i.e.
    'm_Name = " " ' is legal BUT NOT MyName =""
    ' if you haven't set the App.properties then
    'm_Name = ".." 'which is weird but legal

    m_Version = "1.0.0.0" ' Default
    'It doesn't matter to the Manifest file what number you use
    'as long as it is in the '#.#.#.#' format
    'the number is ServicePack number not a built-in VB App.Property
    ' Uncomment next line for the your program's version number
    'm_Version = App.Major & "." & App.Minor & "." & App.Revision & ".0"
    '****(NB the final '& ".0"' which formats things properly)*****
    SettingValue = True

End Sub

Public Sub CommandButton(C As CommandButton)

  ' Call from Sub Form_Load() of the form which hold the control(s)
  ' you can actually call this from several different places in your program
  ' its main purpose is to keep the Caption in sync and to supply the Hwnd
  ' for ShellExecute in Sub LaunchSelf using m_Control.Parent.Hwnd.
  'SPECIALISED CONTROL 'comment out/delete this Sub routine if you are not using a CommandButton

    Set m_Control = C
    m_Control.Caption = Caption(Toggle)
    SettingValue = GetSettingValue

End Sub

Private Function Content() As String

  ' layout shows location of the three values (MyVersion, MyName, MyDescription) you can individualize
  ' in your manifest file. The rest seems to be standardized stuff you should leave alone.
  'DQ is a constant for Double Quotes; NL is a constant for VbNewline

    Content = "<?xml version=" & DQ & "1.0" & DQ & " encoding=" & DQ & "UTF-8" & DQ & _
              " standalone=" & DQ & "yes" & DQ & "?>" & NL & _
              "<assembly xmlns=" & DQ & "urn:schemas-microsoft-com:asm.v1" & DQ & _
              " manifestVersion=" & DQ & "1.0" & DQ & ">" & NL & _
              "<assemblyIdentity  type=" & DQ & "win32" & DQ & " processorArchitecture=" & _
              DQ & "*" & DQ & " version=" & DQ & _
              m_Version _
              & DQ & " name=" & DQ & _
              m_Name _
              & DQ & "/>" & NL & "<description>" & _
              m_Description _
              & "</description>" & NL & "<dependency>" & NL & _
              "<dependentAssembly>type=" & DQ & "win32" & DQ & NL & _
              "<assemblyIdentity type=" & DQ & "win32" & DQ & " name=" & DQ & _
              "Microsoft.Windows.Common-Controls" & DQ & " version=" & DQ & "6.0.0.0" & DQ & _
              " language=" & DQ & "*" & DQ & " processorArchitecture=" & DQ & "*" & DQ _
              & NL & "publicKeyToken=" & DQ & "6595b64144ccf1df" & DQ & "/>" & NL & _
              "</dependentAssembly>" & NL & "</dependency>" & NL & "</assembly>"

End Function

Private Function CreateManifest() As Boolean

  'Quick and Dirty :) file writer you might like to use a more sophisticated/modern file writer

  Dim fnum As Integer

    fnum = FreeFile
    Open ManifestFile For Output As fnum   ' Create file name.
    Print #fnum, Content                   ' Output text.
    Close #fnum                            ' Close file.
    CreateManifest = True

End Function

Private Sub DemoActionResponse(DidIt As Boolean)

  'TO REMOVE DEMO MODE
  'comment out/delete this Sub routine
  'you don't normally want two instances of your program running with different appearances.

    If DidIt Then
        Select Case MsgBox("Changing a manifest file has no effect until you restart the program." & NL & _
               " Do you wish to restart now?" & NL & NL & _
               "(For demo purposes select Cancel and a new instance appears without closing the current one," & NL & _
               "but only the last two instances are sure to do what they say.)", vbDefaultButton1 + vbExclamation + vbYesNoCancel, "ClsManifestation DEMO MODE")

          Case vbYes
            LaunchSelf
            End                 'end current instance of program
          Case vbNo             'Do nothing
          Case vbCancel
            LaunchSelf  'see the difference between Manifest and no Manifest forms
        End Select
    End If

End Sub

'Public Sub DemoControl(C1 As Control, Optional C2 As Control, Optional Item As Integer)
'   'Commented out this was the single entry point in version 1 & 2 but the parameters for
'   'version 3 would have made it to unwieldy for version 3
'  'TO REMOVE DEMO MODE
'  'comment out/delete this Sub routine
'  ' This can set any of the controls but is wasteful to use in your own programs
'  ' Call from Sub Form_Load() of the form which hold the control(s)
'  ' you can actually call this from several different places in your program
'  ' its main purpose is to keep the Caption in sync and to supply the Hwnd
'  ' for ShellExecute in Sub LaunchSelf using m_Control.Parent.Hwnd.
'
'    If DemoValidControls(C1, C2, Item) Then
'        SettingValue = True ' disable Action while setting tickbox
'        If TypeOf C1 Is CommandButton Then
'            SingleToggleControl C1
'          ElseIf TypeOf C1 Is CheckBox Then 'NOT TYPEOF...
'            SingleCheckControl C1
'          ElseIf TypeOf C1 Is OptionButton Or TypeOf C1 Is Menu Then 'NOT TYPEOF...
'            PairedCheckControls C1, C2
'          ElseIf TypeOf C1 Is Listbox Then 'NOT TYPEOF...
'            Listbox C1, Item
'        End If
'        SettingValue = IIf(IsWindowsXP, False, True) ' reenable Action after setting tickbox
'    End If
'
'End Sub

Public Function DemoDisplayManifest() As String

  'TO REMOVE DEMO MODE
  'comment out/delete this Sub routine
  'this is not something you will normally need to see/do but is here for demo purposes

    DemoDisplayManifest = Content & NL & NL & _
                          "_______________________________________________________" & _
                          "CUSTOMISATION" & NL & _
                          "       NAME: " & m_Name & NL & _
                          "    VERSION: " & m_Version & NL & _
                          "DESCRIPTION: " & m_Description

End Function

Public Sub DemoVB6Manifest()

  'TO REMOVE DEMO MODE
  'comment out/delete this Sub routine
  'your end-user's (and probably you) do not need this routine

  Dim fnum As Integer

    m_Name = "Microsoft.Visual Basic 6.VB6.Exe" ' set meaningful content to manifest file
    m_Description = "Allows VB6 to use WindowsXP themes. NOTE you lose Palette selection in Properties."
    m_Version = "6.0.0.0"

    Select Case MsgBox("This routine creates a manifest file you can apply to VB6 itself." & NL & _
           "This Demo program makes no attempt to explore your computer" & NL & _
           "it just writes the 'VB6.Exe.Manifest' file to its own directory." & NL & _
           "Just copy it to the directory containing VB6.Exe." & NL & _
           "NOTE: Manifested VB6 loses the ability to select Palette colours in Properties." & NL & NL & _
           "Continue?", vbDefaultButton1 + vbExclamation + vbYesNo, "Make a Manifest for VB6")

      Case vbNo
        GoTo alldone
      Case vbYes

    End Select

    Select Case MsgBox(Content & NL & NL & _
           "_______________________________________________________" & _
           "CUSTOMISATION" & NL & _
           "       NAME: " & m_Name & NL & _
           "    VERSION: " & m_Version & NL & _
           "DESCRIPTION: " & m_Description & NL & _
           "_______________________________________________________" & _
           NL & NL & _
           "This is the content of VB6.Exe.Manifest." & NL & NL & _
           "Write to disk?", vbDefaultButton1 + vbYesNo + vbExclamation, "Make a Manifest for VB6")
      Case vbNo
        GoTo alldone
      Case vbYes

        fnum = FreeFile
        Open App.Path & "\" & "VB6.exe.manifest" For Output As fnum
        Print #fnum, Content
        Close #fnum
    End Select
    MsgBox "Don't Forget to copy 'VB6.exe.manifest' from:" & NL & NL _
           & App.Path & NL & NL & _
           " to the directory containing VB6.exe", vbInformation, "Make a Manifest for VB6"
alldone:
    Class_Initialize 'reset demo to its own values

End Sub

Private Function GetSettingValue() As Boolean

    If IsWindowsXP Then
        GetSettingValue = False
      Else 'ISWINDOWSXP = FALSE
        GetSettingValue = True
    End If

End Function

Private Function HasAManifest() As Boolean

  ' or use your favourite FileExists routine instead

  Dim Res As String

    If IsWindowsXP Then ' if it isn't XP then it doesn't matter anyway
        On Error Resume Next
            Res = Dir$(ManifestFile)
            If Res = "" Or Err <> 0 Then
                HasAManifest = False
              Else 'NOT RES...
                HasAManifest = True
            End If
        On Error GoTo 0
    End If

End Function

Private Function IsWindowsXP() As Boolean

  'Quick and Dirty :) minimal cut and paste job from other work
  '                you might like to substitute a more complete WindowsVersion routine
  '                or just remove the Static stuff for a simpler one

  Static Done As Boolean
  Static Result As Boolean
  Dim m_os As OSVERSIONINFO

    If Not Done Then ' set the Satic Result
        With m_os
            .dwOSVersionInfoSize = Len(m_os)
            Call GetVersionEx(m_os)
            Result = .dwMajorVersion + .dwMinorVersion / 10 >= 5.1
        End With 'M_OS
        Done = True
    End If
    IsWindowsXP = Result ' To see what happens if not in winXP force this to False

End Function

Private Sub LaunchSelf()

  'Quick and Dirty :) more complex programs may need some safety nets around this
  'Should return x = 0

  Dim x As Long

    x = ShellExecute(m_Control.Parent.hwnd, "open", App.EXEName, "", App.Path, SW_SHOWNORMAL)

End Sub

Public Sub Listbox(L As Listbox, Item As Integer)

  ' Set Listbox (because ListBox uses Selected rather than Value it cannot use Sub SingleCheckControl.
  ' Call from Sub Form_Load() of the form which hold the control(s)
  ' you can call this from several different places in your program
  ' its main purpose is to keep the Caption in sync and to supply the Hwnd
  ' for ShellExecute in Sub LaunchSelf using m_Control.Parent.Hwnd.
  ' If you have more than one control access to this class i.e. (Menu and an Options Screen)
  ' you can call this again after calling Action to keep both control captions coordinated.
  'SPECIALISED CONTROL 'comment out/delete this Sub routine if you are not using a Listbox

    SettingValue = True
    Set m_Control = L
    m_Index1 = Item
    m_Control.List(Item) = Caption(XP) 'WinXP
    m_Control.Selected(Item) = HasAManifest

    SettingValue = GetSettingValue

End Sub

Public Sub Manifest()

  ' Call from Sub Form_Initialize() of first form loaded by your program

    If IsWindowsXP Then         'don't bother if it won't matter
        If HasAManifest Then    ' these tests are not be necessary but I like them just in case
            InitCommonControls
        End If
      Else 'ISWINDOWSXP = FALSE
        'TO REMOVE DEMO MODE
        'delete the 'Else' above and the MsgBox below
        'which is really only for programmers playing with the Demo code
        MsgBox "ClsManifestation is only designed to function in WindowsXP." & NL & _
               " Nothing in it is WindowsXP specific but manifest files are only meaningful to WindowsXP" & NL & _
               "   (Remove the IsWindowsXP test in Sub Manifest if you want to use the class on Non-XP machines," & _
               "    for some reason I don't want to imagine or know.)" _
               , vbInformation, "ClsManifestation"
    End If

End Sub

Private Function ManifestFile() As String

  ' sets the Manifest filename and path

    ManifestFile = App.Path & "\" & App.EXEName & ".exe.manifest"

End Function

Public Sub Menu1item(M As Menu, Optional Chek As Boolean = False)

    m_MenuStyle = 1
    If Chek Then
        SingleCheckControl M
      Else 'CHEK = FALSE
        SingleToggleControl M
    End If

End Sub

Public Sub Menu2items(M1 As Menu, M2 As Menu)

    m_MenuStyle = 2
    PairedCheckControls M1, M2

End Sub

Public Sub OptionButtons(C As OptionButton, C2 As OptionButton)

  ' Call from Sub Form_Load() of the form which hold the control(s)
  ' you can call this many different places in your program
  ' its main purpose is to keep the Caption in sync and to supply the Hwnd
  ' for ShellExecute in Sub LaunchSelf using m_Control.Parent.Hwnd.
  'SPECIALISED CONTROL' comment out/delete this Sub routine if you are not using OptionButtons

    PairedCheckControls C, C2

End Sub

Private Sub PairedCheckControls(C1 As Control, C2 As Control)

  'called from OptionButtons and CheckBox and Menu2items

    SingleCheckControl C1
    SingleCheckControl C2, True 'set 2nd member with inverse caption

End Sub

Private Sub RefreshMe()

  'SPECIALISED CONTROL remove the sections which will not be needed for your program

    SettingValue = True
    If TypeOf m_Control Is CommandButton Then
        m_Control.Caption = Caption(Toggle)
      ElseIf TypeOf m_Control Is CheckBox Then 'NOT TYPEOF...
        SingleCheckControl m_Control
      ElseIf TypeOf m_Control Is OptionButton Then 'NOT TYPEOF...
        PairedCheckControls m_Control, m_Control2
      ElseIf TypeOf m_Control Is Menu Then 'NOT TYPEOF...
        Select Case m_MenuStyle
          Case 1
            Menu1item m_Control, m_Checked
          Case 2
            Menu2items m_Control, m_Control2
        End Select
      ElseIf TypeOf m_Control Is Listbox Then 'NOT TYPEOF...
        Listbox m_Control, CInt(m_Index1)
      ElseIf TypeOf m_Control Is Toolbar Then 'NOT TYPEOF...
        Select Case m_ToolBarStyle
          Case 1
            ToolBarButton m_Control, m_ButtonKey, m_Icn, m_Txt, m_Inverse, m_Index1, m_Index2
          Case 2
            ToolBarButtonMenu m_Control, m_ButtonKey, m_SubToggle, m_Icn, m_Txt, m_Inverse, m_Index1, m_Index2
        End Select

    End If
    SettingValue = GetSettingValue

End Sub

Private Function RemoveManifest() As Boolean

    If HasAManifest Then ' this test is a necessary safety test substitute but don't remove
        Kill ManifestFile
        RemoveManifest = True
    End If

End Function

Private Sub SingleCheckControl(C1 As Control, Optional Inverse As Boolean = False)

  'Called from:
  '           Menu1item
  '           PairedCheckControls where Inverse sets second member

    SettingValue = True
    m_Checked = True
    Set m_Control = C1
    If Inverse Then
        m_Control.Caption = Caption(NoXP)
      Else 'INVERSE = FALSE
        m_Control.Caption = Caption(XP)
    End If

    If TypeOf C1 Is Menu Then
        If Inverse Then
            m_Control.Checked = Not (HasAManifest)
          Else 'INVERSE = FALSE
            m_Control.Checked = HasAManifest
        End If
      ElseIf TypeOf C1 Is CheckBox Then 'NOT TYPEOF...
        If HasAManifest Then
            m_Control.Value = vbChecked
          Else 'HASAMANIFEST = FALSE
            m_Control.Value = vbUnchecked
        End If

      ElseIf TypeOf C1 Is OptionButton Then 'NOT TYPEOF...
        If Inverse Then
            m_Control.Value = Not (HasAManifest)
          Else 'INVERSE = FALSE
            m_Control.Value = HasAManifest
        End If
    End If

    SettingValue = GetSettingValue

End Sub

Private Sub SingleToggleControl(C1 As Control)

    Set m_Control = C1
    m_Control.Caption = Caption(Toggle)

End Sub

Public Sub ToolBarButton(Tb As Toolbar, ButtonKey$, Optional Icn As Boolean = False, Optional txt As Boolean = False, Optional Inverse As Boolean = False, Optional key1 = 1, Optional key2 = 2)

  'Optional   Values
  'ButtonKey$ =key property of button to be used
  'Icn       =False > No Icon on button
  'txt       =False > No text on button
  '          NB if Icn AND txt are False  then Icn is reset to True

    SettingValue = True

    If txt = False And Icn = False Then
        Icn = True                      ' force to icon format
    End If
    m_Txt = txt
    m_Icn = Icn
    m_Index1 = key1
    m_Index2 = key2
    m_Inverse = Inverse
    m_ToolBarStyle = 1
    m_ButtonKey = ButtonKey$
    Set m_Control = Tb
    With m_Control

        If txt Then
            If Inverse Then

                .Buttons(ButtonKey$).Caption = Caption(InVert)
              Else 'INVERSE = FALSE
                .Buttons(ButtonKey$).Caption = Caption(Toggle)
            End If
          Else 'TXT = FALSE
            .Buttons(ButtonKey$).Caption = ""
        End If

        '.Buttons(ButtonKey$).Caption = IIf(txt, IIf(Inverse, Caption(InVert), Caption(Toggle)), "")

        If Icn Then
            If Inverse Then
                If Caption(Toggle) = WinXP Then
                    .Buttons(ButtonKey$).Image = key1
                  Else 'NOT CAPTION(TOGGLE)...
                    .Buttons(ButtonKey$).Image = key2
                End If

                '.Buttons(ButtonKey$).Image = IIf(Caption(Toggle) = WinXP, key1, key2)
              Else 'INVERSE = FALSE

                If Caption(Toggle) = WinXP Then
                    .Buttons(ButtonKey$).Image = key2
                  Else 'NOT CAPTION(TOGGLE)...
                    .Buttons(ButtonKey$).Image = key1
                End If
                '                .Buttons(ButtonKey$).Image = IIf(Caption(Toggle) = WinXP, key2, key1)
            End If
        End If
        .Enabled = IsWindowsXP
    End With 'M_CONTROL
    SettingValue = GetSettingValue

    'SettingValue = IIf(IsWindowsXP, False, True)

End Sub

Public Sub ToolBarButtonMenu(Tb As Toolbar, ButtonKey$, Optional SubToggle As Boolean = False, Optional Icn As Boolean = False, Optional txt As Boolean = False, Optional Inverse As Boolean = True, Optional key1 = 1, Optional key2 = 2)

  'Optional   Values
  'SubToggle =False > Two ButtonMenu members are used with Current Style disabled
  '           True  > Single ButtonMenu member shows Style you will change to.
  'Icn       =False > No Icon on button
  'txt       =False > No text on button
  '          NB if Icn AND txt are False  then Icn is reset to True
  'key1      = 1    > ButtonMenu Indexes; By leaving as Variant you can specify ButtonMenus member(s)
  'key2      = 2    >  using either the Index or Key property.
  '         Defaults to 1 caption=WinXP & 2 caption=Classic if SubToggle=False
  '         1 caption=Style you will change to.
  '         NB if you don't like the default order for the two ButtonMenus (Classic above WinXP)
  '         you can change them by using the Optional Key# paramaters to reverse teh default order

    ToolBarButton Tb, ButtonKey$, Icn, txt, Inverse, key1, key2 ' the button is the same so reuse code
    m_ToolBarStyle = 2
    m_ButtonKey = ButtonKey$
    m_SubToggle = SubToggle
    m_Inverse = Inverse
    SettingValue = True
    With m_Control.Buttons(ButtonKey$).ButtonMenus(key1)
        .Text = Caption(XP)
        If SubToggle Then

            .Enabled = Not (HasAManifest) And IsWindowsXP
          Else 'SUBTOGGLE = FALSE
            .Text = Caption(Toggle)
            .Enabled = IsWindowsXP
        End If
    End With 'M_CONTROL.BUTTONS(BUTTONKEY$).BUTTONMENUS(KEY1)
    If SubToggle Then
        With m_Control.Buttons(ButtonKey$).ButtonMenus(key2)
            .Text = Caption(NoXP)
            .Enabled = HasAManifest And IsWindowsXP
        End With 'M_CONTROL.BUTTONS(BUTTONKEY$).BUTTONMENUS(KEY2)
    End If
    m_Control.Enabled = True
    SettingValue = GetSettingValue

End Sub

':) Ulli's VB Code Formatter V2.13.6 (28/08/2002 2:36:39 PM) 195 + 707 = 902 Lines

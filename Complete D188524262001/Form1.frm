VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "DragonBall Browser"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Navigate"
      Default         =   -1  'True
      Height          =   315
      Left            =   7530
      TabIndex        =   5
      Top             =   345
      Visible         =   0   'False
      Width           =   1395
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   1275
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   2249
      BandCount       =   2
      BackColor       =   14737632
      FixedOrder      =   -1  'True
      _CBWidth        =   9405
      _CBHeight       =   1275
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar1"
      MinHeight1      =   825
      Width1          =   4950
      NewRow1         =   0   'False
      Child2          =   "Combo1"
      MinHeight2      =   360
      Width2          =   7755
      FixedBackground2=   0   'False
      NewRow2         =   -1  'True
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   165
         TabIndex        =   4
         Text            =   "www.pojo.com"
         Top             =   885
         Width           =   9150
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   825
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   1455
         ButtonWidth     =   1296
         ButtonHeight    =   1455
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Navigate"
               ImageIndex      =   1
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "NewWindow"
                     Text            =   "Open In New Window"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Stop"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               ImageIndex      =   1
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Normal"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Expire Check"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Complete Refresh"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Back"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Forward"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exit"
               ImageIndex      =   1
            EndProperty
         EndProperty
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   5550
            Top             =   15
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":0442
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7380
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11033
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "9:24 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "1/1/01"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser Browser 
      Height          =   6135
      Left            =   0
      TabIndex        =   2
      Top             =   1275
      Width           =   9405
      ExtentX         =   16589
      ExtentY         =   10821
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu SaveAs 
         Caption         =   "Save &As"
         Shortcut        =   ^S
      End
      Begin VB.Menu dasher 
         Caption         =   "-"
      End
      Begin VB.Menu Print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu PrintPrev 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu PgSetup 
         Caption         =   "Page Set&up"
      End
      Begin VB.Menu aioshd 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Ops 
      Caption         =   "&Options"
      Begin VB.Menu NavOps 
         Caption         =   "Navigate"
         Begin VB.Menu ReadCache 
            Caption         =   "Read &From Cache"
            Checked         =   -1  'True
         End
         Begin VB.Menu WriteCache 
            Caption         =   "&Write To Cache"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu iojsoijwe 
         Caption         =   "-"
      End
      Begin VB.Menu EnDiBoxes 
         Caption         =   "Ena&ble Popups"
         Checked         =   -1  'True
      End
      Begin VB.Menu poksef 
         Caption         =   "-"
      End
      Begin VB.Menu Online 
         Caption         =   "On&line"
         Checked         =   -1  'True
      End
      Begin VB.Menu plaplsfdp 
         Caption         =   "-"
      End
      Begin VB.Menu SelFont 
         Caption         =   "Select &Font"
         Begin VB.Menu ArNarr 
            Caption         =   "Arial Narro&w"
            Checked         =   -1  'True
         End
         Begin VB.Menu ComSns 
            Caption         =   "Comic Sans M&S"
         End
         Begin VB.Menu Vrd 
            Caption         =   "&Verdana"
         End
      End
      Begin VB.Menu flameageation 
         Caption         =   "-"
      End
      Begin VB.Menu SveOpt 
         Caption         =   "Save Opt&ions"
      End
      Begin VB.Menu LoadOps 
         Caption         =   "Load Optio&ns"
      End
   End
   Begin VB.Menu Favorite 
      Caption         =   "Fa&vorites"
      Begin VB.Menu Favorites 
         Caption         =   "Favorites &List"
      End
      Begin VB.Menu plsdfoks 
         Caption         =   "-"
      End
      Begin VB.Menu AddSite 
         Caption         =   "Add Site To Lis&t"
      End
   End
   Begin VB.Menu Commands 
      Caption         =   "Co&mmands"
      Begin VB.Menu HomePage 
         Caption         =   "Go To &Home"
      End
      Begin VB.Menu SearchPage 
         Caption         =   "Go To Search"
         Begin VB.Menu Lycos 
            Caption         =   "Lycos"
         End
         Begin VB.Menu Altavista 
            Caption         =   "AltaVista"
         End
         Begin VB.Menu Webcrawler 
            Caption         =   "Webcrawler"
         End
         Begin VB.Menu Yahoo 
            Caption         =   "Yahoo"
         End
         Begin VB.Menu AstaLaVista 
            Caption         =   "AstaLaVista"
         End
      End
      Begin VB.Menu alphabetix 
         Caption         =   "-"
      End
      Begin VB.Menu Srce 
         Caption         =   "View Sour&ce"
      End
   End
   Begin VB.Menu Udda 
      Caption         =   "O&ther"
      Begin VB.Menu Props 
         Caption         =   "Pa&ge Properties"
      End
      Begin VB.Menu pokspdo 
         Caption         =   "-"
      End
      Begin VB.Menu Abt 
         Caption         =   "Abou&t"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Dim EnbPops As Boolean

Private Sub Abt_Click()
MsgBox "DragonBall Browser v1.0.0 includes:" & vbNewLine & vbNewLine & "- Status bar" & vbNewLine & "- Favorites/Bookmarks" & vbNewLine & "- Customization settings" & vbNewLine & "- Printing, saving, disabling popups" & vbNewLine & vbNewLine & "Thanks to Douglas, ViO, and Renton." & vbNewLine & vbNewLine & "- Gohan, December 2000" & vbNewLine & "- www.pojo.com" & vbNewLine & vbNewLine & "- This browser was made by a fan of Dragon Ball, and was not intended to infringe upon any copyrights.  If there are any copyright problems with this browser, please contact me at yoshiii@aol.com.", vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub AddSite_Click()
Form2.Show
'Form2.Favorites.AddItem Mid(Form1.Caption, 21, Len(Form1.Caption) - 21) & " - " & Browser.LocationURL
Form2.Favorites.AddItem Browser.LocationName & " - " & Browser.LocationURL
End Sub

Private Sub Altavista_Click()
On Error GoTo ErrHandler
Browser.Navigate "www.altavista.com": Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub ArNarr_Click()
ArNarr.Checked = True: ComSns.Checked = False: Vrd.Checked = False
Combo1.FontName = "Arial Narrow": Status.Font = "Arial Narrow": Combo1.FontSize = 10
End Sub

Private Sub AstaLaVista_Click()
On Error GoTo ErrHandler
Browser.Navigate "www.astalavista.com": Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub Browser_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
Combo1.Text = Browser.LocationURL
End Sub

Private Sub Browser_NewWindow2(ppDisp As Object, Cancel As Boolean)
If EnbPops = False Then
    Cancel = True
ElseIf EnbPops = True Then
    Cancel = False
End If
End Sub

Private Sub Browser_StatusTextChange(ByVal Text As String)
Status.Panels.Item(1).Text = Text
End Sub

Private Sub Browser_TitleChange(ByVal Text As String)
Form1.Caption = "DragonBall Browser - " & Text
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub Command1_Click()
On Error GoTo ErrHandler
If ReadCache.Checked = True And WriteCache.Checked = True Then Browser.Navigate Combo1.Text
If ReadCache.Checked = True And WriteCache.Checked = False Then Browser.Navigate Combo1.Text, 8
If ReadCache.Checked = False And WriteCache.Checked = True Then Browser.Navigate Combo1.Text, 4
If ReadCache.Checked = False And WriteCache.Checked = False Then Browser.Navigate Combo1.Text, 4 Or 8
Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub ComSns_Click()
ArNarr.Checked = False: ComSns.Checked = True: Vrd.Checked = False
Combo1.FontName = "Comic Sans MS": Status.Font = "Comic Sans MS": Combo1.FontSize = 8
End Sub

Private Sub EnDiBoxes_Click()
If EnDiBoxes.Checked = False Then
    EnDiBoxes.Checked = True: EnbPops = True
Else
    EnDiBoxes.Checked = False: EnbPops = False
End If
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Favorites_Click()
Load Form2: Form2.Show
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
EnbPops = True: Browser.Navigate "www.pojo.com"
Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub Form_Resize()
On Error Resume Next
Browser.Height = Form1.Height - (8535 - 6135): Browser.Width = Form1.Width - (9525 - 9405)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandler
Dim Response
Response = MsgBox("Are you sure you wish to exit?", vbInformation Or vbYesNo, "DragonBall Browser v1.0.0")
If Response = vbNo Then Cancel = 1
If Response = vbYes Then Unload Me
Exit Sub
ErrHandler:
MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub HomePage_Click()
On Error GoTo ErrHandler
Browser.GoHome: Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub LoadOps_Click()
On Error GoTo ErrHandler
Dim RetString As String, RetLen As Integer
RetString = String(15, " "): RetLen = Len(RetString)
GetPrivateProfileString "Settings", "Font", vbNullString, RetString, RetLen, App.Path & "\Settings.ini"
If Mid(RetString, 1, 12) = "Arial Narrow" Then
    ArNarr.Checked = True: ComSns.Checked = False: Vrd.Checked = False
    Combo1.FontName = "Arial Narrow": Status.Font = "Arial Narrow": Combo1.FontSize = 10
ElseIf Mid(RetString, 1, 13) = "Comic Sans MS" Then
    ArNarr.Checked = False: ComSns.Checked = True: Vrd.Checked = False
    Combo1.FontName = "Comic Sans MS": Status.Font = "Comic Sans MS": Combo1.FontSize = 8
ElseIf Mid(RetString, 1, 7) = "Verdana" Then
    ArNarr.Checked = False: ComSns.Checked = False: Vrd.Checked = True
    Combo1.FontName = "Verdana": Status.Font = "Verdana": Combo1.FontSize = 8
Else
    GoTo ErrHandler
End If: Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
Combo1.FontName = "Arial Narrow": Status.Font = "Arial Narrow": Combo1.FontSize = 10
ArNarr.Checked = True: ComSns.Checked = False: Vrd.Checked = False
End Sub

Private Sub Lycos_Click()
On Error GoTo ErrHandler
Browser.Navigate "www.lycos.com": Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub Online_Click()
On Error Resume Next
If Online.Checked = True Then
    Online.Checked = False: Browser.Offline = True
Else
    Online.Checked = True: Browser.Offline = False
End If
End Sub

Private Sub PgSetup_Click()
On Error GoTo ErrHandler
Browser.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT: Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub Print_Click()
On Error GoTo ErrHandler
Browser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT: Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub PrintPrev_Click()
On Error GoTo ErrHandler
Browser.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT: Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub Props_Click()
On Error GoTo ErrHandler
Browser.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT: Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub ReadCache_Click()
If ReadCache.Checked = True Then
    ReadCache.Checked = False
Else
    ReadCache.Checked = True
End If
End Sub

Private Sub SaveAs_Click()
On Error GoTo ErrHandler
Browser.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT: Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub Srce_Click()
MsgBox "To view source code for this page, right click the document and selet 'View Source'.", vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub SveOpt_Click()
On Error GoTo ErrHandler
WritePrivateProfileString "Settings", "Font", Combo1.FontName, App.Path & "\Settings.ini"
Exit Sub
ErrHandler: MsgBox Err.Description, "DragonBall Browser v1.0.0"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHandler
If Button.Index = 1 Then Command1_Click
If Button.Index = 3 Then Browser.Stop
If Button.Index = 4 Then Browser.Refresh2 0
If Button.Index = 6 Then Browser.GoBack
If Button.Index = 7 Then Browser.GoForward
If Button.Index = 9 Then Unload Me
Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo ErrHandler
If ButtonMenu.Key = "NewWindow" Then
    If EnbPops = False Then
        EnbPops = True
        Browser.Navigate Combo1.Text, 1
        EnbPops = False
    Else
        Browser.Navigate Combo1.Text, 1
    End If
End If
If ButtonMenu.Text = "Normal" Then Browser.Refresh2 0
If ButtonMenu.Text = "Expire Check" Then Browser.Refresh2 1
If ButtonMenu.Text = "Complete Refresh" Then Browser.Refresh2 3
Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub Vrd_Click()
ArNarr.Checked = False: ComSns.Checked = False: Vrd.Checked = True
Combo1.FontName = "Verdana": Status.Font = "Verdana": Combo1.FontSize = 8
End Sub

Private Sub Webcrawler_Click()
On Error GoTo ErrHandler
Browser.Navigate "www.webcrawler.com": Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

Private Sub WriteCache_Click()
If WriteCache.Checked = True Then
    WriteCache.Checked = False
Else
    WriteCache.Checked = True
End If
End Sub

Private Sub Yahoo_Click()
On Error GoTo ErrHandler
Browser.Navigate "www.yahoo.com": Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0"
End Sub

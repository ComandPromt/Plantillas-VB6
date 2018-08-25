VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Timesheets Lite"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9540
   Icon            =   "frmMDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   1376
      ButtonWidth     =   2778
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ilsIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Login"
            Key             =   "Login"
            Object.ToolTipText     =   "Click to login"
            ImageKey        =   "Login"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Project Maintenance"
            Key             =   "projectMaintenance"
            Object.ToolTipText     =   "Click to see the project maintenance form"
            ImageKey        =   "ProjectMaintenance"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "User Maintenance"
            Key             =   "userMaintenance"
            Object.ToolTipText     =   "Click to see the User Maintenance Form"
            ImageKey        =   "UserMaintenance"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Time Sheet"
            Key             =   "timeSheet"
            Object.ToolTipText     =   "Click to enter the timesheet"
            ImageKey        =   "Timesheet"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporting"
            Key             =   "reporting"
            Object.ToolTipText     =   "Click to enter the reporting screen"
            ImageKey        =   "Reporting"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Key             =   "Options"
            Object.ToolTipText     =   "Click to enter the Options screen"
            ImageKey        =   "Options"
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ilsIcons 
         Left            =   7650
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMDIMain.frx":0442
               Key             =   "ProjectMaintenance"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMDIMain.frx":0894
               Key             =   "UserMaintenance"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMDIMain.frx":0C32
               Key             =   "Reporting"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMDIMain.frx":1084
               Key             =   "Login"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMDIMain.frx":14D6
               Key             =   "Options"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMDIMain.frx":1C28
               Key             =   "Timesheet"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub MDIForm_Terminate()
  globalCode.exitSystem
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  globalCode.exitSystem
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)


  Unload frmProjectMaintenance
  Unload frmUserMaintenance
  Unload frmTimeSheet
  Unload frmReport
  Unload frmOptions
  
  Select Case Button.Key
    Case "Login"
      frmLogin.Show
    Case "projectMaintenance"
      If (usrLoggedIn.bytSecurityLevel = PROJECT_MANAGER Or usrLoggedIn.bytSecurityLevel = SUPER_USER) Then
        If securityCode.usersLoggedIn(usrLoggedIn.lngUserID) = True Then
          MsgBox "There appears to be users logged into the system, all users must be logged out before entering the Project Maintenance screen", , "Other Users Logged In"
        Else
          frmProjectMaintenance.Show
        End If
      Else
        frmLogin.txtMessage = "You are not logged in or do not have permission to view the Project Maintenance screen"
        frmLogin.Show vbModal
        If (usrLoggedIn.bytSecurityLevel = PROJECT_MANAGER Or usrLoggedIn.bytSecurityLevel = SUPER_USER) Then
          If securityCode.usersLoggedIn(usrLoggedIn.lngUserID) = True Then
            MsgBox "There appears to be users logged into the system, all users must be logged out before entering the Project Maintenance screen", , "Other Users Logged In"
          Else
            frmProjectMaintenance.Show
          End If
        Else
          MsgBox "You are not logged in or do not have permission to view the Project Maintenance screen"
        End If
      End If
    Case "userMaintenance"
      If (usrLoggedIn.bytSecurityLevel = PROJECT_MANAGER Or usrLoggedIn.bytSecurityLevel = SUPER_USER) Then
        If securityCode.usersLoggedIn(usrLoggedIn.lngUserID) = True Then
          MsgBox "There appears to be users logged into the system, all users must be logged out before entering the User Maintenance screen", , "Other Users Logged In"
        Else
          frmUserMaintenance.Show
        End If
      Else
        frmLogin.txtMessage = "You are not logged in or do not have permission to view the User Maintenance screen"
        frmLogin.Show vbModal
        If (usrLoggedIn.bytSecurityLevel = PROJECT_MANAGER Or usrLoggedIn.bytSecurityLevel = SUPER_USER) Then
          If securityCode.usersLoggedIn(usrLoggedIn.lngUserID) = True Then
            MsgBox "There appears to be users logged into the system, all users must be logged out before entering the User Maintenance screen", , "Other Users Logged In"
          Else
            frmUserMaintenance.Show
          End If
        Else
          MsgBox "You are not logged in or do not have permission to view the User Maintenance screen"
        End If
      End If
    Case "timeSheet"
      If usrLoggedIn.lngUserID > 0 Then
        frmTimeSheet.Show
      Else
        frmLogin.txtMessage = "You are not logged in"
        frmLogin.Show vbModal
        If usrLoggedIn.lngUserID > 0 Then
          frmTimeSheet.Show
        Else
          MsgBox "You are not logged in"
        End If
      End If
    Case "reporting"
      If (usrLoggedIn.bytSecurityLevel = FINANCIAL_USER Or usrLoggedIn.bytSecurityLevel = SUPER_USER) Then
        If securityCode.usersLoggedIn(usrLoggedIn.lngUserID) = True Then
          MsgBox "There appears to be users logged into the system, all users must be logged out before entering the Reporting screen", , "Other Users Logged In"
        Else
          frmReport.Show
        End If
      Else
        frmLogin.txtMessage = "You are not logged in or do not have permission to view the Reporting screen"
        frmLogin.Show vbModal
        If (usrLoggedIn.bytSecurityLevel = FINANCIAL_USER Or usrLoggedIn.bytSecurityLevel = SUPER_USER) Then
          If securityCode.usersLoggedIn(usrLoggedIn.lngUserID) = True Then
            MsgBox "There appears to be users logged into the system, all users must be logged out before entering the Reporting screen", , "Other Users Logged In"
          Else
            frmReport.Show
          End If
        Else
          MsgBox "You are not logged in or do not have permission to view the Reporting screen"
        End If
      End If
    Case "Options"
      frmOptions.Show vbModal
  End Select
End Sub







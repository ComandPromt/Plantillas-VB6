VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmOS 
   Caption         =   "Operating System Info"
   ClientHeight    =   1260
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1260
   ScaleWidth      =   2955
   Begin VB.ListBox lstInfo 
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   1080
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()

        
        Select Case SysInfo.OSPlatform
                Case 0
                        lstInfo.AddItem "OS Platform = Unknown 32-Bit Windows"
                Case 1
                        lstInfo.AddItem "OS Platform = Windows 95"
                Case 2
                        lstInfo.AddItem "OS Platform = Windows NT"
        End Select
        lstInfo.AddItem "OSVersion = " & SysInfo.OSVersion
        lstInfo.AddItem "OSBuild = " & SysInfo.OSBuild

End Sub



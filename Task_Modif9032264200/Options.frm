VERSION 5.00
Begin VB.Form Options 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7710
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CheckTimer 
      Caption         =   "Search For NonExisting Windows"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "External Application"
      Height          =   1695
      Left            =   2685
      TabIndex        =   2
      Top             =   135
      Width           =   4935
      Begin VB.CommandButton CommandCD 
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtLaunch 
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   3975
      End
      Begin VB.ComboBox ComboPar 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Labelexename 
         Caption         =   "Name of application to launch"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label LabelPar 
         Caption         =   "Command line Parameter"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.CheckBox CheckAuthor 
      Caption         =   "Author Mode"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CheckBox CheckIcons 
      Caption         =   "Show Icons On Treeview (Slower)"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'flag i set if the treeview(Tasktree) needs refreshed after options done
Private Refresh_Task As Boolean
Dim Cd1 As CDg1 'this is our common dialog class

Private Sub CheckAuthor_Click()

    If CheckAuthor.Value Then
        AuthorMode = True
      Else
        AuthorMode = False
    End If
End Sub

Private Sub CheckIcons_Click()

    If CheckIcons.Value Then
        Showicons = True
      Else
        Showicons = False
    End If
    Refresh_Task = True
End Sub

Private Sub CheckTimer_Click()

    If CheckTimer.Value Then
        SearchForWindows = True
      Else
        SearchForWindows = False
    End If
End Sub

Private Sub ComboPar_Click()

    LaunchPar = ComboPar.ListIndex

End Sub

Private Sub CommandCD_Click()
  Set Cd1 = New CDg1
  Cd1.Filter = "Executables|*.exe;*.bat;*.com;*.pif"
  Cd1.InitDir = App.Path
  Cd1.flags = DialogFlags.OFN_NOCHANGEDIR
  Cd1.ShowOpen
  txtLaunch.Text = Cd1.Filename

End Sub

Private Sub Form_Load()

    ComboPar.AddItem "<None>"
    ComboPar.AddItem "Handle"
    ComboPar.AddItem "Parent Handle"
    ComboPar.AddItem "ProcessID"
    ComboPar.AddItem "Exec FileName"
    ComboPar.ListIndex = LaunchPar
    txtLaunch.Text = LaunchFile
    CheckIcons.Value = Abs(Showicons)
    CheckAuthor.Value = Abs(AuthorMode)
    CheckTimer.Value = Abs(SearchForWindows)
    Refresh_Task = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim fdir As String
  Dim LaunchFLen As Integer
    fdir = App.Path & "\taskmod.dat"
    If Dir$(fdir) <> "" Then
        Kill fdir
    End If
    LaunchFile = txtLaunch.Text
    LaunchFLen = Len(LaunchFile)
    Open fdir For Binary As #1
    Put #1, 1, AuthorMode
    Put #1, , Showicons
    Put #1, , SearchForWindows
    Put #1, , LaunchPar
    Put #1, , LaunchFLen
    Put #1, , LaunchFile
    Close #1
    Me.Hide
    frmMain.TaskTree.ZOrder
    If Refresh_Task Then
        frmMain.RefreshTask
    End If
    TaskMenuID = 0
    Unload Me

End Sub


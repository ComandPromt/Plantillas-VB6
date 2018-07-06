VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fun Stuff"
   ClientHeight    =   5205
   ClientLeft      =   2595
   ClientTop       =   1455
   ClientWidth     =   6555
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame9 
      Caption         =   "Mouse"
      Height          =   855
      Left            =   2880
      TabIndex        =   41
      Top             =   4200
      Width           =   2535
      Begin VB.CommandButton Command12 
         Caption         =   "Disable It"
         Height          =   495
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Enable It"
         Height          =   495
         Left            =   1440
         TabIndex        =   42
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Set Local Date"
      Height          =   735
      Left            =   120
      TabIndex        =   36
      Top             =   3360
      Width           =   5295
      Begin VB.CommandButton Command3 
         Caption         =   "Set It"
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox SetYear 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox SetDay 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox SetMonth 
         Height          =   315
         ItemData        =   "Form2.frx":0442
         Left            =   120
         List            =   "Form2.frx":046A
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.ComboBox DayWeek 
      Height          =   315
      ItemData        =   "Form2.frx":04D0
      Left            =   240
      List            =   "Form2.frx":04E9
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ComboBox Month 
      Height          =   315
      ItemData        =   "Form2.frx":052D
      Left            =   2040
      List            =   "Form2.frx":0555
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      Caption         =   "Local Date"
      Height          =   735
      Left            =   120
      TabIndex        =   31
      Top             =   2520
      Width           =   5295
      Begin VB.TextBox Year 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Day 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5760
      Top             =   0
   End
   Begin VB.TextBox Hour 
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   480
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   240
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mouse Position"
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Move Mouse To Position"
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   3855
      Begin VB.CommandButton Command6 
         Caption         =   "Swap Buttons"
         Height          =   495
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Show Mouse"
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Hide Mouse"
         Height          =   495
         Left            =   2040
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Move Mouse"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Local Time"
      Height          =   855
      Left            =   4080
      TabIndex        =   20
      Top             =   0
      Width           =   2295
      Begin VB.TextBox AMPM 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Sec 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Min 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sec"
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Min"
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hour"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Set Local Time"
      Height          =   1335
      Left            =   4080
      TabIndex        =   21
      Top             =   960
      Width           =   2295
      Begin VB.ComboBox SetAMPM 
         Height          =   315
         ItemData        =   "Form2.frx":05BB
         Left            =   1560
         List            =   "Form2.frx":05C5
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Set It"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox SetSec 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox SetMin 
         Height          =   285
         Left            =   600
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox SetHour 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sec"
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Min"
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hour"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Keyboard"
      Height          =   855
      Left            =   120
      TabIndex        =   38
      Top             =   4200
      Width           =   2535
      Begin VB.CommandButton Command8 
         Caption         =   "Enable It"
         Height          =   495
         Left            =   1440
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Disable It"
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label8 
      Caption         =   "These Function at the bottom do not yet work!!"
      Height          =   855
      Left            =   5520
      TabIndex        =   44
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Some Things Take Time to Take Effect"
      Height          =   1575
      Left            =   5520
      TabIndex        =   37
      Top             =   2520
      Width           =   855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim utctime As SYSTEMTIME
Dim SwapMouse As Integer

Private Sub Command1_Click()
On Error GoTo bottom
    retval = SetCursorPos(Val(Text3.Text), Val(Text4.Text))
bottom:
End Sub

Private Sub Command12_Click()
'    Shell "rundll32 mouse,disable"
End Sub

Private Sub Command2_Click()
On Error GoTo bottom
    Dim SETTIME As SYSTEMTIME
    GetLocalTime SETTIME
    'If AMPM.Text = "PM" Then SETTIME.wHour = SETTIME.wHour + 12
    
    If Val(SetHour.Text) > 12 Then GoTo bottom
    
    If SetAMPM.Text = "AM" Then
        SETTIME.wHour = Val(SetHour.Text)
    Else
        SETTIME.wHour = Val(SetHour.Text + 12)
    End If
    SETTIME.wMinute = Val(SetMin.Text)
    SETTIME.wSecond = Val(SetSec.Text)
    SetLocalTime SETTIME
bottom:
End Sub

Private Sub Command3_Click()
    On Error GoTo bottom
    Dim SETTIME As SYSTEMTIME
    GetLocalTime SETTIME
    
    SETTIME.wYear = Val(SetYear.ListIndex + 1972)
    SETTIME.wMonth = Val(SetMonth.ListIndex + 1)
    SETTIME.wDay = Val(SetDay.ListIndex + 1)
    SetLocalTime SETTIME
    Exit Sub
bottom:
    MsgBox "Most likely there's not" & Val(SetDay.ListIndex + 1) & "days in that month", vbOKOnly, "ERROR"
End Sub

Private Sub Command4_Click()
    ShowCursor 0
End Sub

Private Sub Command5_Click()
    ShowCursor 1
End Sub

Private Sub Command6_Click()
    If SwapMouse = 0 Then
        SwapMouseButton 1
        SwapMouse = 1
    Else
        SwapMouseButton 0
        SwapMouse = 0
    End If
End Sub


Private Sub Command7_Click()
'    Shell "rundll32 keyboard,disable"
End Sub

'Private Sub Command4_Click()
'    Shell "rundll32 keyboard,disable"
'End Sub

'Private Sub Command5_Click()
'doesn't work
'Shell "rundll32 keyboard,enable"
'End Sub

Private Sub Form_Activate()
    Timer1.Enabled = True
    Timer2.Enabled = True
End Sub


Private Sub Form_Load()
    For i = 1 To 31
        SetDay.AddItem Str(i), Val(i - 1)
    Next i
    For i = 1972 To 2080
        SetYear.AddItem Str(i), Val(i - 1972)
    Next i
    Dim username As String
    Dim slength As Long
    Dim retval As Long
    username = Space(255)
    slength = 255
    retval = GetUserName(username, slength)
    username = Left(username, slength - 1)
    Form2.Caption = "Fun Stuff With " & username & "'s Computer"
    GetLocalTime utctime
    SetAMPM.Text = "AM"
    SetMonth.Text = "January"
    SetDay.ListIndex = 0
    SetYear.ListIndex = Val(utctime.wYear - 1972)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Timer2.Enabled = False

End Sub

Private Sub Timer1_Timer()
    Dim coord As POINT_TYPE  ' receives coordinates of cursor
    Dim retval As Long  ' return value
    
    retval = GetCursorPos(coord)  ' read cursor location
    Text1.Text = coord.x
    Text2.Text = coord.y
End Sub

Private Sub Timer2_Timer()
    GetLocalTime utctime
    If utctime.wHour > 12 Then
        Hour.Text = utctime.wHour - 12
        AMPM.Text = "PM"
    Else
        Hour.Text = utctime.wHour
        AMPM.Text = "AM"
    End If
    Min.Text = utctime.wMinute
    Sec.Text = utctime.wSecond
    Month.ListIndex = utctime.wMonth - 1
    DayWeek.ListIndex = utctime.wDayOfWeek
    Day.Text = utctime.wDay
    Year.Text = utctime.wYear
End Sub

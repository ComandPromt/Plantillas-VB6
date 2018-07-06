VERSION 5.00
Begin VB.Form frmDate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Date Options"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4140
   Icon            =   "frmDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   300
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   3000
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    'add user's selection to frmMain.RTF
    Select Case List1.ListIndex
        Case 0
            frmMain.Undo.InsertText Format(Date, "DD/MM/YY") + vbCrLf
        Case 1
            frmMain.Undo.InsertText Format(Date, "D/MM/YY") + vbCrLf
        Case 2
            frmMain.Undo.InsertText Format(Date, "DD/MM/YYYY") + vbCrLf
        Case 3
            frmMain.Undo.InsertText Format(Date, "YYYY-MM-DD") + vbCrLf
        Case 4
            frmMain.Undo.InsertText WeekdayName(Weekday(Now, vbUseSystemDayOfWeek)) + "," + Str(Day(Now)) + Chr(32) + MonthName(DatePart("M", (Now))) + Str(Year(Now)) + vbCrLf
        Case 5
            frmMain.Undo.InsertText Format(Time, "HH:MM:SS") + vbCrLf
        Case 6
            frmMain.Undo.InsertText Format(Time, "H:MM:SS") + vbCrLf
        Case 7
            frmMain.Undo.InsertText Trim(Str(Now)) + vbCrLf
    End Select
    Unload Me
End Sub
Private Sub Form_Load()
    'add some date formats for the user to choose
    List1.BackColor = frmMain.RTF.BackColor
    List1.AddItem Format(Date, "DD/MM/YY")
    List1.AddItem Format(Date, "D/MM/YY")
    List1.AddItem Format(Date, "DD/MM/YYYY")
    List1.AddItem Format(Date, "YYYY-MM-DD")
    List1.AddItem WeekdayName(Weekday(Now, vbUseSystemDayOfWeek)) + "," + Str(Day(Now)) + Chr(32) + MonthName(DatePart("M", (Now))) + Str(Year(Now))
    List1.AddItem Format(Time, "HH:MM:SS") + "-(24 hr)"
    List1.AddItem Format(Time, "H:MM:SS")
    List1.AddItem Now
    List1.ListIndex = 0
End Sub
Private Sub List1_DblClick()
    cmdOK_Click
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetAlarm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Alarm"
   ClientHeight    =   2355
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgFiles 
      Left            =   1680
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Files..."
      Filter          =   "Music Files (*.mp3)|*.mp3|Microsoft Sounds (*.wav)|*.wav"
      Orientation     =   2
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Alarm Type"
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
      Begin VB.OptionButton optMusic 
         Caption         =   "Music"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optSilent 
         Caption         =   "Silent"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame frmSetTime 
      Caption         =   "Set Time"
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4815
      Begin VB.OptionButton optPM 
         Caption         =   "PM"
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optAM 
         Caption         =   "AM"
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtMinute 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtHour 
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCurAlarm 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   3120
         TabIndex        =   15
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblAlarmTime 
         Alignment       =   2  'Center
         Caption         =   "Alarm Time:"
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblMinute 
         Caption         =   "Minute"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblHour 
         Caption         =   "Hour"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmSetAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()

' Input: None
' Process: Cancels the current operation
' Output: None

    Call cmdClear_Click ' Clearing all the fields
    Unload Me ' Unloading the form
End Sub

Private Sub cmdClear_Click()

' Input: None
' Process: Clears all the text boxes, all the option values and the captions
' Output: None

    txtHour.Text = "" ' Clearing the text box
    txtMinute.Text = "" ' Clearing the text box
    
    optAM.Value = False ' Clearing the option button
    optPM.Value = False ' Clearing the option button
    
    lblCurAlarm.Caption = "" ' Clearing the label caption
    cmdSet.Tag = "" ' Clearing the flag
    
    frmMain.lblSetTime.Caption = "" ' Changing the caption on the main form to reflect the alarm state
    frmMain.lblTimeSet.Caption = "" ' Changing the caption on the main form to reflect the alarm state
End Sub

Private Sub cmdSet_Click()

' Input: None
' Process: Sets the alarm time to the wanted time
' Output: None

    cmdSet.Tag = txtHour.Text + ":" + txtMinute.Text + ":00" ' Setting the alarm time
    cmdSet.ToolTipText = txtHour.Text + ":" + txtMinute.Text + ":01" ' Setting the backup alarm time
    
    If optAM.Value = True Then ' Checking the time of day
        cmdSet.Tag = cmdSet.Tag + " " + optAM.Caption ' Completing the alarm time
        cmdSet.ToolTipText = cmdSet.ToolTipText + " " + optAM.Caption ' Completing the backup alarm time
    Else
        cmdSet.Tag = cmdSet.Tag + " " + optPM.Caption ' Completing the alarm time
        cmdSet.ToolTipText = cmdSet.ToolTipText + " " + optPM.Caption ' Checking the backup alarm time
    End If
    
    txtHour.Text = "" ' Clearing all the text boxes
    txtMinute.Text = "" ' Clearing all the text boxes
    optAM.Value = False ' Clearing the option buttons
    optPM.Value = False ' Clearing the option buttons
    
    lblCurAlarm.Caption = cmdSet.Tag ' Flagging the alarm time
End Sub

Private Sub Form_Load()

' Input: None
' Process: Setting the alarm caption to whatever is in the alarmtime00 variable
' Output: None

    If frmMain.alarmtime00 = "" Then Exit Sub ' There's nothing there
    lblCurAlarm.Caption = frmMain.alarmtime00 ' Something's there
End Sub

Private Sub OKButton_Click()

' Input: None
' Process: Sets the alarm time and unloads the form
' Output: None

    frmMain.alarmtime00 = cmdSet.Tag ' Flagging the time
    frmMain.alarmtime01 = cmdSet.ToolTipText ' Flagging the backup time
    frmMain.lblSetTime.Caption = "Alarm is set for:" ' Cosmetic adjustments
    If frmMain.alarmtime00 = "" Then frmMain.lblSetTime.Caption = "" ' Cosmetic adjustments
    frmMain.lblTimeSet.Caption = frmMain.alarmtime00 ' More cosmetic adjustments
    Unload Me ' Unloading the form
End Sub

Private Sub optMusic_Click()

' Input: None
' Process: This is what happens when the music button is clicked
' Output: None

    frmMain.tmrChangeTime.Tag = 1 ' Tagging it for the timer to catch
    Call CallWinDir
    dlgFiles.ShowOpen ' Showing the Windows default Open Dialog Box
    filename1 = dlgFiles.filename ' Constructing the filename and path
End Sub

Private Sub optSilent_Click()

' Input: None
' Process: This is what happens when the silent button is clicked
' Output: None

    frmMain.tmrChangeTime.Tag = 0 ' Tagging it for the timer to catch
End Sub

Private Sub txtHour_Change()

' Input: None
' Process: Checking the data range of the text box
' Output: None

    If txtHour.Text = "" Then Exit Sub ' Checking if there is even anything in the text box
    
    If txtHour.Text > 12 Or txtHour.Text < 0 Then ' The data range
        If txtHour.Text = 0 Then txtHour.Text = 12 ' 0 hours is the same as 12 hours
        MsgBox "Invalid Entry!  Please enter a value between 1 and 12", 16, "Invalid Entry..."
        txtHour.Text = "" ' Clearing the text box
        txtHour.SetFocus ' Putting the focus back into the text box
    End If
End Sub

Private Sub txtHour_KeyPress(KeyAscii As Integer)

' Input: None
' Process: More data validation, this time whenever a key is pressed
' Output: None

    If KeyAscii = vbKeyBack Then ' Eliminating the backspace key from the list
        Exit Sub
    End If
    
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then ' Allowing only numbers into the field
        KeyAscii = 0 ' Re-assigning the ascii value of the key pressed
    End If
End Sub

Private Sub txtMinute_Change()

' Input: None
' Process: Checking the data range of the text box
' Output: None

'** NOTE** This sub does the same thing as its counterpart above, only in another text box

    If txtMinute.Text = "" Then Exit Sub
    
    If txtMinute.Text > 60 Or txtMinute.Text < 0 Then
        MsgBox "Invalid Entry!  Please enter a value between 1 and 60", 16, "Invalid Entry..."
        txtMinute.Text = ""
        txtMinute.SetFocus
    End If
End Sub

Private Sub txtMinute_KeyPress(KeyAscii As Integer)

' Input: None
' Process: More data validation
' Output: None

'**NOTE** This sub does the same thing as its counterpart above, just for another text box
    
    If KeyAscii = vbKeyBack Then
        Exit Sub
    End If
    
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
            KeyAscii = 0
    End If
End Sub

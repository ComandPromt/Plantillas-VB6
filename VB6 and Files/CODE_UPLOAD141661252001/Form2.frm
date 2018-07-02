VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Password Maintenance"
   ClientHeight    =   4845
   ClientLeft      =   5010
   ClientTop       =   3585
   ClientWidth     =   4470
   LinkTopic       =   "Form2"
   ScaleHeight     =   4845
   ScaleWidth      =   4470
   Begin VB.CheckBox chkNoReminder 
      Caption         =   "Reminder: SysAdmin password is ALLOW. Do not show this reminder again."
      Height          =   345
      Left            =   345
      TabIndex        =   6
      Top             =   4320
      Width           =   3345
   End
   Begin VB.CheckBox chkPromptPass 
      Caption         =   "Prompt for SysAdmin Password"
      Height          =   360
      Left            =   360
      TabIndex        =   5
      Top             =   3780
      Width           =   3435
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   435
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Permanently delete user from file"
      Top             =   3105
      Width           =   1125
   End
   Begin VB.TextBox Text3 
      Height          =   420
      Left            =   375
      TabIndex        =   2
      Top             =   2475
      Width           =   3720
   End
   Begin VB.CommandButton cmdAddToFile 
      Caption         =   "&Add"
      Height          =   435
      Left            =   2925
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Save entries to file"
      Top             =   3165
      Width           =   1125
   End
   Begin VB.TextBox Text2 
      Height          =   420
      Left            =   390
      TabIndex        =   1
      Top             =   1500
      Width           =   3720
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   375
      TabIndex        =   0
      Top             =   465
      Width           =   3720
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm Password"
      Height          =   240
      Left            =   405
      TabIndex        =   9
      Top             =   2115
      Width           =   2715
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Password"
      Height          =   255
      Left            =   375
      TabIndex        =   8
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Enter User Name"
      Height          =   300
      Left            =   390
      TabIndex        =   7
      Top             =   75
      Width           =   2520
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A program by Legrev3@aol.com
'Submitted for downloading Dec 6, 2000
'Demonstrates usage of module ReadWrite.bas for maintaining .ini file

Option Explicit

Private strUser As String
Private strPass As String


Private Sub Form_Activate()
    Dim strSysAdminPass As String
    Call GetDefaults
    If chkPromptPass.Value = vbChecked Then
        strSysAdminPass$ = InputBox("Only Authorized Users can Add, Edit, or Delete users and Passwords. Please enter the SysAdmin Password: ")
        'cancel button clicked
        If strSysAdminPass = "" Then
            Me.Hide
            Exit Sub
        ElseIf UCase(strSysAdminPass) <> "ALLOW" Then
            MsgBox "Invalid Password"
            Me.Hide
            Form1.Show
            Exit Sub
        End If
     End If
    If Text1 <> "" And Text2 <> "" Then Text3.SetFocus
    
    'set controls backcolor from form backcolor
    Dim i As Integer
    For i = 0 To Me.Controls.Count - 1
  
        If TypeOf Me.Controls(i) Is CommandButton Then
            Me.Controls(i).BackColor = Me.BackColor
        ElseIf TypeOf Me.Controls(i) Is Label Then
            Me.Controls(i).BackColor = Me.BackColor
        ElseIf TypeOf Me.Controls(i) Is CheckBox Then
            Me.Controls(i).BackColor = Me.BackColor
        End If
    Next i

End Sub

Private Sub GetDefaults()
'assign default values to text boxes from first form
    Text3.Text = ""
    If Form1.Text1 = "(enter your username)" Or Form1.Text1 = "" Then
        If strLoginName <> "" Then
            Text1 = strLoginName
            Text2 = strPassword
            strUser = strLoginName
            strPass = strPassword
        Else
            Text1 = "": Text2 = "": Text3 = ""
        End If
    Else
        Text1 = Form1.Text1
        Text2 = Form1.Text2
    End If
'determine if password reminder check box is to be shown
    Dim strChkValue As String
    Dim strChkKey As String
        
    strSection = "User Preferences"
    strChkKey = "chkNoReminder"

    strChkValue = ReadFromFile(strSection, strChkKey)
    
    If strChkValue = "vbChecked" Then
        chkNoReminder.Value = vbChecked
        chkNoReminder.Visible = False
    Else
        chkNoReminder.Value = vbUnchecked
        chkNoReminder.Visible = True
    End If

'determine if check box for adding new users should be checked
    strSection = "User Preferences"
    strChkKey = "chkPromptPass"

    strChkValue = ReadFromFile(strSection, strChkKey)
    
    If strChkValue = "vbChecked" Then
        chkPromptPass.Value = vbChecked
    Else
        chkPromptPass.Value = vbUnchecked
    End If

'get saved backcolor, if any
    If strUser <> "" Then
        strSection = "User Preferences"
        strColor = ReadFromFile(strSection, strUser)
        If strColor = "" Then
            lngColor = &H8000000F             'no user color preferences use system color
        Else
            lngColor = CLng(strColor)
        End If
        Me.BackColor = lngColor
    End If
End Sub

Private Sub cmdAddToFile_Click()
'writes to .ini file viewable in Notepad
    Dim strRetrievedPass As String
    
    strUser = Trim(Text1)
    strPass = Trim(Text2)

'doubble check entries
    If Len(strUser) < 3 Then
        MsgBox "UserName must be at least 3 characters in length.", vbCritical
        Text1.SetFocus
        Exit Sub
    ElseIf Len(strPass) < 3 Then
        MsgBox "Password must be at least 3 characters in length.", vbCritical
        Text2.SetFocus
        Exit Sub
    ElseIf Text2 <> Text3 Then
        MsgBox "Password confirmation failure. Please try again."
        Text3.SetFocus
        Exit Sub
    End If

'see if username is already on file
    strSection = "Password Section"
    strRetrievedPass = ReadFromFile(strSection, strUser)
    If strRetrievedPass <> "" Then
        MsgBox "UserName is on file, duplicates not allowed."
        Form1.Text2 = ""
        Me.Hide
        Form1.Show
        Exit Sub
    End If
    
'ready to write, call function to write
    lngRetVal = WriteToFile(strSection, strUser, strPass)
    If lngRetVal = 0 Then
        lngRetVal = MsgBox("Problem in adding User " & strUser & "to File.", vbRetryCancel)
        If lngRetVal = vbCancel Then
            Text1 = "": Text2 = ""
            Text1.SetFocus
        Else
            cmdAddToFile.SetFocus
        End If
    Else
        MsgBox "User " & strUser & " is now on File."
        Form1.Text1 = strUser
        strLoginName = strUser
        Form1.Text2 = strPass
        strPassword = strPass
        Form2.Hide
        Call Update_List    'updates list boxes for add/edit/delete
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim strRetrievedPass As String
    
    strUser = Trim(Text1)
    strPass = Trim(Text2)

    strSection = "Password Section"
    strRetrievedPass = ReadFromFile(strSection, strUser)

    If strRetrievedPass = "" Then
        MsgBox "User " & strUser & " is not on file."
        Me.Hide
        Form1.Show
        Form1.ZOrder 0
    Else
        lngRetVal = DeleteFromFile(strSection, strUser)
        If lngRetVal = 0 Then
            lngRetVal = MsgBox("Problem in Deleting User " & strUser & "From File.", vbRetryCancel)
            If lngRetVal = vbCancel Then
                Text1.SetFocus
            Else
                cmdDelete.SetFocus
            End If
        Else
            MsgBox "User " & strUser & " has been deleted."
            Form1.Text1 = "": Form2.Text2 = ""
            Text1 = "": Text2 = "": Text3 = ""
            Text1.SetFocus
            Call Update_List    'updates list boxes for add/edit/delete
        End If
    End If
    
End Sub


Private Sub chkNoReminder_Click()
'write user preference to file - do not show reminder
    Dim strChkValue As String
    Dim strChkKey As String
    Dim lngRetVal As Long
    
    If chkNoReminder.Value = vbChecked Then
        strChkValue = "vbChecked"
        chkNoReminder.Visible = True
    Else
        strChkValue = "vbUnchecked"
    End If
    
    strSection = "User Preferences"
    strChkKey = "chkNoReminder"

    lngRetVal = WriteToFile(strSection, strChkKey, strChkValue)
End Sub

Private Sub chkPromptPass_Click()
'write user preference to file - prompt or do not prompt for SysAdmin password
    Dim strChkValue As String
    Dim strChkKey As String
    
    If chkPromptPass.Value = vbChecked Then
        strChkValue = "vbChecked"
    Else
        strChkValue = "vbUnchecked"
    End If
    
    strSection = "User Preferences"
    strChkKey = "chkPromptPass"

    lngRetVal = WriteToFile(strSection, strChkKey, strChkValue)
End Sub

'all Sub Text are not very essential in this demo program and are largely for
'setting focus, for highlighting text, making the Enter key behave like Tab key
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> Asc(vbCr) Then Exit Sub
    Text2.SetFocus
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii <> Asc(vbCr) Then Exit Sub
    Text3.SetFocus
End Sub

Private Sub Text2_GotFocus()
    If Len(Text1) < 3 Then
        MsgBox "UserName must be at least 3 characters in length.", vbCritical
        Text1.SetFocus
    End If
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2)
End Sub


Private Sub Text3_GotFocus()
    If Len(Text2) < 3 Then
        MsgBox "Password must be at least 3 characters in length.", vbCritical
        Text2.SetFocus
    End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii <> Asc(vbCr) Then Exit Sub
    If Text2 <> Text3 Then
        MsgBox "Password confirmation failure. Please try again."
        Text2.SetFocus
    Else
        cmdAddToFile.Value = True
    End If
End Sub

Private Sub Update_List()
    If Form3.lstKeys.Visible = True Then Call Form3.mnuListKeys_Click
    If Form3.lstSections.Visible = True Then Call Form3.mnuListSections_Click
End Sub

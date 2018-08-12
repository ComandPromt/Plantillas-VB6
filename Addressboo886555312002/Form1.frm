VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3165
   ClientLeft      =   -45
   ClientTop       =   -435
   ClientWidth     =   2340
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   2340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtfields 
      Appearance      =   0  'Flat
      DataField       =   "Notes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      ToolTipText     =   "Notes"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtfields 
      Appearance      =   0  'Flat
      DataField       =   "Email"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   3
      ToolTipText     =   "contact email"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtfields 
      Appearance      =   0  'Flat
      DataField       =   "Cell"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "contact phone"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtfields 
      Appearance      =   0  'Flat
      DataField       =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Text            =   "Testentry"
      ToolTipText     =   "contact name"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtsearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "searchfield"
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   2340
      TabIndex        =   5
      Top             =   0
      Width           =   2340
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contacts..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "click me for database info"
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Image imgprint 
      Height          =   240
      Left            =   1200
      Picture         =   "Form1.frx":0442
      Stretch         =   -1  'True
      ToolTipText     =   "prints entry!"
      Top             =   2760
      Width           =   240
   End
   Begin VB.Image imgsearchnext 
      Height          =   240
      Left            =   2040
      Picture         =   "Form1.frx":074C
      Stretch         =   -1  'True
      ToolTipText     =   "find next entry or cicle thru entries"
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgcancel 
      Height          =   240
      Left            =   480
      Picture         =   "Form1.frx":0B8E
      ToolTipText     =   "cancel action"
      Top             =   2760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgedit 
      Height          =   300
      Left            =   480
      Picture         =   "Form1.frx":1118
      ToolTipText     =   "edit entry!"
      Top             =   2760
      Width           =   300
   End
   Begin VB.Image imgdelete 
      Height          =   240
      Left            =   1560
      Picture         =   "Form1.frx":11D3
      ToolTipText     =   "deletes an entry!"
      Top             =   2760
      Width           =   240
   End
   Begin VB.Image imglist 
      Height          =   300
      Left            =   120
      Picture         =   "Form1.frx":175D
      ToolTipText     =   "shows a list of entries"
      Top             =   840
      Width           =   300
   End
   Begin VB.Image imgphone 
      Height          =   300
      Left            =   120
      Picture         =   "Form1.frx":1814
      Top             =   1200
      Width           =   300
   End
   Begin VB.Image imgquit 
      Height          =   300
      Left            =   1920
      Picture         =   "Form1.frx":18B2
      ToolTipText     =   "quit and exit"
      Top             =   2760
      Width           =   300
   End
   Begin VB.Image imgmail 
      Height          =   240
      Left            =   120
      Picture         =   "Form1.frx":19CA
      ToolTipText     =   "opens Outlook"
      Top             =   1560
      Width           =   240
   End
   Begin VB.Image imgnote 
      Height          =   225
      Left            =   120
      Picture         =   "Form1.frx":1B14
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   240
   End
   Begin VB.Image imgnew 
      Height          =   300
      Left            =   120
      Picture         =   "Form1.frx":1F56
      ToolTipText     =   "create a new entry"
      Top             =   2760
      Width           =   300
   End
   Begin VB.Image imgupdate 
      Height          =   240
      Left            =   120
      Picture         =   "Form1.frx":2003
      ToolTipText     =   "update entry!"
      Top             =   2760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgsearch 
      Height          =   240
      Left            =   1680
      Picture         =   "Form1.frx":214D
      ToolTipText     =   "search for an entry!"
      Top             =   360
      Width           =   240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Component  :        frmmain
'   Project    :           smalladdressbook
'
'
'    Author     :      Sumari Arts
'    Modified   :      31/05/2002
'--------------------------------------------------------------------------------

Option Explicit
Dim WithEvents rs As Recordset
Attribute rs.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
 On Error GoTo ErrHandler
 
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\cont.mdb"

  Set rs = New Recordset
  rs.Open "select Name,Cell,Email,Notes from contacts", db, adOpenStatic, adLockOptimistic

Dim oText As TextBox
  'Bind the text boxes to the data provider
For Each oText In Me.txtfields
Set oText.DataSource = rs
Next


'works under win2k and XP
App.TaskVisible = False

'lock textboxes
Call Lockme
'load the list
Call listme

mbDataChanged = False


Exit_:
 Screen.MousePointer = vbNormal
 On Error Resume Next
 Exit Sub

ErrHandler:
 Screen.MousePointer = vbNormal
 MsgBox "Error..." & Err.Number & " in " & Err.Description, vbCritical
 Resume Exit_
End Sub







Private Sub Image5_Click()

End Sub

Private Sub imgdelete_Click()
On Error Resume Next
If txtfields(0).Text = "" Then
Exit Sub
End If
If MsgBox("Delete Entry??", vbExclamation + vbYesNo, "delete?") = vbYes Then
rs.Delete
Call listme
rs.MoveFirst
rs.Update
End If
End Sub

Private Sub imglist_Click()
On Error Resume Next
If List1.Visible = False Then
List1.Visible = True
Else
List1.Visible = False
End If
End Sub

Private Sub imgmail_Click()
Call NewMail
End Sub

Private Sub imgprint_Click()
On Error Resume Next
Dim Answer As String
Answer = MsgBox("Confirm printing on  " & _
Printer.DeviceName, vbYesNo, "print ... ?")
If Answer = vbNo Then Exit Sub
Printer.Print ""
Printer.Print ""
Printer.Print Date & Time
Printer.Print ""
Printer.Print txtfields(0).Text
Printer.Print txtfields(1).Text
Printer.Print txtfields(2).Text
Printer.Print txtfields(3).Text
Printer.EndDoc
End Sub

Private Sub imgquit_Click()
Dim X As Long
Dim inc As Long
inc = 50
On Error Resume Next
For X = Me.Height To 300 Step -inc
    DoEvents
        Me.Move Me.Left, Me.Top + (inc \ 2), Me.Width, X
    Next X
Screen.MousePointer = vbDefault
Unload Me
End
End Sub

Private Sub imgsearch_Click()
On Error Resume Next
rs.MoveFirst
Call FindString
End Sub
Private Sub imgsearchnext_Click()
'also use this button to move from entry to entry (searchfield empty)
   With rs
     .MoveNext
     If .EOF Then
        .MoveLast
        MsgBox "End of database reached. No entry found!", vbCritical, "Contacts"
      Else
        Call FindString
      End If
      End With
End Sub





Private Sub Label1_Click()
Dim c As Long
On Error Resume Next
Open App.Path & "\cont.mdb" For Binary As #1
c = LOF(1)
Close #1
MsgBox "You are currently in database-entry number:  " & CStr(rs.AbsolutePosition) & vbCrLf & _
"Total entries: " & CStr(rs.RecordCount) & vbCrLf & vbCrLf & _
"current database size: " & Format(c, "###,###,###,##0") & "k " & vbCrLf & vbCrLf & _
"Addressbook by (c)2002 Sumari Arts, pls. visit www.planet-source-code", vbInformation, "Contacts"
End Sub

Private Sub List1_Click()
On Error Resume Next
List1.ToolTipText = List1.Text
Call Search(List1.Text, rs, rs.Fields("name"))
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = vbLeftButton Then
  ReleaseCapture
 SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub


Private Sub Lockme()
Dim intFields As Integer

         For intFields = 0 To 3
         txtfields(intFields).Locked = True
         Next
End Sub

Private Sub unLockme()
Dim intFields As Integer

         For intFields = 0 To 3
         txtfields(intFields).Locked = False
         Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call imgquit_Click
End Sub

Private Sub rs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub SetButtons(bVal As Boolean)
  imgnew.Visible = bVal
  imgedit.Visible = bVal
  imgupdate.Visible = Not bVal
  imgcancel.Visible = Not bVal
  imgmail.Enabled = bVal
  imgdelete.Visible = bVal
  imgquit.Visible = bVal
  imglist.Enabled = bVal
  imgprint.Visible = bVal
  imgsearch.Enabled = bVal
  imgsearchnext.Enabled = bVal
End Sub

Private Sub imgnew_Click()
  On Error GoTo AddErr
  With rs
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
  Call unLockme
  List1.Visible = False
  
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub imgupdate_Click()
  On Error GoTo UpdateErr

  rs.UpdateBatch adAffectAll
  Call Lockme
  Call listme

  If mbAddNewFlag Then
    rs.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub
Private Sub imgedit_Click()
  On Error GoTo EditErr

Call unLockme
List1.Visible = False

  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub imgcancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  rs.CancelUpdate
  Call Lockme
  
  If mvBookMark > 0 Then
    rs.Bookmark = mvBookMark
  Else
    rs.MoveFirst
  End If
  mbDataChanged = False

End Sub
Public Sub FindString()

   Dim strFind As String
   Dim intFields As Integer
   
   On Error GoTo FindError
   
   If Trim(txtsearch) <> "" Then
     strFind = Trim(txtsearch)
     With rs
       Do Until .EOF
         For intFields = 0 To 3
           If InStr(1, frmMain.txtfields(intFields), strFind, _
                    vbTextCompare) > 0 Then
              frmMain.txtfields(intFields).SelStart = _
                      InStr(1, frmMain.txtfields(intFields), _
                            strFind, vbTextCompare) - 1
              frmMain.txtfields(intFields).SelLength = Len(strFind)
            frmMain.txtfields(intFields).SetFocus
              Exit Sub
            End If
          Next
          .MoveNext
          DoEvents
        Loop
        MsgBox "No contact found...", vbExclamation, "Contacts"
        .MoveFirst
      End With
     End If
     
     Exit Sub
     
FindError:
   
   MsgBox Err.Description
   Err.Clear
                    
End Sub
Private Sub txtsearch_KeyPress(KeyAscii As Integer)
On Error Resume Next
 If KeyAscii = 13 Then
    KeyAscii = 0 ' no peep
    Call imgsearch_Click
 End If
End Sub

Sub NewMail()
Dim olMail As MailItem
On Error GoTo ErrHandler
 
Set olMail = Application.CreateItem(olMailItem)
With olMail
.To = txtfields(2).Text
'here you can make a complete addon for attachments, text and stuff
'.Subject = "Drink Cola" or Text1.Text
'.Body = "Hi " or Text2.Text
'.Attachments.Add _
'Source:="C:\MyPictures\Cola.jpg" or Text3.Text
.Display
End With

Exit_:
 Screen.MousePointer = vbNormal
 On Error Resume Next
 Exit Sub

ErrHandler:
 Screen.MousePointer = vbNormal
 MsgBox "Error..." & Err.Number & " in " & Err.Description, vbCritical
 Resume Exit_

End Sub

Private Sub listme()
 List1.Clear
 rs.MoveFirst
 While Not rs.EOF
   List1.AddItem rs.Fields("Name")
   rs.MoveNext
 Wend
rs.MoveFirst
End Sub


VERSION 5.00
Object = "{D28F8786-0BB9-402B-92DC-F32DE23A324E}#3.0#0"; "OutlookBar.ocx"
Begin VB.Form Form1 
   Caption         =   "Outlook Bar Sample"
   ClientHeight    =   4752
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   6528
   LinkTopic       =   "Form1"
   ScaleHeight     =   4752
   ScaleWidth      =   6528
   StartUpPosition =   3  'Windows Default
   Begin OutlookBar.ctxOutlookBar ctxOutlookBar1 
      Align           =   3  'Align Left
      Height          =   4752
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   8382
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormatControl   =   "Form1a.frx":0000
      FormatGroup     =   "Form1a.frx":0174
      FormatGroupHover=   "Form1a.frx":0234
      FormatGroupPressed=   "Form1a.frx":02F4
      FormatGroupSelected=   "Form1a.frx":03C8
      FormatItem      =   "Form1a.frx":0474
      FormatItemLargeIcons=   "Form1a.frx":055C
      FormatItemHover =   "Form1a.frx":0658
      FormatItemPressed=   "Form1a.frx":0704
      FormatItemSelected=   "Form1a.frx":07B0
      FormatSmallIcon =   "Form1a.frx":085C
      FormatSmallIconHover=   "Form1a.frx":0944
      FormatSmallIconPressed=   "Form1a.frx":0A40
      FormatSmallIconSelected=   "Form1a.frx":0B3C
      FormatLargeIcon =   "Form1a.frx":0C38
      FormatLargeIconHover=   "Form1a.frx":0D20
      FormatLargeIconPressed=   "Form1a.frx":0E1C
      FormatLargeIconSelected=   "Form1a.frx":0F18
      Groups          =   "Form1a.frx":1014
      OleDragMode     =   1
      OleDropMode     =   1
      AllowGroupDrag  =   -1  'True
      LabelEdit       =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "More Samples"
      Height          =   432
      Left            =   4956
      MousePointer    =   4  'Icon
      TabIndex        =   1
      Top             =   84
      Width           =   1524
   End
   Begin VB.TextBox Text1 
      Height          =   4464
      Left            =   1512
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   3288
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuIcons 
         Caption         =   "Small Icons"
         Index           =   1
      End
      Begin VB.Menu mnuIcons 
         Caption         =   "Large Icons"
         Index           =   2
      End
      Begin VB.Menu mnuIcons 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuIcons 
         Caption         =   "Add Group"
         Index           =   4
      End
      Begin VB.Menu mnuIcons 
         Caption         =   "Remove Group"
         Index           =   5
      End
      Begin VB.Menu mnuIcons 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuIcons 
         Caption         =   "Rename"
         Index           =   7
      End
      Begin VB.Menu mnuIcons 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuIcons 
         Caption         =   "Cancel"
         Index           =   9
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const STR_LINK          As String = "http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=36529&optCodeRatingValue=5&intUserRatingTotal=0&intNumOfUserRatings=0"
Private Const MASK_COLOR        As Long = &HFF00FF

Private m_oImages           As Variant
Private m_oSel              As cButton
Private WithEvents m_oExtender As VBControlExtender
Attribute m_oExtender.VB_VarHelpID = -1
Private m_lIdx              As Long
Private m_hHandCursor       As Long

Private Enum UcsPopupMenu
    ucsMnuSmallIcons = 1
    ucsMnuLargeIcons
    ucsMnuSep1
    ucsMnuAddGroup
    ucsMnuRemoveGroup
    ucsMnuSep2
    ucsMnuRename
    ucsMnuSep3
    ucsMnuCancel
End Enum

Private Sub LogEvent(sText)
    Text1 = Text1 & sText & vbCrLf
    Text1.SelStart = Len(Text1)
End Sub

Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub ctxOutlookBar1_ButtonClick(ByVal oBtn As OutlookBar.cButton)
    If oBtn.Key = "vote" Then
        If MsgBox(vbCrLf & "Do you want to vote for this submission?" & vbCrLf & vbCrLf & vbCrLf & _
                "Note:" & vbTab & "Please, do this only if you feel that this submission is worth it! You will be navigated" & vbCrLf & _
                vbTab & "to the PSC page of this entry where you can validate your vote." & vbCrLf & vbCrLf & _
                vbTab & "Thank you in advance!" & vbCrLf & vbCrLf & _
                vbTab & "</wqw>", vbQuestion Or vbYesNo) = vbYes Then
            ShellExecute 0, "open", STR_LINK, "", "", 5
        End If
    End If
    If oBtn.Class = ucsBtnClassItem Then
'        oBtn.Enabled = Not oBtn.Enabled
    End If
End Sub

Private Sub ctxOutlookBar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim eStyle As UcsIconStyle
    If Button = vbRightButton Then
        eStyle = -1
        If Not ctxOutlookBar1.SelectedGroup Is Nothing Then
            eStyle = ctxOutlookBar1.SelectedGroup.IconsType
            mnuIcons(ucsMnuAddGroup).Enabled = True
            mnuIcons(ucsMnuRemoveGroup).Enabled = ctxOutlookBar1.SelectedGroup.Index > 3
        Else
            mnuIcons(ucsMnuAddGroup).Enabled = False
            mnuIcons(ucsMnuRemoveGroup).Enabled = False
        End If
        mnuIcons(ucsMnuSmallIcons).Checked = (eStyle = ucsIcsSmallIcons)
        mnuIcons(ucsMnuLargeIcons).Checked = (eStyle = ucsIcsLargeIcons)
        Set m_oSel = Nothing
        ctxOutlookBar1.HitTest X, Y, m_oSel
        mnuIcons(ucsMnuRename).Enabled = Not m_oSel Is Nothing
        Me.PopupMenu mnuPopup
    End If
End Sub

Private Sub ctxOutlookBar1_OLEBeforeMove(ByVal oBtn As OutlookBar.cButton, ByVal NewIndex As Long, Cancel As Boolean)
    Debug.Print "oBtn.Selected "; oBtn.Selected; Timer
End Sub

Private Sub ctxOutlookBar1_SelItemChange()
'    Set ctxOutlookBar1.SelectedItem = Nothing
End Sub

Private Sub Form_Load()
    Dim lIdx            As Long
    
    Set m_oExtender = ctxOutlookBar1
    With ctxOutlookBar1.Groups
        .Clear
        With .Add("Outlook Shortcuts", LoadResPicture("sm-Tasks.bmp", vbResBitmap)).GroupItems
            .Add("!!! VOTE !!!", LoadResPicture("la-Tasks.bmp", vbResBitmap), LoadResPicture("sm-Tasks.bmp", vbResBitmap), "vote").Parent.IconsType = ucsIcsLargeIcons
            .Add "My Computer", LoadResPicture("la-MyComputer.bmp", vbResBitmap), LoadResPicture("sm-MyComputer.bmp", vbResBitmap)
            .Add "Outlook Today", LoadResPicture("la-OutlookToday.bmp", vbResBitmap), LoadResPicture("sm-OutlookToday.bmp", vbResBitmap)
            .Add "Inbox", LoadResPicture("la-Inbox.bmp", vbResBitmap), LoadResPicture("sm-Inbox.bmp", vbResBitmap)
            With .Add("Calendar", LoadResPicture("la-Calendar.bmp", vbResBitmap), LoadResPicture("sm-Calendar.bmp", vbResBitmap))
                .Enabled = False
            End With
            .Add "Contacts", LoadResPicture("la-Contacts.bmp", vbResBitmap), LoadResPicture("sm-Contacts.bmp", vbResBitmap)
        End With
        With .Add("My Shortcuts").GroupItems
            .Add "Drafts", LoadResPicture("la-Drafts.bmp", vbResBitmap), LoadResPicture("sm-Drafts.bmp", vbResBitmap)
            .Add "Outbox", LoadResPicture("la-Outbox.bmp", vbResBitmap), LoadResPicture("sm-Outbox.bmp", vbResBitmap)
            .Add "Sent Items", LoadResPicture("la-SentItems.bmp", vbResBitmap), LoadResPicture("sm-SentItems.bmp", vbResBitmap)
            .Add "Journal", LoadResPicture("la-Journal.bmp", vbResBitmap), LoadResPicture("sm-Journal.bmp", vbResBitmap)
            .Add "Outlook Update", LoadResPicture("la-OutlookUpdate.bmp", vbResBitmap), LoadResPicture("sm-OutlookUpdate.bmp", vbResBitmap)
        End With
        With .Add("Other Shortcuts").GroupItems
            .Add("My Computer", LoadResPicture("la-MyComputer.bmp", vbResBitmap), LoadResPicture("sm-MyComputer.bmp", vbResBitmap)).Parent.IconsType = ucsIcsLargeIcons
            .Add("Right click" & vbCrLf & "for context menu", LoadResPicture("la-OutlookToday.bmp", vbResBitmap), LoadResPicture("sm-OutlookToday.bmp", vbResBitmap)).ToolTipText = "This is a two-line button" & vbCrLf & "with a two-line tooltip :-))"
        End With
        With .Add("New Group").GroupItems
            For lIdx = 0 To 19
                Select Case lIdx Mod 3
                Case 0
                    .Add("Item" & vbCrLf & "(" & (lIdx + 1) & ")", LoadResPicture("la-MyComputer.bmp", vbResBitmap), LoadResPicture("sm-MyComputer.bmp", vbResBitmap)).Parent.IconsType = ucsIcsLargeIcons
                Case 1
                    .Add "Journal" & vbCrLf & "(" & (lIdx + 1) & ")", LoadResPicture("la-Journal.bmp", vbResBitmap), LoadResPicture("sm-Journal.bmp", vbResBitmap)
                Case Else
                    .Add "Sent" & vbCrLf & "(" & (lIdx + 1) & ")", LoadResPicture("la-SentItems.bmp", vbResBitmap), LoadResPicture("sm-SentItems.bmp", vbResBitmap)
                End Select
            Next
        End With
    End With
    ctxOutlookBar1.Groups(2).Visible = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Height = ScaleHeight
End Sub

Private Function C2Str(v)
    On Error Resume Next
    C2Str = CStr(v)
End Function

Private Sub m_oExtender_ObjectEvent(Info As EventInfo)
    Dim sEvent As String
    Dim oParam As EventParameter
    
    If Info.Name <> "MouseMove" And Info.Name <> "OLEDragOver" And Info.Name <> "OLEGiveFeedback" Then
        sEvent = Info.Name & "("
        For Each oParam In Info.EventParameters
            sEvent = sEvent & C2Str(oParam.Value) & ", " ' & oParam.Name & ":="
        Next
        If Right(sEvent, 2) = ", " Then
            sEvent = Left(sEvent, Len(sEvent) - 2)
        End If
        LogEvent sEvent & ")"
    End If
End Sub

Private Sub mnuIcons_Click(Index As Integer)
    Select Case Index
    Case ucsMnuSmallIcons
        If Not ctxOutlookBar1.SelectedGroup Is Nothing Then
            ctxOutlookBar1.SelectedGroup.IconsType = ucsIcsSmallIcons
        End If
    Case ucsMnuLargeIcons
        If Not ctxOutlookBar1.SelectedGroup Is Nothing Then
            ctxOutlookBar1.SelectedGroup.IconsType = ucsIcsLargeIcons
        End If
    Case ucsMnuAddGroup
        m_lIdx = m_lIdx + 1
        With ctxOutlookBar1.Groups.Add("New Groups " & m_lIdx)
            .GroupItems.Add "Test " & m_lIdx, LoadResPicture("la-Tasks.bmp", vbResBitmap), LoadResPicture("sm-Tasks.bmp", vbResBitmap)
            .Selected = True
        End With
    Case ucsMnuRemoveGroup
        If Not ctxOutlookBar1.SelectedGroup Is Nothing Then
            ctxOutlookBar1.Groups.Remove ctxOutlookBar1.SelectedGroup.Index
        End If
    Case ucsMnuRename
        ctxOutlookBar1.StartLabelEdit m_oSel
    End Select
End Sub

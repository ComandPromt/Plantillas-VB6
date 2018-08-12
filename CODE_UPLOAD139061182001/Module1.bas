Attribute VB_Name = "Module1"
Option Explicit
'************************************************************************
'RecruiterApp is sort of an address book type app to keep track of
'**Jobs that you apply for that lets you add and delete contacts.
'**Keeps track of jobs that are StillPending (Highlights comments in red)
'**Go to website or send email to your contacts
'Using Access Database
'
'Author: Rick Bales copyright 2000
'rb.sb@gte.net
'Date: 12/21/2000
'**************************************************************************
'New features added on version 2
'query builder enhanced
'saved queries form added
'printing of reports added
'searching with wildcards added
'follow up reminder added
'show all records icon added
'1/18/2001
'***************************************************************************

Declare Sub ReleaseCapture Lib "user32" ()

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
        ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

'flag to stop reloading of the mainform when the splash form is
'unloaded as the about form
Public Started As Boolean
    
Sub CenterForm(frm As Form)
    
    'center child form in the MDI form taking into account the width
    'of the shortcutbar
    If Mainform.Picture1.Visible = True Then
        frm.Left = ((Mainform.Width - Mainform.Picture1.Width) - frm.Width) \ 2
    Else
        frm.Left = ((Mainform.Width) - frm.Width) \ 2
    End If
    
    frm.Top = (Mainform.ScaleHeight - frm.Height) \ 2
    
End Sub

Sub Main()

    'make sure path is correct
    ChDrive App.Path
    ChDir App.Path
    
    frmSplash.Show
    
End Sub

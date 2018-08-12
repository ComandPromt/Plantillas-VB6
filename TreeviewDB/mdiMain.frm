VERSION 5.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Faktuierung"
   ClientHeight    =   5430
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7620
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   2  'Maximiert
   Begin VB.Menu mnuBest 
      Caption         =   "Bestellungen"
      Begin VB.Menu mnuBestöffnen 
         Caption         =   "Bestellungen öffnen"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuBestöffnen_Click()
frmBestellungen.Show
End Sub

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
      Begin VB.Menu mnuBest�ffnen 
         Caption         =   "Bestellungen �ffnen"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuBest�ffnen_Click()
frmBestellungen.Show
End Sub

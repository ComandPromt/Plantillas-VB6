VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBestellungen 
   AutoRedraw      =   -1  'True
   Caption         =   "Bestellung"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   464
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Unten ausrichten
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6585
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11906
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "23.03.02"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TvwBestell 
      Height          =   6195
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   10927
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   14
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBestellungen.frx":0000
            Key             =   "SerieShut"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBestellungen.frx":02F4
            Key             =   "SerieOpen"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBestellungen.frx":05B8
            Key             =   "Teller"
            Object.Tag             =   "3"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabBestell 
      Height          =   6195
      Left            =   3300
      TabIndex        =   2
      Top             =   240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10927
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   794
      MouseIcon       =   "frmBestellungen.frx":08AC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Bestelldaten"
      TabPicture(0)   =   "frmBestellungen.frx":08C8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LvBestell"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picSerie"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.TextBox Text1 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   5760
         TabIndex        =   20
         Top             =   5760
         Width           =   1215
      End
      Begin VB.PictureBox picSerie 
         AutoRedraw      =   -1  'True
         Height          =   2295
         Left            =   240
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   485
         TabIndex        =   3
         Top             =   720
         Width           =   7335
         Begin VB.TextBox txtOrt 
            Height          =   315
            Left            =   4380
            TabIndex        =   19
            Tag             =   "1"
            Top             =   840
            Width           =   2715
         End
         Begin VB.TextBox txtPLZ 
            Height          =   315
            Left            =   3480
            TabIndex        =   18
            Tag             =   "1"
            Top             =   840
            Width           =   795
         End
         Begin VB.TextBox txtStraße 
            Height          =   315
            Left            =   3480
            TabIndex        =   16
            Tag             =   "1"
            Top             =   480
            Width           =   3615
         End
         Begin VB.TextBox txtFrachtkosten 
            Height          =   285
            Left            =   1320
            TabIndex        =   13
            Tag             =   "1"
            Top             =   1200
            Width           =   1035
         End
         Begin VB.TextBox txtVersandDatum 
            Height          =   315
            Left            =   1320
            TabIndex        =   7
            Tag             =   "1"
            Top             =   840
            Width           =   1035
         End
         Begin VB.TextBox txtBestellDatum 
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Tag             =   "1"
            Top             =   480
            Width           =   1035
         End
         Begin VB.TextBox txtEmpfänger 
            Height          =   315
            Left            =   3480
            TabIndex        =   5
            Tag             =   "1"
            Top             =   120
            Width           =   3615
         End
         Begin VB.TextBox txtBestellNr 
            Height          =   315
            Left            =   1320
            TabIndex        =   4
            Tag             =   "1"
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label3 
            Caption         =   "PLZ / Ort"
            Height          =   315
            Left            =   2520
            TabIndex        =   17
            Top             =   900
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Straße"
            Height          =   315
            Left            =   2520
            TabIndex        =   15
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Frachtkosten"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   1260
            Width           =   1155
         End
         Begin VB.Label lblVersandDatum 
            Caption         =   "Versanddatum"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label lblBestellDatum 
            Caption         =   "Bestelldatum"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   540
            Width           =   975
         End
         Begin VB.Label lblEmpfänger 
            Caption         =   "Empfänger "
            Height          =   315
            Left            =   2520
            TabIndex        =   9
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblBestellNr 
            BackStyle       =   0  'Transparent
            Caption         =   "Bestellnummer "
            Height          =   195
            Left            =   60
            TabIndex        =   8
            Top             =   180
            Width           =   1080
         End
      End
      Begin MSComctlLib.ListView LvBestell 
         Height          =   2115
         Left            =   240
         TabIndex        =   12
         Top             =   3540
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3731
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmBestellungen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsKunden          As Recordset
Dim rsBestellDetails  As Recordset

Dim lBestellungKey    As Long
Dim sBestellungNr     As String


Private Sub Form_Activate()
frmBestellungen.Height = 7356
frmBestellungen.Width = 11520
  Call initializeForm
End Sub

Private Sub Form_Load()
If (Not Datenbank()) Then
  MsgBox "Datenbank konnte nicht geöffnet werden !"
  End
End If
End Sub
Public Sub updateTree()
Dim rsBestellung    As Recordset
Dim rsKunden        As Recordset
Dim sqlKunde        As String
Dim sqlBestellung   As String

TvwBestell.Nodes.Clear

sqlKunde = "Select KundenCode, Firma "
sqlKunde = sqlKunde & "From Kunden ORDER BY "
sqlKunde = sqlKunde & "KundenCode"
Set rsKunden = dbNordwind.OpenRecordset(sqlKunde)
    If (rsKunden.RecordCount > 0) Then
  rsKunden.MoveFirst
End If



sqlBestellung = "SELECT BestellNr, Bestelldatum, KundenCode "
sqlBestellung = sqlBestellung & "FROM Bestellungen ORDER BY"
sqlBestellung = sqlBestellung & " Bestelldatum"
Set rsBestellung = dbNordwind.OpenRecordset(sqlBestellung)
If (rsBestellung.RecordCount > 0) Then
  rsBestellung.MoveFirst
End If


Do While Not rsKunden.EOF

TvwBestell.Nodes.Add , , "L" & rsKunden("KundenCode"), rsKunden("KundenCode"), 1, 1
   rsKunden.MoveNext
    Loop
Do While Not rsBestellung.EOF
    TvwBestell.Nodes.Add "L" & rsBestellung!Kundencode, _
    tvwChild, "ID" & CStr(rsBestellung!BestellNr), "Bestellnummer #" & (rsBestellung!BestellNr), 1, 2
    rsBestellung.MoveNext
    If (rsBestellung.EOF) Then
      Exit Do
    End If
  Loop
     DoEvents
sbStatus.Panels.Item(1).Text = "Es sind " & _
rsKunden.RecordCount & " Kunden in der Datenbank"

rsKunden.Close
rsBestellung.Close
End Sub

Public Sub setUpListView()
Dim clmHdr As ColumnHeader
Set clmHdr = LvBestell.ColumnHeaders. _
             Add(, , "Artikel", 3200, lvwColumnLeft)
Set clmHdr = LvBestell.ColumnHeaders. _
             Add(, , "Preis", 800, lvwColumnRight)
Set clmHdr = LvBestell.ColumnHeaders. _
             Add(, , "Anzahl", 800, lvwColumnCenter)
Set clmHdr = LvBestell.ColumnHeaders. _
             Add(, , "Rabatt", 800, lvwColumnRight)
Set clmHdr = LvBestell.ColumnHeaders. _
             Add(, , "Summe", 1000, lvwColumnRight)
Set clmHdr = LvBestell.ColumnHeaders. _
             Add(, , "BestellNr", 0)
        
LvBestell.View = lvwReport
End Sub

Public Sub initializeForm()
  Screen.MousePointer = vbHourglass
  sbStatus.Panels.Item(2).Text = "Laden..."
  tabBestell.Tab = 0
  DoEvents
  Call clearFields
  Call lockFields(True)
  Call updateTree
  Call setUpListView
  Screen.MousePointer = vbDefault
  sbStatus.Panels.Item(2).Text = "Bereit.."
End Sub

Public Sub clearFields()
Dim indx As Integer
Dim tempMask As String
With Me.Controls
  For indx = 0 To .Count - 1
    If Me.Controls(indx).Tag = "1" Then
       If (TypeOf Me.Controls(indx) Is TextBox) Then
           Me.Controls(indx).Text = ""
       ElseIf (TypeOf Me.Controls(indx) Is ComboBox) Then
      End If
    End If
  Next
End With
DoEvents
End Sub

Public Sub populateFields()
Call clearFields
With rsBestellung
If (Not IsNull(!BestellNr)) Then txtBestellNr = !BestellNr
If (Not IsNull(!Bestelldatum)) Then txtBestellDatum = !Bestelldatum
If (Not IsNull(!Empfänger)) Then txtEmpfänger = !Empfänger
If (Not IsNull(!Frachtkosten)) Then txtFrachtkosten = !Frachtkosten
If (Not IsNull(!Versanddatum)) Then txtVersandDatum = !Versanddatum
If (Not IsNull(!Straße)) Then txtStraße = !Straße
If (Not IsNull(!PLZ)) Then txtPLZ = !PLZ
If (Not IsNull(!Ort)) Then txtOrt = !Ort
DoEvents
End With
End Sub

Public Sub lockFields(bDoLock As Boolean)
Dim indx As Integer
For indx = 0 To Me.Controls.Count - 1
  If Me.Controls(indx).Tag = "1" Then
    If (TypeOf Me.Controls(indx) Is TextBox) Then
      If (bDoLock = True) Then
        Me.Controls(indx).Locked = True
        Me.Controls(indx).BackColor = &H808000
        Me.Controls(indx).ForeColor = vbWhite
      Else
        Me.Controls(indx).Locked = False
        Me.Controls(indx).BackColor = vbWhite
        Me.Controls(indx).BackColor = vbBlack
      End If
    ElseIf (TypeOf Me.Controls(indx) Is ComboBox) Then
      If (bDoLock = True) Then
        Me.Controls(indx).Enabled = False
        Me.Controls(indx).BackColor = vbWhite
         Me.Controls(indx).ForeColor = vbBlack
      Else
        Me.Controls(indx).Enabled = True
        Me.Controls(indx).BackColor = vbWhite
        Me.Controls(indx).BackColor = vbBlack
      End If
   End If
 End If
Next
DoEvents
End Sub

Private Sub TvwBestell_NodeClick(ByVal Node As MSComctlLib.Node)
If (Len(Node.Key) = 6) Then Exit Sub
lBestellungKey = CLng(Mid$(Node.Key, 3, Len(Node.Key)))
With rsBestellung
   .Index = "PrimaryKey"
   .Seek "=", lBestellungKey
   If Not .NoMatch Then
     bFieldsPopulated = True
     sBestellungNr = TvwBestell.SelectedItem
     Call populateFields
     Call populateListView
     Call ShowTotal

   Else
     MsgBox ("Ohhhh Nooo")
   End If
End With
End Sub

Public Sub populateListView()
Dim itemToAdd As ListItem
Dim BestellSQL As String
LvBestell.ListItems.Clear
BestellSQL = "SELECT BestellDetails.*,Bestelldetails.ArtikelNr, Artikel.Artikelname, "
BestellSQL = BestellSQL & " ([Bestelldetails].[Einzelpreis]*[Anzahl]*(1-[Rabatt])/100)*100 AS Sum"
BestellSQL = BestellSQL & " FROM Artikel INNER JOIN (Bestellungen INNER JOIN Bestelldetails "
BestellSQL = BestellSQL & " ON Bestellungen.BestellNr = Bestelldetails.BestellNr) "
BestellSQL = BestellSQL & " ON Artikel.[Artikel-Nr] = Bestelldetails.ArtikelNr "
BestellSQL = BestellSQL & " WHERE BestellDetails.BestellNr = " & _
                            lBestellungKey
BestellSQL = BestellSQL & " ORDER BY BestellDetails.BestellNr ASC"

Set rsBestellDetails = dbNordwind.OpenRecordset(BestellSQL)
If (rsBestellDetails.RecordCount > 0) Then
   rsBestellDetails.MoveFirst
    While Not rsBestellDetails.EOF
       Set itemToAdd = LvBestell.ListItems.Add(, , _
          rsBestellDetails!Artikelname)
    itemToAdd.SubItems(1) = Format$(rsBestellDetails!Einzelpreis, "#,##0.00;(#,##0.00)")
    itemToAdd.SubItems(2) = rsBestellDetails!Anzahl
    itemToAdd.SubItems(3) = Format$(rsBestellDetails!Rabatt, "#,##0.00%;(#,##0.00%)")
    itemToAdd.SubItems(4) = Format$(rsBestellDetails!Sum, "#,##0.00;(#,##0.00)")
    itemToAdd.SubItems(5) = CStr(rsBestellDetails!BestellNr)
       rsBestellDetails.MoveNext
   Wend
   sbStatus.Panels.Item(1).Text = " " & _
    rsBestellDetails.RecordCount & " Position(en) zu " & _
    sBestellungNr
Else
   Set itemToAdd = LvBestell.ListItems.Add(, , "Keine ")
   sbStatus.Panels.Item(1).Text = "0 Positionen " _
    & sBestellungNr
End If
LvBestell.SelectedItem = LvBestell.ListItems(1)
Call LvBestell_ItemClick(LvBestell.SelectedItem)
DoEvents
End Sub

Private Sub LvBestell_ItemClick(ByVal Item As MSComctlLib.ListItem)
If (rsBestellDetails.RecordCount > 0) Then
rsBestellDetails.MoveFirst
    rsBestellDetails.FindFirst "BestellNr = " & _
                 LvBestell.ListItems(Item.Index).SubItems(5)

End If
End Sub


Private Sub ShowTotal()
    Dim i As Integer
    Dim cTotal As Currency
    
    With LvBestell
        For i = 1 To .ListItems.Count
            cTotal = cTotal + CCur(.ListItems(i).SubItems(4))
        Next
    End With
Text1 = Format$(cTotal, "#,##0.00;(#,##0.00)")
    sbStatus.Panels.Item(3).Text = "Gesamt : " & Format$(cTotal, "#,##0.00;(#,##0.00)")
End Sub



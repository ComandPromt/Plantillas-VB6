VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1335
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   840
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":055C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1014
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1570
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1ACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2028
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2584
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2AE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":303C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3598
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4050
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":45AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1852
      BandCount       =   7
      FixedOrder      =   -1  'True
      _CBWidth        =   7095
      _CBHeight       =   1050
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar1"
      MinHeight1      =   600
      Width1          =   2235
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Caption2        =   "Address:"
      Child2          =   "Combo1"
      MinHeight2      =   315
      Width2          =   6000
      FixedBackground2=   0   'False
      NewRow2         =   -1  'True
      Caption3        =   "You"
      MinHeight3      =   360
      FixedBackground3=   0   'False
      NewRow3         =   0   'False
      Caption4        =   "can"
      MinHeight4      =   360
      FixedBackground4=   0   'False
      NewRow4         =   0   'False
      Caption5        =   "put"
      MinHeight5      =   360
      FixedBackground5=   0   'False
      NewRow5         =   0   'False
      Caption6        =   "anything"
      MinHeight6      =   360
      FixedBackground6=   0   'False
      NewRow6         =   0   'False
      Caption7        =   "here"
      MinHeight7      =   360
      FixedBackground7=   0   'False
      NewRow7         =   0   'False
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Text            =   "http://www.excite.com"
         Top             =   675
         Width           =   3345
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   600
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1058
         ButtonWidth     =   1323
         ButtonHeight    =   1058
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Back"
               Description     =   "Back"
               Object.Tag             =   "Back"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Forward"
               Description     =   "Forward"
               Object.Tag             =   "Forward"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Stop"
               Description     =   "Stop"
               Object.Tag             =   "Stop"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Description     =   "Refresh"
               Object.Tag             =   "Refresh"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Home"
               Description     =   "Home"
               Object.Tag             =   "Home"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Favorites"
               Description     =   "Favorites"
               Object.Tag             =   "Favorites"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Search"
               Description     =   "Search"
               Object.Tag             =   "Search"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'How to:

' 1) add a form..add the coolbar component
'    (the coolbar defaults with 3 bands on 2 rows)
' 2) drop a toolbar on the first row/first band..
'    this is where your browser buttons will go
' 3) add your imagelists(2 if you plan on having your
'    buttons image change when mouseover)
' 4) drop a combobox on the 2nd row/2nd band
' 5) now go to coolbar's properties, pick the
'    'bands' tab, pick index1(band 1) then in
'    the child dropdown menu, pick your toolbar
'    you added earlier
' 6) index2(band 2)...pick the combobox as the child

'this will embed both the toolbar and combobox into
'each assigned band. now you can go through and tweek
'your buttons the way you like

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Image Converter"
   ClientHeight    =   4455
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd2 
      Left            =   4320
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   120
      Top             =   120
      Width           =   4695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "Image"
      Begin VB.Menu mnuConvert 
         Caption         =   "Convert"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpme 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuReadMe 
         Caption         =   "Read Me"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Path$
Dim Filename$

Private Sub mnuAbout_Click()
MsgBox ("Image Convertor 1.0: Created By Mike Haislop Updates At www.Geocities.com/Viderianentertianment.  Thank You For Using Our Image Converter")
End Sub

Private Sub mnuConvert_Click()
On Error Resume Next
cd1.Filter = "Bmp (*.bmp)|*.bmp|Gif (*.gif)|*.gif|Jpeg (*.jpeg;*.jpg;*.jpe;*.jfif)|*.jpeg;*.jpg;*.jpe;*.jfif|TIFF (*.tiff;*.tif)|*.tiff;*.tif|Png (*.png)|*.png"
cd1.ShowOpen
Path$ = cd1.Filename
Call SavePicture(Image1, Path$)
End Sub

Private Sub mnuOpen_Click()
On Error Resume Next
cd2.Filter = "Image Files (*.*)|*.*"
cd2.ShowOpen
Filename$ = cd2.Filename
Image1.Picture = LoadPicture(Filename$) 'Loads the image into the form
Image1.Refresh
txtheight.Text = Image1.height
txtwidth.Text = Image1.Width
End Sub

Private Sub mnuSave_Click()
On Error Resume Next
cd1.Filter = "Bmp (*.bmp)|*.bmp|Gif (*.gif)|*.gif|Jpeg (*.jpeg;*.jpg;*.jpe;*.jfif)|*.jpeg;*.jpg;*.jpe;*.jfif|TIFF (*.tiff;*.tif)|*.tiff;*.tif|Png (*.png)|*.png"
cd1.ShowOpen
Path$ = cd1.Filename
Call SavePicture(Image1, Path$)
End Sub

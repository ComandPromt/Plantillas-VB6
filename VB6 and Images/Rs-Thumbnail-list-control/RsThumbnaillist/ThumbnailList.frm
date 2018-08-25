VERSION 5.00
Begin VB.Form FrmThumbnaillist 
   ClientHeight    =   6540
   ClientLeft      =   1215
   ClientTop       =   1980
   ClientWidth     =   9540
   Icon            =   "ThumbnailList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   9540
   Begin ThumbnaillistDemo.ThumbNailList ThumbNailList1 
      Height          =   6465
      Left            =   3000
      TabIndex        =   5
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   11404
      Path            =   "D:\A_Vb6_Usercontrols\0__Private\0__Grafisch\RsThumbnaillist"
      ThumbPage_FrameColor=   0
      ThumbNailColor  =   0
      Language        =   0
      CompressToJPG   =   0   'False
      Compress_NO_GIF =   0   'False
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   90
      ScaleHeight     =   285
      ScaleWidth      =   2865
      TabIndex        =   3
      Top             =   2700
      Width           =   2865
      Begin VB.CommandButton CmdShowThumbs 
         Caption         =   "Create Thumbnails -->>"
         Height          =   285
         Left            =   30
         TabIndex        =   4
         Top             =   0
         Width           =   2835
      End
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00000000&
      Height          =   3345
      Left            =   90
      Pattern         =   "*.JPG;*.gif;*.bmp;*.jpeg"
      TabIndex        =   2
      Top             =   2970
      Width           =   2880
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00000000&
      Height          =   2340
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   2880
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2880
   End
   Begin VB.Menu MnuBestand 
      Caption         =   "File"
      Begin VB.Menu MnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu MnuPrintPreview 
         Caption         =   "Print P&review"
      End
      Begin VB.Menu MnuBestandPrintopties 
         Caption         =   "Print &Setup"
         Begin VB.Menu MnuBestandPrintoptiesPrintresolutie 
            Caption         =   "Print Resolution"
            Begin VB.Menu Printresolutie 
               Caption         =   "25  Dpi"
               Index           =   0
            End
            Begin VB.Menu Printresolutie 
               Caption         =   "50  Dpi"
               Index           =   1
            End
            Begin VB.Menu Printresolutie 
               Caption         =   "75  Dpi"
               Index           =   2
            End
            Begin VB.Menu Printresolutie 
               Caption         =   "100  Dpi"
               Index           =   3
            End
            Begin VB.Menu Printresolutie 
               Caption         =   "125  Dpi"
               Index           =   4
            End
            Begin VB.Menu Printresolutie 
               Caption         =   "150  Dpi"
               Index           =   5
            End
            Begin VB.Menu Printresolutie 
               Caption         =   "175  Dpi"
               Index           =   6
            End
            Begin VB.Menu Printresolutie 
               Caption         =   "200  Dpi"
               Index           =   7
            End
            Begin VB.Menu Printresolutie 
               Caption         =   "225  Dpi"
               Index           =   8
            End
            Begin VB.Menu Printresolutie 
               Caption         =   "250 Dpi"
               Index           =   9
            End
            Begin VB.Menu Printresolutie 
               Caption         =   "275  Dpi"
               Index           =   10
            End
            Begin VB.Menu Printresolutie 
               Caption         =   "300  Dpi"
               Index           =   11
            End
         End
         Begin VB.Menu MnuBestandPrintoptiesMenu1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuBestandPrintoptiesThumbnailsperrij 
            Caption         =   "Thumbnails p. Row"
            Begin VB.Menu MnuBestandPrintoptiesThumbnailsperrijAantal 
               Caption         =   "1"
               Index           =   0
            End
            Begin VB.Menu MnuBestandPrintoptiesThumbnailsperrijAantal 
               Caption         =   "2"
               Index           =   1
            End
            Begin VB.Menu MnuBestandPrintoptiesThumbnailsperrijAantal 
               Caption         =   "3"
               Index           =   2
            End
            Begin VB.Menu MnuBestandPrintoptiesThumbnailsperrijAantal 
               Caption         =   "4"
               Index           =   3
            End
            Begin VB.Menu MnuBestandPrintoptiesThumbnailsperrijAantal 
               Caption         =   "5"
               Index           =   4
            End
            Begin VB.Menu MnuBestandPrintoptiesThumbnailsperrijAantal 
               Caption         =   "6"
               Index           =   5
            End
            Begin VB.Menu MnuBestandPrintoptiesThumbnailsperrijAantal 
               Caption         =   "7"
               Index           =   6
            End
            Begin VB.Menu MnuBestandPrintoptiesThumbnailsperrijAantal 
               Caption         =   "8"
               Index           =   7
            End
            Begin VB.Menu MnuBestandPrintoptiesThumbnailsperrijAantal 
               Caption         =   "9"
               Index           =   8
            End
            Begin VB.Menu MnuBestandPrintoptiesThumbnailsperrijAantal 
               Caption         =   "10"
               Index           =   9
            End
         End
         Begin VB.Menu MnuBestandPrintoptiesMenu2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuBestandPrintoptiesZoom 
            Caption         =   "Zoom"
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "10%"
               Index           =   0
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "20%"
               Index           =   1
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "30%"
               Index           =   2
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "40%"
               Index           =   3
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "50%"
               Index           =   4
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "60%"
               Index           =   5
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "70%"
               Index           =   6
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "80%"
               Index           =   7
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "90%"
               Index           =   8
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "100%"
               Index           =   9
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "110%"
               Index           =   10
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "120%"
               Index           =   11
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "130%"
               Index           =   12
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "140%"
               Index           =   13
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "150%"
               Index           =   14
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "160%"
               Index           =   15
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "170%"
               Index           =   16
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "180%"
               Index           =   17
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "190%"
               Index           =   18
            End
            Begin VB.Menu MnuBestandPrintoptiesZoomPercentage 
               Caption         =   "200%"
               Index           =   19
            End
         End
      End
      Begin VB.Menu MnuBestandMenu1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExportHtml 
         Caption         =   "Export to Html"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuBestandMenu2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBestandEinde 
         Caption         =   "End"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuHtml 
      Caption         =   "Html options"
      Begin VB.Menu MnuHtmlKleurpresets 
         Caption         =   "Color Presets (not ready)"
         Enabled         =   0   'False
         Begin VB.Menu Htmlcolorpresets 
            Caption         =   "Zwart"
            Index           =   0
         End
         Begin VB.Menu Htmlcolorpresets 
            Caption         =   "Antraciet"
            Index           =   1
         End
         Begin VB.Menu Htmlcolorpresets 
            Caption         =   "Donkergrijs"
            Index           =   2
         End
         Begin VB.Menu Htmlcolorpresets 
            Caption         =   "Middengrijs"
            Index           =   3
         End
         Begin VB.Menu Htmlcolorpresets 
            Caption         =   "Lichtgrijs"
            Index           =   4
         End
         Begin VB.Menu Htmlcolorpresets 
            Caption         =   "Witgrijs"
            Index           =   5
         End
      End
      Begin VB.Menu MnuHtmlFotorand 
         Caption         =   "Photo Border"
         Begin VB.Menu MnuHtmlFotorandopties 
            Caption         =   "None"
            Index           =   0
            Tag             =   "0"
         End
         Begin VB.Menu MnuHtmlFotorandopties 
            Caption         =   "Very Thin"
            Index           =   1
            Tag             =   "4"
         End
         Begin VB.Menu MnuHtmlFotorandopties 
            Caption         =   "Thin"
            Index           =   2
            Tag             =   "8"
         End
         Begin VB.Menu MnuHtmlFotorandopties 
            Caption         =   "Medium"
            Index           =   3
            Tag             =   "14"
         End
         Begin VB.Menu MnuHtmlFotorandopties 
            Caption         =   "Thick"
            Index           =   4
            Tag             =   "20"
         End
         Begin VB.Menu MnuHtmlFotorandopties 
            Caption         =   "Heavy"
            Index           =   5
            Tag             =   "30"
         End
      End
      Begin VB.Menu stH1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuThumbnailGrootte 
         Caption         =   "Thumbnail Size"
         Index           =   0
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "10 x 10  pixels"
            Index           =   0
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "20 x 20  pixels"
            Index           =   1
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "30 x 30  pixels"
            Index           =   2
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "40 x 40  pixels"
            Index           =   3
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "50 x 50  pixels"
            Index           =   4
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "60 x 60  pixels"
            Index           =   5
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "70 x 70  pixels"
            Index           =   6
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "80 x 80  pixels"
            Index           =   7
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "90 x 90  pixels"
            Index           =   8
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "100 x 100 pixels"
            Index           =   9
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "110 x 110  pixels"
            Index           =   10
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "120 x 120  pixels"
            Index           =   11
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "130 x 130  pixels"
            Index           =   12
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "140 x 140  pixels"
            Index           =   13
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "150 x 150  pixels"
            Index           =   14
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "160 x 160  pixels"
            Index           =   15
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "170 x 170  pixels"
            Index           =   16
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "180 x 180  pixels"
            Index           =   17
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "190 x 190  pixels"
            Index           =   18
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "200 x 200  pixels"
            Index           =   19
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "210 x 210  pixels"
            Index           =   20
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "220 x 220  pixels"
            Index           =   21
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "230 x 230  pixels"
            Index           =   22
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "240 x 240  pixels"
            Index           =   23
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "250 x 250  pixels"
            Index           =   24
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "260 x 260  pixels"
            Index           =   25
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "270 x 270  pixels"
            Index           =   26
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "280 x 280  pixels"
            Index           =   27
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "290 x 290  pixels"
            Index           =   28
         End
         Begin VB.Menu MnuHtmlThumbsize 
            Caption         =   "300 x 300  pixels"
            Index           =   29
         End
      End
   End
   Begin VB.Menu mnuThumbNailOpties 
      Caption         =   "ThumbNail Options"
      Begin VB.Menu MnuThumbnailGrootte2 
         Caption         =   "Thumbnail Size"
         Index           =   0
         Begin VB.Menu MnuThumbsize 
            Caption         =   "10 x 10  pixels"
            Index           =   0
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "20 x 20  pixels"
            Index           =   1
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "30 x 30  pixels"
            Index           =   2
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "40 x 40  pixels"
            Index           =   3
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "50 x 50  pixels"
            Index           =   4
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "60 x 60  pixels"
            Index           =   5
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "70 x 70  pixels"
            Index           =   6
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "80 x 80  pixels"
            Index           =   7
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "90 x 90  pixels"
            Index           =   8
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "100 x 100 pixels"
            Index           =   9
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "110 x 110  pixels"
            Index           =   10
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "120 x 120  pixels"
            Index           =   11
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "130 x 130  pixels"
            Index           =   12
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "140 x 140  pixels"
            Index           =   13
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "150 x 150  pixels"
            Index           =   14
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "160 x 160  pixels"
            Index           =   15
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "170 x 170  pixels"
            Index           =   16
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "180 x 180  pixels"
            Index           =   17
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "190 x 190  pixels"
            Index           =   18
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "200 x 200  pixels"
            Index           =   19
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "210 x 210  pixels"
            Index           =   20
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "220 x 220  pixels"
            Index           =   21
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "230 x 230  pixels"
            Index           =   22
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "240 x 240  pixels"
            Index           =   23
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "250 x 250  pixels"
            Index           =   24
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "260 x 260  pixels"
            Index           =   25
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "270 x 270  pixels"
            Index           =   26
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "280 x 280  pixels"
            Index           =   27
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "290 x 290  pixels"
            Index           =   28
         End
         Begin VB.Menu MnuThumbsize 
            Caption         =   "300 x 300  pixels"
            Index           =   29
         End
      End
      Begin VB.Menu mnuThumbNailOptiesMenu3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExtraHoogte 
         Caption         =   "Extra Height"
         Index           =   0
         Begin VB.Menu MnuExtraheight 
            Caption         =   "None"
            Index           =   30
         End
         Begin VB.Menu Menu1 
            Caption         =   "10  pixels"
            Index           =   0
         End
         Begin VB.Menu Menu2 
            Caption         =   "20  pixels"
            Index           =   1
         End
         Begin VB.Menu Menu3 
            Caption         =   "30  pixels"
            Index           =   2
         End
         Begin VB.Menu Menu4 
            Caption         =   "40  pixels"
            Index           =   3
         End
         Begin VB.Menu Menu5 
            Caption         =   "50  pixels"
            Index           =   4
         End
         Begin VB.Menu Menu6 
            Caption         =   "60  pixels"
            Index           =   5
         End
         Begin VB.Menu Menu7 
            Caption         =   "70  pixels"
            Index           =   6
         End
         Begin VB.Menu Menu8 
            Caption         =   "80  pixels"
            Index           =   7
         End
         Begin VB.Menu Menu9 
            Caption         =   "90  pixels"
            Index           =   8
         End
         Begin VB.Menu Menu10 
            Caption         =   "100 pixels"
            Index           =   9
         End
         Begin VB.Menu Menu11 
            Caption         =   "110  pixels"
            Index           =   10
         End
         Begin VB.Menu Menu12 
            Caption         =   "120  pixels"
            Index           =   11
         End
         Begin VB.Menu Menu13 
            Caption         =   "130  pixels"
            Index           =   12
         End
         Begin VB.Menu Menu14 
            Caption         =   "140  pixels"
            Index           =   13
         End
         Begin VB.Menu Menu15 
            Caption         =   "150 pixels"
            Index           =   14
         End
         Begin VB.Menu Menu16 
            Caption         =   "160  pixels"
            Index           =   15
         End
         Begin VB.Menu Menu17 
            Caption         =   "170  pixels"
            Index           =   16
         End
         Begin VB.Menu Menu18 
            Caption         =   "180  pixels"
            Index           =   17
         End
         Begin VB.Menu Menu19 
            Caption         =   "190  pixels"
            Index           =   18
         End
         Begin VB.Menu Menu20 
            Caption         =   "200  pixels"
            Index           =   19
         End
         Begin VB.Menu Menu21 
            Caption         =   "210  pixels"
            Index           =   20
         End
         Begin VB.Menu Menu22 
            Caption         =   "220  pixels"
            Index           =   21
         End
         Begin VB.Menu Menu23 
            Caption         =   "230  pixels"
            Index           =   22
         End
         Begin VB.Menu Menu24 
            Caption         =   "240  pixels"
            Index           =   23
         End
         Begin VB.Menu Menu25 
            Caption         =   "250  pixels"
            Index           =   24
         End
         Begin VB.Menu Menu26 
            Caption         =   "260  pixels"
            Index           =   25
         End
         Begin VB.Menu Menu27 
            Caption         =   "270  pixels"
            Index           =   26
         End
         Begin VB.Menu Menu28 
            Caption         =   "280  pixels"
            Index           =   27
         End
         Begin VB.Menu Menu29 
            Caption         =   "290  pixels"
            Index           =   28
         End
         Begin VB.Menu Menu30 
            Caption         =   "300  pixels"
            Index           =   29
         End
      End
      Begin VB.Menu MnuExtraBreedte 
         Caption         =   "Extra width"
         Index           =   0
         Begin VB.Menu MnuExtraWidth 
            Caption         =   "none"
            Index           =   30
         End
         Begin VB.Menu Menu31 
            Caption         =   "10  pixels"
            Index           =   0
         End
         Begin VB.Menu Menu32 
            Caption         =   "20  pixels"
            Index           =   1
         End
         Begin VB.Menu Menu33 
            Caption         =   "30  pixels"
            Index           =   2
         End
         Begin VB.Menu Menu34 
            Caption         =   "40  pixels"
            Index           =   3
         End
         Begin VB.Menu Menu35 
            Caption         =   "50  pixels"
            Index           =   4
         End
         Begin VB.Menu Menu36 
            Caption         =   "60  pixels"
            Index           =   5
         End
         Begin VB.Menu Menu37 
            Caption         =   "70  pixels"
            Index           =   6
         End
         Begin VB.Menu Menu38 
            Caption         =   "80  pixels"
            Index           =   7
         End
         Begin VB.Menu Menu39 
            Caption         =   "90  pixels"
            Index           =   8
         End
         Begin VB.Menu Menu40 
            Caption         =   "100 pixels"
            Index           =   9
         End
         Begin VB.Menu Menu41 
            Caption         =   "110  pixels"
            Index           =   10
         End
         Begin VB.Menu Menu42 
            Caption         =   "120  pixels"
            Index           =   11
         End
         Begin VB.Menu Menu43 
            Caption         =   "130  pixels"
            Index           =   12
         End
         Begin VB.Menu Menu44 
            Caption         =   "140  pixels"
            Index           =   13
         End
         Begin VB.Menu Menu45 
            Caption         =   "150 pixels"
            Index           =   14
         End
         Begin VB.Menu Menu46 
            Caption         =   "160  pixels"
            Index           =   15
         End
         Begin VB.Menu Menu47 
            Caption         =   "170  pixels"
            Index           =   16
         End
         Begin VB.Menu Menu48 
            Caption         =   "180  pixels"
            Index           =   17
         End
         Begin VB.Menu Menu49 
            Caption         =   "190  pixels"
            Index           =   18
         End
         Begin VB.Menu Menu50 
            Caption         =   "200  pixels"
            Index           =   19
         End
         Begin VB.Menu Menu51 
            Caption         =   "210  pixels"
            Index           =   20
         End
         Begin VB.Menu Menu52 
            Caption         =   "220  pixels"
            Index           =   21
         End
         Begin VB.Menu Menu53 
            Caption         =   "230  pixels"
            Index           =   22
         End
         Begin VB.Menu Menu54 
            Caption         =   "240  pixels"
            Index           =   23
         End
         Begin VB.Menu Menu55 
            Caption         =   "250  pixels"
            Index           =   24
         End
         Begin VB.Menu Menu56 
            Caption         =   "260  pixels"
            Index           =   25
         End
         Begin VB.Menu Menu57 
            Caption         =   "270  pixels"
            Index           =   26
         End
         Begin VB.Menu Menu58 
            Caption         =   "280  pixels"
            Index           =   27
         End
         Begin VB.Menu Menu59 
            Caption         =   "290  pixels"
            Index           =   28
         End
         Begin VB.Menu Menu60 
            Caption         =   "300  pixels"
            Index           =   29
         End
      End
      Begin VB.Menu streepje2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuThumbNailOptiesMaximumPaginas 
         Caption         =   "Max Pages"
         Begin VB.Menu MnuMaxpage 
            Caption         =   "10 - Pagina's"
            Index           =   0
         End
         Begin VB.Menu MnuMaxpage 
            Caption         =   "20 - Pagina's"
            Index           =   1
         End
         Begin VB.Menu MnuMaxpage 
            Caption         =   "30 - Pagina's"
            Index           =   2
         End
         Begin VB.Menu MnuMaxpage 
            Caption         =   "40 - Pagina's"
            Index           =   3
         End
         Begin VB.Menu MnuMaxpage 
            Caption         =   "50 - Pagina's"
            Index           =   4
         End
         Begin VB.Menu MnuMaxpage 
            Caption         =   "60 - Pagina's"
            Index           =   5
         End
         Begin VB.Menu MnuMaxpage 
            Caption         =   "70 - Pagina's"
            Index           =   6
         End
         Begin VB.Menu MnuMaxpage 
            Caption         =   "80 - Pagina's"
            Index           =   7
         End
         Begin VB.Menu MnuMaxpage 
            Caption         =   "90 - Pagina's"
            Index           =   8
         End
         Begin VB.Menu MnuMaxpage 
            Caption         =   "100 - Pagina's"
            Index           =   9
         End
      End
      Begin VB.Menu mnuThumbNailOptiesMenu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuThumbNailOptiesMenu1 
         Caption         =   "Color Presets"
         Begin VB.Menu ThumbnailColorPreset 
            Caption         =   "Black"
            Index           =   0
         End
         Begin VB.Menu ThumbnailColorPreset 
            Caption         =   "Antraciet"
            Index           =   1
         End
         Begin VB.Menu ThumbnailColorPreset 
            Caption         =   "Dark Gray"
            Index           =   2
         End
         Begin VB.Menu ThumbnailColorPreset 
            Caption         =   "Middle Gray"
            Index           =   3
         End
         Begin VB.Menu ThumbnailColorPreset 
            Caption         =   "Light Gray"
            Index           =   4
         End
         Begin VB.Menu ThumbnailColorPreset 
            Caption         =   "White Gray"
            Index           =   5
         End
      End
   End
End
Attribute VB_Name = "FrmThumbnaillist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim tel
Private Sub CboThumbsize_Click()
  ThumbNailList1.ThumbNailSize = Val(CboThumbsize.Text)
End Sub

Private Sub CboZoom_Click()
   ThumbNailList1.Print_Zoomfactor = CboZoom.Text
End Sub

Private Sub cmdAdd_Click()
    Data1.Recordset.AddNew
    '---------------------
    CmdDelete.Enabled = True
    cmdAdd.Enabled = True
    '--------------------
    If Data1.Recordset.EditMode = 1 Then Exit Sub
    Data1.UpdateRecord
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub

Private Sub CmdDelete_Click()
    CmdDelete.Enabled = True
    cmdAdd.Enabled = True
    With Data1.Recordset
        If .RecordCount = 0 Then Exit Sub
        .Delete
        .MoveNext
        If .RecordCount > 0 And .EOF Then
          .MoveLast
        End If
    End With
     Exit Sub
DeleteErr:
     MsgBox Err.Description

End Sub

Private Sub CheckAutoupdate_Click()
 If CheckAutoupdate.Value = 1 Then
    ThumbNailList1.AutoUpdateOnPathchange = True
 Else
    ThumbNailList1.AutoUpdateOnPathchange = False
 End If
End Sub

Private Sub CmdShowThumbs_Click()
 ThumbNailList1.Makegallery
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
   On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
 pad = File1.Path
 If Right$(pad, 1) <> "\" Then pad = pad & "\"
 Fname = pad & File1.Filename
End Sub


Private Sub File1_DblClick()
   MsgBox "Click on the button << Create Thumbnails >>"
   
End Sub

Private Sub File1_PathChange()
   ThumbNailList1.Path = Dir1.Path
   ' write tothe rigistry
   Call SaveSetting(App.EXEName, "Locations", "drive", Left$(Dir1.Path, 2))
   Call SaveSetting(App.EXEName, "Locations", "Path", Dir1.Path)
   '===================
   If File1.ListCount > 0 Then
      CmdShowThumbs.Enabled = True
   Else
      CmdShowThumbs.Enabled = False
   End If

End Sub

Private Sub Form_Load()
'On Error Resume Next

' Get setting from the windows registry
' Syntax -->> Mysetting = GetSetting(App.EXEName, "Selection", "key", DefaultValue)
' Excuse me for the curious Dutch language in some codeparts here.

Drv$ = GetSetting(App.EXEName, "Locations", "drive", Left$(App.Path, 2))
pad$ = GetSetting(App.EXEName, "Locations", "Path", App.Path)
Thumb_Size = Val(GetSetting(App.EXEName, "Thumbnail", "Size", "100"))                 ' 100 pixels
m_Extraheight = Val(GetSetting(App.EXEName, "Thumbnail", "Extraheight", "0"))         ' 0   pixels
m_ExtraWidth = Val(GetSetting(App.EXEName, "Thumbnail", "ExtraWidth", "0"))           ' 0   pixels
ThumbColorpreset = Val(GetSetting(App.EXEName, "Thumbnail", "Colorpreset", "5"))      ' lightgrey
PrintZoomfactor = Val(GetSetting(App.EXEName, "Print", "Zoomfactor", "100"))          ' 100%
PrintThumbsPerrij = Val(GetSetting(App.EXEName, "Print", "ThumbsPerRow", "5"))        ' 5 thumbnails
Print_resolutie = Val(GetSetting(App.EXEName, "Print", "Resolution", "300"))          ' 300 Dpi
autoUpdate = CBool(GetSetting(App.EXEName, "Control", "autoUpdate", "False"))
HtmlBorderdikte = GetSetting(App.EXEName, "Html", "key", "14")                        ' medium Borderthicknes
Html_Color_Preset = GetSetting(App.EXEName, "Html", "key", "1")                       ' antraciet
htmlThumbsize = Val(GetSetting(App.EXEName, "Html", "key", "200"))                    ' 200 pixels
m_MnuMaxpage = Val(GetSetting(App.EXEName, "Selection", "key", "100"))                ' 100 thumbnail pages maximum (lower system resources)



Drive1.Drive = Drv$
Dir1.Path = pad$
File1.Path = pad$
ThumbNailList1.Html_ThumbnailSize = htmlThumbsize
ThumbNailList1.ThumbNailSize = Thumb_Size
ThumbNailList1.Print_Zoomfactor = PrintZoomfactor
ThumbNailList1.PrintResolution = Print_resolutie
ThumbNailList1.Print_NrThumbCols = PrintThumbsPerrij
ThumbNailList1.Html_PageColorPresets = Html_Color_Preset
ThumbNailList1.ThumbPageColorPresets = ThumbColorpreset
ThumbNailList1.AutoUpdateOnPathchange = CBool(autoUpdate)
ThumbNailList1.Html_ThumbnailBorderThicknes = HtmlBorderdikte
ThumbNailList1.MaxThumbnailpages = m_MnuMaxpage
ThumbNailList1.ThumbnailExtraHeight = m_Extraheight
ThumbNailList1.ThumbnailExtraWidth = m_ExtraWidth


If CBool(autoUpdate) = True Then CheckAutoupdate.Value = 1

' Put the checkmarks in the Menu's
MnuHtmlThumbsize((htmlThumbsize \ 10) - 1).Checked = True
MnuThumbsize((Thumb_Size \ 10) - 1).Checked = True
MnuBestandPrintoptiesZoomPercentage((PrintZoomfactor \ 10) - 1).Checked = True
MnuBestandPrintoptiesThumbnailsperrijAantal(PrintThumbsPerrij - 1).Checked = True
Printresolutie((Print_resolutie \ 25) - 1).Checked = True ' 11

ThumbnailColorPreset(ThumbColorpreset - 1).Checked = True
Htmlcolorpresets(Html_Color_Preset - 1).Checked = True

For I = 0 To 5
  If MnuHtmlFotorandopties(I).Tag = HtmlBorderdikte Then
     MnuHtmlFotorandopties(I).Checked = True
     Exit For
  End If
Next

MnuMaxpage((m_MnuMaxpage \ 10) - 1).Checked = True

If m_Extraheight = 0 Then
   MnuExtraheight(30).Checked = True
Else
   MnuExtraheight((m_Extraheight \ 10) - 1).Checked = True
End If
If m_ExtraWidth = 0 Then
   MnuExtraWidth(30).Checked = True
Else
   MnuExtraWidth((m_ExtraWidth \ 10) - 1).Checked = True
End If


' open a file during the programrun to report errors
Open App.Path & "\Log.txt" For Output As #20

End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub Form_Resize()
 If Me.WindowState = 1 Then Exit Sub
 marge = 20
 
 ch = (Me.ScaleHeight - (Drive1.Height + P1.Height + marge)) \ 2

 Dir1.Height = ch
 File1.Height = ch
 '----------------
 Drive1.Top = 0
 Dir1.Top = Drive1.Height
 P1.Top = Dir1.Top + Dir1.Height
 File1.Top = P1.Top + P1.Height
 
 '-----------------
 Drive1.Left = marge
 Dir1.Left = marge
 File1.Left = marge
 P1.Left = marge
 '------------------
 ThumbNailList1.Height = Me.ScaleHeight
 ThumbNailList1.Left = Dir1.Left + Dir1.Width + marge
 ThumbNailList1.Width = Me.ScaleWidth - ThumbNailList1.Left
End Sub


Private Sub MnuBestandEinde_Click()
Unload Me
End
End Sub

Private Sub MnuBestandPrintoptiesThumbnailsperrijAantal_Click(Index As Integer)
    For I = 0 To 9
        MnuBestandPrintoptiesThumbnailsperrijAantal(I).Checked = False
    Next
    MnuBestandPrintoptiesThumbnailsperrijAantal(Index).Checked = True
    ThumbNailList1.Print_NrThumbCols = Index + 1
    
    Call SaveSetting(App.EXEName, "Print", "ThumbsPerRow", CStr(ThumbNailList1.Print_NrThumbCols))        ' 5 thumbnails
    
    ThumbNailList1.PrintPreview
End Sub

Private Sub MnuBestandPrintoptiesZoomPercentage_Click(Index As Integer)
    For I = 0 To 19
        MnuBestandPrintoptiesZoomPercentage(I).Checked = False
    Next
    MnuBestandPrintoptiesZoomPercentage(Index).Checked = True
    waarde = (Index + 1) * 10
    ThumbNailList1.Print_Zoomfactor = waarde
    
    Call SaveSetting(App.EXEName, "Print", "Zoomfactor", CStr(ThumbNailList1.Print_Zoomfactor))          ' 100%
    
    ThumbNailList1.PrintPreview
End Sub

Private Sub MnuExtraheight_Click(Index As Integer)
    For I = 0 To 30
      MnuExtraheight(I).Checked = False
    Next
    MnuExtraheight(Index).Checked = True
    waarde = (Index + 1) * 10
    If Index = 30 Then waarde = 0
    ThumbNailList1.ThumbnailExtraHeight = waarde
    
    Call SaveSetting(App.EXEName, "Thumbnail", "Extraheight", CStr(ThumbNailList1.ThumbnailExtraHeight))          ' 0   pixels
    
    
    ThumbNailList1.Makegallery
End Sub
Private Sub MnuExtrawidth_Click(Index As Integer)
    For I = 0 To 30
      MnuExtraWidth(I).Checked = False
    Next
    MnuExtraWidth(Index).Checked = True
    waarde = (Index + 1) * 10
    If Index = 30 Then waarde = 0
    ThumbNailList1.ThumbnailExtraWidth = waarde
    
    Call SaveSetting(App.EXEName, "Thumbnail", "ExtraWidth", CStr(ThumbNailList1.ThumbnailExtraWidth))           ' 0   pixels
    
    ThumbNailList1.Makegallery
End Sub

Private Sub MnuHtmlFotorandopties_Click(Index As Integer)
    For I = 0 To 5
        MnuHtmlFotorandopties(I).Checked = False
    Next
    MnuHtmlFotorandopties(Index).Checked = False
    ThumbNailList1.Html_ThumbnailBorderThicknes = MnuHtmlFotorandopties(Index).Tag
    
    Call SaveSetting(App.EXEName, "Html", "key", CStr(ThumbNailList1.Html_ThumbnailBorderThicknes))                        ' medium Borderthicknes
    
    ThumbNailList1.Html_Thumbnail_Preview
End Sub

Private Sub MnuHtmlThumbsize_Click(Index As Integer)
    For I = 0 To 29
        MnuHtmlThumbsize(I).Checked = False
    Next
    MnuHtmlThumbsize(Index).Checked = True
    waarde = (Index + 1) * 10
    ThumbNailList1.Html_ThumbnailSize = waarde
    
    Call SaveSetting(App.EXEName, "Html", "key", CStr(ThumbNailList1.Html_ThumbnailSize))                    ' 200 pixels
    
    ThumbNailList1.Html_Thumbnail_Preview
End Sub

Private Sub MnuMaxpage_Click(Index As Integer)
   For I = 0 To 9
        MnuMaxpage(I).Checked = False
   Next
   MnuMaxpage(Index).Checked = True
   ThumbNailList1.MaxThumbnailpages = (Index + 1) * 10

   Call SaveSetting(App.EXEName, "Selection", "key", CStr(ThumbNailList1.MaxThumbnailpages))                ' 100 thumbnail pages maximum (lower system resources)

End Sub

Private Sub MnuThumbsize_Click(Index As Integer)
    For I = 0 To 29
        MnuThumbsize(I).Checked = False
    Next
    MnuThumbsize(Index).Checked = True
    waarde = (Index + 1) * 10
    ThumbNailList1.ThumbNailSize = waarde

    Call SaveSetting(App.EXEName, "Thumbnail", "Size", CStr(ThumbNailList1.ThumbNailSize))                 ' 100 pixels
    
    Pages = ThumbNailList1.ThumbnailPagesNeeded
    TPPage = ThumbNailList1.ThumbnailsPerPage
    
        
    st = st & "Pages " & Pages
    st = st & "  Thumbnails per page " & TPPage
    Me.Caption = st
    ThumbNailList1.Makegallery
End Sub



Private Sub Printresolutie_Click(Index As Integer)
  MsgBox "Not ready yet"
  Call SaveSetting(App.EXEName, "Print", "Resolution", "300")          ' 300 Dpi
End Sub

Private Sub ThumbnailColorPreset_Click(Index As Integer)
    For I = 0 To 5
        ThumbnailColorPreset(I).Checked = False
    Next
    ThumbnailColorPreset(Index).Checked = True
    ThumbNailList1.ThumbPageColorPresets = Index + 1
    
    Call SaveSetting(App.EXEName, "Thumbnail", "Colorpreset", CStr(ThumbNailList1.ThumbPageColorPresets))      ' lightgrey
    
    ThumbNailList1.Makegallery
End Sub
Private Sub Htmlcolorpresets_Click(Index As Integer)
    For I = 0 To 5
        Htmlcolorpresets(I).Checked = False
    Next
    Htmlcolorpresets(Index).Checked = True
    ThumbNailList1.Html_PageColorPresets = Index + 1
    
    Call SaveSetting(App.EXEName, "Html", "key", CStr(ThumbNailList1.Html_PageColorPresets))                       ' antraciet
    
    ThumbNailList1.Html_Thumbnail_Preview
    MsgBox "This option is not finished yet."
End Sub

Private Sub MnuExportHtml_Click()
 Destination = Left$(File1.Path, 2)
 ' BodyText = " test tekst"
 
 st = st & "<a href=http://storm.prohosting.com/roysmol/>Vb pagina Roy Smol</a> "
 st = st & "<a HREF=http://people.zeelandnet.nl/killroy/CV_Roy/CV_Roy.htm>Cv Roy Smol</a> "
 st = st & "<a HREF=mailto:roysmol@zeelandnet.nl>Email</a> "
 
Customlinks = st
ThumbNailList1.MakeHtmlGallery Destination, BodyText, Customlinks

ThumbNailList1.Makegallery

End Sub

Private Sub MnuPrint_Click()
   ThumbNailList1.PrintPicture
End Sub
Private Sub MnuPrintPreview_Click()
   ThumbNailList1.PrintPreview
End Sub

Private Sub Thumbnail2_ThumbCreateError(ErrorMsg As String, Filename As Variant)
 MsgBox "Fout: " & ErrorMsg & vbCrLf & "Filename: " & Filename
End Sub

Private Sub ThumbNailList1_ThumbClick(Index As Long, Filename As String)
  'you can read the file information here about the original.
  MsgBox "ThumbNailList1_ThumbClick(" & Index & "," & Filename
  
End Sub
Private Sub ThumbNailList1_ThumbDblClick(Index As Long, Filename As String)
   'you can read the file information here about the original.
   MsgBox "ThumbNailList1_ThumbClick(" & Index & "," & Filename
   BigPicture.Pictureview1.Filename = Filename
   '' BigPicture.Show <<- My picture control with scrollbars
   '' not implemented yet
End Sub

Private Sub Printlog(Tekst)
   ' if something is going wrong, then report it here.
   Print #20, Tekst
End Sub

Private Sub ThumbNailList1_ThumbPageStatus(ThumbNailSize As Variant, ThumbsPerPage As Variant, PagesNeeded As Variant, nmfiles)
  st = st & "--------------------------------------------------------" & X & vbCrLf
  st = st & "ThumbNailSize = " & ThumbNailSize & vbCrLf
  st = st & "ThumbsPerPage = " & ThumbsPerPage & vbCrLf
  st = st & "PagesNeeded   = " & PagesNeeded & vbCrLf
  st = st & "Afbeeldingen  = " & nmfiles & vbCrLf
  Call Printlog(st)
End Sub

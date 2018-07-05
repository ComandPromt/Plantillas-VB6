VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ListBox Functions 2001 v2 - by source - www.vbfx.net"
   ClientHeight    =   9375
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command33 
      Caption         =   "E  X  I  T"
      Height          =   495
      Left            =   3000
      TabIndex        =   89
      Top             =   8760
      Width           =   1575
   End
   Begin VB.CommandButton Command32 
      Caption         =   "C O N T A C T"
      Height          =   495
      Left            =   1560
      TabIndex        =   88
      Top             =   8760
      Width           =   1335
   End
   Begin VB.CommandButton Command31 
      Caption         =   "A  B  O  U  T"
      Height          =   495
      Left            =   120
      TabIndex        =   87
      Top             =   8760
      Width           =   1335
   End
   Begin VB.Frame FRAMEsaveload 
      Caption         =   "saving and loading w and w/out com. dial."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3120
      TabIndex        =   9
      Top             =   0
      Width           =   2895
      Begin VB.ListBox List2 
         Height          =   1230
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "save w/ com. dialog"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         ToolTipText     =   "By using common dialog this will prompt you to save list2 located to the left of this button"
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "load w/ com. dialog"
         Height          =   255
         Left            =   960
         TabIndex        =   15
         ToolTipText     =   "By using common dialog this will prompt you to load a listbox, into the list located to the left of this button."
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Clear"
         Height          =   300
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "This clears list2, the list below"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "save with custom path"
         Height          =   255
         Left            =   960
         TabIndex        =   13
         ToolTipText     =   $"frmmain.frx":0000
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   "load with custom path"
         Height          =   300
         Left            =   960
         TabIndex        =   12
         ToolTipText     =   "by entering a file path to load a .lst file from you will load it into the list located on the left of this frame."
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1080
         TabIndex        =   11
         Text            =   "file path to save/load"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame FRAMEloadsysfonts 
      Caption         =   "loading system fonts"
      Height          =   3255
      Left            =   6360
      TabIndex        =   59
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command25 
         Caption         =   "Load System Fonts"
         Height          =   315
         Left            =   1080
         TabIndex        =   65
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Text            =   "Sample Text"
         Top             =   2400
         Width           =   3255
      End
      Begin VB.ListBox List14 
         Height          =   1815
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   61
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox List13 
         Height          =   1815
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   60
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Printer Fonts Here:"
         Height          =   255
         Left            =   1800
         TabIndex        =   63
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Screen Fonts Here:"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame FRAMEsortscore 
      Caption         =   "sort scores"
      Height          =   1695
      Left            =   2040
      TabIndex        =   84
      Top             =   2160
      Width           =   1695
      Begin VB.CommandButton Command30 
         Caption         =   "Sort"
         Height          =   255
         Left            =   480
         TabIndex        =   86
         Top             =   1320
         Width           =   735
      End
      Begin VB.ListBox List21 
         Height          =   1035
         Left            =   120
         TabIndex        =   85
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame FRAMEupordown 
      Caption         =   "up or down"
      Height          =   1695
      Left            =   5640
      TabIndex        =   80
      Top             =   5880
      Width           =   1215
      Begin VB.CommandButton Command29 
         Caption         =   "Down"
         Height          =   255
         Left            =   480
         TabIndex        =   83
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Up"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox List20 
         Height          =   1035
         ItemData        =   "frmmain.frx":00A9
         Left            =   120
         List            =   "frmmain.frx":00C8
         TabIndex        =   81
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FRAMElisttotext 
      Caption         =   "list to text"
      Height          =   1215
      Left            =   3840
      TabIndex        =   77
      Top             =   3360
      Width           =   1815
      Begin VB.TextBox Text9 
         Height          =   855
         Left            =   960
         TabIndex        =   79
         Top             =   240
         Width           =   735
      End
      Begin VB.ListBox List19 
         Height          =   840
         ItemData        =   "frmmain.frx":00E7
         Left            =   120
         List            =   "frmmain.frx":0106
         TabIndex        =   78
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Timer timesclicked 
      Interval        =   3
      Left            =   7440
      Top             =   5880
   End
   Begin VB.Frame FRAMEtimeclicked 
      Caption         =   "timesclicked"
      Height          =   975
      Left            =   5040
      TabIndex        =   72
      Top             =   2160
      Width           =   1095
      Begin VB.ListBox List18 
         Height          =   255
         Left            =   1680
         TabIndex        =   76
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1560
         TabIndex        =   75
         Text            =   "Text8"
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command27 
         Caption         =   "go"
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   600
         Width           =   615
      End
      Begin VB.ListBox List17 
         Height          =   255
         ItemData        =   "frmmain.frx":0136
         Left            =   240
         List            =   "frmmain.frx":013D
         TabIndex        =   73
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame FRAMEallchars 
      Caption         =   "show all characters"
      Height          =   1215
      Left            =   4680
      TabIndex        =   69
      Top             =   4680
      Width           =   2175
      Begin VB.CommandButton Command26 
         Caption         =   "Add All Characters 1 - 255"
         Height          =   855
         Left            =   1080
         TabIndex        =   71
         Top             =   240
         Width           =   975
      End
      Begin VB.ListBox List16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FRAMEselectedintextbox 
      Caption         =   "selected goes to textbox"
      Height          =   1335
      Left            =   2880
      TabIndex        =   66
      Top             =   6120
      Width           =   2535
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   120
         TabIndex        =   68
         Text            =   "Text7"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ListBox List15 
         Height          =   645
         ItemData        =   "frmmain.frx":0144
         Left            =   120
         List            =   "frmmain.frx":0154
         TabIndex        =   67
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame FRAMEmisc 
      Caption         =   "Misc"
      Height          =   1335
      Left            =   6960
      TabIndex        =   55
      Top             =   6360
      Width           =   2415
      Begin VB.CommandButton Command17 
         Caption         =   "Clear All Lists"
         Height          =   300
         Left            =   120
         TabIndex        =   58
         ToolTipText     =   "clears all listbox's extremely fast"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Save everything you see."
         Height          =   300
         Left            =   120
         TabIndex        =   57
         ToolTipText     =   "saves everything the way it is using a sub from the module"
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Load previous saved saved"
         Height          =   285
         Left            =   120
         TabIndex        =   56
         ToolTipText     =   "loads the previous saved form state by using a sub from the module"
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Frame FRAMEmovelistbox 
      Caption         =   "listbox move"
      Height          =   1695
      Left            =   4680
      TabIndex        =   51
      Top             =   7560
      Width           =   2055
      Begin VB.ListBox List11 
         Height          =   840
         ItemData        =   "frmmain.frx":0164
         Left            =   120
         List            =   "frmmain.frx":0174
         TabIndex        =   54
         Top             =   240
         Width           =   855
      End
      Begin VB.ListBox List12 
         Height          =   840
         ItemData        =   "frmmain.frx":0193
         Left            =   1080
         List            =   "frmmain.frx":0195
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command24 
         Caption         =   "move listbox11 to 12"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1320
         Width           =   1815
      End
   End
   Begin VB.Frame FRAMEdraganddrop 
      Caption         =   "drag and drop / listbox search"
      Height          =   4455
      Left            =   120
      TabIndex        =   42
      Top             =   4200
      Width           =   2655
      Begin VB.ListBox List5 
         Height          =   645
         ItemData        =   "frmmain.frx":0197
         Left            =   120
         List            =   "frmmain.frx":0199
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         TabIndex        =   48
         ToolTipText     =   "click and drag a single item from here, to the bottom list. or multiselect from the top list down into the bottom list."
         Top             =   240
         Width           =   2415
      End
      Begin VB.ListBox List6 
         Height          =   645
         ItemData        =   "frmmain.frx":019B
         Left            =   120
         List            =   "frmmain.frx":019D
         OLEDropMode     =   1  'Manual
         TabIndex        =   47
         ToolTipText     =   "double clicking an item in this list will remove it."
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   120
         TabIndex        =   46
         Text            =   "search top list for a string to add to bottom"
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Frame Frame6 
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2640
         Width           =   2415
      End
      Begin VB.CommandButton Command16 
         Caption         =   "select"
         Height          =   300
         Left            =   720
         TabIndex        =   44
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton Command19 
         Caption         =   "clear bottom list"
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Top             =   1575
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   $"frmmain.frx":019F
         Height          =   975
         Left            =   120
         TabIndex        =   50
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "enter a string in the textbox, press select, it will take the string from list5 and add it into list6."
         Height          =   615
         Left            =   120
         TabIndex        =   49
         Top             =   3720
         Width           =   2415
      End
   End
   Begin VB.Frame FRAMEloadfromadd2 
      Caption         =   "load sub from add2"
      Height          =   1335
      Left            =   2880
      TabIndex        =   39
      Top             =   4680
      Width           =   1695
      Begin VB.ListBox List10 
         Height          =   645
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "loads from the sub in the module 'add2'"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Load From: add2"
         Height          =   285
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "loads from the sub in the module 'add2'"
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame FRAMEloadfromadd1 
      Caption         =   "load sub from add1"
      Height          =   1335
      Left            =   2880
      TabIndex        =   36
      Top             =   7440
      Width           =   1695
      Begin VB.ListBox List9 
         Height          =   645
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "loads from the sub in the module 'add1'"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Load From: add1"
         Height          =   285
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   "loads from the sub in the module 'add1'"
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame FRAMEclick4msgbox 
      Caption         =   "click event"
      Height          =   1215
      Left            =   3840
      TabIndex        =   34
      ToolTipText     =   "when you click something it matches it up with what you clicked and will give a msg with what you pressed"
      Top             =   2160
      Width           =   1095
      Begin VB.ListBox List8 
         Height          =   840
         ItemData        =   "frmmain.frx":023D
         Left            =   120
         List            =   "frmmain.frx":024D
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame FRAMEtwolistboxs 
      Caption         =   "two listbox's (save/load/select)"
      Height          =   2895
      Left            =   6960
      TabIndex        =   21
      Top             =   3360
      Width           =   3015
      Begin VB.ListBox List3 
         Height          =   645
         ItemData        =   "frmmain.frx":0262
         Left            =   240
         List            =   "frmmain.frx":0264
         TabIndex        =   30
         ToolTipText     =   "By clicking something in this list, it will selected the apposing item in the list to the right"
         Top             =   480
         Width           =   1095
      End
      Begin VB.ListBox List4 
         Height          =   645
         ItemData        =   "frmmain.frx":0266
         Left            =   1560
         List            =   "frmmain.frx":0268
         TabIndex        =   29
         ToolTipText     =   "By clicking an item in this list, it will select an opposing item in the list to the left"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   240
         TabIndex        =   28
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1560
         TabIndex        =   27
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "add both"
         Height          =   180
         Left            =   1080
         TabIndex        =   26
         ToolTipText     =   "by pressing this it add's the contents of text3(left) and text4(right) into list3(left) and list4(right)."
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "save both"
         Height          =   180
         Left            =   960
         TabIndex        =   25
         ToolTipText     =   "Using common dialog it saves two listbox's, the list 3 and list4 to one file"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command13 
         Caption         =   "load to both"
         Height          =   180
         Left            =   840
         TabIndex        =   24
         ToolTipText     =   "By using common dialog you can load one file into to listbox's (normally for use after you save two list's)"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton Command14 
         Caption         =   "clear ^"
         Height          =   285
         Left            =   240
         TabIndex        =   23
         ToolTipText     =   "clears the above list of all content"
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton Command15 
         Caption         =   "^ clear"
         Height          =   285
         Left            =   2160
         TabIndex        =   22
         ToolTipText     =   "clears the above list of all its content"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "listbox 3:"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "listbox 4:"
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "status"
         Height          =   255
         Left            =   0
         TabIndex        =   31
         ToolTipText     =   "shows whats going on with list item counts, updated by a click of anything inside the frame"
         Top             =   2640
         Width           =   2895
      End
   End
   Begin VB.Frame FRAMEhscroll 
      Caption         =   "horizontal scroll bar"
      Height          =   1935
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   1815
      Begin VB.ListBox List7 
         Height          =   1230
         ItemData        =   "frmmain.frx":026A
         Left            =   120
         List            =   "frmmain.frx":0289
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command18 
         Caption         =   "horizontal scroll bar"
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.Frame FRAMEbasic 
      Caption         =   "Basic And Misc Functions"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "Clear List"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         ToolTipText     =   "This will clear the contents of list1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove Selected"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         ToolTipText     =   "This is remove the selected item that is in list1"
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Copy List To Clipboard"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         ToolTipText     =   $"frmmain.frx":030F
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Kill Duplicates (doubles)"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         ToolTipText     =   $"frmmain.frx":039E
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   840
         TabIndex        =   3
         Text            =   "add to list"
         ToolTipText     =   "add any kind of text to the listbox, press add or press enter to have it add to the list"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "add"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "this will add the contents of the textbox to the left into the listbox 1"
         Top             =   1680
         Width           =   735
      End
      Begin VB.ListBox list1 
         Height          =   1425
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "lst count"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   240
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmmain.frx":0431
      Height          =   1095
      Left            =   6960
      TabIndex        =   90
      Top             =   7800
      Width           =   2895
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ListBox Functions 2001 version 2
'Programmed By Source
'Spaced and Commented By And (c) Source
'www.vbfx.net
'www.terrorfx.com/~source
'www.8op.com/vbfx
'contact
'itzdasource@aol.com
'source@terrorfx.com
'aim
'itzdasource
'hackertaLk
'if you use any of this...if you want you can give me credit
'if not, doesnt matter...ill know when i see my code on your
'app ;) -source
Option Explicit
Private iRet As Integer 'for move up move down forms
Private Sub Command1_Click()
On Error Resume Next 'if theres an error then resume

list1.Clear 'vb command to clear list1

Label1.Caption = list1.ListCount 'sets label1's caption as the listcount of list1
End Sub

Private Sub Command10_Click()
'calls LoadListBox from module, text2 is path to save, list2
'is whats going to be saved.

Call Loadlistbox(Text2, List2)
End Sub

Private Sub Command11_Click()
'add's text3's text to list3's list
List3.AddItem Text3

'adds text4's text to list4's list
List4.AddItem Text4

'sets labels caption to list3's listcount and list4's listcount
Label4.Caption = "list1:" & List3.ListCount & " list2:" & List4.ListCount & ""
End Sub

Private Sub Command12_Click()
'previous work with common dialog was commented(see save/load cmdialog)
    CmDialog1.DialogTitle = "[ListBox Functions 2001] Save Two Lists"
    
    CmDialog1.InitDir = App.Path
    
    CmDialog1.Flags = &H4
    
    CmDialog1.Filter = "text files (*.txt)|*.txt|list files (*.lst)|*.lst|all files (*.*)|*.*"
    
    CmDialog1.ShowSave
    
    'If FileExists(cmDialog1.FileName) = True Then
        'calls sub from module
        Call Save2Lists(List3, List4, CmDialog1.FileName)
    
    'End If
    'sets label's caption to list3's listcount and list4's listcount
    Label4.Caption = "list1:" & List3.ListCount & " list2:" & List4.ListCount & ""
End Sub

Private Sub Command13_Click()
'previous work with common dialog was commented(see save/load cmdialog)
    CmDialog1.DialogTitle = "[ListBox Functions 2001] Load Two Lists"
    
    CmDialog1.InitDir = App.Path
    
    CmDialog1.Flags = &H4
    
    CmDialog1.Filter = "text files (*.txt)|*.txt|list files (*.lst)|*.lst|all files (*.*)|*.*"
    
    CmDialog1.ShowOpen
    
    If FileExists(CmDialog1.FileName) = True Then
        'calls sub from module
        Call Load2Lists(List3, List4, CmDialog1.FileName)
    
    End If
    
    'sets label's caption
   Label4.Caption = "list1:" & List3.ListCount & " list2:" & List4.ListCount & ""

End Sub

Private Sub Command14_Click()
'clears list3
List3.Clear

'sets label4's caption to list3's listcount and list4's listcount
Label4.Caption = "list1:" & List3.ListCount & " list2:" & List4.ListCount & ""
End Sub

Private Sub Command15_Click()
'uses vb's commands it clears list4
List4.Clear

'sets label4's caption to list3's listcount, and list4's listcount
Label4.Caption = "list1:" & List3.ListCount & " list2:" & List4.ListCount & ""
End Sub

Private Sub Command16_Click()
'searches through list5 for text5's text, and adds to list6 as
'long as it isnt already in list6
Dim i As Long

For i = 0 To list1.ListCount - 1
            
            If InStr(LCase(List5.List(i)), LCase(Text5.Text)) > 0 Then
                
                If ListBoxCheckDup(List6, List5.List(i)) = False Then
                 
                 List6.AddItem List5.List(i)
                
                End If
            
            End If
        
        Next

End Sub


Private Sub Command17_Click()
Call ClearListBoxes(frmmain) 'calls sub from module, sets which form
End Sub


Private Sub Command18_Click()
'calls AddHScroll from module, sets it as list7 is where to
'add the horizontal scroll bar
Call AddHScroll(List7)
End Sub

Private Sub Command19_Click()
List6.Clear 'clears list6 my friend
End Sub

Private Sub Command2_Click()
Call xListRemoveSelected(list1) 'calls sub from the module, defines it as (list1)

Label1.Caption = list1.ListCount 'sets label1's caption as list1's listcount
End Sub

Private Sub Command22_Click()
'calls the sub add1, sets it for list9
Call add1(List9)
End Sub

Private Sub Command23_Click()
'more customizeable then the sub add1, this lets you select which
'list you want to add it to, and what to add to it

'calls add2 from the module, makes it list10, and WHATEVERHERE, is what
'will be added to list10.

Call add2(List10, "WHATEVERHERE")
End Sub

Private Sub Command24_Click()
Dim X 'dims the variable x

For X = 0 To List11.ListCount - 1 'preps listbox

List12.AddItem List11.List(X) 'adds all of listbox 11's content to
'listbox 12

Next X 'moves on to next till done

End Sub

Private Sub Command25_Click()
'dims variables l and f which where going to use in the following:
Dim l, f As Integer

'makes l =1 to -1(everything (screen fonts))
For l = 1 To Screen.FontCount - 1

'adds screen fonts
List13.AddItem Screen.Fonts(l)

'goes till done
Next l

'makes f = something, the printable ones
For f = 1 To Printer.FontCount - 1

'adds them all
List14.AddItem Printer.Fonts(f)

'moves till done
Next f

End Sub

Private Sub Command26_Click()
Dim a ' dims dat shizzle
'makes a = 0 to 255 (all characters)
For a = 0 To 255

'sets list16 as A and the next ones
List16.AddItem a & " = " & Chr(a)

'goes on to next till done
Next a
End Sub

Private Sub Command27_Click()
List18.AddItem "" 'adds blank item to list18

List17.Clear 'clears before it adds(take this away if you want then
'numbers down in a row.

List17.List(List17.ListIndex) = frmmain.Text8 + 1
'list17 defines as text8
End Sub

Private Sub Command28_Click()
    iRet = cmdMoveUp_Click(List20)
End Sub

Private Sub Command29_Click()
 iRet = cmdMoveDown_Click(List20)
End Sub

Private Sub Command3_Click()
On Error Resume Next 'if errors then resume

Dim lf As Long, TheList As String 'dims and sets strings

For lf = 0 To list1.ListCount - 1 'if lf is 0 to list1 to -1

If lf = 0 Then ' if its 0 then do the following
    
    TheList = list1.List(lf) 'the list(string) sets it as lf

Else 'if not that then....
    
    TheList = TheList & "," & list1.List(lf) 'resets thelist for list1(lf)

End If 'ends the if statement

Next 'next(after that)

Clipboard.Clear 'clears the clipboard of old content

Clipboard.SetText TheList 'sets if for list1's content

Label1.Caption = list1.ListCount 'label1 becomes list1's listcount

End Sub

Private Sub Command30_Click()
'this was created by patorjk(www.patorjk.com)
Dim i As Integer, i2 As Integer, Hold$
Dim ValPer1%, ValPer2%

For i = 0 To List21.ListCount - 1
    
    For i2 = 0 To List21.ListCount - 1
        ' If we're not looking at the same
        ' two list items then...
        
        If i <> i2 Then
            
            ' get the score of the first item
            
            ValPer1 = Val(Mid$(List21.List(i), InStr(List21.List(i), " -") + 3))
            
            ' get the score of the second item
           
            ValPer2 = Val(Mid$(List21.List(i2), InStr(List21.List(i2), " -") + 3))
            
            ' If the second score is bigger than
            
            ' the first score...
            
            If ValPer1 > ValPer2 Then
                
                '...have them switch places
                
                Hold = List21.List(i)
                
                List21.List(i) = List21.List(i2)
                
                List21.List(i2) = Hold
            
            End If
        
        End If
    
    Next

Next
End Sub

Private Sub Command31_Click()
Call MsgBox("ListBox Functions 2001 v2 By Source : (c) 2001", vbOKOnly, "ListBox Functions 2001 v 2 by Source")
End Sub

Private Sub Command32_Click()
MsgBox "e-mail : itzdasource@aol.com    [or]   aim:hackertaLk"
End Sub

Private Sub Command33_Click()
Unload Me
End Sub

Private Sub Command4_Click()
'sets common dialogs title to the below text
    CmDialog1.DialogTitle = "[listbox functions 2001 v2] Save"
    
    'default/start directory is directory the program thats being run is in
    CmDialog1.InitDir = App.Path
    
    'sets flag to & h4
    CmDialog1.Flags = &H4
    
    'filters the allowed/shown file types
    CmDialog1.Filter = "list files (*.lst)|*.lst|all files or custom (*.*)|*.*"
    
    'shows save buttom
    CmDialog1.ShowSave
    
    'If FileExists(cmDialog1.FileName) = True Then
        
        'calls sub from module, sets name as the one choosen from cmdialog, what to save? list2
        Call xSaveList(CmDialog1.FileName, List2)
    
    'End IfOn Error Resume Next
End Sub

Private Sub Command5_Click()
Call xListKillDupes(list1) 'calls sub from module

Label1.Caption = list1.ListCount 'label1 becomes list1's listcount
End Sub

Private Sub Command6_Click()
list1.AddItem Text1 'adds text1 to list1

Text1.Text = "" 'makes text1 go away

Label1.Caption = list1.ListCount 'label1 becomes list1's listcount
End Sub

Private Sub Command7_Click()
'sets common dialogs title to textbelow
    CmDialog1.DialogTitle = "[ListBox Functions 2001] Load"
    
'directory thats being show is the one the apps directorys at
    CmDialog1.InitDir = App.Path
    
'sets flag to &h4
    CmDialog1.Flags = &H4
    
'filters/allows certain file types
    CmDialog1.Filter = "list files (*.lst)|*.lst|all files or custom(*.*)|*.*"
    
'show the save button, and open it
    CmDialog1.ShowOpen
    
'calls sub from module xLoadList, what to save? list2
    Call xLoadList(CmDialog1.FileName, List2)
End Sub

Private Sub Command8_Click()
On Error Resume Next 'if error's then resume

List2.Clear 'vb command to clear listbox 2
End Sub

Private Sub Command9_Click()
'calls SaveListBox from module, text2 is the directory, list2
'is what is going to be saved.
Call SaveListBox(Text2, List2)
End Sub

Private Sub Form_Load()
list1.AddItem "01"
list1.AddItem "02"
list1.AddItem "03"
list1.AddItem "04"
list1.AddItem "05"
list1.AddItem "06"
list1.AddItem "07"
list1.AddItem "08"
list1.AddItem "09"
list1.AddItem "10"
List2.AddItem "test"
List2.AddItem "test2"
List2.AddItem "test3"
List3.AddItem "sn"
List4.AddItem "pw"
List3.AddItem "sn2"
List4.AddItem "pw2"
List5.AddItem "drag n drop"
List5.AddItem "drag n drop 2"
List5.AddItem "drag n drop 3"
List6.AddItem "drag n droped"
List21.AddItem "Bob - 100"
List21.AddItem "Bill - 20"
List21.AddItem "Steve - 200"
List21.AddItem "Mac - 55"
List21.AddItem "Jim - 75"
List21.AddItem "Jill - 10"
List21.AddItem "Jacko - 500"
List21.AddItem "Mary - 100"
List21.AddItem "Ben - 250"
List21.AddItem "Joe - 150"
List21.AddItem "Fred - 10"
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub List13_Click()
'makes the text6 textbox, =the list13 textname
Text6.FontName = List13.Text
End Sub

Private Sub List15_Click()
Dim X 'dims that shizzle
X = List15.ListIndex 'makes x equal list15's index

Text7 = List15.List(X) 'makes text 7 be what you clicked
End Sub

Private Sub List19_Click()
Dim start 'dims that shizzle
Dim lstindex
start = Text9.SelStart

lstindex = Len(List19.List(List19.ListIndex))

Text9.SelText = List19.List(List19.ListIndex)

Text9.SetFocus

Text9.SelStart = start + lstindex

End Sub

Private Sub List20_Click()

    '-- If more than one item is selected then disable both buttons and then exit
    If List20.SelCount > 1 Then
        Command29.Enabled = False
        Command28.Enabled = False
        Exit Sub
    Else
        Command29.Enabled = True
        Command28.Enabled = True
    End If
    '-- If the first item in the list is selected then disable the move up button
    If List20.Selected(0) Then
        Command28.Enabled = False
    Else
        Command28.Enabled = True
    End If
    '-- If the first item in the list is selected then disable the move down button
    If List20.Selected(List20.ListCount - 1) Then
        Command29.Enabled = False
    Else
        Command29.Enabled = True
    End If
End Sub


Private Sub List8_Click()
'using if & end if statements you can determine what was clicked

'if what was clicked was 0 then...
If List8.List(List8.ListIndex) = "0" Then

'a message box will appear, with what you clicked(can be customizable)
MsgBox " you clicked 0 ", vbOKOnly


End If 'if not the first then move on to second

'if what was clicked was 1 then
If List8.List(List8.ListIndex) = "1" Then

'a message box will appear, with what you clicked(can be customizable)
MsgBox " you clicked 1 ", vbOKOnly


End If 'if not the first or the second then move to third

'if what was clicked was 2 then
If List8.List(List8.ListIndex) = "2" Then

'a message box will appear, with what you clicked(can be customizable)
MsgBox " you clicked 2 ", vbOKOnly


End If 'if not the first or second or third then move to last

'if what you clicked was whuzup then
If List8.List(List8.ListIndex) = "whuzup" Then

'a message box will appear, with what you clicked(can be customizable)
MsgBox " you clicked whuzup ", vbOKOnly


End If 'if it wasn't any of those, then nothing will happen
'unless you put something under here,,like :

'msgbox " what you clicked wasn't here ",vbokonly


End Sub

Private Sub mnuadd1_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = True
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnuadd2_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = True
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnuallcharacters_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = True
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnubasic_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = True
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnuclickevent_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = True
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnudragndrop_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = True
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnuhscroll_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = True
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnulistboxtotextbox_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = True
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnuloadingsystemfonts_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = True
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnumisc_Click()
frmmain.FRAMEmisc.Visible = True
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnumovelistboxs_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = True
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnusaveload_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = True
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnuselected_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = True
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnusortscores_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = True
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnutimesclicked_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = True
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnutwolistboxs_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = True
frmmain.FRAMEupordown.Visible = False
End Sub

Private Sub mnuupordown_Click()
frmmain.FRAMEmisc.Visible = False
frmmain.FRAMEclick4msgbox.Visible = False
frmmain.FRAMEbasic.Visible = False
frmmain.FRAMEallchars.Visible = False
frmmain.FRAMEdraganddrop.Visible = False
frmmain.FRAMEhscroll.Visible = False
frmmain.FRAMElisttotext.Visible = False
frmmain.FRAMEloadfromadd1.Visible = False
frmmain.FRAMEloadfromadd2.Visible = False
frmmain.FRAMEloadsysfonts.Visible = False
frmmain.FRAMEmovelistbox.Visible = False
frmmain.FRAMEsaveload.Visible = False
frmmain.FRAMEselectedintextbox.Visible = False
frmmain.FRAMEsortscore.Visible = False
frmmain.FRAMEtimeclicked.Visible = False
frmmain.FRAMEtwolistboxs.Visible = False
frmmain.FRAMEupordown.Visible = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'if whats pressed on the keyboard is Enter(13) then..

list1.AddItem Text1 'adds text1 to list1

KeyAscii = 0 'sets keyascii to 0 (so it doesnt beep)

Text1.Text = "" 'makes text1 go away

End If 'ends if statement

Label1.Caption = list1.ListCount 'label1 becomes list1's listcount

End Sub
Private Sub List3_Click()

On Error Resume Next

'if list3's index is a negitive one then exit sub
If List3.ListIndex = -1 Then Exit Sub

'line up what was clicked in each list
List4.Selected(List3.ListIndex) = True

End Sub


Private Sub List4_Click()

On Error Resume Next

'if list4's index is a negitive one the exit sub
If List4.ListIndex = -1 Then Exit Sub

'line up what was clicked in each list
List3.Selected(List4.ListIndex) = True

End Sub
Private Sub List5_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
'list5 is the start drag
    Dim i As Long, Temp As String
    
    For i = 0 To List5.ListCount - 1
        
        If List5.Selected(i) = True Then
            
            If Temp = "" Then
                
                Temp = i
            
            Else
                
                Temp = Temp & "|" & i
            
            End If
        
        End If
    
    Next
    
    Data.Clear
    
    Data.SetData Temp, vbCFText

End Sub

Private Sub List6_DblClick()
 
 Dim i As Long
    
    If List6.ListCount = 0 Then Exit Sub
    
    Do
        
        DoEvents
        
        If List6.Selected(i) = True Then
            
            List6.RemoveItem i
        
        Else
            
            i = i + 1
        
        End If
    
    Loop Until i >= List6.ListCount



End Sub



Private Sub List6_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim i As Integer
    
    If KeyCode = 46 Then
            
            If List6.ListCount = 0 Then Exit Sub
    
    Do
        
        DoEvents
        
        If List6.Selected(i) = True Then
            
            List6.RemoveItem i
        
        Else
            
            i = i + 1
        
        End If
    
    Loop Until i >= List6.ListCount


    
    End If



End Sub



Private Sub List6_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i As Long
    
    Dim Temp As String, IndexArray() As String
    
    Temp = Data.GetData(vbCFText)
    
    IndexArray() = Split(Temp, "|")
    
    For i = LBound(IndexArray) To UBound(IndexArray)
        
        If ListBoxCheckDup(List6, List5.List(IndexArray(CInt(i)))) = False Then
            
            List6.AddItem List5.List(IndexArray(CInt(i)))
        
        End If

Next

End Sub

Private Sub selall_Click()
 'for drag and drop
    
    If List5.ListCount = 0 Then Exit Sub
    
    Dim i As Integer
    
    For i = 0 To List5.ListCount - 1
        
        If ListBoxCheckDup(List6, List5.List(i)) = False Then
            
            List6.AddItem List5.List(i)
        
        End If

Next

End Sub
Private Sub Command20_Click()
SaveFormState Me 'calls sub from module
End Sub

Private Sub Command21_Click()
LoadFormState Me 'calls sub from module
End Sub

Private Sub timesclicked_Timer()
'for times clicked sub, sets text 8 as list18's listcount
Text8.Text = List18.ListCount
End Sub
Public Function cmdMoveUp_Click(lstMove As listbox) As Integer
 'not by source
    Dim strTemp1 As String   '-- hold the selected index data temporarily for move
    
    Dim iCnt    As Integer   '-- holds the index of the item to be moved
    
    iCnt = lstMove.ListIndex
    
    If iCnt > -1 Then
         
         strTemp1 = lstMove.List(iCnt)
        
        '-- Add the item selected to one position above the current position
        lstMove.AddItem strTemp1, (iCnt - 1)
        
        '-- remove it from the current position. Note the current position has changed because the add has moved everything down by 1
        lstMove.RemoveItem (iCnt + 1)
        
        '-- Reselect the item that was moved.
             lstMove.Selected(iCnt - 1) = True
    
    End If
End Function
Public Function cmdMoveDown_Click(lstMove As listbox) As Integer
    Dim strTemp1 As String    '-- hold the selected index data temporarily for move
    Dim iCnt    As Integer    '-- holds the index of the item to be moved
        
    '-- Assign the first index
    iCnt = lstMove.ListIndex
    
    If iCnt > -1 Then
         
         strTemp1 = lstMove.List(iCnt)
        
        '-- Add the item selected to below the current position
        lstMove.AddItem strTemp1, (iCnt + 2)
        
        lstMove.RemoveItem (iCnt)
        
        '-- Reselect the item that was moved.
        lstMove.Selected(iCnt + 1) = True
   End If

End Function

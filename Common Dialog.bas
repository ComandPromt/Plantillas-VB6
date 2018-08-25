Attribute VB_Name = "mdlCommonDialog"
'KPD-Team 1999
'URL: http://www.allapi.net
'E-Mail: KPDTeam@Allapi.com

Public Const LF_FACESIZE = 32
Public Const MAX_PATH = 260

'ShowOpen/ShowSave flags:
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 128

'ChooseColor flags:
Public Const CC_ANYCOLOR = &H100
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_RGBINIT = &H1
Public Const CC_SHOWHELP = &H8
Public Const CC_SOLIDCOLOR = &H80

'ChooseFont flags:
Public Const CF_ANSIONLY = &H400&
Public Const CF_APPLY = &H200&
Public Const CF_SCREENFONTS = &H1
Public Const CF_PRINTERFONTS = &H2
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_EFFECTS = &H100&
Public Const CF_DSPTEXT = &H81
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_GDIOBJFIRST = &H300
Public Const CF_GDIOBJLAST = &H3FF
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_NOSIZESEL = &H200000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOVERTFONTS = &H1000000
Public Const CF_OWNERDISPLAY = &H80
Public Const CF_PRIVATEFIRST = &H200
Public Const CF_PRIVATELAST = &H2FF
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_SHOWHELP = &H4&
Public Const CF_TTONLY = &H40000
Public Const CF_USESTYLE = &H80&
Public Const CF_WYSIWYG = &H8000

'PageSetupDlg flags
Public Const PSD_DEFAULTMINMARGINS = &H0
Public Const PSD_DISABLEMARGINS = &H10
Public Const PSD_DISABLEORIENTATION = &H100
Public Const PSD_DISABLEPAGEPAINTING = &H80000
Public Const PSD_DISABLEPAPER = &H200
Public Const PSD_DISABLEPRINTER = &H20
Public Const PSD_ENABLEPAGEPAINTHOOK = &H40000
Public Const PSD_ENABLEPAGESETUPHOOK = &H2000
Public Const PSD_ENABLEPAGESETUPTEMPLATE = &H8000
Public Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE = &H20000
Public Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8
Public Const PSD_INTHOUSANDTHSOFINCHES = &H4
Public Const PSD_INWININIINTLMEASURE = &H0
Public Const PSD_MARGINS = &H2
Public Const PSD_MINMARGINS = &H1
Public Const PSD_NOWARNING = &H80
Public Const PSD_RETURNDEFAULT = &H400
Public Const PSD_SHOWHELP = &H800

'PrintDlg flags:
Public Const PD_ALLPAGES = &H0
Public Const PD_COLLATE = &H10
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000
Public Const PD_NOPAGENUMS = &H8
Public Const PD_NOSELECTION = &H4
Public Const PD_NOWARNING = &H80
Public Const PD_PAGENUMS = &H2
Public Const PD_PRINTSETUP = &H40
Public Const PD_PRINTTOFILE = &H20
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_RETURNIC = &H200
Public Const PD_SELECTION = &H1
Public Const PD_SHOWHELP = &H800
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000

'BrowseForFolder flags:
Public Const BIF_RETURNONLYFSDIRS = &H1       ' For finding a folder to start document searching
Public Const BIF_DONTGOBELOWDOMAIN = &H2      ' For starting the Find Computer
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000   ' Browsing for Computers.
Public Const BIF_BROWSEFORPRINTER = &H2000    ' Browsing for Printers
Public Const BIF_BROWSEINCLUDEFILES = &H4000  ' Browsing for Everything

'Error constants
Public Const CDERR_DIALOGFAILURE = &HFFFF
Public Const CDERR_FINDRESFAILURE = &H6
Public Const CDERR_GENERALCODES = &H0
Public Const CDERR_INITIALIZATION = &H2
Public Const CDERR_LOADRESFAILURE = &H7
Public Const CDERR_LOADSTRFAILURE = &H5
Public Const CDERR_LOCKRESFAILURE = &H8
Public Const CDERR_MEMALLOCFAILURE = &H9
Public Const CDERR_MEMLOCKFAILURE = &HA
Public Const CDERR_NOHINSTANCE = &H4
Public Const CDERR_NOHOOK = &HB
Public Const CDERR_REGISTERMSGFAIL = &HC
Public Const CDERR_NOTEMPLATE = &H3
Public Const CDERR_STRUCTSIZE = &H1

'ShowHelp Enum
Enum enumHelpState
     SW_HIDE = 0
     SW_NORMAL = 1
     SW_MAXIMIZE = 3
     SW_MINIMIZE = 6
     SW_SHOWDEFAULT = 10
End Enum
'Types
Type POINTAPI
    x As Long
    y As Long
End Type
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Type PRINTDLG
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Type PAGESETUPDLG
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type
Type CHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type
Type CHOOSEFONT
    lStructSize As Long
    hWndOwner As Long ' caller's window handle
    hdc As Long ' printer DC/IC or NULL
    lpLogFont As LOGFONT ' ptr. to a LOGFONT struct
    iPointSize As Long ' 10 * size in points of selected font
    flags As Long ' enum. type flags
    rgbColors As Long ' returned text color
    lCustData As Long ' data passed to hook fn.
    lpfnHook As Long ' ptr. to hook function
    lpTemplateName As String ' custom template name
    hInstance As Long ' instance handle of.EXE that
    ' contains cust. dlg. template
    lpszStyle As String ' return the style field here
    ' must be LF_FACESIZE or bigger
    nFontType As Integer ' same value reported to the EnumFonts
    ' call back with the extra FONTTYPE_
    ' bits added
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long ' minimum pt size allowed &
    nSizeMax As Long ' max pt size allowed if
    ' CF_LIMITSIZE is used
End Type
Type BROWSEINFO
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hmem As Long)
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function PRINTDLG Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG) As Long
Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Declare Function GetFileTitleAPI Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global OFName As OPENFILENAME, PD As PRINTDLG, PSD As PAGESETUPDLG, CC As CHOOSECOLOR
Global CFont As CHOOSEFONT, BInfo As BROWSEINFO
Dim CustomColors() As Byte
'Use the OR operator for multiple flags
'   e.g. nFlags = CC_FULLOPEN Or CC_SOLIDCOLOR
Public Function ShowColor(hWndOwner As Long, Optional nFlags As Long) As Long
    Dim Custcolor(16) As Long, lReturn As Long, i As Integer
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte

    For i = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(i) = 0
    Next i

    CC.lStructSize = Len(CC)
    CC.hWndOwner = hWndOwner
    CC.hInstance = App.hInstance
    CC.lpCustColors = StrConv(CustomColors, vbUnicode)
    CC.flags = nFlags

    If CHOOSECOLOR(CC) <> 0 Then
        ShowColor = CC.rgbResult
        CustomColors = StrConv(CC.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
End Function
'Use the vbNullChar character to seperate extensions in the filter
'   e.g.  sFilter = "Text Files (*.txt)" + vbNullChar + "*.txt" + "All Files (*.*)" + vbNullChar + "*.*"
'Use the OR operator for multiple flags
'   e.g. nFlags = OFN_EXPLORER Or OFN_FILEMUSTEXIST
Public Function ShowOpen(hWndOwner As Long, sFilter As String, sTitle As String, Optional nFlags As Long = OFN_EXPLORER) As String
    OFName.lStructSize = Len(OFName)
    OFName.hWndOwner = hWndOwner
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = sFilter
    OFName.lpstrFile = String(254, vbNullChar)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = String(254, vbNullChar)
    OFName.nMaxFileTitle = 255
    OFName.lpstrTitle = sTitle
    OFName.flags = nFlags

    If GetOpenFileName(OFName) Then
        ShowOpen = StripTerminator(OFName.lpstrFile)
    Else
        ShowOpen = ""
    End If
End Function
'Use the vbNullChar character to seperate extensions in the filter
'   e.g.  sFilter = "Text Files (*.txt)" + vbNullChar + "*.txt" + "All Files (*.*)" + vbNullChar + "*.*"
'Use the OR operator for multiple flags
'   e.g. nFlags = OFN_EXPLORER Or OFN_FILEMUSTEXIST
Public Function ShowSave(hWndOwner As Long, sFilter As String, sTitle As String, Optional nFlags As Long = OFN_EXPLORER) As String
    OFName.lStructSize = Len(OFName)
    OFName.hWndOwner = hWndOwner
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = sFilter
    OFName.lpstrFile = String(254, vbNullChar)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = String(254, vbNullChar)
    OFName.nMaxFileTitle = 255
    OFName.lpstrTitle = sTitle
    OFName.flags = nFlags

    If GetSaveFileName(OFName) Then
        ShowSave = StripTerminator(OFName.lpstrFile)
    Else
        ShowSave = ""
    End If
End Function
'Use the OR operator for multiple flags
'   e.g. nFlags = PD_PAGENUMS Or PD_NONETWORKBUTTON
Public Function ShowPrintDlg(hWndOwner As Long, Optional nFlags As Long) As Long
    PD.lStructSize = Len(PD)
    PD.hWndOwner = hWndOwner
    PD.hInstance = App.hInstance
    PD.flags = nFlags

    If PRINTDLG(PD) Then
        ShowPrintDlg = 0
    Else
        ShowPrintDlg = -1
    End If
End Function
'Use the OR operator for multiple flags
'   e.g. nFlags = PSD_DEFAULTMINMARGINS Or PSD_RETURNDEFAULT
Public Function ShowPageSetupDlg(hWndOwner As Long, Optional nFlags As Long) As Long
    PSD.lStructSize = Len(PSD)
    PSD.hWndOwner = hWndOwner
    PSD.hInstance = App.hInstance
    PSD.flags = nFlags

    If PAGESETUPDLG(PSD) Then
        ShowPageSetupDlg = 0
    Else
        ShowPageSetupDlg = -1
    End If
End Function
'Use the OR operator for multiple flags
'   e.g. nFlags = CF_BOTH Or CF_WYSIWYG
Public Function ShowFont(hWndOwner As Long, Optional nFlags As Long = CF_BOTH) As Long
    CFont.lStructSize = Len(CFont)
    CFont.hdc = Printer.hdc
    CFont.hInstance = App.hInstance
    CFont.hWndOwner = hWndOwner
    CFont.flags = nFlags
    If CHOOSEFONT(CFont) Then
        ShowFont = 0
    Else
        ShowFont = -1
    End If
End Function
'Use the OR operator for multiple flags
'   e.g. nFlags = BIF_BROWSEFORCOMPUTER Or BIF_BROWSEFORPRINTER
Public Function BrowseForFolder(hWndOwner As Long, sTitle As String) As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long

    With BInfo
        .hWndOwner = hWndOwner
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(BInfo)
    If lpIDList Then
        BrowseForFolder = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, BrowseForFolder
        CoTaskMemFree lpIDList
        BrowseForFolder = StripTerminator(BrowseForFolder)
    End If
End Function
Public Function GetFileTitle(sFile As String) As String
    GetFileTitle = String(255, vbNullChar)
    GetFileTitleAPI sFile, GetFileTitle, 255
    GetFileTitle = StripTerminator(GetFileTitle)
End Function
Public Function GetErrorString() As String
    Select Case CommDlgExtendedError
        Case CDERR_DIALOGFAILURE
            GetErrorString = "The dialog box could not be created."
        Case CDERR_FINDRESFAILURE
            GetErrorString = "The common dialog box function failed to find a specified resource."
        Case CDERR_INITIALIZATION
            GetErrorString = "The common dialog box function failed during initialization."
        Case CDERR_LOADRESFAILURE
            GetErrorString = "The common dialog box function failed to load a specified resource."
        Case CDERR_LOADSTRFAILURE
            GetErrorString = "The common dialog box function failed to load a specified string."
        Case CDERR_LOCKRESFAILURE
            GetErrorString = "The common dialog box function failed to lock a specified resource."
        Case CDERR_MEMALLOCFAILURE
            GetErrorString = "The common dialog box function was unable to allocate memory for internal structures."
        Case CDERR_MEMLOCKFAILURE
            GetErrorString = "The common dialog box function was unable to lock the memory associated with a handle."
        Case CDERR_NOHINSTANCE
            GetErrorString = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding instance handle."
        Case CDERR_NOHOOK
            GetErrorString = "The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a pointer to a corresponding hook procedure."
        Case CDERR_REGISTERMSGFAIL
            GetErrorString = "The RegisterWindowMessage function returned an error code when it was called by the common dialog box function."
        Case CDERR_NOTEMPLATE
            GetErrorString = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding template."
        Case CDERR_STRUCTSIZE
            GetErrorString = "The lStructSize member of the initialization structure for the corresponding common dialog box is invalid."
        Case Else
            GetErrorString = "Undefined error ..."
    End Select
End Function
Public Sub ShowHelp(hWndOwner As Long, sHelpFile As String, nShowState As enumHelpState)
    ShellExecute hWndOwner, vbNullString, sHelpFile, vbNullString, "C:\", nShowState
End Sub
Private Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    ZeroPos = InStr(1, sInput, vbNullChar)
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

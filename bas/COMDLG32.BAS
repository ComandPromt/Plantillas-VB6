Attribute VB_Name = "basComdlg32"
' Module Name               basComdlg32
' Module Description        Module which contains ALL Declarations
'                           for API calls in the Comdlg32 Library
'
Option Explicit

Declare Function ChooseColor _
    Lib "comdlg32.dll" _
    Alias "ChooseColorA" ( _
        pChoosecolor As ChooseColor) _
    As Long

Declare Function ChooseFont _
    Lib "comdlg32.dll" _
    Alias "ChooseFontA" ( _
        pChoosefont As ChooseFont) _
    As Long

Declare Function CommDlgExtendedError _
    Lib "comdlg32.dll" ( _
        ) _
    As Long

Declare Function FindText _
    Lib "comdlg32.dll" _
    Alias "FindTextA " ( _
        pFindreplace As FINDREPLACE) _
    As Long

Declare Function GetFileTitle _
    Lib "comdlg32.dll" _
    Alias "GetFileTitleA" ( _
        ByVal lpszFile As String, _
        ByVal lpszTitle As String, _
        ByVal cbBuf As Integer) _
    As Integer

Declare Function GetOpenFileName _
    Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" ( _
        pOpenfilename As OPENFILENAME) _
    As Long

Declare Function GetSaveFileName _
    Lib "comdlg32.dll" _
    Alias "GetSaveFileNameA" ( _
        pOpenfilename As OPENFILENAME) _
    As Long

Declare Function PageSetupDlg _
    Lib "comdlg32.dll" _
    Alias "PageSetupDlgA" ( _
        pPagesetupdlg As PageSetupDlg) _
    As Long

Declare Function PrintDlg _
    Lib "comdlg32.dll" _
    Alias "PrintDlgA" ( _
        pPrintdlg As PrintDlg) _
    As Long

Declare Function ReplaceText _
    Lib "comdlg32.dll" _
    Alias "ReplaceTextA" ( _
        pFindreplace As FINDREPLACE) _
    As Long



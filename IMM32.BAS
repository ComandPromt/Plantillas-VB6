Attribute VB_Name = "basIMM32"
' Module Name               basIMM32
' Module Description        Module which contains ALL Declarations
'                           for API calls in the basIMM32 Library
'
Option Explicit

Declare Function ImmAssociateContext _
    Lib "imm32.dll" ( _
        ByVal hWnd As Long, _
        ByVal himc As Long) _
    As Long

Declare Function ImmConfigureIME _
    Lib "imm32.dll" ( _
        ByVal HKL As Long, _
        ByVal hWnd As Long, _
        ByVal dw As Long) _
    As Long

Declare Function ImmCreateContext _
    Lib "imm32.dll" ( _
        ) _
    As Long

Declare Function ImmDestroyContext _
    Lib "imm32.dll" ( _
        ByVal himc As Long) _
    As Long

Declare Function ImmEnumRegisterWord _
    Lib "imm32.dll" _
    Alias "ImmEnumRegisterWordA" ( _
        ByVal HKL As Long, ByVal RegisterWordEnumProc As Long, _
        ByVal lpszReading As String, ByVal dw As Long, _
        ByVal lpszRegister As String, lpv As Any) _
    As Long

Declare Function ImmEscape _
    Lib "imm32.dll" _
    Alias "ImmEscapeA" ( _
        ByVal HKL As Long, ByVal himc As Long, _
        ByVal un As Long, lpv As Any) _
    As Long

Declare Function ImmGetCandidateList _
    Lib "imm32.dll" _
    Alias "ImmGetCandidateListA" ( _
        ByVal himc As Long, ByVal deIndex As Long, _
        lpCandidateList As CANDIDATELIST, ByVal dwBufLen As Long) _
    As Long

Declare Function ImmGetCandidateListCount _
    Lib "imm32.dll" _
    Alias "ImmGetCandidateListCountA" ( _
        ByVal himc As Long, _
        lpdwListCount As Long) _
    As Long

Declare Function ImmGetCandidateWindow _
    Lib "imm32.dll" ( _
        ByVal himc As Long, _
        ByVal dw As Long, _
        lpCandidateForm As CANDIDATEFORM) _
    As Long

Declare Function ImmGetCompositionFont _
    Lib "imm32.dll" _
    Alias "ImmGetCompositionFontA" ( _
        ByVal himc As Long, _
        lpLogFont As LOGFONT) _
    As Long

Declare Function ImmGetCompositionString _
    Lib "imm32.dll" _
    Alias "ImmGetCompositionStringA" ( _
        ByVal himc As Long, ByVal dw As Long, _
        lpv As Any, ByVal dw2 As Long) _
    As Long

Declare Function ImmGetCompositionWindow _
    Lib "imm32.dll" ( _
        ByVal himc As Long, _
        lpCompositionForm As COMPOSITIONFORM) _
    As Long

Declare Function ImmGetContext _
    Lib "imm32.dll" ( _
        ByVal hWnd As Long) _
    As Long

Declare Function ImmGetConversionList _
    Lib "imm32.dll" _
    Alias "ImmGetConversionListA" ( _
        ByVal HKL As Long, ByVal himc As Long, ByVal lpsz As String, _
        lpCandidateList As CANDIDATELIST, ByVal dwBufLen As Long, ByVal uFlag As Long) _
    As Long

Declare Function ImmGetConversionStatus _
    Lib "imm32.dll" ( _
        ByVal himc As Long, _
        lpdw As Long, _
        lpdw2 As Long) _
    As Long

Declare Function ImmGetDefaultIMEWnd _
    Lib "imm32.dll" ( _
        ByVal hWnd As Long) _
    As Long

Declare Function ImmGetDescription _
    Lib "imm32.dll" _
    Alias "ImmGetDescriptionA" ( _
        ByVal HKL As Long, _
        ByVal lpsz As String, _
        ByVal uBufLen As Long) _
    As Long

Declare Function ImmGetGuideLine _
    Lib "imm32.dll" _
    Alias " ImmGetGuideLineA" ( _
        ByVal himc As Long, ByVal dwIndex As Long, _
        ByVal lpStr As String, ByVal dwBufLen As Long) _
    As Long

Declare Function ImmGetIMEFileName _
    Lib "imm32.dll" _
    Alias "ImmGetIMEFileNameA" ( _
        ByVal HKL As Long, _
        ByVal lpStr As String, _
        ByVal uBufLen As Long) _
    As Long

Declare Function ImmGetOpenStatus _
    Lib "imm32.dll" ( _
        ByVal himc As Long) _
    As Long

Declare Function ImmGetProperty _
    Lib "imm32.dll" ( _
        ByVal HKL As Long, _
        ByVal dw As Long) _
    As Long

Declare Function ImmGetRegisterWordStyle _
    Lib "imm32.dll" _
    Alias " ImmGetRegisterWordStyleA" ( _
        ByVal HKL As Long, _
        ByVal nItem As Long, _
        lpStyleBuf As STYLEBUF) _
    As Long

Declare Function ImmGetStatusWindowPos _
    Lib "imm32.dll" ( _
        ByVal himc As Long, _
        lpPoint As POINTAPI) _
    As Long

Declare Function ImmGetVirtualKey _
    Lib "imm32.dll" ( _
        ByVal hWnd As Long) _
    As Long

Declare Function ImmInstallIME _
    Lib "imm32.dll" _
    Alias "ImmInstallIMEA" ( _
        ByVal lpszIMEFileName As String, _
        ByVal lpszLayoutText As String) _
    As Long

Declare Function ImmIsIME _
    Lib "imm32.dll" ( _
        ByVal HKL As Long) _
    As Long

Declare Function ImmIsUIMessage _
    Lib "imm32.dll" _
    Alias "ImmIsUIMessageA" ( _
        ByVal hWnd As Long, ByVal un As Long, _
        ByVal wParam As Long, ByVal lParam As Long) _
    As Long

Declare Function ImmNotifyIME _
    Lib "imm32.dll" ( _
        ByVal himc As Long, _
        ByVal dwAction As Long, _
        ByVal dwIndex As Long, _
        ByVal dwValue As Long) _
    As Long

Declare Function ImmRegisterWord _
    Lib "imm32.dll" _
    Alias "ImmRegisterWordA" ( _
        ByVal HKL As Long, ByVal lpszReading As String, _
        ByVal dw As Long, ByVal lpszRegister As String) _
    As Long

Declare Function ImmReleaseContext _
    Lib "imm32.dll" ( _
        ByVal hWnd As Long, _
        ByVal himc As Long) _
    As Long

Declare Function ImmSetCandidateWindow _
    Lib "imm32.dll" ( _
        ByVal himc As Long, _
        lpCandidateForm As CANDIDATEFORM) _
    As Long

Declare Function ImmSetCompositionFont _
    Lib "imm32.dll" _
    Alias "ImmSetCompositionFontA" ( _
        ByVal himc As Long, _
        lpLogFont As LOGFONT) _
    As Long

Declare Function ImmSetCompositionString _
    Lib "imm32.dll" _
    Alias "ImmSetCompositionStringA" ( _
        ByVal himc As Long, ByVal dwIndex As Long, lpComp As Any, _
        ByVal dw As Long, lpRead As Any, ByVal dw2 As Long) _
    As Long

Declare Function ImmSetCompositionWindow _
    Lib "imm32.dll" ( _
        ByVal himc As Long, _
        lpCompositionForm As COMPOSITIONFORM) _
    As Long

Declare Function ImmSetConversionStatus _
    Lib "imm32.dll" ( _
        ByVal himc As Long, _
        ByVal dw1 As Long, _
        ByVal dw2 As Long) _
    As Long

Declare Function ImmSetOpenStatus _
    Lib "imm32.dll" ( _
        ByVal himc As Long, _
        ByVal b As Long) _
    As Long

Declare Function ImmSetStatusWindowPos _
    Lib "imm32.dll" ( _
        ByVal himc As Long, _
        lpPoint As POINTAPI) _
    As Long

Declare Function ImmSimulateHotKey _
    Lib "imm32.dll" ( _
        ByVal hWnd As Long, _
        ByVal dw As Long) _
    As Long

Declare Function ImmUnregisterWord _
    Lib "imm32.dll" _
    Alias "ImmUnregisterWordA" ( _
        ByVal HKL As Long, ByVal lpszReading As String, _
        ByVal dw As Long, ByVal lpszUnregister As String) _
    As Long



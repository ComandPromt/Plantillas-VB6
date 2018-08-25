Attribute VB_Name = "basWinSpool"
' Module Name               basWinSpool
' Module Description        Module which contains ALL Declarations
'                           for API calls in the basWinSpool Library
'
Option Explicit


Declare Function AbortPrinter _
    Lib "winspool.drv" ( _
        ByVal hPrinter As Long) _
    As Long

Declare Function AddForm _
    Lib "winspool.drv" _
    Alias "AddFormA" ( _
        ByVal hPrinter As Long, _
        ByVal Level As Long, _
        pForm As Byte) _
    As Long

Declare Function AddJob _
    Lib "winspool.drv" _
    Alias "AddJobA" ( _
        ByVal hPrinter As Long, ByVal Level As Long, _
        pData As Byte, ByVal cdBuf As Long, pcbNeeded As Long) _
    As Long

Declare Function AddMonitor _
    Lib "winspool.drv" _
    Alias "AddMonitorA" ( _
        ByVal pName As String, _
        ByVal Level As Long, _
        pMonitors As Byte) _
    As Long

Declare Function AddPort _
    Lib "winspool.drv" _
    Alias "AddPortA" ( _
        ByVal pName As String, _
        ByVal hWnd As Long, _
        ByVal pMonitorName As String) _
    As Long

Declare Function AddPrinter _
    Lib "winspool.drv" _
    Alias "AddPrinterA" ( _
        ByVal pName As String, _
        ByVal Level As Long, _
        pPrinter As Byte) _
    As Long

Declare Function AddPrinterConnection _
    Lib "winspool.drv" _
    Alias "AddPrinterConnectionA" ( _
        ByVal pName As String) _
    As Long

Declare Function AddPrinterDriver _
    Lib "winspool.drv" _
    Alias "AddPrinterDriverA" ( _
        ByVal pName As String, _
        ByVal Level As Long, _
        pDriverInfor As Byte) _
    As Long

Declare Function AddPrintProcessor _
    Lib "winspool.drv" _
    Alias "AddPrintProcessorA" ( _
        ByVal pName As String, _
        ByVal pEnvironment As String, _
        ByVal pPathName As String, _
        ByVal pPrintProcessorName As String) _
    As Long

Declare Function AddPrintProvidor _
    Lib "winspool.drv" _
    Alias "AddPrintProvidorA" ( _
        ByVal pName As String, _
        ByVal Level As Long, _
        pProvidorInfo As Byte) _
    As Long

Declare Function AdvancedDocumentProperties _
    Lib "winspool.drv" _
    Alias "AdvancedDocumentPropertiesA" ( _
        ByVal hWnd As Long, _
        ByVal hPrinter As Long, _
        ByVal pDeviceName As String, _
        pDevModeOutput As DEVMODE, _
        pDevModeInput As DEVMODE) _
    As Long

Declare Function ClosePrinter _
    Lib "winspool.drv" ( _
        ByVal hPrinter As Long) _
    As Long

Declare Function ConfigurePort _
    Lib "winspool.drv" _
    Alias "ConfigurePortA" ( _
        ByVal pName As String, _
        ByVal hWnd As Long, _
        ByVal pPortName As String) _
    As Long

Declare Function ConnectToPrinterDlg _
    Lib "winspool.drv" ( _
        ByVal hWnd As Long, _
        ByVal flags As Long) _
    As Long

Declare Function DeleteForm _
    Lib "winspool.drv" _
    Alias "DeleteFormA" ( _
        ByVal hPrinter As Long, _
        ByVal pFormName As String) _
    As Long

Declare Function DeleteMonitor _
    Lib "winspool.drv" _
    Alias "DeleteMonitorA" ( _
        ByVal pName As String, _
        ByVal pEnvironment As String, _
        ByVal pMonitorName As String) _
    As Long

Declare Function DeletePort _
    Lib "winspool.drv" _
    Alias "DeletePortA" ( _
        ByVal pName As String, _
        ByVal hWnd As Long, _
        ByVal pPortName As String) _
    As Long

Declare Function DeletePrinter _
    Lib "winspool" ( _
        ByVal hPrinter As Long) _
    As Boolean

Declare Function DeletePrinterConnection _
    Lib "winspool.drv" _
    Alias "DeletePrinterConnectionA" ( _
        ByVal pName As String) _
    As Long

Declare Function DeletePrinterDriver _
    Lib "winspool.drv" _
    Alias "DeletePrinterDriverA" ( _
        ByVal pName As String, _
        ByVal pEnvironment As String, _
        ByVal pDriverName As String) _
    As Long

Declare Function DeletePrintProcessor _
    Lib "winspool.drv" _
    Alias "DeletePrintProcessorA" ( _
        ByVal pName As String, _
        ByVal pEnvironment As String, _
        ByVal pPrintProcessorName As String) _
    As Long

Declare Function DeletePrintProvidor _
    Lib "winspool.drv" _
    Alias "DeletePrintProvidorA" ( _
        ByVal pName As String, _
        ByVal pEnvironment As String, _
        ByVal pPrintProvidorName As String) _
    As Long

Declare Function DeviceCapabilities _
    Lib "winspool.drv" _
    Alias "DeviceCapabilitiesA" ( _
        ByVal lpDeviceName As String, _
        ByVal lpPort As String, ByVal iIndex As Long, _
        ByVal lpOutput As String, lpDevMode As DEVMODE) _
    As Long

Declare Function DocumentProperties _
    Lib "winspool.drv" _
    Alias "DocumentPropertiesA" ( _
        ByVal hWnd As Long, ByVal hPrinter As Long, _
        ByVal pDeviceName As String, _
        pDevModeOutput As DEVMODE, pDevModeInput As DEVMODE, _
        ByVal fMode As Long) _
    As Long

Declare Function EndDocPrinter _
    Lib "winspool.drv" ( _
        ByVal hPrinter As Long) _
    As Long

Declare Function EndPagePrinter _
    Lib "winspool.drv" ( _
        ByVal hPrinter As Long) _
    As Long

Declare Function EnumPorts _
    Lib "winspool.drv" _
    Alias "EnumPortsA" ( _
        ByVal pName As String, ByVal Level As Long, _
        ByVal lpbPorts As Long, ByVal cbBuf As Long, _
        pcbNeeded As Long, pcReturned As Long) _
    As Long

Declare Function EnumPrinterDrivers _
    Lib "winspool.drv" _
    Alias "EnumPrinterDriversA" ( _
        ByVal pName As String, ByVal pEnvironment As String, _
        ByVal Level As Long, pDriverInfo As Byte, _
        ByVal cdBuf As Long, pcbNeeded As Long, pcRetruned As Long) _
    As Long

Declare Function EnumPrinterPropertySheets _
    Lib "winspool.drv" ( _
        hPrinter As Long, _
        hWnd As Long, _
        lpfnAdd As Long, _
        ByVal lParam As Long) _
    As Long

Declare Function EnumPrinters _
    Lib "winspool.drv" _
    Alias "EnumPrintersA" ( _
        ByVal flags As Long, ByVal name As String, _
        ByVal Level As Long, pPrinterEnum As Byte, _
        ByVal cdBuf As Long, pcbNeeded As Long, _
        pcReturned As Long) As Long

Declare Function EnumPrintProcessorDatatypes _
    Lib "winspool.drv" _
    Alias "EnumPrintProcessorDatatypesA" ( _
        ByVal pName As String, ByVal pPrintProcessorName As String, _
        ByVal Level As Long, pDatatypes As Byte, _
        ByVal cdBuf As Long, pcbNeeded As Long, pcRetruned As Long) _
    As Long

Declare Function EnumPrintProcessors _
    Lib "winspool.drv" _
    Alias "EnumPrintProcessorsA" ( _
        ByVal pName As String, ByVal pEnvironment As String, _
        ByVal Level As Long, pPrintProcessorInfo As Byte, _
        ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) _
    As Long

Declare Function FindClosePrinterChangeNotification _
    Lib "winspool" ( _
        ByVal hChange As Long) _
    As Boolean

Declare Function FindFirstPrinterChangeNotification _
    Lib "winspool" ( _
        ByVal hPrinter As Long, _
        ByVal fdwFlags As Long, _
        ByVal fdwOptions As Long, _
        ByVal pPrinterNotifyOptions As String) _
    As Long

Declare Function FindNextPrinterChangeNotification _
    Lib "winspool" ( _
        ByVal hChange As Long, _
        pdwChange As Long, _
        ByVal pvReserved As String, _
        ByVal ppPrinterNotifyInfo As Long) _
    As Boolean

Declare Function GetForm _
    Lib "winspool.drv" _
    Alias "GetFormA" ( _
        ByVal hPrinter As Long, ByVal pFormName As String, _
        ByVal Level As Long, pForm As Byte, _
        ByVal cbBuf As Long, pcbNeeded As Long) _
    As Long

Declare Function GetJob _
    Lib "winspool.drv" _
    Alias "GetJobA" ( _
        ByVal hPrinter As Long, ByVal JobId As Long, _
        ByVal Level As Long, pJob As Byte, _
        ByVal cdBuf As Long, pcbNeeded As Long) _
    As Long

Declare Function GetPrinter _
    Lib "winspool" _
    Alias "GetPrinterA" ( _
        ByVal hPrinter As Long, ByVal Level As Long, _
        pPrinter As Byte, ByVal cbBuf As Long, _
        pcbNeeded As Long) _
    As Boolean

Declare Function GetPrinterData _
    Lib "winspool.drv" _
    Alias "GetPrinterDataA" ( _
        ByVal hPrinter As Long, ByVal pValueName As String, _
        pType As Long, pData As Byte, _
        ByVal nSize As Long, pcbNeeded As Long) _
    As Long

Declare Function GetPrinterDriver _
    Lib "winspool.drv" _
    Alias "GetPrinterDriverA" ( _
        ByVal hPrinter As Long, ByVal pEnvironment As String, _
        ByVal Level As Long, pDriverInfo As Byte, _
        ByVal cdBuf As Long, pcbNeeded As Long) _
    As Long

Declare Function GetPrinterDriverDirectory _
    Lib "winspool.drv" _
    Alias "GetPrinterDriverDirectoryA" ( _
        ByVal pName As String, ByVal pEnvironment As String, _
        ByVal Level As Long, pDriverDirectory As Byte, _
        ByVal cdBuf As Long, pcbNeeded As Long) _
    As Long

Declare Function GetPrintProcessorDirectory _
    Lib "winspool.drv" _
    Alias "GetPrintProcessorDirectoryA" ( _
        ByVal pName As String, ByVal pEnvironment As String, _
        ByVal Level As Long, ByVal pPrintProcessorInfo As String, _
        ByVal cdBuf As Long, pcbNeeded As Long) _
    As Long

Declare Function OpenPrinter _
    Lib "winspool.drv" _
    Alias "OpenPrinterA" ( _
        ByVal pPrinterName As String, _
        phPrinter As Long, _
        pDefault As PRINTER_DEFAULTS) _
    As Long

Declare Function PrinterMessageBox _
    Lib "winspool.drv" _
    Alias "PrinterMessageBoxA" ( _
        ByVal hPrinter As Long, ByVal error As Long, _
        ByVal hWnd As Long, ByVal pText As String, _
        ByVal pCaption As String, ByVal dwType As Long) _
    As Long

Declare Function PrinterProperties _
    Lib "winspool.drv" ( _
        ByVal hWnd As Long, _
        ByVal hPrinter As Long) _
    As Long

Declare Function ReadPrinter _
    Lib "winspool.drv" ( _
        ByVal hPrinter As Long, _
        pBuf As Any, _
        ByVal cdBuf As Long, _
        pNoBytesRead As Long) _
    As Long

Declare Function ResetPrinter _
    Lib "winspool.drv" _
    Alias "ResetPrinterA" ( _
        ByVal hPrinter As Long, _
        pDefault As PRINTER_DEFAULTS) _
    As Long

Declare Function SetPrinter _
    Lib "winspool" _
    Alias "SetPrinterA" ( _
        ByVal hPrinter As Long, _
        ByVal Level As Long, _
        pPrinter As Byte, _
        ByVal Command As Long) _
    As Boolean

Declare Function SetPrinterData _
    Lib "winspool.drv" _
    Alias "SetPrinterDataA" ( _
        ByVal hPrinter As Long, ByVal pValueName As String, _
        ByVal dwType As Long, pData As Byte, _
        ByVal cbData As Long) _
    As Long

Declare Function StartDocPrinter _
    Lib "winspool.drv" _
    Alias "StartDocPrinterA" ( _
        ByVal hPrinter As Long, _
        ByVal Level As Long, _
        pDocInfo As Byte) _
    As Long

Declare Function StartPagePrinter _
    Lib "winspool.drv" ( _
        ByVal hPrinter As Long) _
    As Long

Declare Function WaitForPrinterChange _
    Lib "winspool.drv" ( _
        ByVal hPrinter As Long, _
        ByVal flags As Long) _
    As Long



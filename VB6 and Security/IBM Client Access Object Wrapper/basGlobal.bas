Attribute VB_Name = "basGlobal"
Option Explicit

'// Error constants
    Public Const CA_ERR_BASE            As Long = (vbObjectError + 5448766)
    
    Public Const CA_NO_SERVER           As Long = (CA_ERR_BASE + 1)
    Public Const CA_NO_SYSTEM           As Long = (CA_ERR_BASE + 2)
    Public Const CA_NO_PROVIDER         As Long = (CA_ERR_BASE + 3)
    Public Const CA_NO_DATABASE         As Long = (CA_ERR_BASE + 4)

    Public Const CA_LOGON_FAIL          As Long = (CA_ERR_BASE + 5)
    
    

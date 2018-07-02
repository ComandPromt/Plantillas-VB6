Attribute VB_Name = "basAdvapi32"
' Module Name               basAdvapi32
' Module Description        Module which contains ALL Declarations
'                           for API calls in the Advapi32 Library
'
Option Explicit

Declare Function AbortSystemShutdown _
    Lib "advapi32.dll" _
    Alias "AbortSystemShutdownA" ( _
        ByVal lpMachineName As String) _
    As Long

Declare Function AccessCheck _
    Lib "advapi32.dll" ( _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        ByVal ClientToken As Long, ByVal DesiredAccess As Long, _
        GenericMapping As GENERIC_MAPPING, _
        PrivilegeSet As PRIVILEGE_SET, PrivilegeSetLength As Long, _
        GrantedAccess As Long, ByVal Status As Long) _
    As Long

Declare Function AccessCheckAndAuditAlarm _
    Lib "advapi32.dll" _
    Alias "AccessCheckAndAuditAlarmA" ( _
        ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, _
        ByVal ObjectName As String, SecurityDescriptor As SECURITY_DESCRIPTOR, _
        ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, _
        ByVal ObjectCreation As Long, GrantedAccess As Long, ByVal AccessStatus As Long, ByVal pfGenerateOnClose As Long) _
    As Long

Declare Function AddAccessAllowedAce _
    Lib "advapi32.dll" ( _
        pAcl As ACL, _
        ByVal dwAceRevision As Long, _
        ByVal AccessMask As Long, _
        pSid As Any) _
    As Long

Declare Function AddAccessDeniedAce _
    Lib "advapi32.dll" ( _
        pAcl As ACL, _
        ByVal dwAceRevision As Long, _
        ByVal AccessMask As Long, _
        pSid As Any) _
    As Long

Declare Function AddAce _
    Lib "advapi32.dll" ( _
        pAcl As ACL, _
        ByVal dwAceRevision As Long, _
        ByVal dwStartingAceIndex As Long, _
        pAceList As Any, _
        ByVal nAceListLength As Long) _
    As Long

Declare Function AddAuditAccessAce _
    Lib "advapi32.dll" ( _
        pAcl As ACL, _
        ByVal dwAceRevision As Long, ByVal dwAccessMask As Long, _
        pSid As Any, _
        ByVal bAuditSuccess As Long, ByVal bAuditFailure As Long) _
    As Long

Declare Function AdjustTokenGroups _
    Lib "advapi32.dll" ( _
        ByVal TokenHandle As Long, ByVal ResetToDefault As Long, _
        NewState As TOKEN_GROUPS, ByVal BufferLength As Long, _
        PreviousState As TOKEN_GROUPS, ReturnLength As Long) _
    As Long

Declare Function AdjustTokenPrivileges _
    Lib "advapi32.dll" ( _
        ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
        NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
        PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) _
    As Long

Declare Function AllocateAndInitializeSid _
    Lib "advapi32.dll" ( _
        pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, _
        ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, _
        ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, _
        ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, _
        ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) _
    As Long

Declare Function AllocateLocallyUniqueId _
    Lib "advapi32.dll" ( _
        Luid As LARGE_INTEGER) _
    As Long

Declare Function AreAllAccessesGranted _
    Lib "advapi32.dll" ( _
        ByVal GrantedAccess As Long, _
        ByVal DesiredAccess As Long) _
    As Long

Declare Function AreAnyAccessesGranted _
    Lib "advapi32.dll" ( _
        ByVal GrantedAccess As Long, _
        ByVal DesiredAccess As Long) _
    As Long

Declare Function BackupEventLog _
    Lib "advapi32.dll" _
    Alias "BackupEventLogA" ( _
        ByVal hEventLog As Long, _
        ByVal lpBackupFileName As String) _
    As Long

Declare Function ClearEventLog _
    Lib "advapi32.dll" _
    Alias "ClearEventLogA" ( _
        ByVal hEventLog As Long, _
        ByVal lpBackupFileName As String) _
    As Long

Declare Function CloseEventLog _
    Lib "advapi32.dll" ( _
        ByVal hEventLog As Long) _
    As Long

Declare Function CloseServiceHandle _
    Lib "advapi32.dll" ( _
        ByVal hSCObject As Long) _
    As Long

Declare Function ControlService _
    Lib "advapi32.dll" ( _
        ByVal hService As Long, _
        ByVal dwControl As Long, _
        lpServiceStatus As SERVICE_STATUS) _
    As Long

Declare Function CopySid _
    Lib "advapi32.dll" ( _
        ByVal nDestinationSidLength As Long, _
        pDestinationSid As Any, _
        pSourceSid As Any) _
    As Long

Declare Function CreatePrivateObjectSecurity _
    Lib "advapi32.dll" ( _
        ParentDescriptor As SECURITY_DESCRIPTOR, _
        CreatorDescriptor As SECURITY_DESCRIPTOR, _
        NewDescriptor As SECURITY_DESCRIPTOR, _
        ByVal IsDirectoryObject As Long, ByVal Token As Long, GenericMapping As GENERIC_MAPPING) _
    As Long

Declare Function CreateService _
    Lib "advapi32.dll" _
    Alias "CreateServiceA" ( _
        ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, _
        ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, _
        ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, _
        lpdwTagId As Long, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) _
    As Long

Declare Function DeleteAce _
    Lib "advapi32.dll" ( _
        pAcl As ACL, _
        ByVal dwAceIndex As Long) _
    As Long

Declare Function DeleteService _
    Lib "advapi32.dll" ( _
        ByVal hService As Long) _
    As Long

Declare Function DeregisterEventSource _
    Lib "advapi32.dll" ( _
        ByVal hEventLog As Long) _
    As Long

Declare Function DestroyPrivateObjectSecurity _
    Lib "advapi32.dll" ( _
        ObjectDescriptor As SECURITY_DESCRIPTOR) _
    As Long

Declare Function DuplicateToken _
    Lib "advapi32.dll" ( _
        ByVal ExistingTokenHandle As Long, _
        ImpersonationLevel As Integer, _
        DuplicateTokenHandle As Long) _
    As Long

Declare Function EnumDependentServices _
    Lib "advapi32.dll" _
    Alias "EnumDependentServicesA" ( _
        ByVal hService As Long, ByVal dwServiceState As Long, _
        lpServices As ENUM_SERVICE_STATUS, _
        ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long) _
    As Long

Declare Function EnumServicesStatus _
    Lib "advapi32.dll" _
    Alias "EnumServicesStatusA" ( _
        ByVal hSCManager As Long, ByVal dwServiceType As Long, _
        ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, _
        ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long, lpResumeHandle As Long) _
    As Long

Declare Function EqualPrefixSid _
    Lib "advapi32.dll" ( _
        pSid1 As Any, _
        pSid2 As Any) _
    As Long

Declare Function EqualSid _
    Lib "advapi32.dll" ( _
        pSid1 As Any, _
        pSid2 As Any) _
    As Long

Declare Function FindFirstFreeAce _
    Lib "advapi32.dll" ( _
        pAcl As ACL, _
        pAce As Long) _
    As Long

Declare Sub FreeSid _
    Lib "advapi32.dll" ( _
        pSid As Any)

Declare Function GetAce _
    Lib "advapi32.dll" ( _
        pAcl As ACL, _
        ByVal dwAceIndex As Long, _
        pAce As Any) _
    As Long

Declare Function GetAclInformation _
    Lib "advapi32.dll" ( _
        pAcl As ACL, _
        pAclInformation As Any, _
        ByVal nAclInformationLength As Long, _
        ByVal dwAclInformationClass As Integer) _
    As Long

Declare Function GetFileSecurity _
    Lib "advapi32.dll" _
    Alias "GetFileSecurityA" ( _
        ByVal lpFileName As String, ByVal RequestedInformation As Long, _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        ByVal nLength As Long, lpnLengthNeeded As Long) _
    As Long

Declare Function GetKernelObjectSecurity _
    Lib "advapi32.dll" ( _
        ByVal Handle As Long, ByVal RequestedInformation As Long, _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        ByVal nLength As Long, _
        lpnLengthNeeded As Long) _
    As Long

Declare Function GetLengthSid _
    Lib "advapi32.dll" ( _
        pSid As Any) _
    As Long

Declare Function GetOldestEventLogRecord _
    Lib "advapi32.dll" ( _
        ByVal hEventLog As Long, _
        OldestRecord As Long) _
    As Long

Declare Function GetPrivateObjectSecurity _
    Lib "advapi32.dll" ( _
        ObjectDescriptor As SECURITY_DESCRIPTOR, _
        ByVal SecurityInformation As Long, _
        ResultantDescriptor As SECURITY_DESCRIPTOR, _
        ByVal DescriptorLength As Long, ReturnLength As Long) _
    As Long

Declare Function GetSecurityDescriptorControl _
    Lib "advapi32.dll" ( _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        pControl As Integer, _
        lpdwRevision As Long) _
    As Long

Declare Function GetSecurityDescriptorDacl _
    Lib "advapi32.dll" ( _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        lpbDaclPresent As Long, _
        pDacl As ACL, _
        lpbDaclDefaulted As Long) _
    As Long

Declare Function GetSecurityDescriptorGroup _
    Lib "advapi32.dll" ( _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        pGroup As Any, _
        ByVal lpbGroupDefaulted As Long) _
    As Long

Declare Function GetSecurityDescriptorLength _
    Lib "advapi32.dll" ( _
        pSecurityDescriptor As SECURITY_DESCRIPTOR) _
    As Long

Declare Function GetSecurityDescriptorOwner _
    Lib "advapi32.dll" ( _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        pOwner As Any, _
        ByVal lpbOwnerDefaulted As Long) _
    As Long

Declare Function GetSecurityDescriptorSacl _
    Lib "advapi32.dll" ( _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        ByVal lpbSaclPresent As Long, _
        pSacl As ACL, _
        ByVal lpbSaclDefaulted As Long) _
    As Long

Declare Function GetServiceDisplayName _
    Lib "advapi32.dll" _
    Alias "GetServiceDisplayNameA" ( _
        ByVal hSCManager As Long, ByVal lpServiceName As String, _
        ByVal lpDisplayName As String, lpcchBuffer As Long) _
    As Long

Declare Function GetServiceKeyName _
    Lib "advapi32.dll" _
    Alias "GetServiceKeyNameA" ( _
        ByVal hSCManager As Long, _
        ByVal lpDisplayName As String, _
        ByVal lpServiceName As String, _
        lpcchBuffer As Long) _
    As Long

Declare Function GetSidIdentifierAuthority _
    Lib "advapi32.dll" ( _
        pSid As Any) _
    As SID_IDENTIFIER_AUTHORITY

Declare Function GetSidLengthRequired _
    Lib "advapi32.dll" ( _
        ByVal nSubAuthorityCount As Byte) _
    As Long

Declare Function GetSidSubAuthority _
    Lib "advapi32.dll" ( _
        pSid As Any, _
        ByVal nSubAuthority As Long) _
    As Long

Declare Function GetSidSubAuthorityCount _
    Lib "advapi32.dll" ( _
        pSid As Any) _
    As Byte

Declare Function GetTokenInformation _
    Lib "advapi32.dll" ( _
        ByVal TokenHandle As Long, TokenInformationClass As Integer, _
        TokenInformation As Any, ByVal TokenInformationLength As Long, _
        ReturnLength As Long) _
    As Long

Declare Function GetUserName _
    Lib "advapi32.dll" _
    Alias "GetUserNameA" ( _
        ByVal lpBuffer As String, _
        nSize As Long) _
    As Long

Declare Function ImpersonateNamedPipeClient _
    Lib "advapi32.dll" ( _
        ByVal hNamedPipe As Long) _
    As Long

Declare Function ImpersonateSelf _
    Lib "advapi32.dll" ( _
        ImpersonationLevel As Integer) _
    As Long

Declare Function InitializeAcl _
    Lib "advapi32.dll" ( _
        pAcl As ACL, _
        ByVal nAclLength As Long, _
        ByVal dwAclRevision As Long) _
    As Long

Declare Function InitializeSecurityDescriptor _
    Lib "advapi32.dll" ( _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        ByVal dwRevision As Long) _
    As Long

Declare Function InitializeSid _
    Lib "advapi32.dll" ( _
        Sid As Any, _
        pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, _
        ByVal nSubAuthorityCount As Byte) _
    As Long

Declare Function InitiateSystemShutdown _
    Lib "advapi32.dll" _
    Alias "InitiateSystemShutdownA" ( _
        ByVal lpMachineName As String, ByVal lpMessage As String, _
        ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, _
        ByVal bRebootAfterShutdown As Long) _
    As Long

Declare Function LockServiceDatabase _
    Lib "advapi32.dll" ( _
        ByVal hSCManager As Long) _
    As Long

Declare Function LookupAccountName _
    Lib "advapi32.dll" _
    Alias "LookupAccountNameA" ( _
        ByVal lpSystemName As String, ByVal lpAccountName As String, _
        Sid As Long, cbSid As Long, _
        ByVal ReferencedDomainName As String, _
        cbReferencedDomainName As Long, peUse As Integer) _
    As Long

Declare Function LookupAccountSid _
    Lib "advapi32.dll" _
    Alias "LookupAccountSidA" ( _
        ByVal lpSystemName As String, Sid As Any, _
        ByVal name As String, cbName As Long, _
        ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, _
        peUse As Integer) _
    As Long


Declare Function LookupPrivilegeDisplayName _
    Lib "advapi32.dll" _
    Alias "LookupPrivilegeDisplayNameA" ( _
        ByVal lpSystemName As String, ByVal lpName As String, _
        ByVal lpDisplayName As String, cbDisplayName As Long, _
        lpLanguageID As Long) _
    As Long

Declare Function LookupPrivilegeName _
    Lib "advapi32.dll" _
    Alias "LookupPrivilegeNameA" ( _
        ByVal lpSystemName As String, lpLuid As LARGE_INTEGER, _
        ByVal lpName As String, _
        cbName As Long) _
    As Long

Declare Function LookupPrivilegeValue _
    Lib "advapi32.dll" _
    Alias "LookupPrivilegeValueA" ( _
        ByVal lpSystemName As String, _
        ByVal lpName As String, _
        lpLuid As LARGE_INTEGER) _
    As Long

Declare Function NotifyBootConfigStatus _
    Lib "advapi32.dll" ( _
        ByVal BootAcceptable As Long) _
    As Long

Declare Function NotifyChangeEventLog _
    Lib "advapi32" ( _
        ByVal hEventLog As Long, _
        ByVal hEvent As Long) _
    As Boolean

Declare Function ObjectCloseAuditAlarm _
    Lib "advapi32.dll" _
    Alias "ObjectCloseAuditAlarmA" ( _
        ByVal SubsystemName As String, _
        HandleId As Any, _
        ByVal GenerateOnClose As Long) _
    As Long

Declare Function ObjectPrivilegeAuditAlarm _
    Lib "advapi32.dll" _
    Alias "ObjectPrivilegeAuditAlarmA" ( _
        ByVal SubsystemName As String, HandleId As Any, _
        ByVal ClientToken As Long, ByVal DesiredAccess As Long, _
        Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) _
    As Long

Declare Function OpenBackupEventLog _
    Lib "advapi32.dll" _
    Alias "OpenBackupEventLogA" ( _
        ByVal lpUNCServerName As String, _
        ByVal lpFileName As String) _
    As Long

Declare Function OpenEventLog _
    Lib "advapi32.dll" ( _
        ByVal lpUNCServerName As String, _
        ByVal lpSourceName As String) _
    As Long

Declare Function OpenProcessToken _
    Lib "advapi32.dll" ( _
        ByVal ProcessHandle As Long, _
        ByVal DesiredAccess As Long, _
        TokenHandle As Long) _
    As Long

Declare Function OpenSCManager _
    Lib "advapi32.dll" _
    Alias "OpenSCManagerA" ( _
        ByVal lpMachineName As String, _
        ByVal lpDatabaseName As String, _
        ByVal dwDesiredAccess As Long) _
    As Long

Declare Function OpenService _
    Lib "advapi32.dll" _
    Alias "OpenServiceA" ( _
        ByVal hSCManager As Long, _
        ByVal lpServiceName As String, _
        ByVal dwDesiredAccess As Long) _
    As Long

Declare Function OpenThreadToken _
    Lib "advapi32.dll" ( _
        ByVal ThreadHandle As Long, _
        ByVal DesiredAccess As Long, _
        ByVal OpenAsSelf As Long, _
        TokenHandle As Long) _
    As Long

Declare Function PrivilegeCheck _
    Lib "advapi32.dll" ( _
        ByVal ClientToken As Long, _
        RequiredPrivileges As PRIVILEGE_SET, _
        ByVal pfResult As Long) _
    As Long

Declare Function PrivilegedServiceAuditAlarm _
    Lib "advapi32.dll" _
    Alias "PrivilegedServiceAuditAlarmA" ( _
        ByVal SubsystemName As String, ByVal ServiceName As String, _
        ByVal ClientToken As Long, Privileges As PRIVILEGE_SET, _
        ByVal AccessGranted As Long) _
    As Long

Declare Function ReadEventLog _
    Lib "advapi32.dll" _
    Alias "ReadEventLogA" ( _
        ByVal hEventLog As Long, ByVal dwReadFlags As Long, _
        ByVal dwRecordOffset As Long, lpBuffer As EVENTLOGRECORD, _
        ByVal nNumberOfBytesToRead As Long, _
        pnBytesRead As Long, pnMinNumberOfBytesNeeded As Long) _
    As Long

Declare Function RegCloseKey _
    Lib "advapi32.dll" ( _
        ByVal hKey As Long) _
    As Long

Declare Function RegConnectRegistry _
    Lib "advapi32.dll" _
    Alias "RegConnectRegistryA" ( _
        ByVal lpMachineName As String, _
        ByVal hKey As Long, _
        phkResult As Long) _
    As Long

Declare Function RegCreateKey _
    Lib "advapi32.dll" _
    Alias "RegCreateKeyA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        phkResult As Long) _
    As Long

Declare Function RegCreateKeyEx _
    Lib "advapi32.dll" _
    Alias "RegCreateKeyExA" ( _
        ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, _
        ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        phkResult As Long, lpdwDisposition As Long) _
    As Long

Declare Function RegDeleteKey _
    Lib "advapi32.dll" _
    Alias "RegDeleteKeyA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String) _
    As Long

Declare Function RegDeleteValue _
    Lib "advapi32.dll" _
    Alias "RegDeleteValueA" ( _
        ByVal hKey As Long, _
        ByVal lpValueName As String) _
    As Long

Declare Function RegEnumKey _
    Lib "advapi32.dll" _
    Alias "RegEnumKeyA" ( _
        ByVal hKey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpName As String, _
        ByVal cbName As Long) _
    As Long

Declare Function RegEnumKeyEx _
    Lib "advapi32.dll" _
    Alias "RegEnumKeyExA" ( _
        ByVal hKey As Long, ByVal dwIndex As Long, _
        ByVal lpName As String, lpcbName As Long, _
        lpReserved As Long, ByVal lpClass As String, _
        lpcbClass As Long, lpftLastWriteTime As FILETIME) _
    As Long

Declare Function RegEnumValue _
    Lib "advapi32.dll" _
    Alias "RegEnumValueA" ( _
        ByVal hKey As Long, ByVal dwIndex As Long, _
        ByVal lpValueName As String, lpcbValueName As Long, _
        lpReserved As Long, lpType As Long, _
        lpData As Byte, lpcbData As Long) _
    As Long

Declare Function RegFlushKey _
    Lib "advapi32.dll" ( _
        ByVal hKey As Long) _
    As Long

Declare Function RegGetKeySecurity _
    Lib "advapi32.dll" ( _
        ByVal hKey As Long, _
        ByVal SecurityInformation As Long, _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        lpcbSecurityDescriptor As Long) _
    As Long

Declare Function RegisterEventSource _
    Lib "advapi32.dll" _
    Alias "RegisterEventSourceA" ( _
        ByVal lpUNCServerName As String, _
        ByVal lpSourceName As String) _
    As Long

Declare Function RegisterServiceCtrlHandler _
    Lib "advapi32.dll" _
    Alias "RegisterServiceCtrlHandlerA" ( _
        ByVal lpServiceName As String, _
        ByVal lpHandlerProc As Long) _
    As Long

Declare Function RegLoadKey _
    Lib "advapi32.dll" _
    Alias "RegLoadKeyA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal lpFile As String) _
    As Long

Declare Function RegNotifyChangeKeyValue _
    Lib "advapi32.dll" ( _
        ByVal hKey As Long, _
        ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, _
        ByVal hEvent As Long, ByVal fAsynchronus As Long) _
    As Long

Declare Function RegOpenKey _
    Lib "advapi32.dll" _
    Alias "RegOpenKeyA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        phkResult As Long) _
    As Long

Declare Function RegOpenKeyEx _
    Lib "advapi32.dll" _
    Alias "RegOpenKeyExA" ( _
        ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As Long) _
    As Long

Declare Function RegQueryInfoKey _
    Lib "advapi32.dll" _
    Alias "RegQueryInfoKeyA" ( _
        ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, _
        lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, _
        lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, _
        lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) _
    As Long

Declare Function RegQueryValue _
    Lib "advapi32.dll" _
    Alias "RegQueryValueA" ( _
        ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal lpValue As String, lpcbValue As Long) _
    As Long

Declare Function RegQueryValueEx _
    Lib "advapi32.dll" _
    Alias "RegQueryValueExA" ( _
        ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal lpReserved As Long, lpType As Long, lpData As Any, _
        lpcbData As Long) _
    As Long

Declare Function RegReplaceKey _
    Lib "advapi32.dll" _
    Alias "RegReplaceKeyA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal lpNewFile As String, _
        ByVal lpOldFile As String) _
    As Long

Declare Function RegRestoreKey _
    Lib "advapi32.dll" _
    Alias "RegRestoreKeyA" ( _
        ByVal hKey As Long, _
        ByVal lpFile As String, _
        ByVal dwFlags As Long) _
    As Long

Declare Function RegSaveKey _
    Lib "advapi32.dll" _
    Alias "RegSaveKeyA" ( _
        ByVal hKey As Long, _
        ByVal lpFile As String, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES) _
    As Long

Declare Function RegSetKeySecurity _
    Lib "advapi32.dll" ( _
        ByVal hKey As Long, _
        ByVal SecurityInformation As Long, _
        pSecurityDescriptor As SECURITY_DESCRIPTOR) _
    As Long

Declare Function RegSetValue _
    Lib "advapi32.dll" _
    Alias "RegSetValueA" ( _
        ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal dwType As Long, _
        ByVal lpData As String, _
        ByVal cbData As Long) _
    As Long

Declare Function RegSetValueEx _
    Lib "advapi32.dll" _
    Alias "RegSetValueExA" ( _
        ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, _
        lpData As Any, ByVal cbData As Long) _
    As Long

Declare Function RegUnLoadKey _
    Lib "advapi32.dll" _
    Alias "RegUnLoadKeyA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String) _
    As Long

Declare Function ReportEvent _
    Lib "advapi32.dll" _
    Alias "ReportEventA" ( _
        ByVal hEventLog As Long, ByVal wType As Long, ByVal wCategory As Long, _
        ByVal dwEventID As Long, lpUserSid As Any, ByVal wNumStrings As Long, _
        ByVal dwDataSize As Long, ByVal lpStrings As Long, lpRawData As Any) _
    As Long

Declare Function RevertToSelf _
    Lib "advapi32.dll" () _
    As Long

Declare Function SetAclInformation _
    Lib "advapi32.dll" ( _
        pAcl As ACL, _
        pAclInformation As Any, _
        ByVal nAclInformationLength As Long, _
        ByVal dwAclInformationClass As Integer) _
    As Long

Declare Function SetFileSecurity _
    Lib "advapi32.dll" _
    Alias "SetFileSecurityA" ( _
        ByVal lpFileName As String, _
        ByVal SecurityInformation As Long, _
        pSecurityDescriptor As SECURITY_DESCRIPTOR) _
    As Long

Declare Function SetPrivateObjectSecurity _
    Lib "advapi32.dll" ( _
        ByVal SecurityInformation As Long, _
        ModificationDescriptor As SECURITY_DESCRIPTOR, _
        ObjectsSecurityDescriptor As SECURITY_DESCRIPTOR, _
        GenericMapping As GENERIC_MAPPING, _
        ByVal Token As Long) _
    As Long

Declare Function SetSecurityDescriptorDacl _
    Lib "advapi32.dll" ( _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        ByVal bDaclPresent As Long, _
        pDacl As ACL, _
        ByVal bDaclDefaulted As Long) _
    As Long

Declare Function SetSecurityDescriptorGroup _
    Lib "advapi32.dll" ( _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        pGroup As Any, _
        ByVal bGroupDefaulted As Long) _
    As Long

Declare Function SetSecurityDescriptorOwner _
    Lib "advapi32.dll" ( _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        pOwner As Any, _
        ByVal bOwnerDefaulted As Long) _
    As Long

Declare Function SetSecurityDescriptorSacl _
    Lib "advapi32.dll" ( _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, _
        ByVal bSaclPresent As Long, _
        pSacl As ACL, _
        ByVal bSaclDefaulted As Long) _
    As Long

Declare Function SetServiceBits _
    Lib "advapi32" ( _
        ByVal hServiceStatus As Long, _
        ByVal dwServiceBits As Long, _
        ByVal bSetBitsOn As Boolean, _
        ByVal bUpdateImmediately As Boolean) _
    As Boolean

Declare Function SetServiceObjectSecurity _
    Lib "advapi32.dll" ( _
        ByVal hService As Long, _
        ByVal dwSecurityInformation As Long, _
        lpSecurityDescriptor As Any) _
    As Long

Declare Function SetServiceStatus _
    Lib "advapi32.dll" ( _
        ByVal hServiceStatus As Long, _
        lpServiceStatus As SERVICE_STATUS) _
    As Long

Declare Function SetThreadToken _
    Lib "advapi32" ( _
        Thread As Long, _
        ByVal Token As Long) _
    As Boolean

Declare Function SetTokenInformation _
    Lib "advapi32.dll" ( _
        ByVal TokenHandle As Long, _
        TokenInformationClass As Integer, _
        TokenInformation As Any, _
        ByVal TokenInformationLength As Long) _
    As Long

Declare Function StartService _
    Lib "advapi32.dll" _
    Alias "StartServiceA" ( _
        ByVal hService As Long, _
        ByVal dwNumServiceArgs As Long, _
        ByVal lpServiceArgVectors As Long) _
    As Long

Declare Function StartServiceCtrlDispatcher _
    Lib "advapi32.dll" _
    Alias "StartServiceCtrlDispatcherA" ( _
        lpServiceStartTable As SERVICE_TABLE_ENTRY) _
    As Long

Declare Function UnlockServiceDatabase _
    Lib "advapi32.dll" ( _
        ScLock As Any) _
    As Long


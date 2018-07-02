VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCALogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*********************************************************************'
'* Copyright = 5763-XA1 (C) Copyright IBM Corp 1994, 1995.           *'
'*   All rights reserved.                                            *'
'*   Licensed Material - Program property of IBM                     *'
'*   Refer to copyright instructions form number G120-2083.          *'
'*********************************************************************'
'*********************************************************************'
'*                                                                   *'
'* Module:                                                           *'
'*   CWB.BAS                                                         *'
'*                                                                   *'
'* Purpose:                                                          *'
'*   Common declarations for Client Access/400 APIs                  *'
'*                                                                   *'
'* Usage Notes:                                                      *'
'*                                                                   *'
'*********************************************************************'

'*********************************************************************'
'* Common definitions                                                *'
'*********************************************************************'

Private Const CWB_FALSE = 0
Private Const CWB_TRUE = 1

'*********************************************************************'
'* Client Access/400 return codes fall into the following            *'
'* categories:                                                       *'
'* - Global return codes corresponding to operating system errors    *'
'* - Global return codes unique to Client Access/400                 *'
'* - Specific return codes for each Client Access/400 function       *'
'*                                                                   *'
'* The global return codes are defined in this file.  The function   *'
'* specific return codes are defined in the function specific        *'
'* header files.                                                     *'
'*********************************************************************'

'*********************************************************************'
'* Global return codes corresponding to operating system errors      *'
'*********************************************************************'

Private Const CWB_OK = 0
Private Const CWB_INVALID_FUNCTION = 1
Private Const CWB_FILE_NOT_FOUND = 2
Private Const CWB_PATH_NOT_FOUND = 3
Private Const CWB_TOO_MANY_OPEN_FILES = 4
Private Const CWB_ACCESS_DENIED = 5
Private Const CWB_INVALID_HANDLE = 6
Private Const CWB_NOT_ENOUGH_MEMORY = 8
Private Const CWB_INVALID_DRIVE = 15
Private Const CWB_NO_MORE_FILES = 18
Private Const CWB_DRIVE_NOT_READY = 21
Private Const CWB_GENERAL_FAILURE = 31
Private Const CWB_SHARING_VIOLATION = 32
Private Const CWB_LOCK_VIOLATION = 33
Private Const CWB_END_OF_FILE = 38
Private Const CWB_NOT_SUPPORTED = 50
Private Const CWB_BAD_NETWORK_PATH = 53
Private Const CWB_NETWORK_BUSY = 54
Private Const CWB_DEVICE_NOT_EXIST = 55
Private Const CWB_UNEXPECTED_NETWORK_ERROR = 59
Private Const CWB_NETWORK_ACCESS_DENIED = 65
Private Const CWB_FILE_EXISTS = 80
Private Const CWB_ALREADY_ASSIGNED = 85
Private Const CWB_INVALID_PARAMETER = 87
Private Const CWB_NETWORK_WRITE_FAULT = 88
Private Const CWB_OPEN_FAILED = 110
Private Const CWB_BUFFER_OVERFLOW = 111
Private Const CWB_DISK_FULL = 112
Private Const CWB_PROTECTION_VIOLATION = 115
Private Const CWB_INVALID_LEVEL = 124
Private Const CWB_BUSY_DRIVE = 142
Private Const CWB_INVALID_FSD_NAME = 252
Private Const CWB_INVALID_PATH = 253

'*********************************************************************'
'* Global return codes unique to Client Access/400                   *'
'*********************************************************************'
Private Const CWB_START = 4000
Private Const CWB_LAST = 5999

Private Const CWB_USER_CANCELLED_COMMAND = CWB_START
Private Const CWB_CONFIG_ERROR = CWB_START + 1
Private Const CWB_LICENSE_ERROR = CWB_START + 2
Private Const CWB_PROD_OR_COMP_NOT_SET = CWB_START + 3
Private Const CWB_SECURITY_ERROR = CWB_START + 4
Private Const CWB_GLOBAL_CFG_FAILED = CWB_START + 5
Private Const CWB_PROD_RETRIEVE_FAILED = CWB_START + 6
Private Const CWB_COMP_RETRIEVE_FAILED = CWB_START + 7
Private Const CWB_COMP_CFG_FAILED = CWB_START + 8
Private Const CWB_COMP_FIX_LEVEL_UPDATE_FAILED = CWB_START + 9
Private Const CWB_INVALID_API_HANDLE = CWB_START + 10
Private Const CWB_INVALID_API_PARAMETER = CWB_START + 11
Private Const CWB_HOST_NOT_FOUND = CWB_START + 12
Private Const CWB_NOT_COMPATIBLE = CWB_START + 13
Private Const CWB_INVALID_POINTER = CWB_START + 14
Private Const CWB_SERVER_PROGRAM_NOT_FOUND = CWB_START + 15
Private Const CWB_API_ERROR = CWB_START + 16
Private Const CWB_CA_NOT_STARTED = CWB_START + 17
Private Const CWB_FILE_IO_ERROR = CWB_START + 18
Private Const CWB_COMMUNICATIONS_ERROR = CWB_START + 19
Private Const CWB_RUNTIME_CONSTRUCTOR_FAILED = CWB_START + 20
Private Const CWB_DIAGNOSTIC = CWB_START + 21
Private Const CWB_COMM_VERSION_ERROR = CWB_START + 22
Private Const CWB_NO_VIEWER = CWB_START + 23
Private Const CWB_MODULE_NOT_LOADABLE = CWB_START + 24
Private Const CWB_ALREADY_SETUP = CWB_START + 25
Private Const CWB_CANNOT_START_PROCESS = CWB_START + 26



'********************************************************************'
'*                                                                  *'
'* API:                                                             *'
'*   cwbSV_CreateErrHandle                                          *'
'*                                                                  *'
'* Purpose:                                                         *'
'*   This function creates an error message object and returns a    *'
'*   handle to it. This error handle can be passed to Client        *'
'*   Access/400 API's that support it. If an error occurs on one    *'
'*   of these API's, the error handle can be used to retrieve the   *'
'*   error messages text associated with the API error.             *'
'*                                                                  *'
'* Parameters:                                                      *'
'*   cwbSV_ErrHandle * errorHandle -input/output                    *'
'*       Pointer to a cwbSV_ErrHandle where the handle will be      *'
'*       returned.                                                  *'
'*                                                                  *'
'* Return Codes:                                                    *'
'*   The following list shows common return values.                 *'
'*                                                                  *'
'*   CWB_OK - Successful completion.                                *'
'*   CWB_INVALID_POINTER - NULL passed as handle address            *'
'*   CWB_NOT_ENOUGH_MEMORY - Insufficient memory to create handle.  *'
'*                                                                  *'
'* Usage Notes:                                                     *'
'*                                                                  *'
'********************************************************************'
Private Declare Function cwbSV_CreateErrHandle Lib "CWBSV" (ByRef errorHandle As Long) As Long

'/********************************************************************/
'/*                                                                  */
'/* API:                                                             */
'/*   cwbSY_CreateSecurityObj                                        */
'/*                                                                  */
'/* Purpose:                                                         */
'/*   Creates a security object.  When you are finished using the    */
'/*   object you should delete it using the cwbSY_DeleteSecurityObj  */
'/*   API.                                                           */
'/*                                                                  */
'/* Parameters:                                                      */
'/*   cwbSY_SecurityHandle securityHandle - output                   */
'/*        Pointer to a cwbSY_SecurityHandle.  If the create is      */
'/*        successful this will contain a handle to the newly        */
'/*        created security object.  This handle is required on the  */
'/*        other security APIs.                                      */
'/*                                                                  */
'/* Return Codes:                                                    */
'/*   The following list shows common return values.                 */
'/*                                                                  */
'/*   CWB_OK - Successful completion.                                */
'/*   CWB_INVALID_POINTER - A pointer supplied to the API is not     */
'/*                         valid.                                   */
'/*   CWB_NOT_ENOUGH_MEMORY - Not enough memory available to create  */
'/*                           object.                                */
'/*                                                                  */
'/* Usage Notes:                                                     */
'/*   None                                                           */
'/*                                                                  */
'/********************************************************************/
Private Declare Function cwbSY_CreateSecurityObj Lib "cwbsy.dll" (ByRef securityHandle As Long) As Long

'/********************************************************************/
'/*                                                                  */
'/* API:                                                             */
'/*   cwbSY_SetSys                                                   */
'/*                                                                  */
'/* Purpose:                                                         */
'/*   The API sets the name of the system in the specified security  */
'/*   object.  No validation of the system name is done.  The system */
'/*   name is checked when the user ID and password are verified     */
'/*   in the changePassword, verifyUserIDPwd, or the logon APIs.     */
'/*                                                                  */
'/* Parameters:                                                      */
'/*   cwbSY_SecurityHandle securityHandle - input                    */
'/*       A valid handle to a security object.  The handle           */
'/*       is obtained using cwbSY_CreateSecurityObj.                 */
'/*                                                                  */
'/*   const char * systemName - input                                */
'/*       A pointer to an ASCIIZ string containing the name of the   */
'/*       AS/400 system.                                             */
'/*                                                                  */
'/* Return Codes:                                                    */
'/*   The following list shows common return values.                 */
'/*                                                                  */
'/*   CWB_OK - Successful completion.                                */
'/*   CWB_INVALID_HANDLE - The security handle supplied is not       */
'/*                        valid.                                    */
'/*   CWB_INVALID_POINTER - A pointer supplied to the API is not     */
'/*                         valid.                                   */
'/*   CWB_INVALID_PARAMETER - One of the parameters specified is     */
'/*                           invalid.                               */
'/*   CWB_NOT_ENOUGH_MEMORY - Insufficient memory; may have failed   */
'/*                           to allocate temporary buffer.          */
'/*   CWB_NON_REPRESENTABLE_UNICODE_CHAR - One or more input UNICODE */
'/*                                        characters have no        */
'/*                                        representation in the     */
'/*                                        codepage being used.      */
'/*   CWB_API_ERROR - General API failure.                           */
'/*                                                                  */
'/* Usage Notes:                                                     */
'/*   None                                                           */
'/*                                                                  */
'/********************************************************************/
'Public Declare Function cwbSY_SetSys(
'                             cwbSY_SecurityHandle  securityHandle,
'                             const char           *systemName);
Private Declare Function cwbSY_SetSys Lib "cwbsy.dll" (ByVal securityHandle As Long, ByVal systemName As String) As Long

'/********************************************************************/
'/*                                                                  */
'/* API:                                                             */
'/*   cwbSY_LogonUser                                                */
'/*                                                                  */
'/* Purpose:                                                         */
'/*   The API validates the supplied user ID and password without    */
'/*   presenting a sign-on prompt. If they are valid, the            */
'/*   information is stored and subsequent sign-ons and requests for */
'/*   other functions are not prompted for.                          */
'/*                                                                  */
'/*   Note: This API differs from cwbSY_Logon() in that a sign-on    */
'/*         prompt will not be presented because the user ID and     */
'/*         password information is being passed in on the call.     */
'/*         Also, any message dialogs (e.g. password expired) will   */
'/*         not be displayed. The return code will be set to         */
'/*         indicate these type of conditions.                       */
'/*                                                                  */
'/*   This API can be used to ensure that subsequent API function    */
'/*   calls in your program that establish connections to an AS/400  */
'/*   system will not prompt for sign-on information.                */
'/*                                                                  */
'/* Parameters:                                                      */
'/*   cwbSY_SecurityHandle securityHandle - input                    */
'/*       A valid handle to a security object.  The handle           */
'/*       is obtained using cwbSY_CreateSecurityObj.                 */
'/*                                                                  */
'/*   const char * userID - input                                    */
'/*       A pointer to an ASCIIZ string which contains the user ID.  */
'/*       The maximum length is 10 bytes + 1 byte                    */
'/*       for the NULL terminator.                                   */
'/*                                                                  */
'/*   const char * password - input                                  */
'/*       A pointer to an ASCIIZ string which contains the current   */
'/*       password.  The maximum length is 10 bytes + 1 byte for the */
'/*       NULL terminator.                                           */
'/*                                                                  */
'/*   cwbSV_ErrHandle errorHandle - output                           */
'/*     Any returned messages will be written to this object.  It    */
'/*     is created with the cwbSV_CreateErrHandle API.  The          */
'/*     messages may be retrieved through the cwbSV_GetErrText API.  */
'/*     If the parameter is set to zero, no messages will be         */
'/*     retrieved.                                                   */
'/*                                                                  */
'/* Return Codes:                                                    */
'/*   The following list shows common return values.                 */
'/*                                                                  */
'/*   CWB_OK - Successful completion.                                */
'/*   CWB_INVALID_HANDLE - The security handle is not valid.         */
'/*   CWB_INVALID_POINTER - A pointer supplied to the API is not     */
'/*                         valid.                                   */
'/*   CWB_COMMUNICATIONS_ERROR - A communication error occurred      */
'/*                              when attempting to contact the      */
'/*                              system.                             */
'/*   CWBSY_SYSTEM_NOT_CONFIGURED - The system has not been          */
'/*                                 configured.                      */
'/*   CWBSY_UNKNOWN_USERID - The supplied user ID is not             */
'/*                          known to this system.                   */
'/*   CWBSY_WRONG_PASSWORD - Password is not correct.                */
'/*   CWBSY_PASSWORD_EXPIRED - Password has expired.                 */
'/*   CWBSY_LAST_INVALID_PASSWORD - The next invalid password        */
'/*                                 will disable the user profile.   */
'/*   CWBSY_USER_PROFILE_DISABLED - The user ID has been disabled.   */
'/*   CWBSY_INVALID_PASSWORD - One or more characters in the         */
'/*                            password is invalid or the password   */
'/*                            is too long.                          */
'/*   CWBSY_INVALID_USERID - One or more characters in the user ID   */
'/*                          is invalid or the user ID is too long.  */
'/*   CWB_NOT_ENOUGH_MEMORY - Insufficient memory; may have failed   */
'/*                           to allocate temporary buffer.          */
'/*   CWB_NON_REPRESENTABLE_UNICODE_CHAR - One or more input UNICODE */
'/*                                        characters have no        */
'/*                                        representation in the     */
'/*                                        codepage being used.      */
'/*   CWB_API_ERROR - General API failure.                           */
'/*                                                                  */
'/* Usage Notes:                                                     */
'/*   The system name must set prior to calling this API by using    */
'/*   the cwbSY_SetSys API.                                          */
'/*                                                                  */
'/********************************************************************/
'unsigned int CWB_ENTRY cwbSY_LogonUser(
'                             cwbSY_SecurityHandle  securityHandle,
'                             const char           *userID,
'                             const char           *password,
'                             cwbSV_ErrHandle       errorHandle);

Private Declare Function cwbSY_LogonUser Lib "cwbsy.dll" (ByVal securityHandle As Long, ByVal UserID As String, ByVal Password As String, ByVal errorHandle As Long) As Long

'/********************************************************************/
'/*                                                                  */
'/* API:                                                             */
'/*   cwbSY_DeleteSecurityObj                                        */
'/*                                                                  */
'/* Purpose:                                                         */
'/*   Deletes a security object.                                     */
'/*                                                                  */
'/* Parameters:                                                      */
'/*   cwbSY_SecurityHandle securityHandle - input                    */
'/*       A valid handle to a security object.  The handle           */
'/*       is obtained using cwbSY_CreateSecurityObj.                 */
'/*                                                                  */
'/* Return Codes:                                                    */
'/*   The following list shows common return values.                 */
'/*                                                                  */
'/*   CWB_OK - Successful completion.                                */
'/*   CWB_INVALID_HANDLE - The security handle supplied is not       */
'/*                        valid.                                    */
'/*                                                                  */
'/* Usage Notes:                                                     */
'/*   None                                                           */
'/*                                                                  */
'/********************************************************************/
'Public Declare Function cwbSY_DeleteSecurityObj(
'                             cwbSY_SecurityHandle SecurityHandle);
Private Declare Function cwbSY_DeleteSecurityObj Lib "cwbsy.dll" (ByVal securityHandle As Long) As Long

'********************************************************************'
'*                                                                  *'
'* API:                                                             *'
'*   cwbSV_DeleteErrHandle                                          *'
'*                                                                  *'
'* Purpose:                                                         *'
'*   This function deletes the error message object that is         *'
'*   identified by the handle provided.                             *'
'*                                                                  *'
'* Parameters:                                                      *'
'*   cwbSV_ErrHandle errorHandle - output                           *'
'*       Handle that was returned by a previous call to the         *'
'*       cwbSV_CreateErrHandle() function.                          *'
'*                                                                  *'
'* Return Codes:                                                    *'
'*   The following list shows common return values.                 *'
'*                                                                  *'
'*   CWB_OK - Successful completion.                                *'
'*   CWB_INVALID_HANDLE - Invalid handle.                           *'
'*                                                                  *'
'* Usage Notes:                                                     *'
'*   This call should be made when the handle is no longer needed.  *'
'*                                                                  *'
'********************************************************************'
Private Declare Function cwbSV_DeleteErrHandle Lib "CWBSV" (ByVal errorHandle As Long) As Long

'********************************************************************'
'*                                                                  *'
'* API:                                                             *'
'*   cwbSV_GetErrCount                                              *'
'*                                                                  *'
'* Purpose:                                                         *'
'*   Returns the number of messages associated with the error       *'
'*   handle provided.                                               *'
'*                                                                  *'
'* Parameters:                                                      *'
'*   cwbSV_ErrHandle errorHandle - input                            *'
'*       Handle that was returned by a previous call to the         *'
'*       cwbSV_CreateErrHandle() function.                          *'
'*                                                                  *'
'*   unsigned long * errorCount - input/output                      *'
'*       Pointer to variable that receives the number of messages   *'
'*       associated with this error handle. If zero is returned,    *'
'*       no errors are associated with the error handle.            *'
'*                                                                  *'
'* Return Codes:                                                    *'
'*   The following list shows common return values.                 *'
'*                                                                  *'
'*   CWB_OK - Successful completion.                                *'
'*   CWB_INVALID_POINTER - NULL passed on output parameter.         *'
'*   CWB_INVALID_HANDLE - Invalid handle.                           *'
'*                                                                  *'
'* Usage Notes:                                                     *'
'*                                                                  *'
'********************************************************************'
Private Declare Function cwbSV_GetErrCount Lib "CWBSV" (ByVal errorHandle As Long, errorCount As Long) As Long

'********************************************************************'
'*                                                                  *'
'* API:                                                             *'
'*   cwbSV_GetErrText                                               *'
'*                                                                  *'
'* Purpose:                                                         *'
'*   Returns the message text associated with the top-level         *'
'*   (e.g. most recent) error identified by the error handle        *'
'*   provided.                                                      *'
'*                                                                  *'
'* Parameters:                                                    *'
'*   cwbSV_ErrHandle errorHandle - input                            *'
'*       Handle that was returned by a previous call to the         *'
'*       cwbSV_CreateErrHandle() function.                          *'
'*                                                                  *'
'*   char * errorText - input/output                                *'
'*       Pointer to a buffer that will receive the error message    *'
'*       text stored in the error identified by the handle.         *'
'*                                                                  *'
'*   unsigned long errorTextLength - input                          *'
'*       Length of the receive buffer passed in. It should          *'
'*       include space for the terminating null character. If the   *'
'*       buffer is too small, the value will be truncated and       *'
'*       CWB_BUFFER_OVERFLOW and returnLength will be set.          *'
'*                                                                  *'
'*   unsigned long * returnLength - input/output                    *'
'*       Optional, may be NULL. A return address to store the       *'
'*       number of bytes needed to hold the output string if the    *'
'*       receive buffer is too small.                               *'
'*                                                                  *'
'* Return Codes:                                                    *'
'*   The following list shows common return values.                 *'
'*                                                                  *'
'*   CWB_OK - Successful completion.                                *'
'*   CWB_BUFFER_OVERFLOW - Output buffer too small, data truncated. *'
'*   CWB_INVALID_POINTER - NULL passed on output parameter.         *'
'*   CWB_INVALID_HANDLE - Invalid handle.                           *'
'*   CWBSV_NO_ERROR_MESSAGES - No error messages associated with    *'
'*                             error handle.                        *'
'* Usage Notes:                                                     *'
'*                                                                  *'
'********************************************************************'
Private Declare Function cwbSV_GetErrText Lib "CWBSV" (ByVal errorHandle As Long, ByRef errorText As Byte, ByVal errorTextLength As Long, returnLength As Long) As Long

Public Function LogonToMachine(strMachine As String, strUserName As String, strPassword As String) As Long
    
    On Error GoTo Fail

    Dim strErrMessage As String
    Dim rc As Long
    Dim strX As String

    rc = LogonUser(strMachine, strUserName, strPassword, strErrMessage)

    If (rc <> CWB_OK) Then
        strX = "An error occured connecting to  the system.  " _
                & "The information passed was: " & vbCrLf _
                & vbTab & "System Name: " & strMachine & vbCrLf _
                & vbTab & "User ID:     " & strUserName & vbCrLf _
                & vbTab & "Password:    " & strPassword & vbCrLf _
                & "The return code is " & Str(rc) & ".  " & vbCrLf _
                & "The associated error message (if any) is:" & vbCrLf _
                & vbTab & strErrMessage
        Err.Raise 48954, TypeName(Me), strX
    End If

    LogonToMachine = rc

    Exit Function
    
Fail:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Function LogonUser(System As String, UserID As String, Password As String, Optional errMsg) As Long
    
    On Error GoTo Fail

    '***********************************************************************
    '"Prime" the connection with a current user/pwd
    '
    'INPUT: System Name, UserID, Password
    'OUTPUT:
    '  rc   -  API rc
    ' optional errMsg - Err info from the Error Handle
    'DEPENDENCIES:
    '     Byte2String in ByteHelp.bas
    '     cwbsy.bas for API declare and constant defines
    '***********************************************************************
    
    Dim abSysName(8) As Byte
    Dim lSecHandle As Long                                 'Used to hold a handle to a security object.
    Dim lErrHandle As Long                                 'Used to hold a handle to an error object.

    Dim lErrCount As Long
    Dim ErrMsgs() As String                                'Holds all error messages
    Dim S As String
    Dim rc As Long
    Dim I As Long


    '//To access any error messages that are generated, you need to create
    '//an error object:
    rc = cwbSV_CreateErrHandle(lErrHandle)

    '//To use the security APIs you'll first need to create a security
    '//object to work with.
    rc = cwbSY_CreateSecurityObj(lSecHandle)

    '//Now you have a security object.  So, next you need to set the name
    '//of the system.
    rc = cwbSY_SetSys(lSecHandle, System)
    If rc = CWB_OK Then
        rc = cwbSY_LogonUser(lSecHandle, UserID, Password, lErrHandle)
    End If

    If rc <> CWB_OK Then
        'Retrieve Error message from ErrHandle
        If (GetErrorText(lErrHandle, lErrCount, ErrMsgs) = CWB_OK) Then
            If lErrCount > 0 Then
                S = ""
                For I = 0 To lErrCount - 1
                    S = S & ErrMsgs(I) & vbCrLf
                Next I
                errMsg = S
            End If
        End If
    End If

    '//Now we must perform some clean up work so we will delete the
    '//security object when you are finished.

    cwbSY_DeleteSecurityObj (lSecHandle)
    cwbSV_DeleteErrHandle (lErrHandle)
    LogonUser = rc

    Exit Function
    
Fail:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Function GetErrorText(lErrHandle As Long, lErrCount As Long, ErrMsgs() As String) As Long

    On Error GoTo Fail

    '***********************************************************************
    'Retrieve all error messages stored in the Err object
    '
    'INPUT:  lErrHandle - input:  handle to error object
    '        ErrMsgCount -output: ErrMsgCount
    '        ErrMsgs      output: Array of messages
    'OUTPUT:
    '  rc   -  API rc of cwbSV calls
    'DEPENDENCIES:
    '     Byte2String in ByteHelp.bas
    '     cwbsv.bas for API declare and constant defines
    '***********************************************************************
    
    Dim S As String
    Dim temp() As Byte
    Dim lErrMsgLen As Long
    Dim rc As Long                   'Used when retrieving an error message
    Dim I As Integer

'// Get the error message count
    rc = cwbSV_GetErrCount(lErrHandle, lErrCount)
    If rc = CWB_OK Then
        If lErrCount > 0 Then
            ReDim ErrMsgs(lErrCount - 1)
            For I = 0 To lErrCount - 1
'// Get the error msg length
                rc = cwbSV_GetErrText(lErrHandle, 0, 0, lErrMsgLen)
    
'// Size the buffer and get the message
                ReDim temp(lErrMsgLen) As Byte
                rc = cwbSV_GetErrText(lErrHandle, temp(0), lErrMsgLen, lErrMsgLen)
'// Call ViewByteArray(temp, "Error Message Text")
                S = Space$(lErrMsgLen + 1)
                Call Byte2String(temp, S)
                ErrMsgs(I) = S
            Next I
        End If
        GetErrorText = 0
    Else
        lErrCount = 0
        GetErrorText = rc
    End If
    
    Exit Function
    
Fail:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Function Byte2String(InByte() As Byte, OutString As String) As Boolean

    On Error GoTo Fail

    Dim I As Integer

    If (UBound(InByte) = 0) Or (Len(OutString) < (UBound(InByte) + 1)) Then
        Byte2String = False
        Exit Function
    End If

    OutString = StrConv(InByte(), vbUnicode)
    Byte2String = True

    Exit Function
    
Fail:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


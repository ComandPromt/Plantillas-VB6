Attribute VB_Name = "ODBCAPI"
'
' ODBC API declares
'
'

'
' Note: ODBC level 1 and 2 supported. ODBC 3 extensions are
' note included in these definitions.
'

'|========================================================================|
'| ODBC Public Core Definitions                                           |
'|========================================================================|
'
'  ODBC Constants/Types
'
'  generally useful constants
'
Public Const SQL_NTS = -3                  '  NTS = Null Terminated String
Public Const SQL_SQLSTATE_SIZE = 5         '  size of SQLSTATE
Public Const SQL_MAX_MESSAGE_LENGTH = 512  '  message buffer size
Public Const SQL_MAX_DSN_LENGTH = 32       '  maximum data source name size

'  RETCODEs
'
Public Const SQL_ERROR = -1
Public Const SQL_INVALID_HANDLE = -2
Public Const SQL_NEED_DATA = 99
Public Const SQL_NO_DATA_FOUND = 100
Public Const SQL_SUCCESS = 0
Public Const SQL_SUCCESS_WITH_INFO = 1
Public Const SQL_GENERAL_WARNING = "01000"

'
' ODBC 2.0 specific
'
Public Const SQL_CURSOR_TYPE = 6
Public Const SQL_CONCURRENCY = 7
Public Const SQL_KEYSET_SIZE = 8
Public Const SQL_ROWSET_SIZE = 9
Public Const SQL_SIMULATE_CURSOR = 10
Public Const SQL_RETRIEVE_DATA = 11
Public Const SQL_USE_BOOKMARKS = 12
Public Const SQL_GET_BOOKMARK = 13
Public Const SQL_ROW_NUMBER = 14
'Public Const SQL_STMT_OPT_MAX = SQL_ROW_NUMBER
Public Const SQL_CURSOR_FORWARD_ONLY = 0
Public Const SQL_CURSOR_KEYSET_DRIVEN = 1
Public Const SQL_CURSOR_DYNAMIC = 2
Public Const SQL_CURSOR_STATIC = 3
Public Const SQL_CURSOR_TYPE_DEFAULT = SQL_CURSOR_FORWARD_ONLY

'
'  SQLFreeStmt defines
'
Public Const SQL_CLOSE = 0
Public Const SQL_DROP = 1
Public Const SQL_UNBIND = 2
Public Const SQL_RESET_PARAMS = 3

'  SQLSetParam defines
'
Public Const SQL_C_DEFAULT = 99

'  SQLTransact defines
'
Public Const SQL_COMMIT = 0
Public Const SQL_ROLLBACK = 1

'  Standard SQL datatypes, using ANSI type numbering
'
Public Const SQL_CHAR = 1
Public Const SQL_NUMERIC = 2
Public Const SQL_DECIMAL = 3
Public Const SQL_INTEGER = 4
Public Const SQL_SMALLINT = 5
Public Const SQL_FLOAT = 6
Public Const SQL_REAL = 7
Public Const SQL_DOUBLE = 8
Public Const SQL_VARCHAR = 12

Public Const SQL_TYPE_MIN = 1
Public Const SQL_TYPE_NULL = 0
Public Const SQL_TYPE_MAX = 12

'  C datatype to SQL datatype mapping    SQL types
'
Public Const SQL_C_CHAR = SQL_CHAR         '  CHAR, VARCHAR, DECIMAL, NUMERIC
Public Const SQL_C_LONG = SQL_INTEGER      '  INTEGER
Public Const SQL_C_SHORT = SQL_SMALLINT    '  SMALLINT
Public Const SQL_C_FLOAT = SQL_REAL        '  REAL
Public Const SQL_C_DOUBLE = SQL_DOUBLE     '  FLOAT, DOUBLE

'  NULL status constants.  These are used in SQLColumns, SQLColAttributes,
'  SQLDescribeCol, and SQLSpecialColumns to describe the nullablity of a
'  column in a table.  SQL_NULLABLE_UNKNOWN can be returned only by
'  SQLDescribeCol or SQLColAttributes.  It is used when the DBMS's meta-data
'  does not contain this info.
'
Public Const SQL_NO_NULLS = 0
Public Const SQL_NULLABLE = 1
Public Const SQL_NULLABLE_UNKNOWN = 2

'  Special length values
'
Public Const SQL_NULL_DATA = -1
Public Const SQL_DATA_AT_EXEC = -2

'  SQLColAttributes defines
'
Public Const SQL_COLUMN_COUNT = 0
Public Const SQL_COLUMN_NAME = 1
Public Const SQL_COLUMN_TYPE = 2
Public Const SQL_COLUMN_LENGTH = 3
Public Const SQL_COLUMN_PRECISION = 4
Public Const SQL_COLUMN_SCALE = 5
Public Const SQL_COLUMN_DISPLAY_SIZE = 6
Public Const SQL_COLUMN_NULLABLE = 7
Public Const SQL_COLUMN_UNSIGNED = 8
Public Const SQL_COLUMN_MONEY = 9
Public Const SQL_COLUMN_UPDATABLE = 10
Public Const SQL_COLUMN_AUTO_INCREMENT = 11
Public Const SQL_COLUMN_CASE_SENSITIVE = 12
Public Const SQL_COLUMN_SEARCHABLE = 13
Public Const SQL_COLUMN_TYPE_NAME = 14

'  SQLColAttributes subdefines for SQL_COLUMN_UPDATABLE
'
Public Const SQL_ATTR_READONLY = 0
Public Const SQL_ATTR_WRITE = 1
Public Const SQL_ATTR_READWRITE_UNKNOWN = 2

'  SQLColAttributes subdefines for SQL_COLUMN_SEARCHABLE
'  These are also used by SQLGetInfo
'
Public Const SQL_UNSEARCHABLE = 0
Public Const SQL_LIKE_ONLY = 1
Public Const SQL_ALL_EXCEPT_LIKE = 2
Public Const SQL_SEARCHABLE = 3

'  SQLError defines
'
Public Const SQL_NULL_HENV = 0
Public Const SQL_NULL_HDBC = 0
Public Const SQL_NULL_HSTMT = 0

'
'|========================================================================|
'| ODBC Module Core Definitions                                           |
'|========================================================================|
'
'  ODBC Core API's Definitions
'
Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal env As Long, hdbc As Long) As Integer
Declare Function SQLAllocEnv Lib "odbc32.dll" (env As Long) As Integer
Declare Function SQLAllocStmt Lib "odbc32.dll" (ByVal hdbc As Long, hstmt As Long) As Integer
Declare Function SQLBindCol Lib "odbc32.dll" (ByVal hstmt As Long, ByVal icol As Integer, ByVal fCType As Integer, ByVal rgbValue As Any, ByVal cbValueMax As Long, pcbValue As Any) As Integer
Declare Function SQLCancel Lib "odbc32.dll" (ByVal hstmt As Long) As Integer
Declare Function SQLColAttributes Lib "odbc32.dll" (ByVal hstmt&, ByVal icol%, ByVal fDescType%, rgbDesc As Any, ByVal cbDescMax%, pcbDesc%, pfDesc&) As Integer
Declare Function SQLConnect Lib "odbc32.dll" (ByVal hdbc As Long, ByVal Server As String, ByVal serverlen As Integer, ByVal uid As String, ByVal uidlen As Integer, ByVal pwd As String, ByVal pwdlen As Integer) As Integer
Declare Function SQLDescribeCol Lib "odbc32.dll" (ByVal hstmt As Long, ByVal colnum As Integer, ByVal colname As String, ByVal Buflen As Integer, colnamelen As Integer, dtype As Integer, dl As Long, ds As Integer, n As Integer) As Integer
Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hdbc As Long) As Integer
Declare Function SQLError Lib "odbc32.dll" (ByVal env As Long, ByVal hdbc As Long, ByVal hstmt As Long, ByVal SQLState As String, NativeError As Long, ByVal Buffer As String, ByVal Buflen As Integer, Outlen As Integer) As Integer
Declare Function SQLExecDirect Lib "odbc32.dll" (ByVal hstmt As Long, ByVal sqlString As String, ByVal sqlstrlen As Long) As Integer
Declare Function SQLExecute Lib "odbc32.dll" (ByVal hstmt As Long) As Integer
Declare Function SQLFetch Lib "odbc32.dll" (ByVal hstmt As Long) As Integer
Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc As Long) As Integer
Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal env As Long) As Integer
Declare Function SQLFreeStmt Lib "odbc32.dll" (ByVal hstmt As Long, ByVal EndOption As Integer) As Integer
Declare Function SQLGetCursorName Lib "odbc32.dll" (ByVal hstmt As Long, ByVal szCursor As String, ByVal cbCursorMax As Integer, pcbCursor As Long) As Integer
Declare Function SQLNumResultCols Lib "odbc32.dll" (ByVal hstmt As Long, NumCols As Integer) As Integer
Declare Function SQLPrepare Lib "odbc32.dll" (ByVal hstmt As Long, ByVal szSqlStr As String, ByVal cbSqlStr As Long) As Integer
Declare Function SQLRowCount Lib "odbc32.dll" (ByVal hstmt As Long, pcrow As Long) As Integer
Declare Function SQLSetCursorName Lib "odbc32.dll" (ByVal hstmt As Long, ByVal szCursor As String, ByVal cbCursor As Integer) As Integer
Declare Function SQLSetParam Lib "odbc32.dll" (ByVal hstmt&, ByVal ipar%, ByVal fCType%, ByVal fSqlType%, ByVal cbColDef&, ByVal ibScale%, rgbValue As Any, pcbValue As Long) As Integer
Declare Function SQLTransact Lib "odbc32.dll" (ByVal henv As Long, ByVal hdbc As Long, ByVal fType As Integer) As Integer

'
'|========================================================================|
'| ODBC Public Extended Definitions                                       |
'|========================================================================|
'
'  Date/Time/Timestamp Structs
'
Type DATE_STRUCT
  year      As Integer
  month     As Integer
  day       As Integer
End Type

Type TIME_SRUCT
  hour      As Integer
  minute    As Integer
  second    As Integer
End Type

Type TIMESTAMP_STRUCT
  year      As Integer
  month     As Integer
  day       As Integer
  hour      As Integer
  minute    As Integer
  second    As Integer
  fraction  As Long
End Type

' Level 1 Definitions/Functions
' Additional return codes
'
Public Const SQL_STILL_EXECUTING = 2

' SQL extended datatypes
'
Public Const SQL_DATE = 9
Public Const SQL_TIME = 10
Public Const SQL_TIMESTAMP = 11
Public Const SQL_LONGVARCHAR = -1
Public Const SQL_BINARY = -2
Public Const SQL_VARBINARY = -3
Public Const SQL_LONGVARBINARY = -4
Public Const SQL_BIGINT = -5
Public Const SQL_TINYINT = -6
Public Const SQL_BIT = -7

' C datatype to SQL datatype mapping
'
Public Const SQL_C_DATE = SQL_DATE
Public Const SQL_C_TIME = SQL_TIME
Public Const SQL_C_TIMESTAMP = SQL_TIMESTAMP
Public Const SQL_C_BINARY = SQL_BINARY
Public Const SQL_C_BIT = SQL_BIT
Public Const SQL_C_TINYINT = SQL_TINYINT

Public Const SQL_ALL_TYPES = 0

' Access modes
'
Public Const SQL_MODE_READ_WRITE = 0
Public Const SQL_MODE_READ_ONLY = 1

' Options for SQLDriverConnect
'
Public Const SQL_DRIVER_NOPROMPT = 0
Public Const SQL_DRIVER_COMPLETE = 1
Public Const SQL_DRIVER_PROMPT = 2
Public Const SQL_DRIVER_COMPLETE_REQUIRED = 3

' Special return values for SQLGetData
'
Public Const SQL_NO_TOTAL = -4

' Defines for SQLGetFunctions
' Core Functions
'
Public Const SQL_API_SQLALLOCCONNECT = 1
Public Const SQL_API_SQLALLOCENV = 2
Public Const SQL_API_SQLALLOCSTMT = 3
Public Const SQL_API_SQLBINDCOL = 4
Public Const SQL_API_SQLCANCEL = 5
Public Const SQL_API_SQLCOLATTRIBUTES = 6
Public Const SQL_API_SQLCONNECT = 7
Public Const SQL_API_SQLDESCRIBECOL = 8
Public Const SQL_API_SQLDISCONNECT = 9
Public Const SQL_API_SQLERROR = 10
Public Const SQL_API_SQLEXECDIRECT = 11
Public Const SQL_API_SQLEXECUTE = 12
Public Const SQL_API_SQLFETCH = 13
Public Const SQL_API_SQLFREECONNECT = 14
Public Const SQL_API_SQLFREEENV = 15
Public Const SQL_API_SQLFREESTMT = 16
Public Const SQL_API_SQLGETCURSORNAME = 17
Public Const SQL_API_SQLNUMRESULTCOLS = 18
Public Const SQL_API_SQLPREPARE = 19
Public Const SQL_API_SQLROWCOUNT = 20
Public Const SQL_API_SQLSETCURSORNAME = 21
Public Const SQL_API_SQLSETPARAM = 22
Public Const SQL_API_SQLTRANSACT = 23
Public Const SQL_NUM_FUNCTIONS = 23
Public Const SQL_EXT_API_START = 40
Public Const SQL_API_SQLCOLUMNS = 40

' Level 1 Functions
'
Public Const SQL_API_SQLDRIVERCONNECT = 41
Public Const SQL_API_SQLGETCONNECTOPTION = 42
Public Const SQL_API_SQLGETDATA = 43
Public Const SQL_API_SQLGETFUNCTIONS = 44
Public Const SQL_API_SQLGETINFO = 45
Public Const SQL_API_SQLGETSTMTOPTION = 46
Public Const SQL_API_SQLGETTYPEINFO = 47
Public Const SQL_API_SQLPARAMDATA = 48
Public Const SQL_API_SQLPUTDATA = 49
Public Const SQL_API_SQLSETCONNECTOPTION = 50
Public Const SQL_API_SQLSETSTMTOPTION = 51
Public Const SQL_API_SQLSPECIALCOLUMNS = 52
Public Const SQL_API_SQLSTATISTICS = 53
Public Const SQL_API_SQLTABLES = 54

' Level 2 Functions
'
Public Const SQL_API_SQLBROWSECONNECT = 55
Public Const SQL_API_SQLCOLUMNPRIVILEGES = 56
Public Const SQL_API_SQLDATASOURCES = 57
Public Const SQL_API_SQLDESCRIBEPARAM = 58
Public Const SQL_API_SQLEXTENDEDFETCH = 59
Public Const SQL_API_SQLFOREIGNKEYS = 60
Public Const SQL_API_SQLMORERESULTS = 61
Public Const SQL_API_SQLNATIVESQL = 62
Public Const SQL_API_SQLNUMPARAMS = 63
Public Const SQL_API_SQLPARAMOPTIONS = 64
Public Const SQL_API_SQLPRIMARYKEYS = 65
Public Const SQL_API_SQLPROCEDURECOLUMNS = 66
Public Const SQL_API_SQLPROCEDURES = 67
Public Const SQL_API_SQLSETPOS = 68
Public Const SQL_API_SQLSETSCROLLOPTIONS = 69
Public Const SQL_API_SQLTABLEPRIVILEGES = 70
Public Const SQL_EXT_API_LAST = 70

Public Const SQL_NUM_EXTENSIONS = (SQL_EXT_API_LAST - SQL_EXT_API_START + 1)

' Defines for SQLGetInfo
'
Public Const SQL_INFO_FIRST = 0
Public Const SQL_ACTIVE_CONNECTIONS = 0
Public Const SQL_ACTIVE_STATEMENTS = 1
Public Const SQL_DATA_SOURCE_NAME = 2
Public Const SQL_DRIVER_HDBC = 3
Public Const SQL_DRIVER_HENV = 4
Public Const SQL_DRIVER_HSTMT = 5
Public Const SQL_DRIVER_NAME = 6
Public Const SQL_DRIVER_VER = 7
Public Const SQL_FETCH_DIRECTION = 8
Public Const SQL_ODBC_API_CONFORMANCE = 9
Public Const SQL_ODBC_VER = 10
Public Const SQL_ROW_UPDATES = 11
Public Const SQL_ODBC_SAG_CLI_CONFORMANCE = 12
Public Const SQL_SERVER_NAME = 13
Public Const SQL_SEARCH_PATTERN_ESCAPE = 14
Public Const SQL_ODBC_SQL_CONFORMANCE = 15

Public Const SQL_DATABASE_NAME = 16
Public Const SQL_DBMS_NAME = 17
Public Const SQL_DBMS_VER = 18

Public Const SQL_ACCESSIBLE_TABLES = 19
Public Const SQL_ACCESSIBLE_PROCEDURES = 20
Public Const SQL_PROCEDURES = 21
Public Const SQL_CONCAT_NULL_BEHAVIOR = 22
Public Const SQL_CURSOR_COMMIT_BEHAVIOR = 23
Public Const SQL_CURSOR_ROLLBACK_BEHAVIOR = 24
Public Const SQL_DATA_SOURCE_READ_ONLY = 25
Public Const SQL_DEFAULT_TXN_ISOLATION = 26
Public Const SQL_EXPRESSIONS_IN_ORDERBY = 27
Public Const SQL_IDENTIFIER_CASE = 28
Public Const SQL_IDENTIFIER_QUOTE_CHAR = 29
Public Const SQL_MAX_COLUMN_NAME_LEN = 30
Public Const SQL_MAX_CURSOR_NAME_LEN = 31
Public Const SQL_MAX_OWNER_NAME_LEN = 32
Public Const SQL_MAX_PROCEDURE_NAME_LEN = 33
Public Const SQL_MAX_QUALIFIER_NAME_LEN = 34
Public Const SQL_MAX_TABLE_NAME_LEN = 35
Public Const SQL_MULT_RESULT_SETS = 36
Public Const SQL_MULTIPLE_ACTIVE_TXN = 37
Public Const SQL_OUTER_JOINS = 38
Public Const SQL_OWNER_TERM = 39
Public Const SQL_PROCEDURE_TERM = 40
Public Const SQL_QUALIFIER_NAME_SEPARATOR = 41
Public Const SQL_QUALIFIER_TERM = 42
Public Const SQL_SCROLL_CONCURRENCY = 43
Public Const SQL_SCROLL_OPTIONS = 44
Public Const SQL_TABLE_TERM = 45
Public Const SQL_TXN_CAPABLE = 46
Public Const SQL_USER_NAME = 47

Public Const SQL_CONVERT_FUNCTIONS = 48
Public Const SQL_NUMERIC_FUNCTIONS = 49
Public Const SQL_STRING_FUNCTIONS = 50
Public Const SQL_SYSTEM_FUNCTIONS = 51
Public Const SQL_TIMEDATE_FUNCTIONS = 52

Public Const SQL_CONVERT_BIGINT = 53
Public Const SQL_CONVERT_BINARY = 54
Public Const SQL_CONVERT_BIT = 55
Public Const SQL_CONVERT_CHAR = 56
Public Const SQL_CONVERT_DATE = 57
Public Const SQL_CONVERT_DECIMAL = 58
Public Const SQL_CONVERT_DOUBLE = 59
Public Const SQL_CONVERT_FLOAT = 60
Public Const SQL_CONVERT_INTEGER = 61
Public Const SQL_CONVERT_LONGVARCHAR = 62
Public Const SQL_CONVERT_NUMERIC = 63
Public Const SQL_CONVERT_REAL = 64
Public Const SQL_CONVERT_SMALLINT = 65
Public Const SQL_CONVERT_TIME = 66
Public Const SQL_CONVERT_TIMESTAMP = 67
Public Const SQL_CONVERT_TINYINT = 68
Public Const SQL_CONVERT_VARBINARY = 69
Public Const SQL_CONVERT_VARCHAR = 70
Public Const SQL_CONVERT_LONGVARBINARY = 71

Public Const SQL_TXN_ISOLATION_OPTION = 72
Public Const SQL_ODBC_SQL_OPT_IEF = 73

Public Const SQL_INFO_LAST = 73
Public Const SQL_INFO_DRIVER_START = 1000

' "SQL_CONVERT_" return value bitmasks
'
Public Const SQL_CVT_CHAR = &H1&
Public Const SQL_CVT_NUMERIC = &H2&
Public Const SQL_CVT_DECIMAL = &H4&
Public Const SQL_CVT_INTEGER = &H8&
Public Const SQL_CVT_SMALLINT = &H10&
Public Const SQL_CVT_FLOAT = &H20&
Public Const SQL_CVT_REAL = &H40&
Public Const SQL_CVT_DOUBLE = &H80&
Public Const SQL_CVT_VARCHAR = &H100&
Public Const SQL_CVT_LONGVARCHAR = &H200&
Public Const SQL_CVT_BINARY = &H400&
Public Const SQL_CVT_VARBINARY = &H800&
Public Const SQL_CVT_BIT = &H1000&
Public Const SQL_CVT_TINYINT = &H2000&
Public Const SQL_CVT_BIGINT = &H4000&
Public Const SQL_CVT_DATE = &H8000&
Public Const SQL_CVT_TIME = &H10000
Public Const SQL_CVT_TIMESTAMP = &H20000
Public Const SQL_CVT_LONGVARBINARY = &H40000


' Conversion functions
'
Public Const SQL_FN_CVT_CONVERT = &H1&

' String functions
'
Public Const SQL_FN_STR_CONCAT = &H1&
Public Const SQL_FN_STR_INSERT = &H2&
Public Const SQL_FN_STR_LEFT = &H4&
Public Const SQL_FN_STR_LTRIM = &H8&
Public Const SQL_FN_STR_LENGTH = &H10&
Public Const SQL_FN_STR_LOCATE = &H20&
Public Const SQL_FN_STR_LCASE = &H40&
Public Const SQL_FN_STR_REPEAT = &H80&
Public Const SQL_FN_STR_REPLACE = &H100&
Public Const SQL_FN_STR_RIGHT = &H200&
Public Const SQL_FN_STR_RTRIM = &H400&
Public Const SQL_FN_STR_SUBSTRING = &H800&
Public Const SQL_FN_STR_UCASE = &H1000&
Public Const SQL_FN_STR_ASCII = &H2000&
Public Const SQL_FN_STR_CHAR = &H4000&

' Numeric functions
'
Public Const SQL_FN_NUM_ABS = &H1&
Public Const SQL_FN_NUM_ACOS = &H2&
Public Const SQL_FN_NUM_ASIN = &H4&
Public Const SQL_FN_NUM_ATAN = &H8&
Public Const SQL_FN_NUM_ATAN2 = &H10&
Public Const SQL_FN_NUM_CEILING = &H20&
Public Const SQL_FN_NUM_COS = &H40&
Public Const SQL_FN_NUM_COT = &H80&
Public Const SQL_FN_NUM_EXP = &H100&
Public Const SQL_FN_NUM_FLOOR = &H200&
Public Const SQL_FN_NUM_LOG = &H400&
Public Const SQL_FN_NUM_MOD = &H800&
Public Const SQL_FN_NUM_SIGN = &H1000&
Public Const SQL_FN_NUM_SIN = &H2000&
Public Const SQL_FN_NUM_SQRT = &H4000&
Public Const SQL_FN_NUM_TAN = &H8000&
Public Const SQL_FN_NUM_PI = &H10000
Public Const SQL_FN_NUM_RAND = &H20000

' Time/date functions
'
Public Const SQL_FN_TD_NOW = &H1&
Public Const SQL_FN_TD_CURDATE = &H2&
Public Const SQL_FN_TD_DAYOFMONTH = &H4&
Public Const SQL_FN_TD_DAYOFWEEK = &H8&
Public Const SQL_FN_TD_DAYOFYEAR = &H10&
Public Const SQL_FN_TD_MONTH = &H20&
Public Const SQL_FN_TD_QUARTER = &H40&
Public Const SQL_FN_TD_WEEK = &H80&
Public Const SQL_FN_TD_YEAR = &H100&
Public Const SQL_FN_TD_CURTIME = &H200&
Public Const SQL_FN_TD_HOUR = &H400&
Public Const SQL_FN_TD_MINUTE = &H800&
Public Const SQL_FN_TD_SECOND = &H1000&

' System functions
'
Public Const SQL_FN_SYS_USERNAME = &H1&
Public Const SQL_FN_SYS_DBNAME = &H2&
Public Const SQL_FN_SYS_IFNULL = &H4&

' Scroll option masks
'
Public Const SQL_SO_FORWARD_ONLY = &H1&
Public Const SQL_SO_KEYSET_DRIVEN = &H2&
Public Const SQL_SO_DYNAMIC = &H4&
Public Const SQL_SO_MIXED = &H8&

' Scroll concurrency option masks
'
Public Const SQL_SCCO_READ_ONLY = &H1&
Public Const SQL_SCCO_LOCK = &H2&
Public Const SQL_SCCO_OPT_TIMESTAMP = &H4&
Public Const SQL_SCCO_OPT_VALUES = &H8&

' Fetch direction option masks
'
Public Const SQL_FD_FETCH_NEXT = &H1&
Public Const SQL_FD_FETCH_FIRST = &H2&
Public Const SQL_FD_FETCH_LAST = &H4&
Public Const SQL_FD_FETCH_PREV = &H8&
Public Const SQL_FD_FETCH_ABSOLUTE = &H10&
Public Const SQL_FD_FETCH_RELATIVE = &H20&
Public Const SQL_FD_FETCH_RESUME = &H40&

' Transaction isolation option masks
'
Public Const SQL_TXN_READ_UNCOMMITTED = &H1&
Public Const SQL_TXN_READ_COMMITTED = &H2&
Public Const SQL_TXN_REPEATABLE_READ = &H4&
Public Const SQL_TXN_SERIALIZABLE = &H8&
Public Const SQL_TXN_VERSIONING = &H10&

' options for SQLGetStmtOption/SQLSetStmtOption
'
Public Const SQL_QUERY_TIMEOUT = 0
Public Const SQL_MAX_ROWS = 1
Public Const SQL_NOSCAN = 2
Public Const SQL_MAX_LENGTH = 3
Public Const SQL_ASYNC_ENABLE = 4
Public Const SQL_BIND_TYPE = 5

Public Const SQL_BIND_BY_COLUMN = 0

' Suboption for SQL_BIND_TYPE
' options for SQLSetConnectOption/SQLGetConnectOption
'
Public Const SQL_ACCESS_MODE = 101
Public Const SQL_AUTOCOMMIT = 102
Public Const SQL_LOGIN_TIMEOUT = 103
Public Const SQL_OPT_TRACE = 104
Public Const SQL_OPT_TRACEFILE = 105
Public Const SQL_TRANSLATE_DLL = 106
Public Const SQL_TRANSLATE_OPTION = 107
Public Const SQL_TXN_ISOLATION = 108
Public Const SQL_CONNECT_OPT_DRVR_START = 1000

' Column types and scopes in SQLSpecialColumns.
'
Public Const SQL_BEST_ROWID = 1
Public Const SQL_ROWVER = 2

Public Const SQL_SCOPE_CURROW = 0
Public Const SQL_SCOPE_TRANSACTION = 1
Public Const SQL_SCOPE_SESSION = 2

' Level 2 Functions
'
' SQLExtendedFetch "fFetchType" values
'
Public Const SQL_FETCH_NEXT = 1
Public Const SQL_FETCH_FIRST = 2
Public Const SQL_FETCH_LAST = 3
Public Const SQL_FETCH_PREV = 4
Public Const SQL_FETCH_ABSOLUTE = 5
Public Const SQL_FETCH_RELATIVE = 6
Public Const SQL_FETCH_RESUME = 7

' SQLExtendedFetch "rgfRowStatus" element values
'
Public Const SQL_ROW_SUCCESS = 0
Public Const SQL_ROW_DELETED = 1
Public Const SQL_ROW_UPDATED = 2
Public Const SQL_ROW_NOROW = 3

' Defines for SQLForeignKeys (returned in result set)
'
Public Const SQL_CASCADE = 0
Public Const SQL_RESTRICT = 1
Public Const SQL_SET_NULL = 2

' Defines for SQLProcedureColumns (returned in the result set)
'
Public Const SQL_PARAM_TYPE_UNKNOWN = 0
Public Const SQL_PARAM_INPUT = 1
Public Const SQL_PARAM_INPUT_OUTPUT = 2
Public Const SQL_RESULT_COL = 3

' Defines for SQLSetScrollOptions
'
Public Const SQL_CONCUR_READ_ONLY = 1
Public Const SQL_CONCUR_LOCK = 2
Public Const SQL_CONCUR_TIMESTAMP = 3
Public Const SQL_CONCUR_VALUES = 4

Public Const SQL_SCROLL_FORWARD_ONLY = 0
Public Const SQL_SCROLL_KEYSET_DRIVEN = -1
Public Const SQL_SCROLL_DYNAMIC = -2

' Defines for SQLStatistics
'
Public Const SQL_INDEX_UNIQUE = 0
Public Const SQL_INDEX_ALL = 1
Public Const SQL_ENSURE = 1
Public Const SQL_QUICK = 0

' Defines for SQLStatistics (returned in the result set)
'
Public Const SQL_TABLE_STAT = 0
Public Const SQL_INDEX_CLUSTERED = 1
Public Const SQL_INDEX_HASHED = 2
Public Const SQL_INDEX_OTHER = 3

' Defines for SQLSetPos
'
Public Const SQL_ENTIRE_ROWSET = 0

'
'|========================================================================|
'| ODBC Module Extended Definitions                                       |
'|========================================================================|
'
'  ODBC Extended API's Definitions
'
'  Level 1 Prototypes
'
'Declare Function SQLColumns Lib "odbc32.dll" (ByVal hstmt as long, ByVal szTableQualifier as string, ByVal cbTableQualifier%, ByVal szTableOwner$, ByVal cbTableOwner%, ByVal szTableName$, ByVal cbTableName%, ByVal szColumnName$, ByVal cbColumnName%) As Integer
Declare Function SQLDriverConnect Lib "odbc32.dll" (ByVal hdbc As Long, ByVal hwnd As Integer, ByVal szCSIn As String, ByVal cbCSIn As Integer, ByVal szCSOut As String, ByVal cbCSMax As Integer, cbCSOut As Integer, ByVal f As Integer) As Integer
Declare Function SQLGetConnectOption Lib "odbc32.dll" (ByVal hdbc As Long, ByVal fOption As Integer, pvParam As Any) As Integer
Declare Function SQLGetData Lib "odbc32.dll" (ByVal hstmt As Long, ByVal col As Integer, ByVal wConvType As Integer, ByVal lpbBuf As String, ByVal dwbuflen As Long, lpcbout As Long) As Integer
Declare Function SQLGetFunctions Lib "odbc32.dll" (ByVal hdbc As Long, ByVal fFunction As Integer, pfExists As Integer) As Integer
Declare Function SQLGetInfo Lib "odbc32.dll" (ByVal hdbc As Long, ByVal hwnd As Integer, ByVal szInfo As String, ByVal cbInfoMax As Integer, cbInfoOut As Integer) As Integer
Declare Function SQLGetStmtOption Lib "odbc32.dll" (ByVal hstmt As Long, ByVal fOption As Integer, pvParam As Any) As Integer
Declare Function SQLGetTypeInfo Lib "odbc32.dll" (ByVal hstmt As Long, ByVal fSqlType As Integer) As Integer
'Declare Function SQLParamData Lib "odbc32.dll" (ByVal hstmt As Long, prgbValue As Any) As Integer
Declare Function SQLPutData Lib "odbc32.dll" (ByVal hstmt As Long, rgbValue As Any, ByVal cbValue As Long) As Integer
Declare Function SQLSetConnectOption Lib "odbc32.dll" (ByVal hdbc As Long, ByVal fOption As Integer, ByVal vParam As Long) As Integer
Declare Function SQLSetStmtOption Lib "odbc32.dll" (ByVal hstmt As Long, ByVal fOption As Integer, ByVal vParam As Long) As Integer
Declare Function SQLSpecialColumns Lib "odbc32.dll" (ByVal hstmt&, ByVal fColTyp%, ByVal szTblQualifier$, ByVal cbTblQualifier%, ByVal szTblOwner$, ByVal cbTblOwner%, ByVal szTblName$, ByVal cbTblName%, ByVal fScope%, ByVal fNullable%) As Integer
Declare Function SQLStatistics Lib "odbc32.dll" (ByVal hstmt&, ByVal szTblQualifier$, ByVal cbTblQualifier%, ByVal szTblOwner$, ByVal cbTblOwner%, ByVal szTblName$, ByVal cbTblName%, ByVal fUnique%, ByVal fAccuracy%) As Integer
Declare Function SQLTables Lib "odbc32.dll" (ByVal hstmt As Long, ByVal q As Long, ByVal cbq As Integer, ByVal o As Long, ByVal cbo As Integer, ByVal t As Long, ByVal cbt As Integer, ByVal tt As Long, ByVal cbtt As Integer) As Integer
Declare Function SQLColumns Lib "odbc32.dll" (ByVal hstmt As Long, ByVal q As Long, ByVal cbq As Integer, ByVal o As Long, ByVal cbo As Integer, ByVal t As String, ByVal cbt As Integer, ByVal cn As Long, ByVal cbcn As Integer) As Integer

'  Level 2 Prototypes
'
Declare Function SQLBrowseConnect Lib "odbc32.dll" (ByVal hdbc&, ByVal szConnStrIn$, ByVal cbConnStrIn%, ByVal szConnStrOut$, ByVal cbConnStrOutMax%, pcbConnStrOut%) As Integer
Declare Function SQLColumnPrivileges Lib "odbc32.dll" (ByVal hstmt&, ByVal szTQf$, ByVal cbTQf%, ByVal szTOwn$, ByVal cbTOwn%, ByVal szTName$, ByVal cbTName%, ByVal szColName$, ByVal cbColumnName%) As Integer
Declare Function SQLDataSources Lib "odbc32.dll" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Declare Function SQLDescribeParam Lib "odbc32.dll" (ByVal hstmt&, ByVal ipar%, pfSqlType%, pcbColDef&, pibScale%, pfNullable%) As Integer
Declare Function SQLExtendedFetch Lib "odbc32.dll" (ByVal hstmt As Long, ByVal fFetchType As Integer, ByVal irow As Long, pcrow As Long, rgfRowStatus() As Integer) As Integer
Declare Function SQLForeignKeys Lib "odbc32.dll" (ByVal hstmt&, ByVal PTQf$, ByVal PTQf%, ByVal PTO$, ByVal PTO%, ByVal PTName$, ByVal PTName%, ByVal FTQf$, ByVal FTQf%, ByVal FTO$, ByVal FTO%, ByVal FTName$, ByVal FTName%) As Integer
Declare Function SQLMoreResults Lib "odbc32.dll" (ByVal hstmt&) As Integer
Declare Function SQLNativeSql Lib "odbc32.dll" (ByVal hdbc&, ByVal szSqlStrIn$, ByVal cbSqlStrIn&, ByVal szSqlStr$, ByVal cbSqlStrMax&, pcbSqlStr&) As Integer
Declare Function SQLNumParams Lib "odbc32.dll" (ByVal hstmt&, pcpar%) As Integer
Declare Function SQLParamOptions Lib "odbc32.dll" (ByVal hstmt&, ByVal crow%, pirow&) As Integer
Declare Function SQLPrimaryKeys Lib "odbc32.dll" (ByVal hstmt&, ByVal szTableQualifier$, ByVal cbTableQualifier%, ByVal szTableOwner$, ByVal cbTableOwner%, ByVal szTableName$, ByVal cbTableName%) As Integer
Declare Function SQLProcedureColumns Lib "odbc32.dll" (ByVal hstmt&, ByVal szProcQualifier$, ByVal cbProcQualifier%, ByVal szProcOwner$, ByVal cbProcOwner%, ByVal szProcName$, ByVal cbProcName%, ByVal szColumnName$, ByVal cbColumnName%) As Integer
Declare Function SQLProcedures Lib "odbc32.dll" (ByVal hstmt&, ByVal szProcQualifier$, ByVal cbProcQualifier%, ByVal szProcOwner$, ByVal cbProcOwner%, ByVal szProcName$, ByVal cbProcName%) As Integer
Declare Function SQLSetPos Lib "odbc32.dll" (ByVal hstmt&, ByVal irow%, ByVal fRefresh%, ByVal fLock%) As Integer
Declare Function SQLSetScrollOptions Lib "odbc32.dll" (ByVal hstmt&, ByVal fConcurrency%, ByVal crowKeyset&, ByVal crowRowset%) As Integer
Declare Function SQLTablePrivileges Lib "odbc32.dll" (ByVal hstmt&, ByVal szTableQualifier$, ByVal cbTableQualifier%, ByVal szTableOwner$, ByVal cbTableOwner%, ByVal szTableName$, ByVal cbTableName%) As Integer


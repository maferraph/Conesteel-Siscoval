VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Classe_Kernel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ------------------------------------------------------------------------
'
'    WIN32API.TXT -- Win32 API Declarations for Visual Basic
'
'              Copyright (C) 1994 Microsoft Corporation
'
'
'  This file contains only the Const, Type,
' and Declare statements for  Win32 APIs.
'
'  You have a royalty-free right to use, modify, reproduce and distribute
'  this file (and/or any modified version) in any way you find useful,
'  provided that you agree that Microsoft has no warranty, obligation or
'  liability for its contents.  Refer to the Microsoft Windows Programmer's
'  Reference for further information.
'
' ------------------------------------------------------------------------

' Type definitions for Windows' basic types.
Private Const ANYSIZE_ARRAY = 1
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type RECTL
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type POINTAPI
        X As Long
        Y As Long
End Type

Type POINTL
        X As Long
        Y As Long
End Type

Type Size
        cx As Long
        cy As Long
End Type

Type POINTS
        X  As Integer
        Y  As Integer
End Type

Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type


Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000


Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000

Private Const SPECIFIC_RIGHTS_ALL = &HFFFF


Type SID_IDENTIFIER_AUTHORITY
        Value(6) As Byte
End Type

Private Const SID_REVISION = (1)                         '  Current revision level
Private Const SID_MAX_SUB_AUTHORITIES = (15)
Private Const SID_RECOMMENDED_SUB_AUTHORITIES = (1)  ' Will change to around 6 in a future release.

Private Const SidTypeUser = 1
Private Const SidTypeGroup = 2
Private Const SidTypeDomain = 3
Private Const SidTypeAlias = 4
Private Const SidTypeWellKnownGroup = 5
Private Const SidTypeDeletedAccount = 6
Private Const SidTypeInvalid = 7
Private Const SidTypeUnknown = 8

Type SID_AND_ATTRIBUTES
        Sid As Long
        Attributes As Long
End Type

' ///////////////////////////////////////////////////////////////////////////
'                                                                          //
'  Universal well-known SIDs                                               //
'                                                                          //
'      Null SID              S-1-0-0                                       //
'      World                 S-1-1-0                                       //
'      Local                 S-1-2-0                                       //
'      Creator Owner ID      S-1-3-0                                       //
'      Creator Group ID      S-1-3-1                                       //
'                                                                          //
'      (Non-unique IDs)      S-1-4                                         //
'                                                                          //
' ///////////////////////////////////////////////////////////////////////////
Private Const SECURITY_NULL_RID = &H0
Private Const SECURITY_WORLD_RID = &H0
Private Const SECURITY_LOCAL_RID = &H0

Private Const SECURITY_CREATOR_OWNER_RID = &H0
Private Const SECURITY_CREATOR_GROUP_RID = &H1


' ///////////////////////////////////////////////////////////////////////////
'                                                                          //
'  NT well-known SIDs                                                      //
'                                                                          //
'      NT Authority          S-1-5                                         //
'      Dialup                S-1-5-1                                       //
'                                                                          //
'      Network               S-1-5-2                                       //
'      Batch                 S-1-5-3                                       //
'      Interactive           S-1-5-4                                       //
'      Service               S-1-5-6                                       //
'      AnonymousLogon        S-1-5-7       (aka null logon session)        //
'                                                                          //
'      (Logon IDs)           S-1-5-5-X-Y                                   //
'                                                                          //
'      (NT non-unique IDs)   S-1-5-0x15-...                                //
'                                                                          //
'      (Built-in domain)     s-1-5-0x20                                    //
'                                                                          //
' ///////////////////////////////////////////////////////////////////////////

Private Const SECURITY_DIALUP_RID = &H1
Private Const SECURITY_NETWORK_RID = &H2
Private Const SECURITY_BATCH_RID = &H3
Private Const SECURITY_INTERACTIVE_RID = &H4
Private Const SECURITY_SERVICE_RID = &H6
Private Const SECURITY_ANONYMOUS_LOGON_RID = &H7
Private Const SECURITY_LOGON_IDS_RID = &H5
Private Const SECURITY_LOCAL_SYSTEM_RID = &H12
Private Const SECURITY_NT_NON_UNIQUE = &H15
Private Const SECURITY_BUILTIN_DOMAIN_RID = &H20


' ///////////////////////////////////////////////////////////////////////////
'                                                                          //
'  well-known domain relative sub-authority values (RIDs)...               //
'                                                                          //
' ///////////////////////////////////////////////////////////////////////////

Private Const DOMAIN_USER_RID_ADMIN = &H1F4
Private Const DOMAIN_USER_RID_GUEST = &H1F5

Private Const DOMAIN_GROUP_RID_ADMINS = &H200
Private Const DOMAIN_GROUP_RID_USERS = &H201
Private Const DOMAIN_GROUP_RID_GUESTS = &H202


Private Const DOMAIN_ALIAS_RID_ADMINS = &H220
Private Const DOMAIN_ALIAS_RID_USERS = &H221
Private Const DOMAIN_ALIAS_RID_GUESTS = &H222
Private Const DOMAIN_ALIAS_RID_POWER_USERS = &H223
Private Const DOMAIN_ALIAS_RID_ACCOUNT_OPS = &H224
Private Const DOMAIN_ALIAS_RID_SYSTEM_OPS = &H225
Private Const DOMAIN_ALIAS_RID_PRINT_OPS = &H226
Private Const DOMAIN_ALIAS_RID_BACKUP_OPS = &H227
Private Const DOMAIN_ALIAS_RID_REPLICATOR = &H228


'  Allocate the System Luid.  The first 1000 LUIDs are reserved.
'  Use #999 here0x3E7 = 999)

'  end_ntifs

' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                           User and Group related SID attributes     //
'                                                                     //
' //////////////////////////////////////////////////////////////////////

'  Group attributes

Private Const SE_GROUP_MANDATORY = &H1
Private Const SE_GROUP_ENABLED_BY_DEFAULT = &H2
Private Const SE_GROUP_ENABLED = &H4
Private Const SE_GROUP_OWNER = &H8
Private Const SE_GROUP_LOGON_ID = &HC0000000

'  User attributes

'  (None yet defined.)

' ----------------
'  Kernel Section
' ----------------

Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2

Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FILE_FLAG_NO_BUFFERING = &H20000000
Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Private Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Private Const FILE_FLAG_POSIX_SEMANTICS = &H1000000

Private Const CREATE_NEW = 1
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const TRUNCATE_EXISTING = 5

' Define the dwOpenMode values for CreateNamedPipe
Private Const PIPE_ACCESS_INBOUND = &H1
Private Const PIPE_ACCESS_OUTBOUND = &H2
Private Const PIPE_ACCESS_DUPLEX = &H3

' Define the Named Pipe End flags for GetNamedPipeInfo
Private Const PIPE_CLIENT_END = &H0
Private Const PIPE_SERVER_END = &H1

' Define the dwPipeMode values for CreateNamedPipe
Private Const PIPE_WAIT = &H0
Private Const PIPE_NOWAIT = &H1
Private Const PIPE_READMODE_BYTE = &H0
Private Const PIPE_READMODE_MESSAGE = &H2
Private Const PIPE_TYPE_BYTE = &H0
Private Const PIPE_TYPE_MESSAGE = &H4

' Define the well known values for CreateNamedPipe nMaxInstances
Private Const PIPE_UNLIMITED_INSTANCES = 255

' Define the Security Quality of Service bits to be passed
'  into CreateFile
Private Const SECURITY_CONTEXT_TRACKING = &H40000
Private Const SECURITY_EFFECTIVE_ONLY = &H80000

Private Const SECURITY_SQOS_PRESENT = &H100000
Private Const SECURITY_VALID_SQOS_FLAGS = &H1F0000

Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type SystemTime
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

'  Serial provider type.
Private Const SP_SERIALCOMM = &H1&

'  Provider SubTypes
Private Const PST_UNSPECIFIED = &H0&
Private Const PST_RS232 = &H1&
Private Const PST_PARALLELPORT = &H2&
Private Const PST_RS422 = &H3&
Private Const PST_RS423 = &H4&
Private Const PST_RS449 = &H5&
Private Const PST_FAX = &H21&
Private Const PST_SCANNER = &H22&
Private Const PST_NETWORK_BRIDGE = &H100&
Private Const PST_LAT = &H101&
Private Const PST_TCPIP_TELNET = &H102&
Private Const PST_X25 = &H103&

'  Provider capabilities flags.
Private Const PCF_DTRDSR = &H1&
Private Const PCF_RTSCTS = &H2&
Private Const PCF_RLSD = &H4&
Private Const PCF_PARITY_CHECK = &H8&
Private Const PCF_XONXOFF = &H10&
Private Const PCF_SETXCHAR = &H20&
Private Const PCF_TOTALTIMEOUTS = &H40&
Private Const PCF_INTTIMEOUTS = &H80&
Private Const PCF_SPECIALCHARS = &H100&
Private Const PCF_16BITMODE = &H200&

'  Comm provider settable parameters.
Private Const SP_PARITY = &H1&
Private Const SP_BAUD = &H2&
Private Const SP_DATABITS = &H4&
Private Const SP_STOPBITS = &H8&
Private Const SP_HANDSHAKING = &H10&
Private Const SP_PARITY_CHECK = &H20&
Private Const SP_RLSD = &H40&

'  Settable baud rates in the provider.
Private Const BAUD_075 = &H1&
Private Const BAUD_110 = &H2&
Private Const BAUD_134_5 = &H4&
Private Const BAUD_150 = &H8&
Private Const BAUD_300 = &H10&
Private Const BAUD_600 = &H20&
Private Const BAUD_1200 = &H40&
Private Const BAUD_1800 = &H80&
Private Const BAUD_2400 = &H100&
Private Const BAUD_4800 = &H200&
Private Const BAUD_7200 = &H400&
Private Const BAUD_9600 = &H800&
Private Const BAUD_14400 = &H1000&
Private Const BAUD_19200 = &H2000&
Private Const BAUD_38400 = &H4000&
Private Const BAUD_56K = &H8000&
Private Const BAUD_128K = &H10000
Private Const BAUD_115200 = &H20000
Private Const BAUD_57600 = &H40000
Private Const BAUD_USER = &H10000000

'  Settable Data Bits
Private Const DATABITS_5 = &H1&
Private Const DATABITS_6 = &H2&
Private Const DATABITS_7 = &H4&
Private Const DATABITS_8 = &H8&
Private Const DATABITS_16 = &H10&
Private Const DATABITS_16X = &H20&

'  Settable Stop and Parity bits.
Private Const STOPBITS_10 = &H1&
Private Const STOPBITS_15 = &H2&
Private Const STOPBITS_20 = &H4&
Private Const PARITY_NONE = &H100&
Private Const PARITY_ODD = &H200&
Private Const PARITY_EVEN = &H400&
Private Const PARITY_MARK = &H800&
Private Const PARITY_SPACE = &H1000&

Type COMMPROP
        wPacketLength As Integer
        wPacketVersion As Integer
        dwServiceMask As Long
        dwReserved1 As Long
        dwMaxTxQueue As Long
        dwMaxRxQueue As Long
        dwMaxBaud As Long
        dwProvSubType As Long
        dwProvCapabilities As Long
        dwSettableParams As Long
        dwSettableBaud As Long
        wSettableData As Integer
        wSettableStopParity As Integer
        dwCurrentTxQueue As Long
        dwCurrentRxQueue As Long
        dwProvSpec1 As Long
        dwProvSpec2 As Long
        wcProvChar(1) As Integer
End Type

'Type COMSTAT
'        fCtsHold As Long
'        fDsrHold As Long
'        fRlsdHold As Long
'        fXoffHold As Long
'        fXoffSent As Long
'        fEof As Long
'        fTxim As Long
'        fReserved As Long
'        cbInQue As Long
'        cbOutQue As Long
'End Type

Type COMSTAT
        fBitFields As Long 'See Comment in Win32API.Txt
        cbInQue As Long
        cbOutQue As Long
End Type
' The eight actual COMSTAT bit-sized data fields within the four bytes of fBitFields can be manipulated by bitwise logical And/Or operations.
' FieldName     Bit #     Description
' ---------     -----     ---------------------------
' fCtsHold        1       Tx waiting for CTS signal
' fDsrHold        2       Tx waiting for DSR signal
' fRlsdHold       3       Tx waiting for RLSD signal
' fXoffHold       4       Tx waiting, XOFF char rec'd
' fXoffSent       5       Tx waiting, XOFF char sent
' fEof            6       EOF character sent
' fTxim           7       character waiting for Tx
' fReserved       8       reserved (25 bits)

'  DTR Control Flow Values.
Private Const DTR_CONTROL_DISABLE = &H0
Private Const DTR_CONTROL_ENABLE = &H1
Private Const DTR_CONTROL_HANDSHAKE = &H2

'  RTS Control Flow Values
Private Const RTS_CONTROL_DISABLE = &H0
Private Const RTS_CONTROL_ENABLE = &H1
Private Const RTS_CONTROL_HANDSHAKE = &H2
Private Const RTS_CONTROL_TOGGLE = &H3

'Type DCB
'        DCBlength As Long
'        BaudRate As Long
'        fBinary As Long
'        fParity As Long
'        fOutxCtsFlow As Long
'        fOutxDsrFlow As Long
'        fDtrControl As Long
'        fDsrSensitivity As Long
'        fTXContinueOnXoff As Long
'        fOutX As Long
'        fInX As Long
'        fErrorChar As Long
'        fNull As Long
'        fRtsControl As Long
'        fAbortOnError As Long
'        fDummy2 As Long
'        wReserved As Integer
'        XonLim As Integer
'        XoffLim As Integer
'        ByteSize As Byte
'        Parity As Byte
'        StopBits As Byte
'        XonChar As Byte
'        XoffChar As Byte
'        ErrorChar As Byte
'        EofChar As Byte
'        EvtChar As Byte
'End Type

Type DCB
        DCBlength As Long
        BaudRate As Long
        fBitFields As Long 'See Comments in Win32API.Txt
        wReserved As Integer
        XonLim As Integer
        XoffLim As Integer
        ByteSize As Byte
        Parity As Byte
        StopBits As Byte
        XonChar As Byte
        XoffChar As Byte
        ErrorChar As Byte
        EofChar As Byte
        EvtChar As Byte
        wReserved1 As Integer 'Reserved; Do Not Use
End Type
' The fourteen actual DCB bit-sized data fields within the four bytes of fBitFields can be manipulated by bitwise logical And/Or operations.
' FieldName             Bit #     Description
' -----------------     -----     ------------------------------
' fBinary                 1       binary mode, no EOF check
' fParity                 2       enable parity checking
' fOutxCtsFlow            3       CTS output flow control
' fOutxDsrFlow            4       DSR output flow control
' fDtrControl             5       DTR flow control type (2 bits)
' fDsrSensitivity         7       DSR sensitivity
' fTXContinueOnXoff       8       XOFF continues Tx
' fOutX                   9       XON/XOFF out flow control
' fInX                   10       XON/XOFF in flow control
' fErrorChar             11       enable error replacement
' fNull                  12       enable null stripping
' fRtsControl            13       RTS flow control (2 bits)
' fAbortOnError          15       abort reads/writes on error
' fDummy2                16       reserved

Type COMMTIMEOUTS
        ReadIntervalTimeout As Long
        ReadTotalTimeoutMultiplier As Long
        ReadTotalTimeoutConstant As Long
        WriteTotalTimeoutMultiplier As Long
        WriteTotalTimeoutConstant As Long
End Type

Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

' Global Memory Flags
Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_ZEROINIT = &H40
Private Const GMEM_MODIFY = &H80
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_INVALID_HANDLE = &H8000

Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

' Flags returned by GlobalFlags (in addition to GMEM_DISCARDABLE)
Private Const GMEM_DISCARDED = &H4000
Private Const GMEM_LOCKCOUNT = &HFF

Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

' Local Memory Flags
Private Const LMEM_FIXED = &H0
Private Const LMEM_MOVEABLE = &H2
Private Const LMEM_NOCOMPACT = &H10
Private Const LMEM_NODISCARD = &H20
Private Const LMEM_ZEROINIT = &H40
Private Const LMEM_MODIFY = &H80
Private Const LMEM_DISCARDABLE = &HF00
Private Const LMEM_VALID_FLAGS = &HF72
Private Const LMEM_INVALID_HANDLE = &H8000

Private Const LHND = (LMEM_MOVEABLE + LMEM_ZEROINIT)
Private Const LPTR = (LMEM_FIXED + LMEM_ZEROINIT)

Private Const NONZEROLHND = (LMEM_MOVEABLE)
Private Const NONZEROLPTR = (LMEM_FIXED)

' Flags returned by LocalFlags (in addition to LMEM_DISCARDABLE)
Private Const LMEM_DISCARDED = &H4000
Private Const LMEM_LOCKCOUNT = &HFF

'  dwCreationFlag values

Private Const DEBUG_PROCESS = &H1
Private Const DEBUG_ONLY_THIS_PROCESS = &H2

Private Const CREATE_SUSPENDED = &H4

Private Const DETACHED_PROCESS = &H8

Private Const CREATE_NEW_CONSOLE = &H10

Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const IDLE_PRIORITY_CLASS = &H40
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const REALTIME_PRIORITY_CLASS = &H100

Private Const CREATE_NEW_PROCESS_GROUP = &H200

Private Const CREATE_NO_WINDOW = &H8000000

Private Const PROFILE_USER = &H10000000
Private Const PROFILE_KERNEL = &H20000000
Private Const PROFILE_SERVER = &H40000000

Private Const MAXLONG = &H7FFFFFFF
Private Const THREAD_BASE_PRIORITY_MIN = -2
Private Const THREAD_BASE_PRIORITY_MAX = 2
Private Const THREAD_BASE_PRIORITY_LOWRT = 15
Private Const THREAD_BASE_PRIORITY_IDLE = -15
Private Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
Private Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
Private Const THREAD_PRIORITY_NORMAL = 0
Private Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Private Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Private Const THREAD_PRIORITY_ERROR_RETURN = (MAXLONG)

Private Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
Private Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE

' ++ BUILD Version: 0093     Increment this if a change has global effects

' Copyright (c) 1990-1995  Microsoft Corporation

' Module Name:

'     winnt.h

' Abstract:

'     This module defines the 32-Bit Windows types and constants that are
'     defined by NT, but exposed through the Win32 API.

' Revision History:
Private Const APPLICATION_ERROR_MASK = &H20000000
Private Const ERROR_SEVERITY_SUCCESS = &H0
Private Const ERROR_SEVERITY_INFORMATIONAL = &H40000000
Private Const ERROR_SEVERITY_WARNING = &H80000000
Private Const ERROR_SEVERITY_ERROR = &HC0000000


Private Const MINCHAR = &H80
Private Const MAXCHAR = &H7F
Private Const MINSHORT = &H8000
Private Const MAXSHORT = &H7FFF
Private Const MINLONG = &H80000000
Private Const MAXByte = &HFF
Private Const MAXWORD = &HFFFF
Private Const MAXDWORD = &HFFFF
'
'  Calculate the byte offset of a field in a structure of type type.
'  *  Language IDs.
'  *
'  *  The following two combinations of primary language ID and
'  *  sublanguage ID have special semantics:
'  *
'  *    Primary Language ID   Sublanguage ID      Result
'  *    -------------------   ---------------     ------------------------
'  *    LANG_NEUTRAL          SUBLANG_NEUTRAL     Language neutral
'  *    LANG_NEUTRAL          SUBLANG_DEFAULT     User default language
'  *    LANG_NEUTRAL          SUBLANG_SYS_DEFAULT System default language
'  */
'
'  *  Primary language IDs.
'  */
Private Const LANG_NEUTRAL = &H0

Private Const LANG_BULGARIAN = &H2
Private Const LANG_CHINESE = &H4
Private Const LANG_CROATIAN = &H1A
Private Const LANG_CZECH = &H5
Private Const LANG_DANISH = &H6
Private Const LANG_DUTCH = &H13
Private Const LANG_ENGLISH = &H9
Private Const LANG_FINNISH = &HB
Private Const LANG_FRENCH = &HC
Private Const LANG_GERMAN = &H7
Private Const LANG_GREEK = &H8
Private Const LANG_HUNGARIAN = &HE
Private Const LANG_ICELANDIC = &HF
Private Const LANG_ITALIAN = &H10
Private Const LANG_JAPANESE = &H11
Private Const LANG_KOREAN = &H12
Private Const LANG_NORWEGIAN = &H14
Private Const LANG_POLISH = &H15
Private Const LANG_PORTUGUESE = &H16
Private Const LANG_ROMANIAN = &H18
Private Const LANG_RUSSIAN = &H19
Private Const LANG_SLOVAK = &H1B
Private Const LANG_SLOVENIAN = &H24
Private Const LANG_SPANISH = &HA
Private Const LANG_SWEDISH = &H1D
Private Const LANG_TURKISH = &H1F

'
'  *  Sublanguage IDs.
'  *
'  *  The name immediately following SUBLANG_ dictates which primary
'  *  language ID that sublanguage ID can be combined with to form a
'  *  valid language ID.
'  */
Private Const SUBLANG_NEUTRAL = &H0                       '  language neutral
Private Const SUBLANG_DEFAULT = &H1                       '  user default
Private Const SUBLANG_SYS_DEFAULT = &H2                   '  system default

Private Const SUBLANG_CHINESE_TRADITIONAL = &H1           '  Chinese (Taiwan)
Private Const SUBLANG_CHINESE_SIMPLIFIED = &H2            '  Chinese (PR China)
Private Const SUBLANG_CHINESE_HONGKONG = &H3              '  Chinese (Hong Kong)
Private Const SUBLANG_CHINESE_SINGAPORE = &H4             '  Chinese (Singapore)
Private Const SUBLANG_DUTCH = &H1                         '  Dutch
Private Const SUBLANG_DUTCH_BELGIAN = &H2                 '  Dutch (Belgian)
Private Const SUBLANG_ENGLISH_US = &H1                    '  English (USA)
Private Const SUBLANG_ENGLISH_UK = &H2                    '  English (UK)
Private Const SUBLANG_ENGLISH_AUS = &H3                   '  English (Australian)
Private Const SUBLANG_ENGLISH_CAN = &H4                   '  English (Canadian)
Private Const SUBLANG_ENGLISH_NZ = &H5                    '  English (New Zealand)
Private Const SUBLANG_ENGLISH_EIRE = &H6                  '  English (Irish)
Private Const SUBLANG_FRENCH = &H1                        '  French
Private Const SUBLANG_FRENCH_BELGIAN = &H2                '  French (Belgian)
Private Const SUBLANG_FRENCH_CANADIAN = &H3               '  French (Canadian)
Private Const SUBLANG_FRENCH_SWISS = &H4                  '  French (Swiss)
Private Const SUBLANG_GERMAN = &H1                        '  German
Private Const SUBLANG_GERMAN_SWISS = &H2                  '  German (Swiss)
Private Const SUBLANG_GERMAN_AUSTRIAN = &H3               '  German (Austrian)
Private Const SUBLANG_ITALIAN = &H1                       '  Italian
Private Const SUBLANG_ITALIAN_SWISS = &H2                 '  Italian (Swiss)
Private Const SUBLANG_NORWEGIAN_BOKMAL = &H1              '  Norwegian (Bokma
Private Const SUBLANG_NORWEGIAN_NYNORSK = &H2             '  Norwegian (Nynorsk)
Private Const SUBLANG_PORTUGUESE = &H2                    '  Portuguese
Private Const SUBLANG_PORTUGUESE_BRAZILIAN = &H1          '  Portuguese (Brazilian)
Private Const SUBLANG_SPANISH = &H1                       '  Spanish (Castilian)
Private Const SUBLANG_SPANISH_MEXICAN = &H2               '  Spanish (Mexican)
Private Const SUBLANG_SPANISH_MODERN = &H3                '  Spanish (Modern)

'
'  *  Sorting IDs.
'  *
'  */
Private Const SORT_DEFAULT = &H0                          '  sorting default

Private Const SORT_JAPANESE_XJIS = &H0                    '  Japanese0xJIS order
Private Const SORT_JAPANESE_UNICODE = &H1                 '  Japanese Unicode order

Private Const SORT_CHINESE_BIG5 = &H0                     '  Chinese BIG5 order
Private Const SORT_CHINESE_UNICODE = &H1                  '  Chinese Unicode order

Private Const SORT_KOREAN_KSC = &H0                       '  Korean KSC order
Private Const SORT_KOREAN_UNICODE = &H1                   '  Korean Unicode order

'  The FILE_READ_DATA and FILE_WRITE_DATA constants are also defined in
'  devioctl.h as FILE_READ_ACCESS and FILE_WRITE_ACCESS. The values for these
'  constants *MUST* always be in sync.
'  The values are redefined in devioctl.h because they must be available to
'  both DOS and NT.
'

Private Const FILE_READ_DATA = (&H1)                     '  file pipe
Private Const FILE_LIST_DIRECTORY = (&H1)                '  directory

Private Const FILE_WRITE_DATA = (&H2)                    '  file pipe
Private Const FILE_ADD_FILE = (&H2)                      '  directory

Private Const FILE_APPEND_DATA = (&H4)                   '  file
Private Const FILE_ADD_SUBDIRECTORY = (&H4)              '  directory
Private Const FILE_CREATE_PIPE_INSTANCE = (&H4)          '  named pipe

Private Const FILE_READ_EA = (&H8)                       '  file directory
Private Const FILE_READ_PROPERTIES = FILE_READ_EA

Private Const FILE_WRITE_EA = (&H10)                     '  file directory
Private Const FILE_WRITE_PROPERTIES = FILE_WRITE_EA

Private Const FILE_EXECUTE = (&H20)                      '  file
Private Const FILE_TRAVERSE = (&H20)                     '  directory

Private Const FILE_DELETE_CHILD = (&H40)                 '  directory

Private Const FILE_READ_ATTRIBUTES = (&H80)              '  all

Private Const FILE_WRITE_ATTRIBUTES = (&H100)            '  all

Private Const FILE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H1FF)

Private Const FILE_GENERIC_READ = (STANDARD_RIGHTS_READ Or FILE_READ_DATA Or FILE_READ_ATTRIBUTES Or FILE_READ_EA Or SYNCHRONIZE)


Private Const FILE_GENERIC_WRITE = (STANDARD_RIGHTS_WRITE Or FILE_WRITE_DATA Or FILE_WRITE_ATTRIBUTES Or FILE_WRITE_EA Or FILE_APPEND_DATA Or SYNCHRONIZE)


Private Const FILE_GENERIC_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or FILE_READ_ATTRIBUTES Or FILE_EXECUTE Or SYNCHRONIZE)

Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_NOTIFY_CHANGE_FILE_NAME = &H1
Private Const FILE_NOTIFY_CHANGE_DIR_NAME = &H2
Private Const FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4
Private Const FILE_NOTIFY_CHANGE_SIZE = &H8
Private Const FILE_NOTIFY_CHANGE_LAST_WRITE = &H10
Private Const FILE_NOTIFY_CHANGE_SECURITY = &H100
Private Const MAILSLOT_NO_MESSAGE = (-1)
Private Const MAILSLOT_WAIT_FOREVER = (-1)
Private Const FILE_CASE_SENSITIVE_SEARCH = &H1
Private Const FILE_CASE_PRESERVED_NAMES = &H2
Private Const FILE_UNICODE_ON_DISK = &H4
Private Const FILE_PERSISTENT_ACLS = &H8
Private Const FILE_FILE_COMPRESSION = &H10
Private Const FILE_VOLUME_IS_COMPRESSED = &H8000
Private Const IO_COMPLETION_MODIFY_STATE = &H2
Private Const IO_COMPLETION_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3)
Private Const DUPLICATE_CLOSE_SOURCE = &H1
Private Const DUPLICATE_SAME_ACCESS = &H2

' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                              ACCESS MASK                            //
'                                                                     //
' //////////////////////////////////////////////////////////////////////

'
'   Define the access mask as a longword sized structure divided up as
'   follows:

'       typedef struct _ACCESS_MASK {
'           WORD   SpecificRights;
'           Byte  StandardRights;
'           Byte  AccessSystemAcl : 1;
'           Byte  Reserved : 3;
'           Byte  GenericAll : 1;
'           Byte  GenericExecute : 1;
'           Byte  GenericWrite : 1;
'           Byte  GenericRead : 1;
'       } ACCESS_MASK;
'       typedef ACCESS_MASK *PACCESS_MASK;
'
'   but to make life simple for programmer's we'll allow them to specify
'   a desired access mask by simply OR'ing together mulitple single rights
'   and treat an access mask as a DWORD.  For example
'
'       DesiredAccess = DELETE  Or  READ_CONTROL
'
'   So we'll declare ACCESS_MASK as DWORD
'

'  begin_ntddk begin_nthal begin_ntifs

' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                              ACCESS TYPES                           //
'                                                                     //
' //////////////////////////////////////////////////////////////////////


'  begin_ntddk begin_nthal begin_ntifs
'
'   The following are masks for the predefined standard access types

'  AccessSystemAcl access type

Private Const ACCESS_SYSTEM_SECURITY = &H1000000

'  MaximumAllowed access type

Private Const MAXIMUM_ALLOWED = &H2000000

'   These are the generic rights.

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_ALL = &H10000000

'   Define the generic mapping array.  This is used to denote the
'   mapping of each generic access right to a specific access mask.

Type GENERIC_MAPPING
        GenericRead As Long
        GenericWrite As Long
        GenericExecute As Long
        GenericAll As Long
End Type


' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                         LUID_AND_ATTRIBUTES                         //
'                                                                     //
' //////////////////////////////////////////////////////////////////////
'
Type Luid
    lowpart As Long
    highpart As Long
End Type


Type LUID_AND_ATTRIBUTES
        pLuid As Luid
        Attributes As Long
End Type

' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                          ACL  and  ACE                              //
'                                                                     //
' //////////////////////////////////////////////////////////////////////

'
'   Define an ACL and the ACE format.  The structure of an ACL header
'   followed by one or more ACEs.  Pictorally the structure of an ACL header
'   is as follows:
'
'   The current AclRevision is defined to be ACL_REVISION.
'
'   AclSize is the size, in bytes, allocated for the ACL.  This includes
'   the ACL header, ACES, and remaining free space in the buffer.
'
'   AceCount is the number of ACES in the ACL.
'

'  begin_ntddk begin_ntifs
'  This is the *current* ACL revision

Private Const ACL_REVISION = (2)

'  This is the history of ACL revisions.  Add a new one whenever
'  ACL_REVISION is updated

Private Const ACL_REVISION1 = (1)
Private Const ACL_REVISION2 = (2)

Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type

' typedef ACL *PACL;

'  end_ntddk

'   The structure of an ACE is a common ace header followed by ace type
'   specific data.  Pictorally the structure of the common ace header is
'   as follows:

'   AceType denotes the type of the ace, there are some predefined ace
'   types
'
'   AceSize is the size, in bytes, of ace.
'
'   AceFlags are the Ace flags for audit and inheritance, defined Integerly.

Type ACE_HEADER
        AceType As Byte
        AceFlags As Byte
        AceSize As Long
End Type
'
'   The following are the predefined ace types that go into the AceType
'   field of an Ace header.

Private Const ACCESS_ALLOWED_ACE_TYPE = &H0
Private Const ACCESS_DENIED_ACE_TYPE = &H1
Private Const SYSTEM_AUDIT_ACE_TYPE = &H2
Private Const SYSTEM_ALARM_ACE_TYPE = &H3

'   The following are the inherit flags that go into the AceFlags field
'   of an Ace header.

Private Const OBJECT_INHERIT_ACE = &H1
Private Const CONTAINER_INHERIT_ACE = &H2
Private Const NO_PROPAGATE_INHERIT_ACE = &H4
Private Const INHERIT_ONLY_ACE = &H8
Private Const VALID_INHERIT_FLAGS = &HF


'   The following are the currently defined ACE flags that go into the
'   AceFlags field of an ACE header.  Each ACE type has its own set of
'   AceFlags.
'
'   SUCCESSFUL_ACCESS_ACE_FLAG - used only with system audit and alarm ACE
'   types to indicate that a message is generated for successful accesses.
'
'   FAILED_ACCESS_ACE_FLAG - used only with system audit and alarm ACE types
'   to indicate that a message is generated for failed accesses.

'   SYSTEM_AUDIT and SYSTEM_ALARM AceFlags
'
'   These control the signaling of audit and alarms for success or failure.

Private Const SUCCESSFUL_ACCESS_ACE_FLAG = &H40
Private Const FAILED_ACCESS_ACE_FLAG = &H80


'
'   We'll define the structure of the predefined ACE types.  Pictorally
'   the structure of the predefined ACE's is as follows:

'   Mask is the access mask associated with the ACE.  This is either the
'   access allowed, access denied, audit, or alarm mask.
'
'   Sid is the Sid associated with the ACE.
'
'   The following are the four predefined ACE types.
'   Examine the AceType field in the Header to determine
'   which structure is appropriate to use for casting.


Type ACCESS_ALLOWED_ACE
        Header As ACE_HEADER
        Mask As Long
        SidStart As Long
End Type

Type ACCESS_DENIED_ACE
        Header As ACE_HEADER
        Mask As Long
        SidStart As Long
End Type


Type SYSTEM_AUDIT_ACE
        Header As ACE_HEADER
        Mask As Long
        SidStart As Long
End Type

Type SYSTEM_ALARM_ACE
        Header As ACE_HEADER
        Mask As Long
        SidStart As Long
End Type

'   The following declarations are used for setting and querying information
'   about and ACL.  First are the various information classes available to
'   the user.
'

Private Const AclRevisionInformation = 1
Private Const AclSizeInformation = 2
'
'   This record is returned/sent if the user is requesting/setting the
'   AclRevisionInformation
'

Type ACL_REVISION_INFORMATION
        AclRevision As Long
End Type

'
'   This record is returned if the user is requesting AclSizeInformation
'

Type ACL_SIZE_INFORMATION
        AceCount As Long
        AclBytesInUse As Long
        AclBytesFree As Long
End Type

' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                              SECURITY_DESCRIPTOR                    //
'                                                                     //
' //////////////////////////////////////////////////////////////////////
'
'   Define the Security Descriptor and related data types.
'   This is an opaque data structure.
'

'  begin_ntddk begin_ntifs
'
'  Current security descriptor revision value
'

Private Const SECURITY_DESCRIPTOR_REVISION = (1)
Private Const SECURITY_DESCRIPTOR_REVISION1 = (1)

'  end_ntddk

'
'  Minimum length, in bytes, needed to build a security descriptor
'  (NOTE: This must manually be kept consistent with the)
'  (sizeof(SECURITY_DESCRIPTOR)                         )
'

Private Const SECURITY_DESCRIPTOR_MIN_LENGTH = (20)



Private Const SE_OWNER_DEFAULTED = &H1
Private Const SE_GROUP_DEFAULTED = &H2
Private Const SE_DACL_PRESENT = &H4
Private Const SE_DACL_DEFAULTED = &H8
Private Const SE_SACL_PRESENT = &H10
Private Const SE_SACL_DEFAULTED = &H20
Private Const SE_SELF_RELATIVE = &H8000

'
'   Where:
'
'       SE_OWNER_DEFAULTED - This boolean flag, when set, indicates that the
'           SID pointed to by the Owner field was provided by a
'           defaulting mechanism rather than explicitly provided by the
'           original provider of the security descriptor.  This may
'           affect the treatment of the SID with respect to inheritence
'           of an owner.
'
'       SE_GROUP_DEFAULTED - This boolean flag, when set, indicates that the
'           SID in the Group field was provided by a defaulting mechanism
'           rather than explicitly provided by the original provider of
'           the security descriptor.  This may affect the treatment of
'           the SID with respect to inheritence of a primary group.
'
'       SE_DACL_PRESENT - This boolean flag, when set, indicates that the
'           security descriptor contains a discretionary ACL.  If this
'           flag is set and the Dacl field of the SECURITY_DESCRIPTOR is
'           null, then a null ACL is explicitly being specified.
'
'       SE_DACL_DEFAULTED - This boolean flag, when set, indicates that the
'           ACL pointed to by the Dacl field was provided by a defaulting
'           mechanism rather than explicitly provided by the original
'           provider of the security descriptor.  This may affect the
'           treatment of the ACL with respect to inheritence of an ACL.
'           This flag is ignored if the DaclPresent flag is not set.
'
'       SE_SACL_PRESENT - This boolean flag, when set,  indicates that the
'           security descriptor contains a system ACL pointed to by the
'           Sacl field.  If this flag is set and the Sacl field of the
'           SECURITY_DESCRIPTOR is null, then an empty (but present)
'           ACL is being specified.
'
'       SE_SACL_DEFAULTED - This boolean flag, when set, indicates that the
'           ACL pointed to by the Sacl field was provided by a defaulting
'           mechanism rather than explicitly provided by the original
'           provider of the security descriptor.  This may affect the
'           treatment of the ACL with respect to inheritence of an ACL.
'           This flag is ignored if the SaclPresent flag is not set.
'
'       SE_SELF_RELATIVE - This boolean flag, when set, indicates that the
'           security descriptor is in self-relative form.  In this form,
'           all fields of the security descriptor are contiguous in memory
'           and all pointer fields are expressed as offsets from the
'           beginning of the security descriptor.  This form is useful
'           for treating security descriptors as opaque data structures
'           for transmission in communication protocol or for storage on
'           secondary media.
'
'
'
'  In general, this data structure should be treated opaquely to ensure future
'  compatibility.
'
'

Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As ACL
        Dacl As ACL
End Type


'  Where:
'
'      Revision - Contains the revision level of the security
'          descriptor.  This allows this structure to be passed between
'          systems or stored on disk even though it is expected to
'          change in the future.
'
'      Control - A set of flags which qualify the meaning of the
'          security descriptor or individual fields of the security
'          descriptor.
'
'      Owner - is a pointer to an SID representing an object's owner.
'          If this field is null, then no owner SID is present in the
'          security descriptor.  If the security descriptor is in
'          self-relative form, then this field contains an offset to
'          the SID, rather than a pointer.
'
'      Group - is a pointer to an SID representing an object's primary
'          group.  If this field is null, then no primary group SID is
'          present in the security descriptor.  If the security descriptor
'          is in self-relative form, then this field contains an offset to
'          the SID, rather than a pointer.
'
'      Sacl - is a pointer to a system ACL.  This field value is only
'          valid if the DaclPresent control flag is set.  If the
'          SaclPresent flag is set and this field is null, then a null
'          ACL  is specified.  If the security descriptor is in
'          self-relative form, then this field contains an offset to
'          the ACL, rather than a pointer.
'
'      Dacl - is a pointer to a discretionary ACL.  This field value is
'          only valid if the DaclPresent control flag is set.  If the
'          DaclPresent flag is set and this field is null, then a null
'          ACL (unconditionally granting access) is specified.  If the
'          security descriptor is in self-relative form, then this field
'          contains an offset to the ACL, rather than a pointer.
'



' //////////////////////////////////////////////////////////////////////
'                                                                     //
'                Privilege Related Data Structures                    //
'                                                                     //
' //////////////////////////////////////////////////////////////////////


'  Privilege attributes
'

Private Const SE_PRIVILEGE_ENABLED_BY_DEFAULT = &H1
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const SE_PRIVILEGE_USED_FOR_ACCESS = &H80000000

'
'  Privilege Set Control flags
'

Private Const PRIVILEGE_SET_ALL_NECESSARY = (1)

'
'   Privilege Set - This is defined for a privilege set of one.
'                   If more than one privilege is needed, then this structure
'                   will need to be allocated with more space.
'
'   Note: don't change this structure without fixing the INITIAL_PRIVILEGE_SET
'   structure (defined in se.h)
'

Type PRIVILEGE_SET
        PrivilegeCount As Long
        Control As Long
        Privilege(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type


'//////////////////////////////////////////////////////////////////////
'                                                                     //
'                NT Defined Privileges                                //
'                                                                     //
' //////////////////////////////////////////////////////////////////////

Private Const SE_CREATE_TOKEN_NAME = "SeCreateTokenPrivilege"
Private Const SE_ASSIGNPRIMARYTOKEN_NAME = "SeAssignPrimaryTokenPrivilege"
Private Const SE_LOCK_MEMORY_NAME = "SeLockMemoryPrivilege"
Private Const SE_INCREASE_QUOTA_NAME = "SeIncreaseQuotaPrivilege"
Private Const SE_UNSOLICITED_INPUT_NAME = "SeUnsolicitedInputPrivilege"
Private Const SE_MACHINE_ACCOUNT_NAME = "SeMachineAccountPrivilege"
Private Const SE_TCB_NAME = "SeTcbPrivilege"
Private Const SE_SECURITY_NAME = "SeSecurityPrivilege"
Private Const SE_TAKE_OWNERSHIP_NAME = "SeTakeOwnershipPrivilege"
Private Const SE_LOAD_DRIVER_NAME = "SeLoadDriverPrivilege"
Private Const SE_SYSTEM_PROFILE_NAME = "SeSystemProfilePrivilege"
Private Const SE_SYSTEMTIME_NAME = "SeSystemtimePrivilege"
Private Const SE_PROF_SINGLE_PROCESS_NAME = "SeProfileSingleProcessPrivilege"
Private Const SE_INC_BASE_PRIORITY_NAME = "SeIncreaseBasePriorityPrivilege"
Private Const SE_CREATE_PAGEFILE_NAME = "SeCreatePagefilePrivilege"
Private Const SE_CREATE_PERMANENT_NAME = "SeCreatePermanentPrivilege"
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"
Private Const SE_RESTORE_NAME = "SeRestorePrivilege"
Private Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Private Const SE_DEBUG_NAME = "SeDebugPrivilege"
Private Const SE_AUDIT_NAME = "SeAuditPrivilege"
Private Const SE_SYSTEM_ENVIRONMENT_NAME = "SeSystemEnvironmentPrivilege"
Private Const SE_CHANGE_NOTIFY_NAME = "SeChangeNotifyPrivilege"
Private Const SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege"


' //////////////////////////////////////////////////////////////////
'                                                                 //
'            Security Quality Of Service                          //
'                                                                 //
'                                                                 //
' //////////////////////////////////////////////////////////////////

'  begin_ntddk begin_nthal begin_ntifs
'
'  Impersonation Level
'
'  Impersonation level is represented by a pair of bits in Windows.
'  If a new impersonation level is added or lowest value is changed from
'  0 to something else, fix the Windows CreateFile call.
'

Private Const SecurityAnonymous = 1
Private Const SecurityIdentification = 2

'//////////////////////////////////////////////////////////////////////
'                                                                     //
'                Registry API Constants                                //
'                                                                     //
' //////////////////////////////////////////////////////////////////////

' Reg Data Types...
Private Const REG_NONE = 0                       ' No value type
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Private Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Private Const REG_LINK = 6                       ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Private Const REG_RESOURCE_LIST = 8              ' Resource list in the resource map
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Private Const REG_CREATED_NEW_KEY = &H1                      ' New Registry Key created
Private Const REG_OPENED_EXISTING_KEY = &H2                      ' Existing Key opened
Private Const REG_WHOLE_HIVE_VOLATILE = &H1                      ' Restore whole hive volatile
Private Const REG_REFRESH_HIVE = &H2                      ' Unwind changes to last flush
Private Const REG_NOTIFY_CHANGE_NAME = &H1                      ' Create or delete (child)
Private Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Private Const REG_NOTIFY_CHANGE_LAST_SET = &H4                      ' Time stamp
Private Const REG_NOTIFY_CHANGE_SECURITY = &H8
Private Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Private Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)

' Reg Create Type Values...
Private Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Private Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
Private Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore

' Reg Key Security Options
Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

' end winnt.txt

' Debug APIs
Private Const EXCEPTION_DEBUG_EVENT = 1
Private Const CREATE_THREAD_DEBUG_EVENT = 2
Private Const CREATE_PROCESS_DEBUG_EVENT = 3
Private Const EXIT_THREAD_DEBUG_EVENT = 4
Private Const EXIT_PROCESS_DEBUG_EVENT = 5
Private Const LOAD_DLL_DEBUG_EVENT = 6
Private Const UNLOAD_DLL_DEBUG_EVENT = 7
Private Const OUTPUT_DEBUG_STRING_EVENT = 8
Private Const RIP_EVENT = 9

Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15

Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long    ' Pointer to an EXCEPTION_RECORD structure
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type

Type EXCEPTION_DEBUG_INFO
        pExceptionRecord As EXCEPTION_RECORD
        dwFirstChance As Long
End Type

Type CREATE_THREAD_DEBUG_INFO
        hThread As Long
        lpThreadLocalBase As Long
        lpStartAddress As Long
End Type

Type CREATE_PROCESS_DEBUG_INFO
        hFile As Long
        hProcess As Long
        hThread As Long
        lpBaseOfImage As Long
        dwDebugInfoFileOffset As Long
        nDebugInfoSize As Long
        lpThreadLocalBase As Long
        lpStartAddress As Long
        lpImageName As Long
        fUnicode As Integer
End Type

Type EXIT_THREAD_DEBUG_INFO
        dwExitCode As Long
End Type

Type EXIT_PROCESS_DEBUG_INFO
        dwExitCode As Long
End Type

Type LOAD_DLL_DEBUG_INFO
        hFile As Long
        lpBaseOfDll As Long
        dwDebugInfoFileOffset As Long
        nDebugInfoSize As Long
        lpImageName As Long
        fUnicode As Integer
End Type

Type UNLOAD_DLL_DEBUG_INFO
        lpBaseOfDll As Long
End Type

Type OUTPUT_DEBUG_STRING_INFO
        lpDebugStringData As String
        fUnicode As Integer
        nDebugStringLength As Integer
End Type

Type RIP_INFO
        dwError As Long
        dwType As Long
End Type

' GetDriveType return values
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

Private Const FILE_TYPE_UNKNOWN = &H0
Private Const FILE_TYPE_DISK = &H1
Private Const FILE_TYPE_CHAR = &H2
Private Const FILE_TYPE_PIPE = &H3
Private Const FILE_TYPE_REMOTE = &H8000

Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&

Private Const NOPARITY = 0
Private Const ODDPARITY = 1
Private Const EVENPARITY = 2
Private Const MARKPARITY = 3
Private Const SPACEPARITY = 4

Private Const ONESTOPBIT = 0
Private Const ONE5STOPBITS = 1
Private Const TWOSTOPBITS = 2

Private Const IGNORE = 0 '  Ignore signal
Private Const INFINITE = &HFFFF      '  Infinite timeout

' Comm Baud Rate indices
Private Const CBR_110 = 110
Private Const CBR_300 = 300
Private Const CBR_600 = 600
Private Const CBR_1200 = 1200
Private Const CBR_2400 = 2400
Private Const CBR_4800 = 4800
Private Const CBR_9600 = 9600
Private Const CBR_14400 = 14400
Private Const CBR_19200 = 19200
Private Const CBR_38400 = 38400
Private Const CBR_56000 = 56000
Private Const CBR_57600 = 57600
Private Const CBR_115200 = 115200
Private Const CBR_128000 = 128000
Private Const CBR_256000 = 256000

' Error Flags
Private Const CE_RXOVER = &H1                '  Receive Queue overflow
Private Const CE_OVERRUN = &H2               '  Receive Overrun Error
Private Const CE_RXPARITY = &H4              '  Receive Parity Error
Private Const CE_FRAME = &H8                 '  Receive Framing error
Private Const CE_BREAK = &H10                '  Break Detected
Private Const CE_TXFULL = &H100              '  TX Queue is full
Private Const CE_PTO = &H200                 '  LPTx Timeout
Private Const CE_IOE = &H400                 '  LPTx I/O Error
Private Const CE_DNS = &H800                 '  LPTx Device not selected
Private Const CE_OOP = &H1000                '  LPTx Out-Of-Paper
Private Const CE_MODE = &H8000               '  Requested mode unsupported

Private Const IE_BADID = (-1)                '  Invalid or unsupported id
Private Const IE_OPEN = (-2)                 '  Device Already Open
Private Const IE_NOPEN = (-3)                '  Device Not Open
Private Const IE_MEMORY = (-4)               '  Unable to allocate queues
Private Const IE_DEFAULT = (-5)              '  Error in default parameters
Private Const IE_HARDWARE = (-10)            '  Hardware Not Present
Private Const IE_BYTESIZE = (-11)            '  Illegal Byte Size
Private Const IE_BAUDRATE = (-12)            '  Unsupported BaudRate

' Events
Private Const EV_RXCHAR = &H1                '  Any Character received
Private Const EV_RXFLAG = &H2                '  Received certain character
Private Const EV_TXEMPTY = &H4               '  Transmitt Queue Empty
Private Const EV_CTS = &H8                   '  CTS changed state
Private Const EV_DSR = &H10                  '  DSR changed state
Private Const EV_RLSD = &H20                 '  RLSD changed state
Private Const EV_BREAK = &H40                '  BREAK received
Private Const EV_ERR = &H80                  '  Line status error occurred
Private Const EV_RING = &H100                '  Ring signal detected
Private Const EV_PERR = &H200                '  Printer error occured
Private Const EV_RX80FULL = &H400            '  Receive buffer is 80 percent full
Private Const EV_EVENT1 = &H800              '  Provider specific event 1
Private Const EV_EVENT2 = &H1000             '  Provider specific event 2

' Escape Functions
Private Const SETXOFF = 1  '  Simulate XOFF received
Private Const SETXON = 2 '  Simulate XON received
Private Const SETRTS = 3 '  Set RTS high
Private Const CLRRTS = 4 '  Set RTS low
Private Const SETDTR = 5 '  Set DTR high
Private Const CLRDTR = 6 '  Set DTR low
Private Const RESETDEV = 7       '  Reset device if possible
Private Const SETBREAK = 8  'Set the device break line
Private Const CLRBREAK = 9 ' Clear the device break line

'  PURGE function flags.
Private Const PURGE_TXABORT = &H1     '  Kill the pending/current writes to the comm port.
Private Const PURGE_RXABORT = &H2     '  Kill the pending/current reads to the comm port.
Private Const PURGE_TXCLEAR = &H4     '  Kill the transmit queue if there.
Private Const PURGE_RXCLEAR = &H8     '  Kill the typeahead buffer if there.

Private Const LPTx = &H80        '  Set if ID is for LPT device

'  Modem Status Flags
Private Const MS_CTS_ON = &H10&
Private Const MS_DSR_ON = &H20&
Private Const MS_RING_ON = &H40&
Private Const MS_RLSD_ON = &H80&

' WaitSoundState() Constants
Private Const S_QUEUEEMPTY = 0
Private Const S_THRESHOLD = 1
Private Const S_ALLTHRESHOLD = 2

' Accent Modes
Private Const S_NORMAL = 0
Private Const S_LEGATO = 1
Private Const S_STACCATO = 2

' SetSoundNoise() Sources
Private Const S_PERIOD512 = 0    '  Freq = N/512 high pitch, less coarse hiss
Private Const S_PERIOD1024 = 1   '  Freq = N/1024
Private Const S_PERIOD2048 = 2   '  Freq = N/2048 low pitch, more coarse hiss
Private Const S_PERIODVOICE = 3  '  Source is frequency from voice channel (3)
Private Const S_WHITE512 = 4     '  Freq = N/512 high pitch, less coarse hiss
Private Const S_WHITE1024 = 5    '  Freq = N/1024
Private Const S_WHITE2048 = 6    '  Freq = N/2048 low pitch, more coarse hiss
Private Const S_WHITEVOICE = 7   '  Source is frequency from voice channel (3)

Private Const S_SERDVNA = (-1)   '  Device not available
Private Const S_SEROFM = (-2)    '  Out of memory
Private Const S_SERMACT = (-3)   '  Music active
Private Const S_SERQFUL = (-4)   '  Queue full
Private Const S_SERBDNT = (-5)   '  Invalid note
Private Const S_SERDLN = (-6)    '  Invalid note length
Private Const S_SERDCC = (-7)    '  Invalid note count
Private Const S_SERDTP = (-8)    '  Invalid tempo
Private Const S_SERDVL = (-9)    '  Invalid volume
Private Const S_SERDMD = (-10)   '  Invalid mode
Private Const S_SERDSH = (-11)   '  Invalid shape
Private Const S_SERDPT = (-12)   '  Invalid pitch
Private Const S_SERDFQ = (-13)   '  Invalid frequency
Private Const S_SERDDR = (-14)   '  Invalid duration
Private Const S_SERDSR = (-15)   '  Invalid source
Private Const S_SERDST = (-16)   '  Invalid state

Private Const NMPWAIT_WAIT_FOREVER = &HFFFF
Private Const NMPWAIT_NOWAIT = &H1
Private Const NMPWAIT_USE_DEFAULT_WAIT = &H0
Private Const FS_CASE_IS_PRESERVED = FILE_CASE_PRESERVED_NAMES
Private Const FS_CASE_SENSITIVE = FILE_CASE_SENSITIVE_SEARCH
Private Const FS_UNICODE_STORED_ON_DISK = FILE_UNICODE_ON_DISK
Private Const FS_PERSISTENT_ACLS = FILE_PERSISTENT_ACLS

Private Const SECTION_QUERY = &H1
Private Const SECTION_MAP_WRITE = &H2
Private Const SECTION_MAP_READ = &H4
Private Const SECTION_MAP_EXECUTE = &H8
Private Const SECTION_EXTEND_SIZE = &H10
Private Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE

Private Const FILE_MAP_COPY = SECTION_QUERY
Private Const FILE_MAP_WRITE = SECTION_MAP_WRITE
Private Const FILE_MAP_READ = SECTION_MAP_READ
Private Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS

' OpenFile() Flags
Private Const OF_READ = &H0
Private Const OF_WRITE = &H1
Private Const OF_READWRITE = &H2
Private Const OF_SHARE_COMPAT = &H0
Private Const OF_SHARE_EXCLUSIVE = &H10
Private Const OF_SHARE_DENY_WRITE = &H20
Private Const OF_SHARE_DENY_READ = &H30
Private Const OF_SHARE_DENY_NONE = &H40
Private Const OF_PARSE = &H100
Private Const OF_DELETE = &H200
Private Const OF_VERIFY = &H400
Private Const OF_CANCEL = &H800
Private Const OF_CREATE = &H1000
Private Const OF_PROMPT = &H2000
Private Const OF_EXIST = &H4000
Private Const OF_REOPEN = &H8000

Private Const OFS_MAXPATHNAME = 128

' OpenFile() Structure
Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type

Declare Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Declare Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
Declare Function InterlockedExchange Lib "kernel32" (Target As Long, ByVal Value As Long) As Long

' Loader Routines
Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Declare Function SetProcessShutdownParameters Lib "kernel32" (ByVal dwLevel As Long, ByVal dwFlags As Long) As Long
Declare Function GetProcessShutdownParameters Lib "kernel32" (lpdwLevel As Long, lpdwFlags As Long) As Long

Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As String
Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long

Private Const DONT_RESOLVE_DLL_REFERENCES = &H1

Declare Function LoadModule Lib "kernel32" (ByVal lpModuleName As String, lpParameterBlock As Any) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Declare Sub DebugBreak Lib "kernel32" ()
Declare Function ContinueDebugEvent Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwThreadId As Long, ByVal dwContinueStatus As Long) As Long
Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long

Type CRITICAL_SECTION
    dummy As Long
End Type

Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)

Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Declare Function PulseEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Declare Function ReleaseSemaphore Lib "kernel32" (ByVal hSemaphore As Long, ByVal lReleaseCount As Long, lpPreviousCount As Long) As Long
Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Declare Function GetVersion Lib "kernel32" () As Long

Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

' GetTempFileName() Flags
'
Private Const TF_FORCEDRIVE = &H80

Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function SetHandleCount Lib "kernel32" (ByVal wNumber As Long) As Long
Declare Function GetLogicalDrives Lib "kernel32" () As Long
Declare Function LockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Declare Function UnlockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long
Declare Function LockFileEx Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long, ByVal dwReserved As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long, lpOverlapped As OVERLAPPED) As Long

Private Const LOCKFILE_FAIL_IMMEDIATELY = &H1
Private Const LOCKFILE_EXCLUSIVE_LOCK = &H2

Declare Function UnlockFileEx Lib "kernel32" (ByVal hFile As Long, ByVal dwReserved As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long, lpOverlapped As OVERLAPPED) As Long

Type BY_HANDLE_FILE_INFORMATION
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        dwVolumeSerialNumber As Long
        nFileSizeHigh As Long
        nFileSizeLow As Long
        nNumberOfLinks As Long
        nFileIndexHigh As Long
        nFileIndexLow As Long
End Type

Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function SetStdHandle Lib "kernel32" (ByVal nStdHandle As Long, ByVal nHandle As Long) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long

Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalHandle Lib "kernel32" (wMem As Any) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long

Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalFlags Lib "kernel32" (ByVal hMem As Long) As Long
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Const LNOTIFY_OUTOFMEM = 0
Private Const LNOTIFY_MOVE = 1
Private Const LNOTIFY_DISCARD = 2

Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalHandle Lib "kernel32" (wMem As Any) As Long
Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal wBytes As Long, ByVal wFlags As Long) As Long

Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalFlags Lib "kernel32" (ByVal hMem As Long) As Long

Type MEMORY_BASIC_INFORMATION
     BaseAddress As Long
     AllocationBase As Long
     AllocationProtect As Long
     RegionSize As Long
     State As Long
     Protect As Long
     lType As Long
End Type

Declare Function FlushInstructionCache Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, ByVal dwSize As Long) As Long
Declare Function VirtualAlloc Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Declare Function VirtualFree Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Declare Function VirtualQuery Lib "kernel32" (lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Declare Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Declare Function GetLastError Lib "kernel32" () As Long
Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Private Const SLE_ERROR = &H1
Private Const SLE_MINORERROR = &H2
Private Const SLE_WARNING = &H3

Declare Sub SetLastErrorEx Lib "user32" (ByVal dwErrCode As Long, ByVal dwType As Long)
Declare Function GetOverlappedResult Lib "kernel32" (ByVal hFile As Long, lpOverlapped As OVERLAPPED, lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long

Private Const SEM_FAILCRITICALERRORS = &H1
Private Const SEM_NOGPFAULTERRORBOX = &H2
Private Const SEM_NOOPENFILEERRORBOX = &H8000

Declare Sub SetDebugErrorLevel Lib "user32" (ByVal dwLevel As Long)
Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long

Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function GetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
Declare Function SetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long

Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long) As Long
Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long
Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long

' Predefined Resource Types
Private Const RT_CURSOR = 1&
Private Const RT_BITMAP = 2&
Private Const RT_ICON = 3&
Private Const RT_MENU = 4&
Private Const RT_DIALOG = 5&
Private Const RT_STRING = 6&
Private Const RT_FONTDIR = 7&
Private Const RT_FONT = 8&
Private Const RT_ACCELERATOR = 9&
Private Const RT_RCDATA = 10&

Declare Function InitAtomTable Lib "kernel32" (ByVal nSize As Long) As Long
Declare Function AddAtom Lib "kernel32" Alias "AddAtomA" (ByVal lpString As String) As Integer
Declare Function DeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Declare Function FindAtom Lib "kernel32" Alias "FindAtomA" (ByVal lpString As String) As Integer
Declare Function GetAtomName Lib "kernel32" Alias "GetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Declare Function GlobalFindAtom Lib "kernel32" Alias "GlobalFindAtomA" (ByVal lpString As String) As Integer
Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

' User Profile Routines
' NOTE: The lpKeyName argument for GetProfileString, WriteProfileString,
'       GetPrivateProfileString, and WritePrivateProfileString can be either
'       a string or NULL.  This is why the argument is defined as "As Any".
'          For example, to pass a string specify   ByVal "wallpaper"
'          To pass NULL specify                    ByVal 0&
'       You can also pass NULL for the lpString argument for WriteProfileString
'       and WritePrivateProfileString
Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long

Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Declare Function GetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function WriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Declare Function GetCurrentDirectory Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

Private Const DDD_RAW_TARGET_PATH = &H1
Private Const DDD_REMOVE_DEFINITION = &H2
Private Const DDD_EXACT_MATCH_ON_REMOVE = &H4

Private Const MAX_PATH = 260

Declare Function DefineDosDevice Lib "kernel32" Alias "DefineDosDeviceA" (ByVal dwFlags As Long, ByVal lpDeviceName As String, ByVal lpTargetPath As String) As Long
Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long

Private Const MOVEFILE_REPLACE_EXISTING = &H1
Private Const MOVEFILE_COPY_ALLOWED = &H2
Private Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4

Type EVENTLOGRECORD
     Length As Long     '  Length of full record
     Reserved As Long     '  Used by the service
     RecordNumber As Long     '  Absolute record number
     TimeGenerated As Long     '  Seconds since 1-1-1970
     TimeWritten As Long     'Seconds since 1-1-1970
     EventID As Long
     EventType As Integer
     NumStrings As Integer
     EventCategory As Integer
     ReservedFlags As Integer     '  For use with paired events (auditing)
     ClosingRecordNumber As Long     'For use with paired events (auditing)
     StringOffset As Long     '  Offset from beginning of record
     UserSidLength As Long
     UserSidOffset As Long
     DataLength As Long
     DataOffset As Long     '  Offset from beginning of record
End Type

Declare Function CreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeA" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function GetNamedPipeHandleState Lib "kernel32" Alias "GetNamedPipeHandleStateA" (ByVal hNamedPipe As Long, lpState As Long, lpCurInstances As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long, ByVal lpUserName As String, ByVal nMaxUserNameSize As Long) As Long
Declare Function CallNamedPipe Lib "kernel32" Alias "CallNamedPipeA" (ByVal lpNamedPipeName As String, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, ByVal nTimeOut As Long) As Long
Declare Function WaitNamedPipe Lib "kernel32" Alias "WaitNamedPipeA" (ByVal lpNamedPipeName As String, ByVal nTimeOut As Long) As Long
Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Declare Sub SetFileApisToOEM Lib "kernel32" ()
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Declare Function ClearEventLog Lib "ADVAPI32.DLL" Alias "ClearEventLogA" (ByVal hEventLog As Long, ByVal lpBackupFileName As String) As Long
Declare Function BackupEventLog Lib "ADVAPI32.DLL" Alias "BackupEventLogA" (ByVal hEventLog As Long, ByVal lpBackupFileName As String) As Long
Declare Function CloseEventLog Lib "ADVAPI32.DLL" (ByVal hEventLog As Long) As Long
Declare Function DeregisterEventSource Lib "ADVAPI32.DLL" (ByVal hEventLog As Long) As Long
Declare Function GetNumberOfEventLogRecords Lib "ADVAPI32.DLL" (ByVal hEventLog As Long, NumberOfRecords As Long) As Long
Declare Function GetOldestEventLogRecord Lib "ADVAPI32.DLL" (ByVal hEventLog As Long, OldestRecord As Long) As Long
Declare Function OpenEventLog Lib "ADVAPI32.DLL" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Declare Function RegisterEventSource Lib "ADVAPI32.DLL" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Declare Function OpenBackupEventLog Lib "ADVAPI32.DLL" Alias "OpenBackupEventLogA" (ByVal lpUNCServerName As String, ByVal lpFileName As String) As Long
Declare Function ReadEventLog Lib "ADVAPI32.DLL" Alias "ReadEventLogA" (ByVal hEventLog As Long, ByVal dwReadFlags As Long, ByVal dwRecordOffset As Long, lpBuffer As EVENTLOGRECORD, ByVal nNumberOfBytesToRead As Long, pnBytesRead As Long, pnMinNumberOfBytesNeeded As Long) As Long
Declare Function ReportEvent Lib "ADVAPI32.DLL" Alias "ReportEventA" (ByVal hEventLog As Long, ByVal wType As Long, ByVal wCategory As Long, ByVal dwEventID As Long, lpUserSid As Any, ByVal wNumStrings As Long, ByVal dwDataSize As Long, ByVal lpStrings As Long, lpRawData As Any) As Long

' Security APIs
Private Const TokenUser = 1
Private Const TokenGroups = 2
Private Const TokenPrivileges = 3
Private Const TokenOwner = 4
Private Const TokenPrimaryGroup = 5
Private Const TokenDefaultDacl = 6
Private Const TokenSource = 7
Private Const TokenType = 8
Private Const TokenImpersonationLevel = 9
Private Const TokenStatistics = 10

Type TOKEN_GROUPS
    GroupCount As Long
    Groups(ANYSIZE_ARRAY) As SID_AND_ATTRIBUTES
End Type

Declare Function DuplicateToken Lib "ADVAPI32.DLL" (ByVal ExistingTokenHandle As Long, Impersonationlevel As Integer, DuplicateTokenHandle As Long) As Long
Declare Function GetKernelObjectSecurity Lib "ADVAPI32.DLL" (ByVal handle As Long, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Declare Function ImpersonateNamedPipeClient Lib "ADVAPI32.DLL" (ByVal hNamedPipe As Long) As Long
Declare Function ImpersonateSelf Lib "ADVAPI32.DLL" (Impersonationlevel As Integer) As Long
Declare Function RevertToSelf Lib "ADVAPI32.DLL" () As Long
Declare Function AccessCheck Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As Long, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, PrivilegeSet As PRIVILEGE_SET, PrivilegeSetLength As Long, GrantedAccess As Long, ByVal Status As Long) As Long

Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Declare Function OpenProcessToken Lib "ADVAPI32.DLL" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Declare Function OpenThreadToken Lib "ADVAPI32.DLL" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long

Declare Function GetTokenInformation Lib "ADVAPI32.DLL" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Declare Function SetTokenInformation Lib "ADVAPI32.DLL" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long) As Long
Declare Function AdjustTokenPrivileges Lib "ADVAPI32.DLL" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Declare Function AdjustTokenGroups Lib "ADVAPI32.DLL" (ByVal TokenHandle As Long, ByVal ResetToDefault As Long, NewState As TOKEN_GROUPS, ByVal BufferLength As Long, PreviousState As TOKEN_GROUPS, ReturnLength As Long) As Long
Declare Function PrivilegeCheck Lib "ADVAPI32.DLL" (ByVal ClientToken As Long, RequiredPrivileges As PRIVILEGE_SET, ByVal pfResult As Long) As Long
Declare Function AccessCheckAndAuditAlarm Lib "ADVAPI32.DLL" Alias "AccessCheckAndAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, SecurityDescriptor As SECURITY_DESCRIPTOR, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, ByVal ObjectCreation As Long, GrantedAccess As Long, ByVal AccessStatus As Long, ByVal pfGenerateOnClose As Long) As Long
Declare Function ObjectOpenAuditAlarm Lib "kernel32" Alias "ObjectOpenAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As Long, ByVal DesiredAccess As Long, ByVal GrantedAccess As Long, Privileges As PRIVILEGE_SET, ByVal ObjectCreation As Long, ByVal AccessGranted As Long, ByVal GenerateOnClose As Long) As Long
Declare Function ObjectPrivilegeAuditAlarm Lib "ADVAPI32.DLL" Alias "ObjectPrivilegeAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ClientToken As Long, ByVal DesiredAccess As Long, Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) As Long
Declare Function ObjectCloseAuditAlarm Lib "ADVAPI32.DLL" Alias "ObjectCloseAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal GenerateOnClose As Long) As Long
Declare Function PrivilegedServiceAuditAlarm Lib "ADVAPI32.DLL" Alias "PrivilegedServiceAuditAlarmA" (ByVal SubsystemName As String, ByVal ServiceName As String, ByVal ClientToken As Long, Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) As Long
Declare Function IsValidSid Lib "ADVAPI32.DLL" (pSid As Any) As Long
Declare Function EqualSid Lib "ADVAPI32.DLL" (pSid1 As Any, pSid2 As Any) As Long
Declare Function EqualPrefixSid Lib "ADVAPI32.DLL" (pSid1 As Any, pSid2 As Any) As Long
Declare Function GetSidLengthRequired Lib "ADVAPI32.DLL" (ByVal nSubAuthorityCount As Byte) As Long
Declare Function AllocateAndInitializeSid Lib "ADVAPI32.DLL" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Declare Sub FreeSid Lib "ADVAPI32.DLL" (pSid As Any)
Declare Function InitializeSid Lib "ADVAPI32.DLL" (Sid As Any, pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte) As Long
Declare Function GetSidIdentifierAuthority Lib "ADVAPI32.DLL" (pSid As Any) As SID_IDENTIFIER_AUTHORITY
Declare Function GetSidSubAuthority Lib "ADVAPI32.DLL" (pSid As Any, ByVal nSubAuthority As Long) As Long
Declare Function GetSidSubAuthorityCount Lib "ADVAPI32.DLL" (pSid As Any) As Byte
Declare Function GetLengthSid Lib "ADVAPI32.DLL" (pSid As Any) As Long
Declare Function CopySid Lib "ADVAPI32.DLL" (ByVal nDestinationSidLength As Long, pDestinationSid As Any, pSourceSid As Any) As Long
Declare Function AreAllAccessesGranted Lib "ADVAPI32.DLL" (ByVal GrantedAccess As Long, ByVal DesiredAccess As Long) As Long
Declare Function AreAnyAccessesGranted Lib "ADVAPI32.DLL" (ByVal GrantedAccess As Long, ByVal DesiredAccess As Long) As Long
Declare Sub MapGenericMask Lib "ADVAPI32.DLL" (AccessMask As Long, GenericMapping As GENERIC_MAPPING)
Declare Function IsValidAcl Lib "ADVAPI32.DLL" (pAcl As ACL) As Long
Declare Function InitializeAcl Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
Declare Function GetAclInformation Lib "ADVAPI32.DLL" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
Declare Function SetAclInformation Lib "ADVAPI32.DLL" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
Declare Function AddAce Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwStartingAceIndex As Long, pAceList As Any, ByVal nAceListLength As Long) As Long
Declare Function DeleteAce Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal dwAceIndex As Long) As Long
Declare Function GetAce Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal dwAceIndex As Long, pAce As Any) As Long
Declare Function AddAccessAllowedAce Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
Declare Function AddAccessDeniedAce Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
Declare Function AddAuditAccessAce Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwAccessMask As Long, pSid As Any, ByVal bAuditSuccess As Long, ByVal bAuditFailure As Long) As Long
Declare Function FindFirstFreeAce Lib "ADVAPI32.DLL" (pAcl As ACL, pAce As Long) As Long
Declare Function InitializeSecurityDescriptor Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Long) As Long
Declare Function IsValidSecurityDescriptor Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function GetSecurityDescriptorLength Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long

Declare Function GetSecurityDescriptorControl Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pControl As Integer, lpdwRevision As Long) As Long
Declare Function SetSecurityDescriptorDacl Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bDaclPresent As Long, pDacl As ACL, ByVal bDaclDefaulted As Long) As Long
Declare Function GetSecurityDescriptorDacl Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, lpbDaclPresent As Long, pDacl As ACL, lpbDaclDefaulted As Long) As Long
Declare Function SetSecurityDescriptorSacl Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bSaclPresent As Long, pSacl As ACL, ByVal bSaclDefaulted As Long) As Long
Declare Function GetSecurityDescriptorSacl Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal lpbSaclPresent As Long, pSacl As ACL, ByVal lpbSaclDefaulted As Long) As Long
Declare Function SetSecurityDescriptorOwner Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal bOwnerDefaulted As Long) As Long
Declare Function GetSecurityDescriptorOwner Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal lpbOwnerDefaulted As Long) As Long
Declare Function SetSecurityDescriptorGroup Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pGroup As Any, ByVal bGroupDefaulted As Long) As Long
Declare Function GetSecurityDescriptorGroup Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pGroup As Any, ByVal lpbGroupDefaulted As Long) As Long
Declare Function CreatePrivateObjectSecurity Lib "ADVAPI32.DLL" (ParentDescriptor As SECURITY_DESCRIPTOR, CreatorDescriptor As SECURITY_DESCRIPTOR, NewDescriptor As SECURITY_DESCRIPTOR, ByVal IsDirectoryObject As Long, ByVal Token As Long, GenericMapping As GENERIC_MAPPING) As Long
Declare Function SetPrivateObjectSecurity Lib "ADVAPI32.DLL" (ByVal SecurityInformation As Long, ModificationDescriptor As SECURITY_DESCRIPTOR, ObjectsSecurityDescriptor As SECURITY_DESCRIPTOR, GenericMapping As GENERIC_MAPPING, ByVal Token As Long) As Long
Declare Function GetPrivateObjectSecurity Lib "ADVAPI32.DLL" (ObjectDescriptor As SECURITY_DESCRIPTOR, ByVal SecurityInformation As Long, ResultantDescriptor As SECURITY_DESCRIPTOR, ByVal DescriptorLength As Long, ReturnLength As Long) As Long
Declare Function DestroyPrivateObjectSecurity Lib "ADVAPI32.DLL" (ObjectDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function MakeSelfRelativeSD Lib "ADVAPI32.DLL" (pAbsoluteSecurityDescriptor As SECURITY_DESCRIPTOR, pSelfRelativeSecurityDescriptor As SECURITY_DESCRIPTOR, lpdwBufferLength As Long) As Long
Declare Function MakeAbsoluteSD Lib "ADVAPI32.DLL" (pSelfRelativeSecurityDescriptor As SECURITY_DESCRIPTOR, pAbsoluteSecurityDescriptor As SECURITY_DESCRIPTOR, lpdwAbsoluteSecurityDescriptorSize As Long, pDacl As ACL, lpdwDaclSize As Long, pSacl As ACL, lpdwSaclSize As Long, pOwner As Any, lpdwOwnerSize As Long, pPrimaryGroup As Any, lpdwPrimaryGroupSize As Long) As Long
Declare Function SetFileSecurity Lib "ADVAPI32.DLL" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function GetFileSecurity Lib "ADVAPI32.DLL" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Declare Function SetKernelObjectSecurity Lib "ADVAPI32.DLL" (ByVal handle As Long, ByVal SecurityInformation As Long, SecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function FindFirstChangeNotification Lib "kernel32" Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As Long
Declare Function FindNextChangeNotification Lib "kernel32" (ByVal hChangeHandle As Long) As Long
Declare Function FindCloseChangeNotification Lib "kernel32" (ByVal hChangeHandle As Long) As Long
Declare Function VirtualLock Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long) As Long
Declare Function VirtualUnlock Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long) As Long
Declare Function MapViewOfFileEx Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long, lpBaseAddress As Any) As Long
Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long

Type CONTEXT
    FltF0 As Double
    FltF1 As Double
    FltF2 As Double
    FltF3 As Double
    FltF4 As Double
    FltF5 As Double
    FltF6 As Double
    FltF7 As Double
    FltF8 As Double
    FltF9 As Double
    FltF10 As Double
    FltF11 As Double
    FltF12 As Double
    FltF13 As Double
    FltF14 As Double
    FltF15 As Double
    FltF16 As Double
    FltF17 As Double
    FltF18 As Double
    FltF19 As Double
    FltF20 As Double
    FltF21 As Double
    FltF22 As Double
    FltF23 As Double
    FltF24 As Double
    FltF25 As Double
    FltF26 As Double
    FltF27 As Double
    FltF28 As Double
    FltF29 As Double
    FltF30 As Double
    FltF31 As Double

    IntV0 As Double
    IntT0 As Double
    IntT1 As Double
    IntT2 As Double
    IntT3 As Double
    IntT4 As Double
    IntT5 As Double
    IntT6 As Double
    IntT7 As Double
    IntS0 As Double
    IntS1 As Double
    IntS2 As Double
    IntS3 As Double
    IntS4 As Double
    IntS5 As Double
    IntFp As Double
    IntA0 As Double
    IntA1 As Double
    IntA2 As Double
    IntA3 As Double
    IntA4 As Double
    IntA5 As Double
    IntT8 As Double
    IntT9 As Double
    IntT10 As Double
    IntT11 As Double
    IntRa As Double
    IntT12 As Double
    IntAt As Double
    IntGp As Double
    IntSp As Double
    IntZero As Double

    Fpcr As Double
    SoftFpcr As Double

    Fir As Double
    Psr As Long

    ContextFlags As Long
    Fill(4) As Long
End Type

Type EXCEPTION_POINTERS
    pExceptionRecord As EXCEPTION_RECORD
    ContextRecord As CONTEXT
End Type

Type LDT_BYTES  ' Defined for use in LDT_ENTRY Type
    BaseMid As Byte
    Flags1 As Byte
    Flags2 As Byte
    BaseHi As Byte
End Type

Type LDT_ENTRY
    LimitLow As Integer
    BaseLow As Integer
    HighWord As Long    ' Can use LDT_BYTES Type
End Type

Declare Sub FatalExit Lib "kernel32" (ByVal code As Long)
Declare Function GetEnvironmentStrings Lib "kernel32" Alias "GetEnvironmentStringsA" () As String
Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)
Declare Function UnhandledExceptionFilter Lib "kernel32" (ExceptionInfo As EXCEPTION_POINTERS) As Long

Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Declare Function GetCurrentThread Lib "kernel32" () As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Declare Function GetThreadTimes Lib "kernel32" (ByVal hThread As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Declare Function GetThreadSelectorEntry Lib "kernel32" (ByVal hThread As Long, ByVal dwSelector As Long, lpSelectorEntry As LDT_ENTRY) As Long

' COMM declarations
Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As DCB) As Long
Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Declare Function GetCommState Lib "kernel32" (ByVal nCid As Long, lpDCB As DCB) As Long
Declare Function GetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Declare Function PurgeComm Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long) As Long
Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, lpDCB As DCB) As Long
Declare Function BuildCommDCBAndTimeouts Lib "kernel32" Alias "BuildCommDCBAndTimeoutsA" (ByVal lpDef As String, lpDCB As DCB, lpCommTimeouts As COMMTIMEOUTS) As Long
Declare Function TransmitCommChar Lib "kernel32" (ByVal nCid As Long, ByVal cChar As Byte) As Long
Declare Function SetCommBreak Lib "kernel32" (ByVal nCid As Long) As Long
Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
Declare Function ClearCommBreak Lib "kernel32" (ByVal nCid As Long) As Long
Declare Function ClearCommError Lib "kernel32" (ByVal hFile As Long, lpErrors As Long, lpStat As COMSTAT) As Long
Declare Function SetupComm Lib "kernel32" (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
Declare Function EscapeCommFunction Lib "kernel32" (ByVal nCid As Long, ByVal nFunc As Long) As Long
Declare Function GetCommMask Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long) As Long
Declare Function GetCommProperties Lib "kernel32" (ByVal hFile As Long, lpCommProp As COMMPROP) As Long
Declare Function GetCommModemStatus Lib "kernel32" (ByVal hFile As Long, lpModemStat As Long) As Long
Declare Function WaitCommEvent Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As OVERLAPPED) As Long

Declare Function SetTapePosition Lib "kernel32" (ByVal hDevice As Long, ByVal dwPositionMethod As Long, ByVal dwPartition As Long, ByVal dwOffsetLow As Long, ByVal dwOffsetHigh As Long, ByVal bimmediate As Long) As Long
Declare Function GetTapePosition Lib "kernel32" (ByVal hDevice As Long, ByVal dwPositionType As Long, lpdwPartition As Long, lpdwOffsetLow As Long, lpdwOffsetHigh As Long) As Long
Declare Function PrepareTape Lib "kernel32" (ByVal hDevice As Long, ByVal dwOperation As Long, ByVal bimmediate As Long) As Long
Declare Function EraseTape Lib "kernel32" (ByVal hDevice As Long, ByVal dwEraseType As Long, ByVal bimmediate As Long) As Long
Declare Function CreateTapePartition Lib "kernel32" (ByVal hDevice As Long, ByVal dwPartitionMethod As Long, ByVal dwCount As Long, ByVal dwSize As Long) As Long
Declare Function WriteTapemark Lib "kernel32" (ByVal hDevice As Long, ByVal dwTapemarkType As Long, ByVal dwTapemarkCount As Long, ByVal bimmediate As Long) As Long
Declare Function GetTapeStatus Lib "kernel32" (ByVal hDevice As Long) As Long
Declare Function GetTapeParameters Lib "kernel32" (ByVal hDevice As Long, ByVal dwOperation As Long, lpdwSize As Long, lpTapeInformation As Any) As Long

Private Const GET_TAPE_MEDIA_INFORMATION = 0
Private Const GET_TAPE_DRIVE_INFORMATION = 1

Declare Function SetTapeParameters Lib "kernel32" (ByVal hDevice As Long, ByVal dwOperation As Long, lpTapeInformation As Any) As Long

Private Const SET_TAPE_MEDIA_INFORMATION = 0
Private Const SET_TAPE_DRIVE_INFORMATION = 1

Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SystemTime)
Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SystemTime) As Long
Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SystemTime)
Declare Function SetLocalTime Lib "kernel32" (lpSystemTime As SystemTime) As Long
Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(32) As Integer
        StandardDate As SystemTime
        StandardBias As Long
        DaylightName(32) As Integer
        DaylightDate As SystemTime
        DaylightBias As Long
End Type

Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Declare Function SetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

' Routines to convert back and forth
' between system time and file time

Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SystemTime, lpFileTime As FILETIME) As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SystemTime) As Long
Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long
Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long
Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Declare Function ConnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function DisconnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long) As Long
Declare Function SetNamedPipeHandleState Lib "kernel32" (ByVal hNamedPipe As Long, lpMode As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long) As Long
Declare Function GetNamedPipeInfo Lib "kernel32" (ByVal hNamedPipe As Long, lpFlags As Long, lpOutBufferSize As Long, lpInBufferSize As Long, lpMaxInstances As Long) As Long
Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Declare Function TransactNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, lpOverlapped As OVERLAPPED) As Long

Declare Function CreateMailslot Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function GetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Declare Function SetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, ByVal lReadTimeout As Long) As Long
Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Declare Function FlushViewOfFile Lib "kernel32" (lpBaseAddress As Any, ByVal dwNumberOfBytesToFlush As Long) As Long
Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long

Declare Function lstrcmp Lib "kernel32" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Declare Function lcreat Lib "kernel32" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Long) As Long
Declare Function llseek Lib "kernel32" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function lread Lib "kernel32" Alias "_lread" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
Declare Function lwrite Lib "kernel32" Alias "_lwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal wBytes As Long) As Long

Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Declare Function hwrite Lib "kernel32" Alias "_hwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal lBytes As Long) As Long

Declare Function TlsAlloc Lib "kernel32" () As Long

Private Const TLS_OUT_OF_INDEXES = &HFFFF

Declare Function TlsGetValue Lib "kernel32" (ByVal dwTlsIndex As Long) As Long
Declare Function TlsSetValue Lib "kernel32" (ByVal dwTlsIndex As Long, lpTlsValue As Any) As Long
Declare Function TlsFree Lib "kernel32" (ByVal dwTlsIndex As Long) As Long
Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Declare Function WaitForSingleObjectEx Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Declare Function WaitForMultipleObjectsEx Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Declare Function BackupRead Lib "kernel32" (ByVal hFile As Long, lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, lpContext As Any) As Long
Declare Function BackupSeek Lib "kernel32" (ByVal hFile As Long, ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, lpdwLowByteSeeked As Long, lpdwHighByteSeeked As Long, lpContext As Long) As Long
Declare Function BackupWrite Lib "kernel32" (ByVal hFile As Long, lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, lpContext As Long) As Long

' Stream ID type
Type WIN32_STREAM_ID
        dwStreamID As Long
        dwStreamAttributes As Long
        dwStreamSizeLow As Long
        dwStreamSizeHigh As Long
        dwStreamNameSize As Long
        cStreamName As Byte
End Type

' Stream IDs
Private Const BACKUP_DATA = &H1
Private Const BACKUP_EA_DATA = &H2
Private Const BACKUP_SECURITY_DATA = &H3
Private Const BACKUP_ALTERNATE_DATA = &H4
Private Const BACKUP_LINK = &H5

'   Stream Attributes
Private Const STREAM_MODIFIED_WHEN_READ = &H1
Private Const STREAM_CONTAINS_SECURITY = &H2

'  Dual Mode API below this line. Dual Mode Types also included.

Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESIZE = &H2
Private Const STARTF_USEPOSITION = &H4
Private Const STARTF_USECOUNTCHARS = &H8
Private Const STARTF_USEFILLATTRIBUTE = &H10
Private Const STARTF_RUNFULLSCREEN = &H20        '  ignored for non-x86 platforms
Private Const STARTF_FORCEONFEEDBACK = &H40
Private Const STARTF_FORCEOFFFEEDBACK = &H80
Private Const STARTF_USESTDHANDLES = &H100

Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Byte
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Private Const SHUTDOWN_NORETRY = &H1

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Declare Function OpenEvent Lib "kernel32" Alias "OpenEventA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function CreateSemaphore Lib "kernel32" Alias "CreateSemaphoreA" (lpSemaphoreAttributes As SECURITY_ATTRIBUTES, ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) As Long
Declare Function OpenSemaphore Lib "kernel32" Alias "OpenSemaphoreA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Declare Function IsBadReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Declare Function IsBadWritePtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Declare Function IsBadStringPtr Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As String, ByVal ucchMax As Long) As Long
Declare Function IsBadHugeReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Declare Function IsBadHugeWritePtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long

Declare Function LookupAccountSid Lib "ADVAPI32.DLL" Alias "LookupAccountSidA" (ByVal lpSystemName As String, Sid As Any, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long

Declare Function LookupAccountName Lib "ADVAPI32.DLL" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, Sid As Long, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Declare Function LookupPrivilegeValue Lib "ADVAPI32.DLL" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Declare Function LookupPrivilegeName Lib "ADVAPI32.DLL" Alias "LookupPrivilegeNameA" (ByVal lpSystemName As String, lpLuid As LARGE_INTEGER, ByVal lpName As String, cbName As Long) As Long
Declare Function LookupPrivilegeDisplayName Lib "ADVAPI32.DLL" Alias "LookupPrivilegeDisplayNameA" (ByVal lpSystemName As String, ByVal lpName As String, ByVal lpDisplayName As String, cbDisplayName As Long, lpLanguageID As Long) As Long
Declare Function AllocateLocallyUniqueId Lib "ADVAPI32.DLL" (Luid As LARGE_INTEGER) As Long

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Declare Function GetUserName Lib "ADVAPI32.DLL" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

' Performance counter API's
Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

'  Abnormal termination codes
Private Const TC_NORMAL = 0
Private Const TC_HARDERR = 1
Private Const TC_GP_TRAP = 2
Private Const TC_SIGNAL = 3


' Procedure declarations, constant definitions, and macros
' for the NLS component

' String Length Maximums
Private Const MAX_LEADBYTES = 12  '  5 ranges, 2 bytes ea., 0 term.

' MBCS and Unicode Translation Flags.

Private Const MB_PRECOMPOSED = &H1         '  use precomposed chars
Private Const MB_COMPOSITE = &H2         '  use composite chars
Private Const MB_USEGLYPHCHARS = &H4         '  use glyph chars, not ctrl chars

Private Const WC_DEFAULTCHECK = &H100       '  check for default char
Private Const WC_COMPOSITECHECK = &H200       '  convert composite to precomposed
Private Const WC_DISCARDNS = &H10        '  discard non-spacing chars
Private Const WC_SEPCHARS = &H20        '  generate separate chars
Private Const WC_DEFAULTCHAR = &H40        '  replace w/ default char

' Character Type Flags.
Private Const CT_CTYPE1 = &H1         '  ctype 1 information
Private Const CT_CTYPE2 = &H2         '  ctype 2 information
Private Const CT_CTYPE3 = &H4         '  ctype 3 information

' CType 1 Flag Bits.
Private Const C1_UPPER = &H1     '  upper case
Private Const C1_LOWER = &H2     '  lower case
Private Const C1_DIGIT = &H4     '  decimal digits
Private Const C1_SPACE = &H8     '  spacing characters
Private Const C1_PUNCT = &H10    '  punctuation characters
Private Const C1_CNTRL = &H20    '  control characters
Private Const C1_BLANK = &H40    '  blank characters
Private Const C1_XDIGIT = &H80    '  other digits
Private Const C1_ALPHA = &H100   '  any letter

' CType 2 Flag Bits.

Private Const C2_LEFTTORIGHT = &H1     '  left to right
Private Const C2_RIGHTTOLEFT = &H2     '  right to left

Private Const C2_EUROPENUMBER = &H3     '  European number, digit
Private Const C2_EUROPESEPARATOR = &H4     '  European numeric separator
Private Const C2_EUROPETERMINATOR = &H5     '  European numeric terminator
Private Const C2_ARABICNUMBER = &H6     '  Arabic number
Private Const C2_COMMONSEPARATOR = &H7     '  common numeric separator

Private Const C2_BLOCKSEPARATOR = &H8     '  block separator
Private Const C2_SEGMENTSEPARATOR = &H9     '  segment separator
Private Const C2_WHITESPACE = &HA     '  white space
Private Const C2_OTHERNEUTRAL = &HB     '  other neutrals

Private Const C2_NOTAPPLICABLE = &H0     '  no implicit directionality

' CType 3 Flag Bits.
Private Const C3_NONSPACING = &H1     '  nonspacing character
Private Const C3_DIACRITIC = &H2     '  diacritic mark
Private Const C3_VOWELMARK = &H4     '  vowel mark
Private Const C3_SYMBOL = &H8     '  symbols

Private Const C3_NOTAPPLICABLE = &H0     '  ctype 3 is not applicable

' String Flags.
Private Const NORM_IGNORECASE = &H1         '  ignore case
Private Const NORM_IGNORENONSPACE = &H2         '  ignore nonspacing chars
Private Const NORM_IGNORESYMBOLS = &H4         '  ignore symbols

' Locale Independent Mapping Flags.
Private Const MAP_FOLDCZONE = &H10        '  fold compatibility zone chars
Private Const MAP_PRECOMPOSED = &H20        '  convert to precomposed chars
Private Const MAP_COMPOSITE = &H40        '  convert to composite chars
Private Const MAP_FOLDDIGITS = &H80        '  all digits to ASCII 0-9

' Locale Dependent Mapping Flags.
Private Const LCMAP_LOWERCASE = &H100       '  lower case letters
Private Const LCMAP_UPPERCASE = &H200       '  upper case letters
Private Const LCMAP_SORTKEY = &H400       '  WC sort key (normalize)
Private Const LCMAP_BYTEREV = &H800       '  byte reversal

' Sorting Flags.
Private Const SORT_STRINGSORT = &H1000      '  use string sort method

' Code Page Default Values.
Private Const CP_ACP = 0  '  default to ANSI code page
Private Const CP_OEMCP = 1  '  default to OEM  code page

' Country Codes.

Private Const CTRY_DEFAULT = 0

Private Const CTRY_AUSTRALIA = 61  '  Australia
Private Const CTRY_AUSTRIA = 43  '  Austria
Private Const CTRY_BELGIUM = 32  '  Belgium
Private Const CTRY_BRAZIL = 55  '  Brazil
Private Const CTRY_CANADA = 2  '  Canada
Private Const CTRY_DENMARK = 45  '  Denmark
Private Const CTRY_FINLAND = 358  '  Finland
Private Const CTRY_FRANCE = 33  '  France
Private Const CTRY_GERMANY = 49  '  Germany
Private Const CTRY_ICELAND = 354  '  Iceland
Private Const CTRY_IRELAND = 353  '  Ireland
Private Const CTRY_ITALY = 39  '  Italy
Private Const CTRY_JAPAN = 81  '  Japan
Private Const CTRY_MEXICO = 52  '  Mexico
Private Const CTRY_NETHERLANDS = 31  '  Netherlands
Private Const CTRY_NEW_ZEALAND = 64  '  New Zealand
Private Const CTRY_NORWAY = 47  '  Norway
Private Const CTRY_PORTUGAL = 351  '  Portugal
Private Const CTRY_PRCHINA = 86  '  PR China
Private Const CTRY_SOUTH_KOREA = 82  '  South Korea
Private Const CTRY_SPAIN = 34  '  Spain
Private Const CTRY_SWEDEN = 46  '  Sweden
Private Const CTRY_SWITZERLAND = 41  '  Switzerland
Private Const CTRY_TAIWAN = 886  '  Taiwan
Private Const CTRY_UNITED_KINGDOM = 44  '  United Kingdom
Private Const CTRY_UNITED_STATES = 1  '  United States

' Locale Types.
' These types are used for the GetLocaleInfoW NLS API routine.

' LOCALE_NOUSEROVERRIDE is also used in GetTimeFormatW and GetDateFormatW.
Private Const LOCALE_NOUSEROVERRIDE = &H80000000  '  do not use user overrides

Private Const LOCALE_ILANGUAGE = &H1         '  language id
Private Const LOCALE_SLANGUAGE = &H2         '  localized name of language
Private Const LOCALE_SENGLANGUAGE = &H1001      '  English name of language
Private Const LOCALE_SABBREVLANGNAME = &H3         '  abbreviated language name
Private Const LOCALE_SNATIVELANGNAME = &H4         '  native name of language
Private Const LOCALE_ICOUNTRY = &H5         '  country code
Private Const LOCALE_SCOUNTRY = &H6         '  localized name of country
Private Const LOCALE_SENGCOUNTRY = &H1002      '  English name of country
Private Const LOCALE_SABBREVCTRYNAME = &H7         '  abbreviated country name
Private Const LOCALE_SNATIVECTRYNAME = &H8         '  native name of country
Private Const LOCALE_IDEFAULTLANGUAGE = &H9         '  default language id
Private Const LOCALE_IDEFAULTCOUNTRY = &HA         '  default country code
Private Const LOCALE_IDEFAULTCODEPAGE = &HB         '  default code page

Private Const LOCALE_SLIST = &HC         '  list item separator
Private Const LOCALE_IMEASURE = &HD         '  0 = metric, 1 = US

Private Const LOCALE_SDECIMAL = &HE         '  decimal separator
Private Const LOCALE_STHOUSAND = &HF         '  thousand separator
Private Const LOCALE_SGROUPING = &H10        '  digit grouping
Private Const LOCALE_IDIGITS = &H11        '  number of fractional digits
Private Const LOCALE_ILZERO = &H12        '  leading zeros for decimal
Private Const LOCALE_SNATIVEDIGITS = &H13        '  native ascii 0-9

Private Const LOCALE_SCURRENCY = &H14        '  local monetary symbol
Private Const LOCALE_SINTLSYMBOL = &H15        '  intl monetary symbol
Private Const LOCALE_SMONDECIMALSEP = &H16        '  monetary decimal separator
Private Const LOCALE_SMONTHOUSANDSEP = &H17        '  monetary thousand separator
Private Const LOCALE_SMONGROUPING = &H18        '  monetary grouping
Private Const LOCALE_ICURRDIGITS = &H19        '  # local monetary digits
Private Const LOCALE_IINTLCURRDIGITS = &H1A        '  # intl monetary digits
Private Const LOCALE_ICURRENCY = &H1B        '  positive currency mode
Private Const LOCALE_INEGCURR = &H1C        '  negative currency mode

Private Const LOCALE_SDATE = &H1D        '  date separator
Private Const LOCALE_STIME = &H1E        '  time separator
Private Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Private Const LOCALE_SLONGDATE = &H20        '  long date format string
Private Const LOCALE_STIMEFORMAT = &H1003      '  time format string
Private Const LOCALE_IDATE = &H21        '  short date format ordering
Private Const LOCALE_ILDATE = &H22        '  long date format ordering
Private Const LOCALE_ITIME = &H23        '  time format specifier
Private Const LOCALE_ICENTURY = &H24        '  century format specifier
Private Const LOCALE_ITLZERO = &H25        '  leading zeros in time field
Private Const LOCALE_IDAYLZERO = &H26        '  leading zeros in day field
Private Const LOCALE_IMONLZERO = &H27        '  leading zeros in month field
Private Const LOCALE_S1159 = &H28        '  AM designator
Private Const LOCALE_S2359 = &H29        '  PM designator

Private Const LOCALE_SDAYNAME1 = &H2A        '  long name for Monday
Private Const LOCALE_SDAYNAME2 = &H2B        '  long name for Tuesday
Private Const LOCALE_SDAYNAME3 = &H2C        '  long name for Wednesday
Private Const LOCALE_SDAYNAME4 = &H2D        '  long name for Thursday
Private Const LOCALE_SDAYNAME5 = &H2E        '  long name for Friday
Private Const LOCALE_SDAYNAME6 = &H2F        '  long name for Saturday
Private Const LOCALE_SDAYNAME7 = &H30        '  long name for Sunday
Private Const LOCALE_SABBREVDAYNAME1 = &H31        '  abbreviated name for Monday
Private Const LOCALE_SABBREVDAYNAME2 = &H32        '  abbreviated name for Tuesday
Private Const LOCALE_SABBREVDAYNAME3 = &H33        '  abbreviated name for Wednesday
Private Const LOCALE_SABBREVDAYNAME4 = &H34        '  abbreviated name for Thursday
Private Const LOCALE_SABBREVDAYNAME5 = &H35        '  abbreviated name for Friday
Private Const LOCALE_SABBREVDAYNAME6 = &H36        '  abbreviated name for Saturday
Private Const LOCALE_SABBREVDAYNAME7 = &H37        '  abbreviated name for Sunday
Private Const LOCALE_SMONTHNAME1 = &H38        '  long name for January
Private Const LOCALE_SMONTHNAME2 = &H39        '  long name for February
Private Const LOCALE_SMONTHNAME3 = &H3A        '  long name for March
Private Const LOCALE_SMONTHNAME4 = &H3B        '  long name for April
Private Const LOCALE_SMONTHNAME5 = &H3C        '  long name for May
Private Const LOCALE_SMONTHNAME6 = &H3D        '  long name for June
Private Const LOCALE_SMONTHNAME7 = &H3E        '  long name for July
Private Const LOCALE_SMONTHNAME8 = &H3F        '  long name for August
Private Const LOCALE_SMONTHNAME9 = &H40        '  long name for September
Private Const LOCALE_SMONTHNAME10 = &H41        '  long name for October
Private Const LOCALE_SMONTHNAME11 = &H42        '  long name for November
Private Const LOCALE_SMONTHNAME12 = &H43        '  long name for December
Private Const LOCALE_SABBREVMONTHNAME1 = &H44        '  abbreviated name for January
Private Const LOCALE_SABBREVMONTHNAME2 = &H45        '  abbreviated name for February
Private Const LOCALE_SABBREVMONTHNAME3 = &H46        '  abbreviated name for March
Private Const LOCALE_SABBREVMONTHNAME4 = &H47        '  abbreviated name for April
Private Const LOCALE_SABBREVMONTHNAME5 = &H48        '  abbreviated name for May
Private Const LOCALE_SABBREVMONTHNAME6 = &H49        '  abbreviated name for June
Private Const LOCALE_SABBREVMONTHNAME7 = &H4A        '  abbreviated name for July
Private Const LOCALE_SABBREVMONTHNAME8 = &H4B        '  abbreviated name for August
Private Const LOCALE_SABBREVMONTHNAME9 = &H4C        '  abbreviated name for September
Private Const LOCALE_SABBREVMONTHNAME10 = &H4D        '  abbreviated name for October
Private Const LOCALE_SABBREVMONTHNAME11 = &H4E        '  abbreviated name for November
Private Const LOCALE_SABBREVMONTHNAME12 = &H4F        '  abbreviated name for December
Private Const LOCALE_SABBREVMONTHNAME13 = &H100F

Private Const LOCALE_SPOSITIVESIGN = &H50        '  positive sign
Private Const LOCALE_SNEGATIVESIGN = &H51        '  negative sign
Private Const LOCALE_IPOSSIGNPOSN = &H52        '  positive sign position
Private Const LOCALE_INEGSIGNPOSN = &H53        '  negative sign position
Private Const LOCALE_IPOSSYMPRECEDES = &H54        '  mon sym precedes pos amt
Private Const LOCALE_IPOSSEPBYSPACE = &H55        '  mon sym sep by space from pos amt
Private Const LOCALE_INEGSYMPRECEDES = &H56        '  mon sym precedes neg amt
Private Const LOCALE_INEGSEPBYSPACE = &H57        '  mon sym sep by space from neg amt

' Time Flags for GetTimeFormatW.
Private Const TIME_NOMINUTESORSECONDS = &H1         '  do not use minutes or seconds
Private Const TIME_NOSECONDS = &H2         '  do not use seconds
Private Const TIME_NOTIMEMARKER = &H4         '  do not use time marker
Private Const TIME_FORCE24HOURFORMAT = &H8         '  always use 24 hour format

' Date Flags for GetDateFormatW.
Private Const DATE_SHORTDATE = &H1         '  use short date picture
Private Const DATE_LONGDATE = &H2         '  use long date picture


' Code Page Dependent APIs

Declare Function IsValidCodePage Lib "kernel32" (ByVal CodePage As Long) As Long
Declare Function GetACP Lib "kernel32" () As Long
Declare Function GetOEMCP Lib "kernel32" () As Long
Declare Function GetCPInfo Lib "kernel32" (ByVal CodePage As Long, lpCPInfo As CPINFO) As Long
Declare Function IsDBCSLeadByte Lib "kernel32" (ByVal bTestChar As Byte) As Long
Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long

' Locale Dependent APIs

Declare Function CompareString Lib "kernel32" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SystemTime, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SystemTime, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Declare Function SetThreadLocale Lib "kernel32" (ByVal Locale As Long) As Long
Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

' Locale Independent APIs

Declare Function GetStringTypeA Lib "kernel32" (ByVal lcid As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Long) As Long
Declare Function FoldString Lib "kernel32" Alias "FoldStringA" (ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long

' *************************************************************************
' *                                                                         *
' * winnls.h -- NLS procedure declarations, constant definitions and macros *
' *                                                                         *
' * Copyright (c) 1991-1995, Microsoft Corp. All rights reserved.           *
' *                                                                         *
' **************************************************************************/

' *  Calendar Types.
'  *
'  *  These types are used for the GetALTCalendarInfoW NLS API routine.
'  */
Private Const MAX_DEFAULTCHAR = 2
Private Const CAL_ICALINTVALUE = &H1                     '  calendar type
Private Const CAL_SCALNAME = &H2                         '  native name of calendar
Private Const CAL_IYEAROFFSETRANGE = &H3                 '  starting years of eras
Private Const CAL_SERASTRING = &H4                       '  era name for IYearOffsetRanges
Private Const CAL_SSHORTDATE = &H5                       '  Integer date format string
Private Const CAL_SLONGDATE = &H6                        '  long date format string
Private Const CAL_SDAYNAME1 = &H7                        '  native name for Monday
Private Const CAL_SDAYNAME2 = &H8                        '  native name for Tuesday
Private Const CAL_SDAYNAME3 = &H9                        '  native name for Wednesday
Private Const CAL_SDAYNAME4 = &HA                        '  native name for Thursday
Private Const CAL_SDAYNAME5 = &HB                        '  native name for Friday
Private Const CAL_SDAYNAME6 = &HC                        '  native name for Saturday
Private Const CAL_SDAYNAME7 = &HD                        '  native name for Sunday
Private Const CAL_SABBREVDAYNAME1 = &HE                  '  abbreviated name for Monday
Private Const CAL_SABBREVDAYNAME2 = &HF                  '  abbreviated name for Tuesday
Private Const CAL_SABBREVDAYNAME3 = &H10                 '  abbreviated name for Wednesday
Private Const CAL_SABBREVDAYNAME4 = &H11                 '  abbreviated name for Thursday
Private Const CAL_SABBREVDAYNAME5 = &H12                 '  abbreviated name for Friday
Private Const CAL_SABBREVDAYNAME6 = &H13                 '  abbreviated name for Saturday
Private Const CAL_SABBREVDAYNAME7 = &H14                 '  abbreviated name for Sunday
Private Const CAL_SMONTHNAME1 = &H15                     '  native name for January
Private Const CAL_SMONTHNAME2 = &H16                     '  native name for February
Private Const CAL_SMONTHNAME3 = &H17                     '  native name for March
Private Const CAL_SMONTHNAME4 = &H18                     '  native name for April
Private Const CAL_SMONTHNAME5 = &H19                     '  native name for May
Private Const CAL_SMONTHNAME6 = &H1A                     '  native name for June
Private Const CAL_SMONTHNAME7 = &H1B                     '  native name for July
Private Const CAL_SMONTHNAME8 = &H1C                     '  native name for August
Private Const CAL_SMONTHNAME9 = &H1D                     '  native name for September
Private Const CAL_SMONTHNAME10 = &H1E                    '  native name for October
Private Const CAL_SMONTHNAME11 = &H1F                    '  native name for November
Private Const CAL_SMONTHNAME12 = &H20                    '  native name for December
Private Const CAL_SMONTHNAME13 = &H21                    '  native name for 13th month (if any)
Private Const CAL_SABBREVMONTHNAME1 = &H22               '  abbreviated name for January
Private Const CAL_SABBREVMONTHNAME2 = &H23               '  abbreviated name for February
Private Const CAL_SABBREVMONTHNAME3 = &H24               '  abbreviated name for March
Private Const CAL_SABBREVMONTHNAME4 = &H25               '  abbreviated name for April
Private Const CAL_SABBREVMONTHNAME5 = &H26               '  abbreviated name for May
Private Const CAL_SABBREVMONTHNAME6 = &H27               '  abbreviated name for June
Private Const CAL_SABBREVMONTHNAME7 = &H28               '  abbreviated name for July
Private Const CAL_SABBREVMONTHNAME8 = &H29               '  abbreviated name for August
Private Const CAL_SABBREVMONTHNAME9 = &H2A               '  abbreviated name for September
Private Const CAL_SABBREVMONTHNAME10 = &H2B              '  abbreviated name for October
Private Const CAL_SABBREVMONTHNAME11 = &H2C              '  abbreviated name for November
Private Const CAL_SABBREVMONTHNAME12 = &H2D              '  abbreviated name for December
Private Const CAL_SABBREVMONTHNAME13 = &H2E              '  abbreviated name for 13th month (if any)

'
'  *  Calendar Enumeration Value.
'  */
Private Const ENUM_ALL_CALENDARS = &HFFFF                '  enumerate all calendars
'
'  *  Calendar ID Values.
'  */
Private Const CAL_GREGORIAN = 1                 '  Gregorian (localized) calendar
Private Const CAL_GREGORIAN_US = 2              '  Gregorian (U.S.) calendar
Private Const CAL_JAPAN = 3                     '  Japanese Emperor Era calendar
Private Const CAL_TAIWAN = 4                    '  RePrivate of China Era calendar
Private Const CAL_KOREA = 5                     '  Korean Tangun Era calendar

' *************************************************************************** Typedefs
' *
' * Define all types for the NLS component here.
' \***************************************************************************/
'
'  *  CP Info.
'  */

Type CPINFO
        MaxCharSize As Long                    '  max length (Byte) of a char
        DefaultChar(MAX_DEFAULTCHAR) As Byte   '  default character
        LeadByte(MAX_LEADBYTES) As Byte        '  lead byte ranges
End Type

Type NUMBERFMT
        NumDigits As Long                 '  number of decimal digits
        LeadingZero As Long '  if leading zero in decimal fields
        Grouping As Long '  group size left of decimal
        lpDecimalSep As String              '  ptr to decimal separator string
        lpThousandSep As String             '  ptr to thousand separator string
        NegativeOrder As Long '  negative number ordering
End Type
'
'  *  Currency format.
'  */

Type CURRENCYFMT
        NumDigits As Long '  number of decimal digits
        LeadingZero As Long '  if leading zero in decimal fields
        Grouping As Long '  group size left of decimal
        lpDecimalSep As String              '  ptr to decimal separator string
        lpThousandSep As String             '  ptr to thousand separator string
        NegativeOrder As Long '  negative currency ordering
        PositiveOrder As Long '  positive currency ordering
        lpCurrencySymbol As String          '  ptr to currency symbol string
End Type

Declare Function EnumTimeFormats Lib "kernel32" (ByVal lpTimeFmtEnumProc As Long, ByVal Locale As Long, ByVal dwFlags As Long) As Long
Declare Function EnumDateFormats Lib "kernel32" (ByVal lpDateFmtEnumProc As Long, ByVal Locale As Long, ByVal dwFlags As Long) As Long
Declare Function IsValidLocale Lib "kernel32" (ByVal Locale As Long, ByVal dwFlags As Long) As Long
Declare Function ConvertDefaultLocale Lib "kernel32" (ByVal Locale As Long) As Long
Declare Function GetThreadLocale Lib "kernel32" () As Long
Declare Function EnumSystemLocales Lib "kernel32" (ByVal lpLocaleEnumProc As Long, ByVal dwFlags As Long) As Long
Declare Function EnumSystemCodePages Lib "kernel32" (ByVal lpCodePageEnumProc As Long, ByVal dwFlags As Long) As Long

' The following section contains the Private data structures, data types,
' and procedures exported by the NT console subsystem.

Type COORD
        X As Integer
        Y As Integer
End Type

Type SMALL_RECT
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
End Type

Type KEY_EVENT_RECORD
        bKeyDown As Long
        wRepeatCount As Integer
        wVirtualKeyCode As Integer
        wVirtualScanCode As Integer
        uChar As Integer
        dwControlKeyState As Long
End Type

'  ControlKeyState flags
Private Const RIGHT_ALT_PRESSED = &H1     '  the right alt key is pressed.
Private Const LEFT_ALT_PRESSED = &H2     '  the left alt key is pressed.
Private Const RIGHT_CTRL_PRESSED = &H4     '  the right ctrl key is pressed.
Private Const LEFT_CTRL_PRESSED = &H8     '  the left ctrl key is pressed.
Private Const SHIFT_PRESSED = &H10    '  the shift key is pressed.
Private Const NUMLOCK_ON = &H20    '  the numlock light is on.
Private Const SCROLLLOCK_ON = &H40    '  the scrolllock light is on.
Private Const CAPSLOCK_ON = &H80    '  the capslock light is on.
Private Const ENHANCED_KEY = &H100   '  the key is enhanced.

Type MOUSE_EVENT_RECORD
        dwMousePosition As COORD
        dwButtonState As Long
        dwControlKeyState As Long
        dwEventFlags As Long
End Type

'  ButtonState flags
Private Const FROM_LEFT_1ST_BUTTON_PRESSED = &H1
Private Const RIGHTMOST_BUTTON_PRESSED = &H2
Private Const FROM_LEFT_2ND_BUTTON_PRESSED = &H4
Private Const FROM_LEFT_3RD_BUTTON_PRESSED = &H8
Private Const FROM_LEFT_4TH_BUTTON_PRESSED = &H10

'  EventFlags
Private Const MOUSE_MOVED = &H1
Private Const DOUBLE_CLICK = &H2

Type WINDOW_BUFFER_SIZE_RECORD
        dwSize As COORD
End Type

Type MENU_EVENT_RECORD
        dwCommandId As Long
End Type

Type FOCUS_EVENT_RECORD
        bSetFocus As Long
End Type

'   EventType flags:
Private Const KEY_EVENT = &H1     '  Event contains key event record
Private Const mouse_eventC = &H2     '  Event contains mouse event record
Private Const WINDOW_BUFFER_SIZE_EVENT = &H4     '  Event contains window change event record
Private Const MENU_EVENT = &H8     '  Event contains menu event record
Private Const FOCUS_EVENT = &H10    '  event contains focus change

Type CHAR_INFO
        Char As Integer
        Attributes As Integer
End Type

'  Attributes flags:
Private Const FOREGROUND_BLUE = &H1     '  text color contains blue.
Private Const FOREGROUND_GREEN = &H2     '  text color contains green.
Private Const FOREGROUND_RED = &H4     '  text color contains red.
Private Const FOREGROUND_INTENSITY = &H8     '  text color is intensified.
Private Const BACKGROUND_BLUE = &H10    '  background color contains blue.
Private Const BACKGROUND_GREEN = &H20    '  background color contains green.
Private Const BACKGROUND_RED = &H40    '  background color contains red.
Private Const BACKGROUND_INTENSITY = &H80    '  background color is intensified.

Type CONSOLE_SCREEN_BUFFER_INFO
        dwSize As COORD
        dwCursorPosition As COORD
        wAttributes As Integer
        srWindow As SMALL_RECT
        dwMaximumWindowSize As COORD
End Type

Type CONSOLE_CURSOR_INFO
        dwSize As Long
        bVisible As Long
End Type

Private Const CTRL_C_EVENT = 0
Private Const CTRL_BREAK_EVENT = 1
Private Const CTRL_CLOSE_EVENT = 2
'  3 is reserved!
'  4 is reserved!
Private Const CTRL_LOGOFF_EVENT = 5
Private Const CTRL_SHUTDOWN_EVENT = 6

' Input Mode flags:
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_WINDOW_INPUT = &H8
Private Const ENABLE_MOUSE_INPUT = &H10

' Output Mode flags:
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2

Declare Function ReadConsoleOutput Lib "kernel32" Alias "ReadConsoleOutputA" (ByVal hConsoleOutput As Long, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpReadRegion As SMALL_RECT) As Long
Declare Function WriteConsoleOutput Lib "kernel32" Alias "WriteConsoleOutputA" (ByVal hConsoleOutput As Long, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpWriteRegion As SMALL_RECT) As Long
Declare Function ReadConsoleOutputCharacter Lib "kernel32" Alias "ReadConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfCharsRead As Long) As Long
Declare Function ReadConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, lpAttribute As Long, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfAttrsRead As Long) As Long
Declare Function WriteConsoleOutputCharacter Lib "kernel32" Alias "WriteConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long

Declare Function WriteConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, lpAttribute As Integer, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
Declare Function FillConsoleOutputCharacter Lib "kernel32" Alias "FillConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long
Declare Function FillConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttribute As Long, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
Declare Function GetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, lpMode As Long) As Long
Declare Function GetNumberOfConsoleInputEvents Lib "kernel32" (ByVal hConsoleInput As Long, lpNumberOfEvents As Long) As Long
Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Declare Function GetLargestConsoleWindowSize Lib "kernel32" (ByVal hConsoleOutput As Long) As COORD
Declare Function GetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Declare Function GetNumberOfConsoleMouseButtons Lib "kernel32" (lpNumberOfMouseButtons As Long) As Long
Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, ByVal dwMode As Long) As Long
Declare Function SetConsoleActiveScreenBuffer Lib "kernel32" (ByVal hConsoleOutput As Long) As Long
Declare Function FlushConsoleInputBuffer Lib "kernel32" (ByVal hConsoleInput As Long) As Long
Declare Function SetConsoleScreenBufferSize Lib "kernel32" (ByVal hConsoleOutput As Long, dwSize As COORD) As Long
Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, dwCursorPosition As COORD) As Long
Declare Function SetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Declare Function ScrollConsoleScreenBuffer Lib "kernel32" Alias "ScrollConsoleScreenBufferA" (ByVal hConsoleOutput As Long, lpScrollRectangle As SMALL_RECT, lpClipRectangle As SMALL_RECT, dwDestinationOrigin As COORD, lpFill As CHAR_INFO) As Long
Declare Function SetConsoleWindowInfo Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal bAbsolute As Long, lpConsoleWindow As SMALL_RECT) As Long
Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Declare Function SetConsoleCtrlHandler Lib "kernel32" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long
Declare Function GenerateConsoleCtrlEvent Lib "kernel32" (ByVal dwCtrlEvent As Long, ByVal dwProcessGroupId As Long) As Long
Declare Function AllocConsole Lib "kernel32" () As Long
Declare Function FreeConsole Lib "kernel32" () As Long
Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, lpBuffer As Any, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long

Private Const CONSOLE_TEXTMODE_BUFFER = 1

Declare Function CreateConsoleScreenBuffer Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwFlags As Long, lpScreenBufferData As Any) As Long
Declare Function GetConsoleCP Lib "kernel32" () As Long
Declare Function SetConsoleCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
Declare Function GetConsoleOutputCP Lib "kernel32" () As Long
Declare Function SetConsoleOutputCP Lib "kernel32" (ByVal wCodePageID As Long) As Long


' -------------
'  GDI Section
' -------------

' Binary raster ops
Private Const R2_BLACK = 1       '   0
Private Const R2_NOTMERGEPEN = 2 '  DPon
Private Const R2_MASKNOTPEN = 3  '  DPna
Private Const R2_NOTCOPYPEN = 4  '  PN
Private Const R2_MASKPENNOT = 5  '  PDna
Private Const R2_NOT = 6 '  Dn
Private Const R2_XORPEN = 7      '  DPx
Private Const R2_NOTMASKPEN = 8  '  DPan
Private Const R2_MASKPEN = 9     '  DPa
Private Const R2_NOTXORPEN = 10  '  DPxn
Private Const R2_NOP = 11        '  D
Private Const R2_MERGENOTPEN = 12        '  DPno
Private Const R2_COPYPEN = 13    '  P
Private Const R2_MERGEPENNOT = 14        '  PDno
Private Const R2_MERGEPEN = 15   '  DPo
Private Const R2_WHITE = 16      '   1
Private Const R2_LAST = 16

'  Ternary raster operations
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Private Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Private Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Private Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Private Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Private Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Private Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Private Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Private Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Private Const BLACKNESS = &H42 ' (DWORD) dest = BLACK
Private Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE

Private Const GDI_ERROR = &HFFFF
Private Const HGDI_ERROR = &HFFFF

' Region Flags
Private Const ERRORAPI = 0
Private Const NULLREGION = 1
Private Const SIMPLEREGION = 2
Private Const COMPLEXREGION = 3

' CombineRgn() Styles
Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5
Private Const RGN_MIN = RGN_AND
Private Const RGN_MAX = RGN_COPY

' StretchBlt() Modes
Private Const BLACKONWHITE = 1
Private Const WHITEONBLACK = 2
Private Const COLORONCOLOR = 3
Private Const HALFTONE = 4
Private Const MAXSTRETCHBLTMODE = 4

' PolyFill() Modes
Private Const ALTERNATE = 1
Private Const WINDING = 2
Private Const POLYFILL_LAST = 2

' Text Alignment Options
Private Const TA_NOUPDATECP = 0
Private Const TA_UPDATECP = 1

Private Const TA_LEFT = 0
Private Const TA_RIGHT = 2
Private Const TA_CENTER = 6

Private Const TA_TOP = 0
Private Const TA_BOTTOM = 8
Private Const TA_BASELINE = 24
Private Const TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)

Private Const VTA_BASELINE = TA_BASELINE
Private Const VTA_LEFT = TA_BOTTOM
Private Const VTA_RIGHT = TA_TOP
Private Const VTA_CENTER = TA_CENTER
Private Const VTA_BOTTOM = TA_RIGHT
Private Const VTA_TOP = TA_LEFT

Private Const ETO_GRAYED = 1
Private Const ETO_OPAQUE = 2
Private Const ETO_CLIPPED = 4

Private Const ASPECT_FILTERING = &H1

Private Const DCB_RESET = &H1
Private Const DCB_ACCUMULATE = &H2
Private Const DCB_DIRTY = DCB_ACCUMULATE
Private Const DCB_SET = (DCB_RESET Or DCB_ACCUMULATE)
Private Const DCB_ENABLE = &H4
Private Const DCB_DISABLE = &H8

' Metafile Functions
Private Const META_SETBKCOLOR = &H201
Private Const META_SETBKMODE = &H102
Private Const META_SETMAPMODE = &H103
Private Const META_SETROP2 = &H104
Private Const META_SETRELABS = &H105
Private Const META_SETPOLYFILLMODE = &H106
Private Const META_SETSTRETCHBLTMODE = &H107
Private Const META_SETTEXTCHAREXTRA = &H108
Private Const META_SETTEXTCOLOR = &H209
Private Const META_SETTEXTJUSTIFICATION = &H20A
Private Const META_SETWINDOWORG = &H20B
Private Const META_SETWINDOWEXT = &H20C
Private Const META_SETVIEWPORTORG = &H20D
Private Const META_SETVIEWPORTEXT = &H20E
Private Const META_OFFSETWINDOWORG = &H20F
Private Const META_SCALEWINDOWEXT = &H410
Private Const META_OFFSETVIEWPORTORG = &H211
Private Const META_SCALEVIEWPORTEXT = &H412
Private Const META_LINETO = &H213
Private Const META_MOVETO = &H214
Private Const META_EXCLUDECLIPRECT = &H415
Private Const META_INTERSECTCLIPRECT = &H416
Private Const META_ARC = &H817
Private Const META_ELLIPSE = &H418
Private Const META_FLOODFILL = &H419
Private Const META_PIE = &H81A
Private Const META_RECTANGLE = &H41B
Private Const META_ROUNDRECT = &H61C
Private Const META_PATBLT = &H61D
Private Const META_SAVEDC = &H1E
Private Const META_SETPIXEL = &H41F
Private Const META_OFFSETCLIPRGN = &H220
Private Const META_TEXTOUT = &H521
Private Const META_BITBLT = &H922
Private Const META_STRETCHBLT = &HB23
Private Const META_POLYGON = &H324
Private Const META_POLYLINE = &H325
Private Const META_ESCAPE = &H626
Private Const META_RESTOREDC = &H127
Private Const META_FILLREGION = &H228
Private Const META_FRAMEREGION = &H429
Private Const META_INVERTREGION = &H12A
Private Const META_PAINTREGION = &H12B
Private Const META_SELECTCLIPREGION = &H12C
Private Const META_SELECTOBJECT = &H12D
Private Const META_SETTEXTALIGN = &H12E
Private Const META_CHORD = &H830
Private Const META_SETMAPPERFLAGS = &H231
Private Const META_EXTTEXTOUT = &HA32
Private Const META_SETDIBTODEV = &HD33
Private Const META_SELECTPALETTE = &H234
Private Const META_REALIZEPALETTE = &H35
Private Const META_ANIMATEPALETTE = &H436
Private Const META_SETPALENTRIES = &H37
Private Const META_POLYPOLYGON = &H538
Private Const META_RESIZEPALETTE = &H139
Private Const META_DIBBITBLT = &H940
Private Const META_DIBSTRETCHBLT = &HB41
Private Const META_DIBCREATEPATTERNBRUSH = &H142
Private Const META_STRETCHDIB = &HF43
Private Const META_EXTFLOODFILL = &H548
Private Const META_DELETEOBJECT = &H1F0
Private Const META_CREATEPALETTE = &HF7
Private Const META_CREATEPATTERNBRUSH = &H1F9
Private Const META_CREATEPENINDIRECT = &H2FA
Private Const META_CREATEFONTINDIRECT = &H2FB
Private Const META_CREATEBRUSHINDIRECT = &H2FC
Private Const META_CREATEREGION = &H6FF


' GDI Escapes
Private Const NEWFRAME = 1
Private Const AbortDocC = 2
Private Const NEXTBAND = 3
Private Const SETCOLORTABLE = 4
Private Const GETCOLORTABLE = 5
Private Const FLUSHOUTPUT = 6
Private Const DRAFTMODE = 7
Private Const QUERYESCSUPPORT = 8
Private Const SetAbortProc = 9
Private Const StartDocC = 10
Private Const EndDocC = 11
Private Const GETPHYSPAGESIZE = 12
Private Const GETPRINTINGOFFSET = 13
Private Const GETSCALINGFACTOR = 14
Private Const MFCOMMENT = 15
Private Const GETPENWIDTH = 16
Private Const SETCOPYCOUNT = 17
Private Const SELECTPAPERSOURCE = 18
Private Const DEVICEDATA = 19
Private Const PASSTHROUGH = 19
Private Const GETTECHNOLGY = 20
Private Const GETTECHNOLOGY = 20
Private Const SETLINECAP = 21
Private Const SETLINEJOIN = 22
Private Const SetMiterLimitC = 23
Private Const BANDINFO = 24
Private Const DRAWPATTERNRECT = 25
Private Const GETVECTORPENSIZE = 26
Private Const GETVECTORBRUSHSIZE = 27
Private Const ENABLEDUPLEX = 28
Private Const GETSETPAPERBINS = 29
Private Const GETSETPRINTORIENT = 30
Private Const ENUMPAPERBINS = 31
Private Const SETDIBSCALING = 32
Private Const EPSPRINTING = 33
Private Const ENUMPAPERMETRICS = 34
Private Const GETSETPAPERMETRICS = 35
Private Const POSTSCRIPT_DATA = 37
Private Const POSTSCRIPT_IGNORE = 38
Private Const MOUSETRAILS = 39
Private Const GETDEVICEUNITS = 42

Private Const GETEXTENDEDTEXTMETRICS = 256
Private Const GETEXTENTTABLE = 257
Private Const GETPAIRKERNTABLE = 258
Private Const GETTRACKKERNTABLE = 259
Private Const ExtTextOutC = 512
Private Const GETFACENAME = 513
Private Const DOWNLOADFACE = 514
Private Const ENABLERELATIVEWIDTHS = 768
Private Const ENABLEPAIRKERNING = 769
Private Const SETKERNTRACK = 770
Private Const SETALLJUSTVALUES = 771
Private Const SETCHARSET = 772

Private Const StretchBltC = 2048
Private Const GETSETSCREENPARAMS = 3072
Private Const BEGIN_PATH = 4096
Private Const CLIP_TO_PATH = 4097
Private Const END_PATH = 4098
Private Const EXT_DEVICE_CAPS = 4099
Private Const RESTORE_CTM = 4100
Private Const SAVE_CTM = 4101
Private Const SET_ARC_DIRECTION = 4102
Private Const SET_BACKGROUND_COLOR = 4103
Private Const SET_POLY_MODE = 4104
Private Const SET_SCREEN_ANGLE = 4105
Private Const SET_SPREAD = 4106
Private Const TRANSFORM_CTM = 4107
Private Const SET_CLIP_BOX = 4108
Private Const SET_BOUNDS = 4109
Private Const SET_MIRROR_MODE = 4110
Private Const OPENCHANNEL = 4110
Private Const DOWNLOADHEADER = 4111
Private Const CLOSECHANNEL = 4112
Private Const POSTSCRIPT_PASSTHROUGH = 4115
Private Const ENCAPSULATED_POSTSCRIPT = 4116

' Spooler Error Codes
Private Const SP_NOTREPORTED = &H4000
Private Const SP_ERROR = (-1)
Private Const SP_APPABORT = (-2)
Private Const SP_USERABORT = (-3)
Private Const SP_OUTOFDISK = (-4)
Private Const SP_OUTOFMEMORY = (-5)

Private Const PR_JOBSTATUS = &H0

'  Object Definitions for EnumObjects()
Private Const OBJ_PEN = 1
Private Const OBJ_BRUSH = 2
Private Const OBJ_DC = 3
Private Const OBJ_METADC = 4
Private Const OBJ_PAL = 5
Private Const OBJ_FONT = 6
Private Const OBJ_BITMAP = 7
Private Const OBJ_REGION = 8
Private Const OBJ_METAFILE = 9
Private Const OBJ_MEMDC = 10
Private Const OBJ_EXTPEN = 11
Private Const OBJ_ENHMETADC = 12
Private Const OBJ_ENHMETAFILE = 13

'  xform stuff
Private Const MWT_IDENTITY = 1
Private Const MWT_LEFTMULTIPLY = 2
Private Const MWT_RIGHTMULTIPLY = 3

Private Const MWT_MIN = MWT_IDENTITY
Private Const MWT_MAX = MWT_RIGHTMULTIPLY

Type xform
        eM11 As Double
        eM12 As Double
        eM21 As Double
        eM22 As Double
        eDx As Double
        eDy As Double
End Type

' Bitmap Header Definition
Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Type RGBTRIPLE
        rgbtBlue As Byte
        rgbtGreen As Byte
        rgbtRed As Byte
End Type

Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

' structures for defining DIBs
Type BITMAPCOREHEADER '12 bytes
        bcSize As Long
        bcWidth As Integer
        bcHeight As Integer
        bcPlanes As Integer
        bcBitCount As Integer
End Type

Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

' constants for the biCompression field
Private Const BI_RGB = 0&
Private Const BI_RLE8 = 1&
Private Const BI_RLE4 = 2&
Private Const BI_bitfields = 3&

Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Type BITMAPCOREINFO
        bmciHeader As BITMAPCOREHEADER
        bmciColors As RGBTRIPLE
End Type

Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type


' Clipboard Metafile Picture Structure
Type HANDLETABLE
        objectHandle(1) As Long
End Type

Type METARECORD
        rdSize As Long
        rdFunction As Integer
        rdParm(1) As Integer
End Type


Type METAFILEPICT
        mm As Long
        xExt As Long
        yExt As Long
        hMF As Long
End Type

Type METAHEADER
        mtType As Integer
        mtHeaderSize As Integer
        mtVersion As Integer
        mtSize As Long
        mtNoObjects As Integer
        mtMaxRecord As Long
        mtNoParameters As Integer
End Type

Type ENHMETARECORD
        iType As Long
        nSize As Long
        dParm(1) As Long
End Type

Type SIZEL
    cx As Long
    cy As Long
End Type

Type ENHMETAHEADER
        iType As Long
        nSize As Long
        rclBounds As RECTL
        rclFrame As RECTL
        dSignature As Long
        nVersion As Long
        nBytes As Long
        nRecords As Long
        nHandles As Integer
        sReserved As Integer
        nDescription As Long
        offDescription As Long
        nPalEntries As Long
        szlDevice As SIZEL
        szlMillimeters As SIZEL
End Type

Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

' ntmFlags field flags
Private Const NTM_REGULAR = &H40&
Private Const NTM_BOLD = &H20&
Private Const NTM_ITALIC = &H1&

' Structure passed to FONTENUMPROC
' NOTE: NEWTEXTMETRIC is the same as TEXTMETRIC plus 4 new fields
Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type

'  tmPitchAndFamily flags
Private Const TMPF_FIXED_PITCH = &H1
Private Const TMPF_VECTOR = &H2
Private Const TMPF_DEVICE = &H8
Private Const TMPF_TRUETYPE = &H4


' GDI Logical Objects:

Type PELARRAY
        paXCount As Long
        paYCount As Long
        paXExt As Long
        paYExt As Long
        paRGBs As Integer
End Type

' Logical Brush (or Pattern)
Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

' Logical Pen
Type LOGPEN
        lopnStyle As Long
        lopnWidth As POINTAPI
        lopnColor As Long
End Type

Type EXTLOGPEN
        elpPenStyle As Long
        elpWidth As Long
        elpBrushStyle As Long
        elpColor As Long
        elpHatch As Long
        elpNumEntries As Long
        elpStyleEntry(1) As Long
End Type

Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type

' Logical Palette
Type LOGPALETTE
        palVersion As Integer
        palNumEntries As Integer
        palPalEntry(1) As PALETTEENTRY
End Type

' Logical Font
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

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

Type NONCLIENTMETRICS
        cbSize As Long
        iBorderWidth As Long
        iScrollWidth As Long
        iScrollHeight As Long
        iCaptionWidth As Long
        iCaptionHeight As Long
        lfCaptionFont As LOGFONT
        iSMCaptionWidth As Long
        iSMCaptionHeight As Long
        lfSMCaptionFont As LOGFONT
        iMenuWidth As Long
        iMenuHeight As Long
        lfMenuFont As LOGFONT
        lfStatusFont As LOGFONT
        lfMessageFont As LOGFONT
End Type

Type ENUMLOGFONT
        elfLogFont As LOGFONT
        elfFullName(LF_FULLFACESIZE) As Byte
        elfStyle(LF_FACESIZE) As Byte
End Type

Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_STRING_PRECIS = 1
Private Const OUT_CHARACTER_PRECIS = 2
Private Const OUT_STROKE_PRECIS = 3
Private Const OUT_TT_PRECIS = 4
Private Const OUT_DEVICE_PRECIS = 5
Private Const OUT_RASTER_PRECIS = 6
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const OUT_OUTLINE_PRECIS = 8

Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_STROKE_PRECIS = 2
Private Const CLIP_MASK = &HF
Private Const CLIP_LH_ANGLES = 16
Private Const CLIP_TT_ALWAYS = 32
Private Const CLIP_EMBEDDED = 128

Private Const DEFAULT_QUALITY = 0
Private Const DRAFT_QUALITY = 1
Private Const PROOF_QUALITY = 2

Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2

Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const CHINESEBIG5_CHARSET = 136
Private Const OEM_CHARSET = 255

' Font Families
'
Private Const FF_DONTCARE = 0    '  Don't care or don't know.
Private Const FF_ROMAN = 16      '  Variable stroke width, serifed.

' Times Roman, Century Schoolbook, etc.
Private Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.

' Helvetica, Swiss, etc.
Private Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.

' Pica, Elite, Courier, etc.
Private Const FF_SCRIPT = 64     '  Cursive, etc.
Private Const FF_DECORATIVE = 80 '  Old English, etc.

' Font Weights
Private Const FW_DONTCARE = 0
Private Const FW_THIN = 100
Private Const FW_EXTRALIGHT = 200
Private Const FW_LIGHT = 300
Private Const FW_NORMAL = 400
Private Const FW_MEDIUM = 500
Private Const FW_SEMIBOLD = 600
Private Const FW_BOLD = 700
Private Const FW_EXTRABOLD = 800
Private Const FW_HEAVY = 900

Private Const FW_ULTRALIGHT = FW_EXTRALIGHT
Private Const FW_REGULAR = FW_NORMAL
Private Const FW_DEMIBOLD = FW_SEMIBOLD
Private Const FW_ULTRABOLD = FW_EXTRABOLD
Private Const FW_BLACK = FW_HEAVY

Private Const PANOSE_COUNT = 10
Private Const PAN_FAMILYTYPE_INDEX = 0
Private Const PAN_SERIFSTYLE_INDEX = 1
Private Const PAN_WEIGHT_INDEX = 2
Private Const PAN_PROPORTION_INDEX = 3
Private Const PAN_CONTRAST_INDEX = 4
Private Const PAN_STROKEVARIATION_INDEX = 5
Private Const PAN_ARMSTYLE_INDEX = 6
Private Const PAN_LETTERFORM_INDEX = 7
Private Const PAN_MIDLINE_INDEX = 8
Private Const PAN_XHEIGHT_INDEX = 9

Private Const PAN_CULTURE_LATIN = 0

Type PANOSE
        ulculture As Long
        bFamilyType As Byte
        bSerifStyle As Byte
        bWeight As Byte
        bProportion As Byte
        bContrast As Byte
        bStrokeVariation As Byte
        bArmStyle As Byte
        bLetterform As Byte
        bMidline As Byte
        bXHeight As Byte
End Type

Private Const PAN_ANY = 0  '  Any
Private Const PAN_NO_FIT = 1  '  No Fit

Private Const PAN_FAMILY_TEXT_DISPLAY = 2  '  Text and Display
Private Const PAN_FAMILY_SCRIPT = 3  '  Script
Private Const PAN_FAMILY_DECORATIVE = 4  '  Decorative
Private Const PAN_FAMILY_PICTORIAL = 5  '  Pictorial

Private Const PAN_SERIF_COVE = 2  '  Cove
Private Const PAN_SERIF_OBTUSE_COVE = 3  '  Obtuse Cove
Private Const PAN_SERIF_SQUARE_COVE = 4  '  Square Cove
Private Const PAN_SERIF_OBTUSE_SQUARE_COVE = 5  '  Obtuse Square Cove
Private Const PAN_SERIF_SQUARE = 6  '  Square
Private Const PAN_SERIF_THIN = 7  '  Thin
Private Const PAN_SERIF_BONE = 8  '  Bone
Private Const PAN_SERIF_EXAGGERATED = 9  '  Exaggerated
Private Const PAN_SERIF_TRIANGLE = 10  '  Triangle
Private Const PAN_SERIF_NORMAL_SANS = 11  '  Normal Sans
Private Const PAN_SERIF_OBTUSE_SANS = 12  '  Obtuse Sans
Private Const PAN_SERIF_PERP_SANS = 13  '  Prep Sans
Private Const PAN_SERIF_FLARED = 14  '  Flared
Private Const PAN_SERIF_ROUNDED = 15  '  Rounded

Private Const PAN_WEIGHT_VERY_LIGHT = 2  '  Very Light
Private Const PAN_WEIGHT_LIGHT = 3  '  Light
Private Const PAN_WEIGHT_THIN = 4  '  Thin
Private Const PAN_WEIGHT_BOOK = 5  '  Book
Private Const PAN_WEIGHT_MEDIUM = 6  '  Medium
Private Const PAN_WEIGHT_DEMI = 7  '  Demi
Private Const PAN_WEIGHT_BOLD = 8  '  Bold
Private Const PAN_WEIGHT_HEAVY = 9  '  Heavy
Private Const PAN_WEIGHT_BLACK = 10  '  Black
Private Const PAN_WEIGHT_NORD = 11  '  Nord

Private Const PAN_PROP_OLD_STYLE = 2  '  Old Style
Private Const PAN_PROP_MODERN = 3  '  Modern
Private Const PAN_PROP_EVEN_WIDTH = 4  '  Even Width
Private Const PAN_PROP_EXPANDED = 5  '  Expanded
Private Const PAN_PROP_CONDENSED = 6  '  Condensed
Private Const PAN_PROP_VERY_EXPANDED = 7  '  Very Expanded
Private Const PAN_PROP_VERY_CONDENSED = 8  '  Very Condensed
Private Const PAN_PROP_MONOSPACED = 9  '  Monospaced

Private Const PAN_CONTRAST_NONE = 2  '  None
Private Const PAN_CONTRAST_VERY_LOW = 3  '  Very Low
Private Const PAN_CONTRAST_LOW = 4  '  Low
Private Const PAN_CONTRAST_MEDIUM_LOW = 5  '  Medium Low
Private Const PAN_CONTRAST_MEDIUM = 6  '  Medium
Private Const PAN_CONTRAST_MEDIUM_HIGH = 7  '  Mediim High
Private Const PAN_CONTRAST_HIGH = 8  '  High
Private Const PAN_CONTRAST_VERY_HIGH = 9  '  Very High

Private Const PAN_STROKE_GRADUAL_DIAG = 2  '  Gradual/Diagonal
Private Const PAN_STROKE_GRADUAL_TRAN = 3  '  Gradual/Transitional
Private Const PAN_STROKE_GRADUAL_VERT = 4  '  Gradual/Vertical
Private Const PAN_STROKE_GRADUAL_HORZ = 5  '  Gradual/Horizontal
Private Const PAN_STROKE_RAPID_VERT = 6  '  Rapid/Vertical
Private Const PAN_STROKE_RAPID_HORZ = 7  '  Rapid/Horizontal
Private Const PAN_STROKE_INSTANT_VERT = 8  '  Instant/Vertical

Private Const PAN_STRAIGHT_ARMS_HORZ = 2  '  Straight Arms/Horizontal
Private Const PAN_STRAIGHT_ARMS_WEDGE = 3  '  Straight Arms/Wedge
Private Const PAN_STRAIGHT_ARMS_VERT = 4  '  Straight Arms/Vertical
Private Const PAN_STRAIGHT_ARMS_SINGLE_SERIF = 5 '  Straight Arms/Single-Serif
Private Const PAN_STRAIGHT_ARMS_DOUBLE_SERIF = 6 '  Straight Arms/Double-Serif
Private Const PAN_BENT_ARMS_HORZ = 7  '  Non-Straight Arms/Horizontal
Private Const PAN_BENT_ARMS_WEDGE = 8  '  Non-Straight Arms/Wedge
Private Const PAN_BENT_ARMS_VERT = 9  '  Non-Straight Arms/Vertical
Private Const PAN_BENT_ARMS_SINGLE_SERIF = 10  '  Non-Straight Arms/Single-Serif
Private Const PAN_BENT_ARMS_DOUBLE_SERIF = 11  '  Non-Straight Arms/Double-Serif

Private Const PAN_LETT_NORMAL_CONTACT = 2  '  Normal/Contact
Private Const PAN_LETT_NORMAL_WEIGHTED = 3  '  Normal/Weighted
Private Const PAN_LETT_NORMAL_BOXED = 4  '  Normal/Boxed
Private Const PAN_LETT_NORMAL_FLATTENED = 5  '  Normal/Flattened
Private Const PAN_LETT_NORMAL_ROUNDED = 6  '  Normal/Rounded
Private Const PAN_LETT_NORMAL_OFF_CENTER = 7  '  Normal/Off Center
Private Const PAN_LETT_NORMAL_SQUARE = 8  '  Normal/Square
Private Const PAN_LETT_OBLIQUE_CONTACT = 9  '  Oblique/Contact
Private Const PAN_LETT_OBLIQUE_WEIGHTED = 10  '  Oblique/Weighted
Private Const PAN_LETT_OBLIQUE_BOXED = 11  '  Oblique/Boxed
Private Const PAN_LETT_OBLIQUE_FLATTENED = 12  '  Oblique/Flattened
Private Const PAN_LETT_OBLIQUE_ROUNDED = 13  '  Oblique/Rounded
Private Const PAN_LETT_OBLIQUE_OFF_CENTER = 14  '  Oblique/Off Center
Private Const PAN_LETT_OBLIQUE_SQUARE = 15  '  Oblique/Square

Private Const PAN_MIDLINE_STANDARD_TRIMMED = 2  '  Standard/Trimmed
Private Const PAN_MIDLINE_STANDARD_POINTED = 3  '  Standard/Pointed
Private Const PAN_MIDLINE_STANDARD_SERIFED = 4  '  Standard/Serifed
Private Const PAN_MIDLINE_HIGH_TRIMMED = 5  '  High/Trimmed
Private Const PAN_MIDLINE_HIGH_POINTED = 6  '  High/Pointed
Private Const PAN_MIDLINE_HIGH_SERIFED = 7  '  High/Serifed
Private Const PAN_MIDLINE_CONSTANT_TRIMMED = 8  '  Constant/Trimmed
Private Const PAN_MIDLINE_CONSTANT_POINTED = 9  '  Constant/Pointed
Private Const PAN_MIDLINE_CONSTANT_SERIFED = 10  '  Constant/Serifed
Private Const PAN_MIDLINE_LOW_TRIMMED = 11  '  Low/Trimmed
Private Const PAN_MIDLINE_LOW_POINTED = 12  '  Low/Pointed
Private Const PAN_MIDLINE_LOW_SERIFED = 13  '  Low/Serifed

Private Const PAN_XHEIGHT_CONSTANT_SMALL = 2  '  Constant/Small
Private Const PAN_XHEIGHT_CONSTANT_STD = 3  '  Constant/Standard
Private Const PAN_XHEIGHT_CONSTANT_LARGE = 4  '  Constant/Large
Private Const PAN_XHEIGHT_DUCKING_SMALL = 5  '  Ducking/Small
Private Const PAN_XHEIGHT_DUCKING_STD = 6  '  Ducking/Standard
Private Const PAN_XHEIGHT_DUCKING_LARGE = 7  '  Ducking/Large

Private Const ELF_VENDOR_SIZE = 4

Type EXTLOGFONT
        elfLogFont  As LOGFONT
        elfFullName(LF_FULLFACESIZE) As Byte
        elfStyle(LF_FACESIZE) As Byte
        elfVersion As Long
        elfStyleSize As Long
        elfMatch As Long
        elfReserved As Long
        elfVendorId(ELF_VENDOR_SIZE) As Byte
        elfCulture As Long
        elfPanose As PANOSE
End Type

Private Const ELF_VERSION = 0
Private Const ELF_CULTURE_LATIN = 0

'  EnumFonts Masks
Private Const RASTER_FONTTYPE = &H1
Private Const DEVICE_FONTTYPE = &H2
Private Const TRUETYPE_FONTTYPE = &H4

' palette entry flags
Private Const PC_RESERVED = &H1  '  palette index used for animation
Private Const PC_EXPLICIT = &H2  '  palette index is explicit to device
Private Const PC_NOCOLLAPSE = &H4        '  do not match color to system palette

' Background Modes
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Private Const BKMODE_LAST = 2

'  Graphics Modes
Private Const GM_COMPATIBLE = 1
Private Const GM_ADVANCED = 2
Private Const GM_LAST = 2

'  PolyDraw and GetPath point types
Private Const PT_CLOSEFIGURE = &H1
Private Const PT_LINETO = &H2
Private Const PT_BEZIERTO = &H4
Private Const PT_MOVETO = &H6

'  Mapping Modes
Private Const MM_TEXT = 1
Private Const MM_LOMETRIC = 2
Private Const MM_HIMETRIC = 3
Private Const MM_LOENGLISH = 4
Private Const MM_HIENGLISH = 5
Private Const MM_TWIPS = 6
Private Const MM_ISOTROPIC = 7
Private Const MM_ANISOTROPIC = 8

'  Min and Max Mapping Mode values
Private Const MM_MIN = MM_TEXT
Private Const MM_MAX = MM_ANISOTROPIC
Private Const MM_MAX_FIXEDSCALE = MM_TWIPS

' Coordinate Modes
Private Const ABSOLUTE = 1
Private Const RELATIVE = 2

' Stock Logical Objects
Private Const WHITE_BRUSH = 0
Private Const LTGRAY_BRUSH = 1
Private Const GRAY_BRUSH = 2
Private Const DKGRAY_BRUSH = 3
Private Const BLACK_BRUSH = 4
Private Const NULL_BRUSH = 5
Private Const HOLLOW_BRUSH = NULL_BRUSH
Private Const WHITE_PEN = 6
Private Const BLACK_PEN = 7
Private Const NULL_PEN = 8
Private Const OEM_FIXED_FONT = 10
Private Const ANSI_FIXED_FONT = 11
Private Const ANSI_VAR_FONT = 12
Private Const SYSTEM_FONT = 13
Private Const DEVICE_DEFAULT_FONT = 14
Private Const DEFAULT_PALETTE = 15
Private Const SYSTEM_FIXED_FONT = 16
Private Const STOCK_LAST = 16

Private Const CLR_INVALID = &HFFFF

' Brush Styles
Private Const BS_SOLID = 0
Private Const BS_NULL = 1
Private Const BS_HOLLOW = BS_NULL
Private Const BS_HATCHED = 2
Private Const BS_PATTERN = 3
Private Const BS_INDEXED = 4
Private Const BS_DIBPATTERN = 5
Private Const BS_DIBPATTERNPT = 6
Private Const BS_PATTERN8X8 = 7
Private Const BS_DIBPATTERN8X8 = 8

'  Hatch Styles
Private Const HS_HORIZONTAL = 0              '  -----
Private Const HS_VERTICAL = 1                '  |||||
Private Const HS_FDIAGONAL = 2               '  \\\\\
Private Const HS_BDIAGONAL = 3               '  /////
Private Const HS_CROSS = 4                   '  +++++
Private Const HS_DIAGCROSS = 5               '  xxxxx
Private Const HS_FDIAGONAL1 = 6
Private Const HS_BDIAGONAL1 = 7
Private Const HS_SOLID = 8
Private Const HS_DENSE1 = 9
Private Const HS_DENSE2 = 10
Private Const HS_DENSE3 = 11
Private Const HS_DENSE4 = 12
Private Const HS_DENSE5 = 13
Private Const HS_DENSE6 = 14
Private Const HS_DENSE7 = 15
Private Const HS_DENSE8 = 16
Private Const HS_NOSHADE = 17
Private Const HS_HALFTONE = 18
Private Const HS_SOLIDCLR = 19
Private Const HS_DITHEREDCLR = 20
Private Const HS_SOLIDTEXTCLR = 21
Private Const HS_DITHEREDTEXTCLR = 22
Private Const HS_SOLIDBKCLR = 23
Private Const HS_DITHEREDBKCLR = 24
Private Const HS_API_MAX = 25

'  Pen Styles
Private Const PS_SOLID = 0
Private Const PS_DASH = 1                    '  -------
Private Const PS_DOT = 2                     '  .......
Private Const PS_DASHDOT = 3                 '  _._._._
Private Const PS_DASHDOTDOT = 4              '  _.._.._
Private Const PS_NULL = 5
Private Const PS_INSIDEFRAME = 6
Private Const PS_USERSTYLE = 7
Private Const PS_ALTERNATE = 8
Private Const PS_STYLE_MASK = &HF

Private Const PS_ENDCAP_ROUND = &H0
Private Const PS_ENDCAP_SQUARE = &H100
Private Const PS_ENDCAP_FLAT = &H200
Private Const PS_ENDCAP_MASK = &HF00

Private Const PS_JOIN_ROUND = &H0
Private Const PS_JOIN_BEVEL = &H1000
Private Const PS_JOIN_MITER = &H2000
Private Const PS_JOIN_MASK = &HF000

Private Const PS_COSMETIC = &H0
Private Const PS_GEOMETRIC = &H10000
Private Const PS_TYPE_MASK = &HF0000

Private Const AD_COUNTERCLOCKWISE = 1
Private Const AD_CLOCKWISE = 2

'  Device Parameters for GetDeviceCaps()
Private Const DRIVERVERSION = 0      '  Device driver version
Private Const TECHNOLOGY = 2         '  Device classification
Private Const HORZSIZE = 4           '  Horizontal size in millimeters
Private Const VERTSIZE = 6           '  Vertical size in millimeters
Private Const HORZRES = 8            '  Horizontal width in pixels
Private Const VERTRES = 10           '  Vertical width in pixels
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PLANES = 14            '  Number of planes
Private Const NUMBRUSHES = 16        '  Number of brushes the device has
Private Const NUMPENS = 18           '  Number of pens the device has
Private Const NUMMARKERS = 20        '  Number of markers the device has
Private Const NUMFONTS = 22          '  Number of fonts the device has
Private Const NUMCOLORS = 24         '  Number of colors the device supports
Private Const PDEVICESIZE = 26       '  Size required for device descriptor
Private Const CURVECAPS = 28         '  Curve capabilities
Private Const LINECAPS = 30          '  Line capabilities
Private Const POLYGONALCAPS = 32     '  Polygonal capabilities
Private Const TEXTCAPS = 34          '  Text capabilities
Private Const CLIPCAPS = 36          '  Clipping capabilities
Private Const RASTERCAPS = 38        '  Bitblt capabilities
Private Const ASPECTX = 40           '  Length of the X leg
Private Const ASPECTY = 42           '  Length of the Y leg
Private Const ASPECTXY = 44          '  Length of the hypotenuse

Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Private Const SIZEPALETTE = 104      '  Number of entries in physical palette
Private Const NUMRESERVED = 106      '  Number of reserved entries in palette
Private Const COLORRES = 108         '  Actual color resolution

'  Printing related DeviceCaps. These replace the appropriate Escapes
Private Const PHYSICALWIDTH = 110 '  Physical Width in device units
Private Const PHYSICALHEIGHT = 111 '  Physical Height in device units
Private Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
Private Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin
Private Const SCALINGFACTORX = 114 '  Scaling factor x
Private Const SCALINGFACTORY = 115 '  Scaling factor y

'  Device Capability Masks:

'  Device Technologies
Private Const DT_PLOTTER = 0             '  Vector plotter
Private Const DT_RASDISPLAY = 1          '  Raster display
Private Const DT_RASPRINTER = 2          '  Raster printer
Private Const DT_RASCAMERA = 3           '  Raster camera
Private Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Private Const DT_METAFILE = 5            '  Metafile, VDM
Private Const DT_DISPFILE = 6            '  Display-file

'  Curve Capabilities
Private Const CC_NONE = 0                '  Curves not supported
Private Const CC_CIRCLES = 1             '  Can do circles
Private Const CC_PIE = 2                 '  Can do pie wedges
Private Const CC_CHORD = 4               '  Can do chord arcs
Private Const CC_ELLIPSES = 8            '  Can do ellipese
Private Const CC_WIDE = 16               '  Can do wide lines
Private Const CC_STYLED = 32             '  Can do styled lines
Private Const CC_WIDESTYLED = 64         '  Can do wide styled lines
Private Const CC_INTERIORS = 128 '  Can do interiors
Private Const CC_ROUNDRECT = 256 '

'  Line Capabilities
Private Const LC_NONE = 0                '  Lines not supported
Private Const LC_POLYLINE = 2            '  Can do polylines
Private Const LC_MARKER = 4              '  Can do markers
Private Const LC_POLYMARKER = 8          '  Can do polymarkers
Private Const LC_WIDE = 16               '  Can do wide lines
Private Const LC_STYLED = 32             '  Can do styled lines
Private Const LC_WIDESTYLED = 64         '  Can do wide styled lines
Private Const LC_INTERIORS = 128 '  Can do interiors

'  Polygonal Capabilities
Private Const PC_NONE = 0                '  Polygonals not supported
Private Const PC_POLYGON = 1             '  Can do polygons
Private Const PC_RECTANGLE = 2           '  Can do rectangles
Private Const PC_WINDPOLYGON = 4         '  Can do winding polygons
Private Const PC_TRAPEZOID = 4           '  Can do trapezoids
Private Const PC_SCANLINE = 8            '  Can do scanlines
Private Const PC_WIDE = 16               '  Can do wide borders
Private Const PC_STYLED = 32             '  Can do styled borders
Private Const PC_WIDESTYLED = 64         '  Can do wide styled borders
Private Const PC_INTERIORS = 128 '  Can do interiors

'  Polygonal Capabilities
Private Const CP_NONE = 0                '  No clipping of output
Private Const CP_RECTANGLE = 1           '  Output clipped to rects
Private Const CP_REGION = 2              '

'  Text Capabilities
Private Const TC_OP_CHARACTER = &H1              '  Can do OutputPrecision   CHARACTER
Private Const TC_OP_STROKE = &H2                 '  Can do OutputPrecision   STROKE
Private Const TC_CP_STROKE = &H4                 '  Can do ClipPrecision     STROKE
Private Const TC_CR_90 = &H8                     '  Can do CharRotAbility    90
Private Const TC_CR_ANY = &H10                   '  Can do CharRotAbility    ANY
Private Const TC_SF_X_YINDEP = &H20              '  Can do ScaleFreedom      X_YINDEPENDENT
Private Const TC_SA_DOUBLE = &H40                '  Can do ScaleAbility      DOUBLE
Private Const TC_SA_INTEGER = &H80               '  Can do ScaleAbility      INTEGER
Private Const TC_SA_CONTIN = &H100               '  Can do ScaleAbility      CONTINUOUS
Private Const TC_EA_DOUBLE = &H200               '  Can do EmboldenAbility   DOUBLE
Private Const TC_IA_ABLE = &H400                 '  Can do ItalisizeAbility  ABLE
Private Const TC_UA_ABLE = &H800                 '  Can do UnderlineAbility  ABLE
Private Const TC_SO_ABLE = &H1000                '  Can do StrikeOutAbility  ABLE
Private Const TC_RA_ABLE = &H2000                '  Can do RasterFontAble    ABLE
Private Const TC_VA_ABLE = &H4000                '  Can do VectorFontAble    ABLE
Private Const TC_RESERVED = &H8000
Private Const TC_SCROLLBLT = &H10000             '  do text scroll with blt

'  Raster Capabilities
Private Const RC_NONE = 0
Private Const RC_BITBLT = 1                  '  Can do standard BLT.
Private Const RC_BANDING = 2                 '  Device requires banding support
Private Const RC_SCALING = 4                 '  Device requires scaling support
Private Const RC_BITMAP64 = 8                '  Device can support >64K bitmap
Private Const RC_GDI20_OUTPUT = &H10             '  has 2.0 output calls
Private Const RC_GDI20_STATE = &H20
Private Const RC_SAVEBITMAP = &H40
Private Const RC_DI_BITMAP = &H80                '  supports DIB to memory
Private Const RC_PALETTE = &H100                 '  supports a palette
Private Const RC_DIBTODEV = &H200                '  supports DIBitsToDevice
Private Const RC_BIGFONT = &H400                 '  supports >64K fonts
Private Const RC_STRETCHBLT = &H800              '  supports StretchBlt
Private Const RC_FLOODFILL = &H1000              '  supports FloodFill
Private Const RC_STRETCHDIB = &H2000             '  supports StretchDIBits
Private Const RC_OP_DX_OUTPUT = &H4000
Private Const RC_DEVBITS = &H8000

' DIB color table identifiers
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
Private Const DIB_PAL_COLORS = 1 '  color table in palette indices
Private Const DIB_PAL_INDICES = 2 '  No color table indices into surf palette
Private Const DIB_PAL_PHYSINDICES = 2 '  No color table indices into surf palette
Private Const DIB_PAL_LOGINDICES = 4 '  No color table indices into DC palette

' constants for Get/SetSystemPaletteUse()
Private Const SYSPAL_ERROR = 0
Private Const SYSPAL_STATIC = 1
Private Const SYSPAL_NOSTATIC = 2

' constants for CreateDIBitmap
Private Const CBM_CREATEDIB = &H2      '  create DIB bitmap
Private Const CBM_INIT = &H4           '  initialize bitmap

' ExtFloodFill style flags
Private Const FLOODFILLBORDER = 0
Private Const FLOODFILLSURFACE = 1

'  size of a device name string
Private Const CCHDEVICENAME = 32

'  size of a form name string
Private Const CCHFORMNAME = 32

Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

' current version of specification
Private Const DM_SPECVERSION = &H320

' field selection bits
Private Const DM_ORIENTATION = &H1&
Private Const DM_PAPERSIZE = &H2&
Private Const DM_PAPERLENGTH = &H4&
Private Const DM_PAPERWIDTH = &H8&
Private Const DM_SCALE = &H10&
Private Const DM_COPIES = &H100&
Private Const DM_DEFAULTSOURCE = &H200&
Private Const DM_PRINTQUALITY = &H400&
Private Const DM_COLOR = &H800&
Private Const DM_DUPLEX = &H1000&
Private Const DM_YRESOLUTION = &H2000&
Private Const DM_TTOPTION = &H4000&
Private Const DM_COLLATE As Long = &H8000
Private Const DM_FORMNAME As Long = &H10000

'  orientation selections
Private Const DMORIENT_PORTRAIT = 1
Private Const DMORIENT_LANDSCAPE = 2

'  paper selections
Private Const DMPAPER_LETTER = 1
Private Const DMPAPER_FIRST = DMPAPER_LETTER
               '  Letter 8 1/2 x 11 in
Private Const DMPAPER_LETTERSMALL = 2            '  Letter Small 8 1/2 x 11 in
Private Const DMPAPER_TABLOID = 3                '  Tabloid 11 x 17 in
Private Const DMPAPER_LEDGER = 4                 '  Ledger 17 x 11 in
Private Const DMPAPER_LEGAL = 5                  '  Legal 8 1/2 x 14 in
Private Const DMPAPER_STATEMENT = 6              '  Statement 5 1/2 x 8 1/2 in
Private Const DMPAPER_EXECUTIVE = 7              '  Executive 7 1/4 x 10 1/2 in
Private Const DMPAPER_A3 = 8                     '  A3 297 x 420 mm
Private Const DMPAPER_A4 = 9                     '  A4 210 x 297 mm
Private Const DMPAPER_A4SMALL = 10               '  A4 Small 210 x 297 mm
Private Const DMPAPER_A5 = 11                    '  A5 148 x 210 mm
Private Const DMPAPER_B4 = 12                    '  B4 250 x 354
Private Const DMPAPER_B5 = 13                    '  B5 182 x 257 mm
Private Const DMPAPER_FOLIO = 14                 '  Folio 8 1/2 x 13 in
Private Const DMPAPER_QUARTO = 15                '  Quarto 215 x 275 mm
Private Const DMPAPER_10X14 = 16                 '  10x14 in
Private Const DMPAPER_11X17 = 17                 '  11x17 in
Private Const DMPAPER_NOTE = 18                  '  Note 8 1/2 x 11 in
Private Const DMPAPER_ENV_9 = 19                 '  Envelope #9 3 7/8 x 8 7/8
Private Const DMPAPER_ENV_10 = 20                '  Envelope #10 4 1/8 x 9 1/2
Private Const DMPAPER_ENV_11 = 21                '  Envelope #11 4 1/2 x 10 3/8
Private Const DMPAPER_ENV_12 = 22                '  Envelope #12 4 \276 x 11
Private Const DMPAPER_ENV_14 = 23                '  Envelope #14 5 x 11 1/2
Private Const DMPAPER_CSHEET = 24                '  C size sheet
Private Const DMPAPER_DSHEET = 25                '  D size sheet
Private Const DMPAPER_ESHEET = 26                '  E size sheet
Private Const DMPAPER_ENV_DL = 27                '  Envelope DL 110 x 220mm
Private Const DMPAPER_ENV_C5 = 28                '  Envelope C5 162 x 229 mm
Private Const DMPAPER_ENV_C3 = 29                '  Envelope C3  324 x 458 mm
Private Const DMPAPER_ENV_C4 = 30                '  Envelope C4  229 x 324 mm
Private Const DMPAPER_ENV_C6 = 31                '  Envelope C6  114 x 162 mm
Private Const DMPAPER_ENV_C65 = 32               '  Envelope C65 114 x 229 mm
Private Const DMPAPER_ENV_B4 = 33                '  Envelope B4  250 x 353 mm
Private Const DMPAPER_ENV_B5 = 34                '  Envelope B5  176 x 250 mm
Private Const DMPAPER_ENV_B6 = 35                '  Envelope B6  176 x 125 mm
Private Const DMPAPER_ENV_ITALY = 36             '  Envelope 110 x 230 mm
Private Const DMPAPER_ENV_MONARCH = 37           '  Envelope Monarch 3.875 x 7.5 in
Private Const DMPAPER_ENV_PERSONAL = 38          '  6 3/4 Envelope 3 5/8 x 6 1/2 in
Private Const DMPAPER_FANFOLD_US = 39            '  US Std Fanfold 14 7/8 x 11 in
Private Const DMPAPER_FANFOLD_STD_GERMAN = 40    '  German Std Fanfold 8 1/2 x 12 in
Private Const DMPAPER_FANFOLD_LGL_GERMAN = 41    '  German Legal Fanfold 8 1/2 x 13 in

Private Const DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN

Private Const DMPAPER_USER = 256

'  bin selections
Private Const DMBIN_UPPER = 1
Private Const DMBIN_FIRST = DMBIN_UPPER

Private Const DMBIN_ONLYONE = 1
Private Const DMBIN_LOWER = 2
Private Const DMBIN_MIDDLE = 3
Private Const DMBIN_MANUAL = 4
Private Const DMBIN_ENVELOPE = 5
Private Const DMBIN_ENVMANUAL = 6
Private Const DMBIN_AUTO = 7
Private Const DMBIN_TRACTOR = 8
Private Const DMBIN_SMALLFMT = 9
Private Const DMBIN_LARGEFMT = 10
Private Const DMBIN_LARGECAPACITY = 11
Private Const DMBIN_CASSETTE = 14
Private Const DMBIN_LAST = DMBIN_CASSETTE

Private Const DMBIN_USER = 256               '  device specific bins start here

'  print qualities
Private Const DMRES_DRAFT = (-1)
Private Const DMRES_LOW = (-2)
Private Const DMRES_MEDIUM = (-3)
Private Const DMRES_HIGH = (-4)

'  color enable/disable for color printers
Private Const DMCOLOR_MONOCHROME = 1
Private Const DMCOLOR_COLOR = 2

'  duplex enable
Private Const DMDUP_SIMPLEX = 1
Private Const DMDUP_VERTICAL = 2
Private Const DMDUP_HORIZONTAL = 3

'  TrueType options
Private Const DMTT_BITMAP = 1            '  print TT fonts as graphics
Private Const DMTT_DOWNLOAD = 2          '  download TT fonts as soft fonts
Private Const DMTT_SUBDEV = 3            '  substitute device fonts for TT fonts

'  Collation selections
Private Const DMCOLLATE_FALSE = 0
Private Const DMCOLLATE_TRUE = 1

'  DEVMODE dmDisplayFlags flags

Private Const DM_GRAYSCALE = &H1
Private Const DM_INTERLACED = &H2

'  GetRegionData/ExtCreateRegion

Private Const RDH_RECTANGLES = 1

Type RGNDATAHEADER
        dwSize As Long
        iType As Long
        nCount As Long
        nRgnSize As Long
        rcBound As RECT
End Type

Type RgnData
        rdh As RGNDATAHEADER
        Buffer As Byte
End Type

Type ABC
        abcA As Long
        abcB As Long
        abcC As Long
End Type

Type ABCFLOAT
        abcfA As Double
        abcfB As Double
        abcfC As Double
End Type

Type OUTLINETEXTMETRIC
        otmSize As Long
        otmTextMetrics As TEXTMETRIC
        otmFiller As Byte
        otmPanoseNumber As PANOSE
        otmfsSelection As Long
        otmfsType As Long
        otmsCharSlopeRise As Long
        otmsCharSlopeRun As Long
        otmItalicAngle As Long
        otmEMSquare As Long
        otmAscent As Long
        otmDescent As Long
        otmLineGap As Long
        otmsCapEmHeight As Long
        otmsXHeight As Long
        otmrcFontBox As RECT
        otmMacAscent As Long
        otmMacDescent As Long
        otmMacLineGap As Long
        otmusMinimumPPEM As Long
        otmptSubscriptSize As POINTAPI
        otmptSubscriptOffset As POINTAPI
        otmptSuperscriptSize As POINTAPI
        otmptSuperscriptOffset As POINTAPI
        otmsStrikeoutSize As Long
        otmsStrikeoutPosition As Long
        otmsUnderscorePosition As Long
        otmsUnderscoreSize As Long
        otmpFamilyName As String
        otmpFaceName As String
        otmpStyleName As String
        otmpFullName As String
End Type

Type POLYTEXT
        X As Long
        Y As Long
        n As Long
        lpStr As String
        uiFlags As Long
        rcl As RECT
        pdx As Long
End Type

Type FIXED
        fract As Integer
        Value As Integer
End Type

Type MAT2
        eM11 As FIXED
        eM12 As FIXED
        eM21 As FIXED
        eM22 As FIXED
End Type

Type GLYPHMETRICS
        gmBlackBoxX As Long
        gmBlackBoxY As Long
        gmptGlyphOrigin As POINTAPI
        gmCellIncX As Integer
        gmCellIncY As Integer
End Type


' GetGlyphOutline constants
Private Const GGO_METRICS = 0
Private Const GGO_BITMAP = 1
Private Const GGO_NATIVE = 2

Private Const TT_POLYGON_TYPE = 24

Private Const TT_PRIM_LINE = 1
Private Const TT_PRIM_QSPLINE = 2

Type POINTFX
        X As FIXED
        Y As FIXED
End Type

Type TTPOLYCURVE
        wType As Integer
        cpfx As Integer
        apfx As POINTFX
End Type

Type TTPOLYGONHEADER
        cb As Long
        dwType As Long
        pfxStart As POINTFX
End Type

Type RASTERIZER_STATUS
        nSize As Integer
        wFlags As Integer
        nLanguageID As Integer
End Type

' bits defined in wFlags of RASTERIZER_STATUS
Private Const TT_AVAILABLE = &H1
Private Const TT_ENABLED = &H2

Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long

Declare Function AnimatePalette Lib "gdi32" Alias "AnimatePaletteA" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteColors As PALETTEENTRY) As Long
Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CancelDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function Chord Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function CloseMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function CopyMetaFile Lib "gdi32" Alias "CopyMetaFileA" (ByVal hMF As Long, ByVal lpFileName As String) As Long

Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateDiscardableBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long

Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function CreateDIBPatternBrush Lib "gdi32" (ByVal hPackedDIB As Long, ByVal wUsage As Long) As Long
Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long

Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long

Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long

Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function CreateScalableFontResource Lib "gdi32" Alias "CreateScalableFontResourceA" (ByVal fHidden As Long, ByVal lpszResourceFile As String, ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long

'  mode selections for the device mode function
Private Const DM_UPDATE = 1
Private Const DM_COPY = 2
Private Const DM_PROMPT = 4
Private Const DM_MODIFY = 8

Private Const DM_IN_BUFFER = DM_MODIFY
Private Const DM_IN_PROMPT = DM_PROMPT
Private Const DM_OUT_BUFFER = DM_COPY
Private Const DM_OUT_DEFAULT = DM_UPDATE

'  device capabilities indices
Private Const DC_FIELDS = 1
Private Const DC_PAPERS = 2
Private Const DC_PAPERSIZE = 3
Private Const DC_MINEXTENT = 4
Private Const DC_MAXEXTENT = 5
Private Const DC_BINS = 6
Private Const DC_DUPLEX = 7
Private Const DC_SIZE = 8
Private Const DC_EXTRA = 9
Private Const DC_VERSION = 10
Private Const DC_DRIVER = 11
Private Const DC_BINNAMES = 12
Private Const DC_ENUMRESOLUTIONS = 13
Private Const DC_FILEDEPENDENCIES = 14
Private Const DC_TRUETYPE = 15
Private Const DC_PAPERNAMES = 16
Private Const DC_ORIENTATION = 17
Private Const DC_COPIES = 18

'  bit fields of the return value (DWORD) for DC_TRUETYPE
Private Const DCTT_BITMAP = &H1&
Private Const DCTT_DOWNLOAD = &H2&
Private Const DCTT_SUBDEV = &H4&

Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As DEVMODE) As Long

Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Declare Function EqualRgn Lib "gdi32" (ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long) As Long
Declare Function Escape Lib "gdi32" (ByVal hdc As Long, ByVal nEscape As Long, ByVal nCount As Long, ByVal lpInData As String, lpOutData As Any) As Long
Declare Function ExtEscape Lib "gdi32" (ByVal hdc As Long, ByVal nEscape As Long, ByVal cbInput As Long, ByVal lpszInData As String, ByVal cbOutput As Long, ByVal lpszOutData As String) As Long
Declare Function DrawEscape Lib "gdi32" (ByVal hdc As Long, ByVal nEscape As Long, ByVal cbInput As Long, ByVal lpszInData As String) As Long
Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function ExtCreateRegion Lib "gdi32" (lpXform As xform, ByVal nCount As Long, lpRgnData As RgnData) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function GetROP2 Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetAspectRatioFilterEx Lib "gdi32" (ByVal hdc As Long, lpAspectRatio As Size) As Long
Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetBkMode Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Declare Function GetBitmapDimensionEx Lib "gdi32" (ByVal hBitmap As Long, lpDimension As Size) As Long
Declare Function GetBoundsRect Lib "gdi32" (ByVal hdc As Long, lprcBounds As RECT, ByVal Flags As Long) As Long

 Declare Function GetCharWidth Lib "gdi32" Alias "GetCharWidthA" (ByVal hdc As Long, ByVal wFirstChar As Long, ByVal wLastChar As Long, lpBuffer As Long) As Long
Declare Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpBuffer As Long) As Long
Declare Function GetCharWidthFloat Lib "gdi32" Alias "GetCharWidthFloatA" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, pxBuffer As Double) As Long

Declare Function GetCharABCWidths Lib "gdi32" Alias "GetCharABCWidthsA" (ByVal hdc As Long, ByVal uFirstChar As Long, ByVal uLastChar As Long, lpabc As ABC) As Long
Declare Function GetCharABCWidthsFloat Lib "gdi32" Alias "GetCharABCWidthsFloatA" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpABCF As ABCFLOAT) As Long

Declare Function GetClipBox Lib "gdi32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function GetMetaRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function GetFontData Lib "gdi32" Alias "GetFontDataA" (ByVal hdc As Long, ByVal dwTable As Long, ByVal dwOffset As Long, lpvBuffer As Any, ByVal cbData As Long) As Long
Declare Function GetGlyphOutline Lib "gdi32" Alias "GetGlyphOutlineA" (ByVal hdc As Long, ByVal uChar As Long, ByVal fuFormat As Long, lpgm As GLYPHMETRICS, ByVal cbBuffer As Long, lpBuffer As Any, lpmat2 As MAT2) As Long
Declare Function GetGraphicsMode Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hMF As Long, ByVal nSize As Long, lpvData As Any) As Long
Declare Function GetMetaFile Lib "gdi32" Alias "GetMetaFileA" (ByVal lpFileName As String) As Long
Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long

Declare Function GetOutlineTextMetrics Lib "gdi32" Alias "GetOutlineTextMetricsA" (ByVal hdc As Long, ByVal cbData As Long, lpotm As OUTLINETEXTMETRIC) As Long

Declare Function GetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetPolyFillMode Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetRasterizerCaps Lib "gdi32" (lpraststat As RASTERIZER_STATUS, ByVal cb As Long) As Long
Declare Function GetRegionData Lib "gdi32" Alias "GetRegionDataA" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As RgnData) As Long
Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Declare Function GetStretchBltMode Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Declare Function GetSystemPaletteUse Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetTextCharacterExtra Lib "gdi32" Alias "GetTextCharacterExtraA" (ByVal hdc As Long) As Long
Declare Function GetTextAlign Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long

Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As Size) As Long
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Declare Function GetTextExtentExPoint Lib "gdi32" Alias "GetTextExtentExPointA" (ByVal hdc As Long, ByVal lpszStr As String, ByVal cchString As Long, ByVal nMaxExtent As Long, lpnFit As Long, alpDx As Long, lpSize As Size) As Long

Declare Function GetViewportExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As Size) As Long
Declare Function GetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Declare Function GetWindowExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As Size) As Long
Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long

Declare Function IntersectClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function InvertRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function MaskBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long, ByVal dwRop As Long) As Long
Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long

Declare Function OffsetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function PlayMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hMF As Long) As Long
Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function PtVisible Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function RectVisible Lib "gdi32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hdc As Long, lpInitData As DEVMODE) As Long
Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function ResizePalette Lib "gdi32" (ByVal hPalette As Long, ByVal nNumEntries As Long) As Long

Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function ExtSelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal fnMode As Long) As Long
Declare Function SetMetaRgn Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

Declare Function SetBoundsRect Lib "gdi32" (ByVal hdc As Long, lprcBounds As RECT, ByVal Flags As Long) As Long
Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function SetMapperFlags Lib "gdi32" (ByVal hdc As Long, ByVal dwFlag As Long) As Long
Declare Function SetGraphicsMode Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Declare Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Byte) As Long
Declare Function SetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function SetRectRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Declare Function SetSystemPaletteUse Lib "gdi32" (ByVal hdc As Long, ByVal wUsage As Long) As Long
Declare Function SetTextCharacterExtra Lib "gdi32" Alias "SetTextCharacterExtraA" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Declare Function SetTextJustification Lib "gdi32" (ByVal hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Declare Function UpdateColors Lib "gdi32" (ByVal hdc As Long) As Long

Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Declare Function PlayMetaFileRecord Lib "gdi32" (ByVal hdc As Long, lpHandletable As HANDLETABLE, lpMetaRecord As METARECORD, ByVal nHandles As Long) As Long

Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As String, lpRect As RECT, ByVal lpDescription As String) As Long
Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As Long) As Long
Declare Function EnumEnhMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hemf As Long, ByVal lpEnhMetaFunc As Long, lpData As Any, lpRect As RECT) As Long
Declare Function GetEnhMetaFile Lib "gdi32" Alias "GetEnhMetaFileA" (ByVal lpszMetaFile As String) As Long
Declare Function GetEnhMetaFileBits Lib "gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpbBuffer As Byte) As Long
Declare Function GetEnhMetaFileDescription Lib "gdi32" Alias "GetEnhMetaFileDescriptionA" (ByVal hemf As Long, ByVal cchBuffer As Long, ByVal lpszDescription As String) As Long
Declare Function GetEnhMetaFileHeader Lib "gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpemh As ENHMETAHEADER) As Long
Declare Function GetEnhMetaFilePaletteEntries Lib "gdi32" (ByVal hemf As Long, ByVal cEntries As Long, lppe As PALETTEENTRY) As Long
Declare Function GetWinMetaFileBits Lib "gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal fnMapMode As Long, ByVal hdcRef As Long) As Long
Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hemf As Long, lpRect As RECT) As Long
Declare Function PlayEnhMetaFileRecord Lib "gdi32" (ByVal hdc As Long, lpHandletable As HANDLETABLE, lpEnhMetaRecord As ENHMETARECORD, ByVal nHandles As Long) As Long
Declare Function SetEnhMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpData As Byte) As Long
Declare Function SetWinMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As Long, lpmfp As METAFILEPICT) As Long
Declare Function GdiComment Lib "gdi32" (ByVal hdc As Long, ByVal cbSize As Long, lpData As Byte) As Long

Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long

Declare Function AngleArc Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dwRadius As Long, ByVal eStartAngle As Double, ByVal eSweepAngle As Double) As Long
Declare Function PolyPolyline Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
Declare Function GetWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As xform) As Long
Declare Function SetWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As xform) As Long
Declare Function ModifyWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As xform, ByVal iMode As Long) As Long
Declare Function CombineTransform Lib "gdi32" (lpxformResult As xform, lpxform1 As xform, lpxform2 As xform) As Long

'  Flags value for COLORADJUSTMENT
Private Const CA_NEGATIVE = &H1
Private Const CA_LOG_FILTER = &H2

'  IlluminantIndex values
Private Const ILLUMINANT_DEVICE_DEFAULT = 0
Private Const ILLUMINANT_A = 1
Private Const ILLUMINANT_B = 2
Private Const ILLUMINANT_C = 3
Private Const ILLUMINANT_D50 = 4
Private Const ILLUMINANT_D55 = 5
Private Const ILLUMINANT_D65 = 6
Private Const ILLUMINANT_D75 = 7
Private Const ILLUMINANT_F2 = 8
Private Const ILLUMINANT_MAX_INDEX = ILLUMINANT_F2

Private Const ILLUMINANT_TUNGSTEN = ILLUMINANT_A
Private Const ILLUMINANT_DAYLIGHT = ILLUMINANT_C
Private Const ILLUMINANT_FLUORESCENT = ILLUMINANT_F2
Private Const ILLUMINANT_NTSC = ILLUMINANT_C

'  Min and max for RedGamma, GreenGamma, BlueGamma
Private Const RGB_GAMMA_MIN = 2500 'words
Private Const RGB_GAMMA_MAX = 65000

'  Min and max for ReferenceBlack and ReferenceWhite
Private Const REFERENCE_WHITE_MIN = 6000 'words
Private Const REFERENCE_WHITE_MAX = 10000
Private Const REFERENCE_BLACK_MIN = 0
Private Const REFERENCE_BLACK_MAX = 4000

'  Min and max for Contrast, Brightness, Colorfulness, RedGreenTint
Private Const COLOR_ADJ_MIN = -100 'shorts
Private Const COLOR_ADJ_MAX = 100

Type ColorAdjustment
        caSize As Integer
        caFlags As Integer
        caIlluminantIndex As Integer
        caRedGamma As Integer
        caGreenGamma As Integer
        caBlueGamma As Integer
        caReferenceBlack As Integer
        caReferenceWhite As Integer
        caContrast As Integer
        caBrightness As Integer
        caColorfulness As Integer
        caRedGreenTint As Integer
End Type

Declare Function SetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As ColorAdjustment) As Long
Declare Function GetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As ColorAdjustment) As Long
Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long

Type DOCINFO
        cbSize As Long
        lpszDocName As String
        lpszOutput As String
End Type

Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndDoc Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function AbortDoc Lib "gdi32" (ByVal hdc As Long) As Long

Declare Function AbortPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function ArcTo Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CloseFigure Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function FlattenPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetPath Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpTypes As Byte, ByVal nSize As Long) As Long
Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function PolyDraw Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long
Declare Function SelectClipPath Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Declare Function SetArcDirection Lib "gdi32" (ByVal hdc As Long, ByVal ArcDirection As Long) As Long
Declare Function SetMiterLimit Lib "gdi32" (ByVal hdc As Long, ByVal eNewLimit As Double, peOldLimit As Double) As Long
Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function WidenPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long
Declare Function GetMiterLimit Lib "gdi32" (ByVal hdc As Long, peLimit As Double) As Long
Declare Function GetArcDirection Lib "gdi32" (ByVal hdc As Long) As Long

Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDX As Long) As Long
Declare Function PolyTextOut Lib "gdi32" Alias "PolyTextOutA" (ByVal hdc As Long, pptxt As POLYTEXT, cStrings As Long) As Long

Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function DPtoLP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function LPtoDP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Declare Function PolyBezier Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Declare Function PolyBezierTo Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Declare Function PolylineTo Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long

Declare Function SetViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long

Declare Function OffsetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Declare Function OffsetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Declare Function ScaleWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As Size) As Long
Declare Function ScaleViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As Size) As Long
Declare Function SetBitmapDimensionEx Lib "gdi32" (ByVal hbm As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long

Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long

Private Const FONTMAPPER_MAX = 10

Type KERNINGPAIR
        wFirst As Integer
        wSecond As Integer
        iKernAmount As Long
End Type

Declare Function GetKerningPairs Lib "gdi32" Alias "GetKerningPairsA" (ByVal hdc As Long, ByVal cPairs As Long, lpkrnpair As KERNINGPAIR) As Long

Declare Function GetDCOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long

Declare Function GdiFlush Lib "gdi32" () As Long
Declare Function GdiSetBatchLimit Lib "gdi32" (ByVal dwLimit As Long) As Long
Declare Function GdiGetBatchLimit Lib "gdi32" () As Long

' Enhanced metafile constants

Private Const ENHMETA_SIGNATURE = &H464D4520

'  Stock object flag used in the object handle
' index in the enhanced metafile records.
'  E.g. The object handle index (META_STOCK_OBJECT Or BLACK_BRUSH)
'  represents the stock object BLACK_BRUSH.

Private Const ENHMETA_STOCK_OBJECT = &H80000000

'  Enhanced metafile record types.

Private Const EMR_HEADER = 1
Private Const EMR_POLYBEZIER = 2
Private Const EMR_POLYGON = 3
Private Const EMR_POLYLINE = 4
Private Const EMR_POLYBEZIERTO = 5
Private Const EMR_POLYLINETO = 6
Private Const EMR_POLYPOLYLINE = 7
Private Const EMR_POLYPOLYGON = 8
Private Const EMR_SETWINDOWEXTEX = 9
Private Const EMR_SETWINDOWORGEX = 10
Private Const EMR_SETVIEWPORTEXTEX = 11
Private Const EMR_SETVIEWPORTORGEX = 12
Private Const EMR_SETBRUSHORGEX = 13
Private Const EMR_EOF = 14
Private Const EMR_SETPIXELV = 15
Private Const EMR_SETMAPPERFLAGS = 16
Private Const EMR_SETMAPMODE = 17
Private Const EMR_SETBKMODE = 18
Private Const EMR_SETPOLYFILLMODE = 19
Private Const EMR_SETROP2 = 20
Private Const EMR_SETSTRETCHBLTMODE = 21
Private Const EMR_SETTEXTALIGN = 22
Private Const EMR_SETCOLORADJUSTMENT = 23
Private Const EMR_SETTEXTCOLOR = 24
Private Const EMR_SETBKCOLOR = 25
Private Const EMR_OFFSETCLIPRGN = 26
Private Const EMR_MOVETOEX = 27
Private Const EMR_SETMETARGN = 28
Private Const EMR_EXCLUDECLIPRECT = 29
Private Const EMR_INTERSECTCLIPRECT = 30
Private Const EMR_SCALEVIEWPORTEXTEX = 31
Private Const EMR_SCALEWINDOWEXTEX = 32
Private Const EMR_SAVEDC = 33
Private Const EMR_RESTOREDC = 34
Private Const EMR_SETWORLDTRANSFORM = 35
Private Const EMR_MODIFYWORLDTRANSFORM = 36
Private Const EMR_SELECTOBJECT = 37
Private Const EMR_CREATEPEN = 38
Private Const EMR_CREATEBRUSHINDIRECT = 39
Private Const EMR_DELETEOBJECT = 40
Private Const EMR_ANGLEARC = 41
Private Const EMR_ELLIPSE = 42
Private Const EMR_RECTANGLE = 43
Private Const EMR_ROUNDRECT = 44
Private Const EMR_ARC = 45
Private Const EMR_CHORD = 46
Private Const EMR_PIE = 47
Private Const EMR_SELECTPALETTE = 48
Private Const EMR_CREATEPALETTE = 49
Private Const EMR_SETPALETTEENTRIES = 50
Private Const EMR_RESIZEPALETTE = 51
Private Const EMR_REALIZEPALETTE = 52
Private Const EMR_EXTFLOODFILL = 53
Private Const EMR_LINETO = 54
Private Const EMR_ARCTO = 55
Private Const EMR_POLYDRAW = 56
Private Const EMR_SETARCDIRECTION = 57
Private Const EMR_SETMITERLIMIT = 58
Private Const EMR_BEGINPATH = 59
Private Const EMR_ENDPATH = 60
Private Const EMR_CLOSEFIGURE = 61
Private Const EMR_FILLPATH = 62
Private Const EMR_STROKEANDFILLPATH = 63
Private Const EMR_STROKEPATH = 64
Private Const EMR_FLATTENPATH = 65
Private Const EMR_WIDENPATH = 66
Private Const EMR_SELECTCLIPPATH = 67
Private Const EMR_ABORTPATH = 68

Private Const EMR_GDICOMMENT = 70
Private Const EMR_FILLRGN = 71
Private Const EMR_FRAMERGN = 72
Private Const EMR_INVERTRGN = 73
Private Const EMR_PAINTRGN = 74
Private Const EMR_EXTSELECTCLIPRGN = 75
Private Const EMR_BITBLT = 76
Private Const EMR_STRETCHBLT = 77
Private Const EMR_MASKBLT = 78
Private Const EMR_PLGBLT = 79
Private Const EMR_SETDIBITSTODEVICE = 80
Private Const EMR_STRETCHDIBITS = 81
Private Const EMR_EXTCREATEFONTINDIRECTW = 82
Private Const EMR_EXTTEXTOUTA = 83
Private Const EMR_EXTTEXTOUTW = 84
Private Const EMR_POLYBEZIER16 = 85
Private Const EMR_POLYGON16 = 86
Private Const EMR_POLYLINE16 = 87
Private Const EMR_POLYBEZIERTO16 = 88
Private Const EMR_POLYLINETO16 = 89
Private Const EMR_POLYPOLYLINE16 = 90
Private Const EMR_POLYPOLYGON16 = 91
Private Const EMR_POLYDRAW16 = 92
Private Const EMR_CREATEMONOBRUSH = 93
Private Const EMR_CREATEDIBPATTERNBRUSHPT = 94
Private Const EMR_EXTCREATEPEN = 95
Private Const EMR_POLYTEXTOUTA = 96
Private Const EMR_POLYTEXTOUTW = 97

Private Const EMR_MIN = 1
Private Const EMR_MAX = 97

Type emr
        iType As Long
        nSize As Long
End Type

Type emrtext
        ptlReference As POINTL
        nchars As Long
        offString As Long
        fOptions As Long
        rcl As RECTL
        offDx As Long
End Type

Type EMRABORTPATH
        pEmr As emr
End Type

Type EMRBEGINPATH
        pEmr As emr
End Type

Type EMRENDPATH
        pEmr As emr
End Type

Type EMRCLOSEFIGURE
        pEmr As emr
End Type

Type EMRFLATTENPATH
        pEmr As emr
End Type

Type EMRWIDENPATH
        pEmr As emr
End Type

Type EMRSETMETARGN
        pEmr As emr
End Type

Type EMREMRSAVEDC
        pEmr As emr
End Type

Type EMRREALIZEPALETTE
        pEmr As emr
End Type

Type EMRSELECTCLIPPATH
        pEmr As emr
        iMode As Long
End Type

Type EMRSETBKMODE
        pEmr As emr
        iMode As Long
End Type

Type EMRSETMAPMODE
        pEmr As emr
        iMode As Long
End Type

Type EMRSETPOLYFILLMODE
        pEmr As emr
        iMode As Long
End Type

Type EMRSETROP2
        pEmr As emr
        iMode As Long
End Type

Type EMRSETSTRETCHBLTMODE
        pEmr As emr
        iMode As Long
End Type

Type EMRSETTEXTALIGN
        pEmr As emr
        iMode As Long
End Type

Type EMRSETMITERLIMIT
        pEmr As emr
        eMiterLimit As Double
End Type

Type EMRRESTOREDC
        pEmr As emr
        iRelative As Long
End Type

Type EMRSETARCDIRECTION
        pEmr As emr
        iArcDirection As Long
End Type

Type EMRSETMAPPERFLAGS
        pEmr As emr
        dwFlags As Long
End Type

Type EMRSETTEXTCOLOR
        pEmr As emr
        crColor As Long
End Type

Type EMRSETBKCOLOR
        pEmr As emr
        crColor As Long
End Type

Type EMRSELECTOBJECT
        pEmr As emr
        ihObject As Long
End Type

Type EMRDELETEOBJECT
        pEmr As emr
        ihObject As Long
End Type

Type EMRSELECTPALETTE
        pEmr As emr
        ihPal As Long
End Type

Type EMRRESIZEPALETTE
        pEmr As emr
        ihPal As Long
        cEntries As Long
End Type

Type EMRSETPALETTEENTRIES
        pEmr As emr
        ihPal As Long
        iStart As Long
        cEntries As Long
        aPalEntries(1) As PALETTEENTRY
End Type

Type EMRSETCOLORADJUSTMENT
        pEmr As emr
        ColorAdjustment As ColorAdjustment
End Type

Type EMRGDICOMMENT
        pEmr As emr
        cbData As Long
        Data(1) As Integer
End Type

Type EMREOF
        pEmr As emr
        nPalEntries As Long
        offPalEntries As Long
        nSizeLast As Long
End Type

Type EMRLINETO
        pEmr As emr
        ptl As POINTL
End Type

Type EMRMOVETOEX
        pEmr As emr
        ptl As POINTL
End Type

Type EMROFFSETCLIPRGN
        pEmr As emr
        ptlOffset As POINTL
End Type

Type EMRFILLPATH
        pEmr As emr
        rclBounds As RECTL
End Type

Type EMRSTROKEANDFILLPATH
        pEmr As emr
        rclBounds As RECTL
End Type

Type EMRSTROKEPATH
        pEmr As emr
        rclBounds As RECTL
End Type

Type EMREXCLUDECLIPRECT
        pEmr As emr
        rclClip As RECTL
End Type

Type EMRINTERSECTCLIPRECT
        pEmr As emr
        rclClip As RECTL
End Type

Type EMRSETVIEWPORTORGEX
        pEmr As emr
        ptlOrigin As POINTL
End Type

Type EMRSETWINDOWORGEX
        pEmr As emr
        ptlOrigin As POINTL
End Type

Type EMRSETBRUSHORGEX
        pEmr As emr
        ptlOrigin As POINTL
End Type

Type EMRSETVIEWPORTEXTEX
        pEmr As emr
        szlExtent As SIZEL
End Type

Type EMRSETWINDOWEXTEX
        pEmr As emr
        szlExtent As SIZEL
End Type

Type EMRSCALEVIEWPORTEXTEX
        pEmr As emr
        xNum As Long
        xDenom As Long
        yNum As Long
        yDemon As Long
End Type

Type EMRSCALEWINDOWEXTEX
        pEmr As emr
        xNum As Long
        xDenom As Long
        yNum As Long
        yDemon As Long
End Type

Type EMRSETWORLDTRANSFORM
        pEmr As emr
        xform As xform
End Type

Type EMRMODIFYWORLDTRANSFORM
        pEmr As emr
        xform As xform
        iMode As Long
End Type

Type EMRSETPIXELV
        pEmr As emr
        ptlPixel As POINTL
        crColor As Long
End Type

Type EMREXTFLOODFILL
        pEmr As emr
        ptlStart As POINTL
        crColor As Long
        iMode As Long
End Type

Type EMRELLIPSE
        pEmr As emr
        rclBox As RECTL
End Type

Type EMRRECTANGLE
        pEmr As emr
        rclBox As RECTL
End Type

Type EMRROUNDRECT
        pEmr As emr
        rclBox As RECTL
        szlCorner As SIZEL
End Type

Type EMRARC
        pEmr As emr
        rclBox As RECTL
        ptlStart As POINTL
        ptlEnd As POINTL
End Type

Type EMRARCTO
        pEmr As emr
        rclBox As RECTL
        ptlStart As POINTL
        ptlEnd As POINTL
End Type

Type EMRCHORD
        pEmr As emr
        rclBox As RECTL
        ptlStart As POINTL
        ptlEnd As POINTL
End Type

Type EMRPIE
        pEmr As emr
        rclBox As RECTL
        ptlStart As POINTL
        ptlEnd As POINTL
End Type

Type EMRANGLEARC
        pEmr As emr
        ptlCenter As POINTL
        nRadius As Long
        eStartAngle As Double
        eSweepAngle As Double
End Type

Type EMRPOLYLINE
        pEmr As emr
        rclBounds As RECTL
        cptl As Long
        aptl(1) As POINTL
End Type

Type EMRPOLYBEZIER
        pEmr As emr
        rclBounds As RECTL
        cptl As Long
        aptl(1) As POINTL
End Type

Type EMRPOLYGON
        pEmr As emr
        rclBounds As RECTL
        cptl As Long
        aptl(1) As POINTL
End Type

Type EMRPOLYBEZIERTO
        pEmr As emr
        rclBounds As RECTL
        cptl As Long
        aptl(1) As POINTL
End Type

Type EMRPOLYLINE16
        pEmr As emr
        rclBounds As RECTL
        cpts As Long
        apts(1) As POINTS
End Type

Type EMRPOLYBEZIER16
        pEmr As emr
        rclBounds As RECTL
        cpts As Long
        apts(1) As POINTS
End Type

Type EMRPOLYGON16
        pEmr As emr
        rclBounds As RECTL
        cpts As Long
        apts(1) As POINTS
End Type

Type EMRPLOYBEZIERTO16
        pEmr As emr
        rclBounds As RECTL
        cpts As Long
        apts(1) As POINTS
End Type

Type EMRPOLYLINETO16
        pEmr As emr
        rclBounds As RECTL
        cpts As Long
        apts(1) As POINTS
End Type

Type EMRPOLYDRAW
        pEmr As emr
        rclBounds As RECTL
        cptl As Long
        aptl(1) As POINTL
        abTypes(1) As Integer
End Type

Type EMRPOLYDRAW16
        pEmr As emr
        rclBounds As RECTL
        cpts As Long
        apts(1) As POINTS
        abTypes(1) As Integer
End Type

Type EMRPOLYPOLYLINE
        pEmr As emr
        rclBounds As RECTL
        nPolys As Long
        cptl As Long
        aPolyCounts(1) As Long
        aptl(1) As POINTL
End Type

Type EMRPOLYPOLYGON
        pEmr As emr
        rclBounds As RECTL
        nPolys As Long
        cptl As Long
        aPolyCounts(1) As Long
        aptl(1) As POINTL
End Type

Type EMRPOLYPOLYLINE16
        pEmr As emr
        rclBounds As RECTL
        nPolys As Long
        cpts As Long
        aPolyCounts(1) As Long
        apts(1) As POINTS
End Type

Type EMRPOLYPOLYGON16
        pEmr As emr
        rclBounds As RECTL
        nPolys As Long
        cpts As Long
        aPolyCounts(1) As Long
        apts(1) As POINTS
End Type

Type EMRINVERTRGN
        pEmr As emr
        rclBounds As RECTL
        cbRgnData As Long
        RgnData(1) As Integer
End Type

Type EMRPAINTRGN
        pEmr As emr
        rclBounds As RECTL
        cbRgnData As Long
        RgnData(1) As Integer
End Type

Type EMRFILLRGN
        pEmr As emr
        rclBounds As RECTL
        cbRgnData As Long
        ihBrush As Long
        RgnData(1) As Integer
End Type

Type EMRFRAMERGN
        pEmr As emr
        rclBounds As RECTL
        cbRgnData As Long
        ihBrush As Long
        szlStroke As SIZEL
        RgnData(1) As Integer
End Type

Type EMREXTSELECTCLIPRGN
        pEmr As emr
        cbRgnData As Long
        iMode As Long
        RgnData(1) As Integer
End Type

Type EMREXTTEXTOUT
        pEmr As emr
        rclBounds As RECTL
        iGraphicsMode As Long
        exScale As Double
        eyScale As Double
        emrtext As emrtext
End Type

Type EMRBITBLT
        pEmr As emr
        rclBounds As RECTL
        xDest As Long
        yDest As Long
        cxDest As Long
        cyDest As Long
        dwRop As Long
        xSrc As Long
        ySrc As Long
        xformSrc As xform
        crBkColorSrc As Long
        iUsageSrc As Long
        offBmiSrc As Long
        cbBmiSrc As Long
        offBitsSrc As Long
        cbBitsSrc As Long
End Type

Type EMRSTRETCHBLT
        pEmr As emr
        rclBounds As RECTL
        xDest As Long
        yDest As Long
        cxDest As Long
        cyDest As Long
        dwRop As Long
        xSrc As Long
        ySrc As Long
        xformSrc As xform
        crBkColorSrc As Long
        iUsageSrc As Long
        offBmiSrc As Long
        cbBmiSrc As Long
        offBitsSrc As Long
        cbBitsSrc As Long
        cxSrc As Long
        cySrc As Long
End Type

Type EMRMASKBLT
        pEmr As emr
        rclBounds As RECTL
        xDest As Long
        yDest As Long
        cxDest As Long
        cyDest As Long
        dwRop As Long
        xSrc2 As Long
        cyDest2 As Long
        dwRop2 As Long
        xSrc As Long
        ySrc As Long
        xformSrc As xform
        crBkColorSrc As Long
        iUsageSrc As Long
        offBmiSrc As Long
        cbBmiSrc As Long
        offBitsSrc As Long
        cbBitsSrc As Long
        xMask As Long
        yMask As Long
        iUsageMask As Long
        offBmiMask As Long
        cbBmiMask As Long
        offBitsMask As Long
        cbBitsMask As Long
End Type

Type EMRPLGBLT
        pEmr As emr
        rclBounds As RECTL
        aptlDest(3) As POINTL
        xSrc As Long
        ySrc As Long
        cxSrc As Long
        cySrc As Long
        xformSrc As xform
        crBkColorSrc As Long
        iUsageSrc As Long
        offBmiSrc As Long
        cbBmiSrc As Long
        offBitsSrc As Long
        cbBitsSrc As Long
        xMask As Long
        yMask As Long
        iUsageMask As Long
        offBmiMask As Long
        cbBmiMask As Long
        offBitsMask As Long
        cbBitsMask As Long
End Type

Type EMRSETDIBITSTODEVICE
        pEmr As emr
        rclBounds As RECTL
        xDest As Long
        yDest As Long
        xSrc As Long
        ySrc As Long
        cxSrc As Long
        cySrc As Long
        offBmiSrc As Long
        cbBmiSrc As Long
        offBitsSrc As Long
        cbBitsSrc As Long
        iUsageSrc As Long
        iStartScan As Long
        cScans As Long
End Type

Type EMRSTRETCHDIBITS
        pEmr As emr
        rclBounds As RECTL
        xDest As Long
        yDest As Long
        xSrc As Long
        ySrc As Long
        cxSrc As Long
        cySrc As Long
        offBmiSrc As Long
        cbBmiSrc As Long
        offBitsSrc As Long
        cbBitsSrc As Long
        iUsageSrc As Long
        dwRop As Long
        cxDest As Long
        cyDest As Long
End Type

Type EMREXTCREATEFONTINDIRECT
        pEmr As emr
        ihFont As Long
        elfw As EXTLOGFONT
End Type

Type EMRCREATEPALETTE
        pEmr As emr
        ihPal As Long
        lgpl As LOGPALETTE
End Type

Type EMRCREATEPEN
        pEmr As emr
        ihPen As Long
        lopn As LOGPEN
End Type

Type EMREXTCREATEPEN
        pEmr As emr
        ihPen As Long
        offBmi As Long
        cbBmi As Long
        offBits As Long
        cbBits As Long
        elp As EXTLOGPEN
End Type

Type EMRCREATEBRUSHINDIRECT
        pEmr As emr
        ihBrush As Long
        lb As LOGBRUSH
End Type

Type EMRCREATEMONOBRUSH
        pEmr As emr
        ihBrush As Long
        iUsage As Long
        offBmi As Long
        cbBmi As Long
        offBits As Long
        cbBits As Long
End Type

Type EMRCREATEDIBPATTERNBRUSHPT
        pEmr As emr
        ihBursh As Long
        iUsage As Long
        offBmi As Long
        cbBmi As Long
        offBits As Long
        cbBits As Long
End Type

' new wingdi
' *************************************************************************
' *                                                                         *
' * wingdi.h -- GDI procedure declarations, constant definitions and macros *
' *                                                                         *
' * Copyright (c) 1985-1995, Microsoft Corp. All rights reserved.           *
' *                                                                         *
' **************************************************************************/

'  StretchBlt() Modes
Private Const STRETCH_ANDSCANS = 1
Private Const STRETCH_ORSCANS = 2
Private Const STRETCH_DELETESCANS = 3
Private Const STRETCH_HALFTONE = 4

Type BITMAPV4HEADER
        bV4Size As Long
        bV4Width As Long
        bV4Height As Long
        bV4Planes As Integer
        bV4BitCount As Integer
        bV4V4Compression As Long
        bV4SizeImage As Long
        bV4XPelsPerMeter As Long
        bV4YPelsPerMeter As Long
        bV4ClrUsed As Long
        bV4ClrImportant As Long
        bV4RedMask As Long
        bV4GreenMask As Long
        bV4BlueMask As Long
        bV4AlphaMask As Long
        bV4CSType As Long
        bV4Endpoints As Long
        bV4GammaRed As Long
        bV4GammaGreen As Long
        bV4GammaBlue As Long
End Type

Type FONTSIGNATURE
        fsUsb(4) As Long
        fsCsb(2) As Long
End Type

Type CHARSETINFO
        ciCharset As Long
        ciACP As Long
        fs As FONTSIGNATURE
End Type

Private Const TCI_SRCCHARSET = 1
Private Const TCI_SRCCODEPAGE = 2
Private Const TCI_SRCFONTSIG = 3

Type LOCALESIGNATURE
        lsUsb(4) As Long
        lsCsbDefault(2) As Long
        lsCsbSupported(2) As Long
End Type

Type NEWTEXTMETRICEX
        ntmTm As NEWTEXTMETRIC
        ntmFontSig As FONTSIGNATURE
End Type

Type ENUMLOGFONTEX
        elfLogFont As LOGFONT
        elfFullName(LF_FULLFACESIZE) As Byte
        elfStyle(LF_FACESIZE) As Byte
        elfScript(LF_FACESIZE) As Byte
End Type

Private Const MONO_FONT = 8
Private Const JOHAB_CHARSET = 130
Private Const HEBREW_CHARSET = 177
Private Const ARABIC_CHARSET = 178
Private Const GREEK_CHARSET = 161
Private Const TURKISH_CHARSET = 162
Private Const THAI_CHARSET = 222
Private Const EASTEUROPE_CHARSET = 238
Private Const RUSSIAN_CHARSET = 204

Private Const MAC_CHARSET = 77
Private Const BALTIC_CHARSET = 186

Private Const FS_LATIN1 = &H1&
Private Const FS_LATIN2 = &H2&
Private Const FS_CYRILLIC = &H4&
Private Const FS_GREEK = &H8&
Private Const FS_TURKISH = &H10&
Private Const FS_HEBREW = &H20&
Private Const FS_ARABIC = &H40&
Private Const FS_BALTIC = &H80&
Private Const FS_THAI = &H10000
Private Const FS_JISJAPAN = &H20000
Private Const FS_CHINESESIMP = &H40000
Private Const FS_WANSUNG = &H80000
Private Const FS_CHINESETRAD = &H100000
Private Const FS_JOHAB = &H200000
Private Const FS_SYMBOL = &H80000000
Private Const DEFAULT_GUI_FONT = 17
'  current version of specification
Private Const DM_RESERVED1 = &H800000
Private Const DM_RESERVED2 = &H1000000
Private Const DM_ICMMETHOD = &H2000000
Private Const DM_ICMINTENT = &H4000000
Private Const DM_MEDIATYPE = &H8000000
Private Const DM_DITHERTYPE = &H10000000
Private Const DMPAPER_ISO_B4 = 42                '  B4 (ISO) 250 x 353 mm
Private Const DMPAPER_JAPANESE_POSTCARD = 43     '  Japanese Postcard 100 x 148 mm
Private Const DMPAPER_9X11 = 44                  '  9 x 11 in
Private Const DMPAPER_10X11 = 45                 '  10 x 11 in
Private Const DMPAPER_15X11 = 46                 '  15 x 11 in
Private Const DMPAPER_ENV_INVITE = 47            '  Envelope Invite 220 x 220 mm
Private Const DMPAPER_RESERVED_48 = 48           '  RESERVED--DO NOT USE
Private Const DMPAPER_RESERVED_49 = 49           '  RESERVED--DO NOT USE
Private Const DMPAPER_LETTER_EXTRA = 50              '  Letter Extra 9 \275 x 12 in
Private Const DMPAPER_LEGAL_EXTRA = 51               '  Legal Extra 9 \275 x 15 in
Private Const DMPAPER_TABLOID_EXTRA = 52              '  Tabloid Extra 11.69 x 18 in
Private Const DMPAPER_A4_EXTRA = 53                   '  A4 Extra 9.27 x 12.69 in
Private Const DMPAPER_LETTER_TRANSVERSE = 54     '  Letter Transverse 8 \275 x 11 in
Private Const DMPAPER_A4_TRANSVERSE = 55         '  A4 Transverse 210 x 297 mm
Private Const DMPAPER_LETTER_EXTRA_TRANSVERSE = 56 '  Letter Extra Transverse 9\275 x 12 in
Private Const DMPAPER_A_PLUS = 57                '  SuperA/SuperA/A4 227 x 356 mm
Private Const DMPAPER_B_PLUS = 58                '  SuperB/SuperB/A3 305 x 487 mm
Private Const DMPAPER_LETTER_PLUS = 59           '  Letter Plus 8.5 x 12.69 in
Private Const DMPAPER_A4_PLUS = 60               '  A4 Plus 210 x 330 mm
Private Const DMPAPER_A5_TRANSVERSE = 61         '  A5 Transverse 148 x 210 mm
Private Const DMPAPER_B5_TRANSVERSE = 62         '  B5 (JIS) Transverse 182 x 257 mm
Private Const DMPAPER_A3_EXTRA = 63              '  A3 Extra 322 x 445 mm
Private Const DMPAPER_A5_EXTRA = 64              '  A5 Extra 174 x 235 mm
Private Const DMPAPER_B5_EXTRA = 65              '  B5 (ISO) Extra 201 x 276 mm
Private Const DMPAPER_A2 = 66                    '  A2 420 x 594 mm
Private Const DMPAPER_A3_TRANSVERSE = 67         '  A3 Transverse 297 x 420 mm
Private Const DMPAPER_A3_EXTRA_TRANSVERSE = 68   '  A3 Extra Transverse 322 x 445 mm
Private Const DMTT_DOWNLOAD_OUTLINE = 4 '  download TT fonts as outline soft fonts

'  ICM methods
Private Const DMICMMETHOD_NONE = 1       '  ICM disabled
Private Const DMICMMETHOD_SYSTEM = 2     '  ICM handled by system
Private Const DMICMMETHOD_DRIVER = 3     '  ICM handled by driver
Private Const DMICMMETHOD_DEVICE = 4     '  ICM handled by device
Private Const DMICMMETHOD_USER = 256     '  Device-specific methods start here

'  ICM Intents
Private Const DMICM_SATURATE = 1         '  Maximize color saturation
Private Const DMICM_CONTRAST = 2         '  Maximize color contrast
Private Const DMICM_COLORMETRIC = 3      '  Use specific color metric
Private Const DMICM_USER = 256           '  Device-specific intents start here

'  Media types
Private Const DMMEDIA_STANDARD = 1         '  Standard paper
Private Const DMMEDIA_GLOSSY = 2           '  Glossy paper
Private Const DMMEDIA_TRANSPARENCY = 3     '  Transparency

Private Const DMMEDIA_USER = 256           '  Device-specific media start here

'  Dither types
Private Const DMDITHER_NONE = 1          '  No dithering
Private Const DMDITHER_COARSE = 2        '  Dither with a coarse brush
Private Const DMDITHER_FINE = 3          '  Dither with a fine brush
Private Const DMDITHER_LINEART = 4       '  LineArt dithering
Private Const DMDITHER_GRAYSCALE = 5     '  Device does grayscaling
Private Const DMDITHER_USER = 256        '  Device-specific dithers start here
Private Const GGO_GRAY2_BITMAP = 4
Private Const GGO_GRAY4_BITMAP = 5
Private Const GGO_GRAY8_BITMAP = 6
Private Const GGO_GLYPH_INDEX = &H80
Private Const GCP_DBCS = &H1
Private Const GCP_REORDER = &H2
Private Const GCP_USEKERNING = &H8
Private Const GCP_GLYPHSHAPE = &H10
Private Const GCP_LIGATE = &H20
Private Const GCP_DIACRITIC = &H100
Private Const GCP_KASHIDA = &H400
Private Const GCP_ERROR = &H8000
Private Const FLI_MASK = &H103B
Private Const GCP_JUSTIFY = &H10000
Private Const GCP_NODIACRITICS = &H20000
Private Const FLI_GLYPHS = &H40000
Private Const GCP_CLASSIN = &H80000
Private Const GCP_MAXEXTENT = &H100000
Private Const GCP_JUSTIFYIN = &H200000
Private Const GCP_DISPLAYZWG = &H400000
Private Const GCP_SYMSWAPOFF = &H800000
Private Const GCP_NUMERICOVERRIDE = &H1000000
Private Const GCP_NEUTRALOVERRIDE = &H2000000
Private Const GCP_NUMERICSLATIN = &H4000000
Private Const GCP_NUMERICSLOCAL = &H8000000
Private Const GCPCLASS_LATIN = 1
Private Const GCPCLASS_HEBREW = 2
Private Const GCPCLASS_ARABIC = 2
Private Const GCPCLASS_NEUTRAL = 3
Private Const GCPCLASS_LOCALNUMBER = 4
Private Const GCPCLASS_LATINNUMBER = 5
Private Const GCPCLASS_LATINNUMERICTERMINATOR = 6
Private Const GCPCLASS_LATINNUMERICSEPARATOR = 7
Private Const GCPCLASS_NUMERICSEPARATOR = 8
Private Const GCPCLASS_PREBOUNDRTL = &H80
Private Const GCPCLASS_PREBOUNDLTR = &H40



Type GCP_RESULTS
        lStructSize As Long
        lpOutString As String
        lpOrder As Long
        lpDX As Long
        lpCaretPos As Long
        lpClass As String
        lpGlyphs As String
        nGlyphs As Long
        nMaxFit As Long
End Type

Private Const DC_BINADJUST = 19
Private Const DC_EMF_COMPLIANT = 20
Private Const DC_DATATYPE_PRODUCED = 21
Private Const DC_COLLATE = 22

Private Const DCTT_DOWNLOAD_OUTLINE = &H8&

'  return values for DC_BINADJUST
Private Const DCBA_FACEUPNONE = &H0
Private Const DCBA_FACEUPCENTER = &H1
Private Const DCBA_FACEUPLEFT = &H2
Private Const DCBA_FACEUPRIGHT = &H3
Private Const DCBA_FACEDOWNNONE = &H100
Private Const DCBA_FACEDOWNCENTER = &H101
Private Const DCBA_FACEDOWNLEFT = &H102
Private Const DCBA_FACEDOWNRIGHT = &H103

Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, ByVal lParam As Long) As Long
Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long
Declare Function GetTextCharset Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetTextCharsetInfo Lib "gdi32" (ByVal hdc As Long, lpSig As FONTSIGNATURE, ByVal dwFlags As Long) As Long

Declare Function TranslateCharsetInfo Lib "gdi32" (lpSrc As Long, lpcs As CHARSETINFO, ByVal dwFlags As Long) As Long
Declare Function GetFontLanguageInfo Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetCharacterPlacement Lib "gdi32" Alias " GetCharacterPlacementA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n1 As Long, ByVal n2 As Long, lpGcpResults As GCP_RESULTS, ByVal dw As Long) As Long

Private Const ICM_OFF = 1
Private Const ICM_ON = 2
Private Const ICM_QUERY = 3

Type CIEXYZ
    ciexyzX As Long
    ciexyzY As Long
    ciexyzZ As Long
End Type

Type CIEXYZTRIPLE
    ciexyzRed As CIEXYZ
    ciexyzGreen As CIEXYZ
    ciexyBlue As CIEXYZ
End Type

Type LOGCOLORSPACE
    lcsSignature As Long
    lcsVersion As Long
    lcsSize As Long
    lcsCSType As Long
    lcsIntent As Long
    lcsEndPoints As CIEXYZTRIPLE
    lcsGammaRed As Long
    lcsGammaGreen As Long
    lcsGammaBlue As Long
    lcsFileName As String * MAX_PATH
End Type

Declare Function SetICMMode Lib "gdi32" (ByVal hdc As Long, ByVal n As Long) As Long
Declare Function CheckColorsInGamut Lib "gdi32" (ByVal hdc As Long, lpv As Any, lpv2 As Any, ByVal dw As Long) As Long
Declare Function GetLogColorSpace Lib "gdi32" Alias "GetLogColorSpaceA" (ByVal hcolorspace As Long, ByVal lplogcolorspace As LOGCOLORSPACE, ByVal dw As Long) As Long
Declare Function GetColorSpace Lib "gdi32" (ByVal hdc As Long) As Long

Declare Function CreateColorSpace Lib "gdi32" Alias "CreateColorSpaceA" (lplogcolorspace As LOGCOLORSPACE) As Long

Declare Function SetColorSpace Lib "gdi32" (ByVal hdc As Long, ByVal hcolorspace As Long) As Long
Declare Function DeleteColorSpace Lib "gdi32" (ByVal hcolorspace As Long) As Long
Declare Function GetICMProfile Lib "gdi32" Alias "GetICMProfileA" (ByVal hdc As Long, ByVal dw As Long, ByVal lpStr As String) As Long
Declare Function SetICMProfile Lib "gdi32" Alias "SetICMProfileA" (ByVal hdc As Long, ByVal lpStr As String) As Long
Declare Function GetDeviceGammaRamp Lib "gdi32" (ByVal hdc As Long, lpv As Any) As Long
Declare Function SetDeviceGammaRamp Lib "gdi32" (ByVal hdc As Long, lpv As Any) As Long
Declare Function ColorMatchToTarget Lib "gdi32" (ByVal hdc As Long, ByVal hdc2 As Long, ByVal dw As Long) As Long

Declare Function EnumICMProfiles Lib "gdi32" Alias "EnumICMProfilesA" (ByVal hdc As Long, ByVal icmEnumProc As Long, ByVal lParam As Long) As Long

Private Const EMR_SETICMMODE = 98
Private Const EMR_CREATECOLORSPACE = 99
Private Const EMR_SETCOLORSPACE = 100
Private Const EMR_DELETECOLORSPACE = 101

Type EMRSELECTCOLORSPACE
        pEmr As emr
        ihCS As Long               '  ColorSpace handle index
End Type

Type EMRCREATECOLORSPACE
        pEmr As emr
        ihCS As Long        '  ColorSpace handle index
        lcs As LOGCOLORSPACE
End Type


' --------------
'  USER Section
' --------------

' Scroll Bar Constants
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const SB_CTL = 2
Private Const SB_BOTH = 3

' Scroll Bar Commands
Private Const SB_LINEUP = 0
Private Const SB_LINELEFT = 0
Private Const SB_LINEDOWN = 1
Private Const SB_LINERIGHT = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGELEFT = 2
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGERIGHT = 3
Private Const SB_THUMBPOSITION = 4
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_LEFT = 6
Private Const SB_BOTTOM = 7
Private Const SB_RIGHT = 7
Private Const SB_ENDSCROLL = 8

' ShowWindow() Commands
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10

' Old ShowWindow() Commands
Private Const HIDE_WINDOW = 0
Private Const SHOW_OPENWINDOW = 1
Private Const SHOW_ICONWINDOW = 2
Private Const SHOW_FULLSCREEN = 3
Private Const SHOW_OPENNOACTIVATE = 4

' Identifiers for the WM_SHOWWINDOW message
Private Const SW_PARENTCLOSING = 1
Private Const SW_OTHERZOOM = 2
Private Const SW_PARENTOPENING = 3
Private Const SW_OTHERUNZOOM = 4

' WM_KEYUP/DOWN/CHAR HIWORD(lParam) flags
Private Const KF_EXTENDED = &H100
Private Const KF_DLGMODE = &H800
Private Const KF_MENUMODE = &H1000
Private Const KF_ALTDOWN = &H2000
Private Const KF_REPEAT = &H4000
Private Const KF_UP = &H8000

' Virtual Keys, Standard Set
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const VK_CANCEL = &H3
Private Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON

Private Const VK_BACK = &H8
Private Const VK_TAB = &H9

Private Const VK_CLEAR = &HC
Private Const VK_RETURN = &HD

Private Const VK_SHIFT = &H10
Private Const VK_CONTROL = &H11
Private Const VK_MENU = &H12
Private Const VK_PAUSE = &H13
Private Const VK_CAPITAL = &H14

Private Const VK_ESCAPE = &H1B

Private Const VK_SPACE = &H20
Private Const VK_PRIOR = &H21
Private Const VK_NEXT = &H22
Private Const VK_END = &H23
Private Const VK_HOME = &H24
Private Const VK_LEFT = &H25
Private Const VK_UP = &H26
Private Const VK_RIGHT = &H27
Private Const VK_DOWN = &H28
Private Const VK_SELECT = &H29
Private Const VK_PRINT = &H2A
Private Const VK_EXECUTE = &H2B
Private Const VK_SNAPSHOT = &H2C
Private Const VK_INSERT = &H2D
Private Const VK_DELETE = &H2E
Private Const VK_HELP = &H2F

' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'

Private Const VK_NUMPAD0 = &H60
Private Const VK_NUMPAD1 = &H61
Private Const VK_NUMPAD2 = &H62
Private Const VK_NUMPAD3 = &H63
Private Const VK_NUMPAD4 = &H64
Private Const VK_NUMPAD5 = &H65
Private Const VK_NUMPAD6 = &H66
Private Const VK_NUMPAD7 = &H67
Private Const VK_NUMPAD8 = &H68
Private Const VK_NUMPAD9 = &H69
Private Const VK_MULTIPLY = &H6A
Private Const VK_ADD = &H6B
Private Const VK_SEPARATOR = &H6C
Private Const VK_SUBTRACT = &H6D
Private Const VK_DECIMAL = &H6E
Private Const VK_DIVIDE = &H6F
Private Const VK_F1 = &H70
Private Const VK_F2 = &H71
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B
Private Const VK_F13 = &H7C
Private Const VK_F14 = &H7D
Private Const VK_F15 = &H7E
Private Const VK_F16 = &H7F
Private Const VK_F17 = &H80
Private Const VK_F18 = &H81
Private Const VK_F19 = &H82
Private Const VK_F20 = &H83
Private Const VK_F21 = &H84
Private Const VK_F22 = &H85
Private Const VK_F23 = &H86
Private Const VK_F24 = &H87

Private Const VK_NUMLOCK = &H90
Private Const VK_SCROLL = &H91

'
'   VK_L VK_R - left and right Alt, Ctrl and Shift virtual keys.
'   Used only as parameters to GetAsyncKeyState() and GetKeyState().
'   No other API or message will distinguish left and right keys in this way.
'  /
Private Const VK_LSHIFT = &HA0
Private Const VK_RSHIFT = &HA1
Private Const VK_LCONTROL = &HA2
Private Const VK_RCONTROL = &HA3
Private Const VK_LMENU = &HA4
Private Const VK_RMENU = &HA5

Private Const VK_ATTN = &HF6
Private Const VK_CRSEL = &HF7
Private Const VK_EXSEL = &HF8
Private Const VK_EREOF = &HF9
Private Const VK_PLAY = &HFA
Private Const VK_ZOOM = &HFB
Private Const VK_NONAME = &HFC
Private Const VK_PA1 = &HFD
Private Const VK_OEM_CLEAR = &HFE

' SetWindowsHook() codes
Private Const WH_MIN = (-1)
Private Const WH_MSGFILTER = (-1)
Private Const WH_JOURNALRECORD = 0
Private Const WH_JOURNALPLAYBACK = 1
Private Const WH_KEYBOARD = 2
Private Const WH_GETMESSAGE = 3
Private Const WH_CALLWNDPROC = 4
Private Const WH_CBT = 5
Private Const WH_SYSMSGFILTER = 6
Private Const WH_MOUSE = 7
Private Const WH_HARDWARE = 8
Private Const WH_DEBUG = 9
Private Const WH_SHELL = 10
Private Const WH_FOREGROUNDIDLE = 11
Private Const WH_MAX = 11

' Hook Codes
Private Const HC_ACTION = 0
Private Const HC_GETNEXT = 1
Private Const HC_SKIP = 2
Private Const HC_NOREMOVE = 3
Private Const HC_NOREM = HC_NOREMOVE
Private Const HC_SYSMODALON = 4
Private Const HC_SYSMODALOFF = 5

' CBT Hook Codes
Private Const HCBT_MOVESIZE = 0
Private Const HCBT_MINMAX = 1
Private Const HCBT_QS = 2
Private Const HCBT_CREATEWND = 3
Private Const HCBT_DESTROYWND = 4
Private Const HCBT_ACTIVATE = 5
Private Const HCBT_CLICKSKIPPED = 6
Private Const HCBT_KEYSKIPPED = 7
Private Const HCBT_SYSCOMMAND = 8
Private Const HCBT_SETFOCUS = 9

' HCBT_ACTIVATE structure pointed to by lParam
Type CBTACTIVATESTRUCT
        fMouse As Long
        hWndActive As Long
End Type

' WH_MSGFILTER Filter Proc Codes
Private Const MSGF_DIALOGBOX = 0
Private Const MSGF_MESSAGEBOX = 1
Private Const MSGF_MENU = 2
Private Const MSGF_MOVE = 3
Private Const MSGF_SIZE = 4
Private Const MSGF_SCROLLBAR = 5
Private Const MSGF_NEXTWINDOW = 6
Private Const MSGF_MAINLOOP = 8
Private Const MSGF_MAX = 8
Private Const MSGF_USER = 4096

Private Const HSHELL_WINDOWCREATED = 1
Private Const HSHELL_WINDOWDESTROYED = 2
Private Const HSHELL_ACTIVATESHELLWINDOW = 3

' Message Structure used in Journaling
Type EVENTMSG
        message As Long
        paramL As Long
        paramH As Long
        time As Long
        hWnd As Long
End Type

Type CWPSTRUCT
        lParam As Long
        wParam As Long
        message As Long
        hWnd As Long
End Type

Type DEBUGHOOKINFO
        hModuleHook As Long
        Reserved As Long
        lParam As Long
        wParam As Long
        code As Long
End Type


Type MOUSEHOOKSTRUCT
        pt As POINTAPI
        hWnd As Long
        wHitTestCode As Long
        dwExtraInfo As Long
End Type

' Keyboard Layout API
Private Const HKL_PREV = 0
Private Const HKL_NEXT = 1

Private Const KLF_ACTIVATE = &H1
Private Const KLF_SUBSTITUTE_OK = &H2
Private Const KLF_UNLOADPREVIOUS = &H4
Private Const KLF_REORDER = &H8

' Size of KeyboardLayoutName (number of characters), including nul terminator
Private Const KL_NAMELENGTH = 9

Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal Flags As Long) As Long
Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal Flags As Long) As Long
Declare Function UnloadKeyboardLayout Lib "user32" (ByVal hkl As Long) As Long
Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long

' Desktop-specific access flags
Private Const DESKTOP_READOBJECTS = &H1&
Private Const DESKTOP_CREATEWINDOW = &H2&
Private Const DESKTOP_CREATEMENU = &H4&
Private Const DESKTOP_HOOKCONTROL = &H8&
Private Const DESKTOP_JOURNALRECORD = &H10&
Private Const DESKTOP_JOURNALPLAYBACK = &H20&
Private Const DESKTOP_ENUMERATE = &H40&
Private Const DESKTOP_WRITEOBJECTS = &H80&

Declare Function GetThreadDesktop Lib "user32" (ByVal dwThread As Long) As Long

' Windowstation-specific access flags
Private Const WINSTA_ENUMDESKTOPS = &H1&
Private Const WINSTA_READATTRIBUTES = &H2&
Private Const WINSTA_ACCESSCLIPBOARD = &H4&
Private Const WINSTA_CREATEDESKTOP = &H8&
Private Const WINSTA_WRITEATTRIBUTES = &H10&
Private Const WINSTA_ACCESSPrivateATOMS = &H20&
Private Const WINSTA_EXITWINDOWS = &H40&
Private Const WINSTA_ENUMERATE = &H100&
Private Const WINSTA_READSCREEN = &H200&

Declare Function GetProcessWindowStation Lib "user32" () As Long
Declare Function SetUserObjectSecurity Lib "user32" (ByVal hObj As Long, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR) As Long
Declare Function GetUserObjectSecurity Lib "user32" (ByVal hObj As Long, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long

' Message structure

' Window field offsets for GetWindowLong() and GetWindowWord()
Private Const GWL_WNDPROC = (-4)
Private Const GWL_HINSTANCE = (-6)
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_USERDATA = (-21)
Private Const GWL_ID = (-12)

' Class field offsets for GetClassLong() and GetClassWord()
Private Const GCL_MENUNAME = (-8)
Private Const GCL_HBRBACKGROUND = (-10)
Private Const GCL_HCURSOR = (-12)
Private Const GCL_HICON = (-14)
Private Const GCL_HMODULE = (-16)
Private Const GCL_CBWNDEXTRA = (-18)
Private Const GCL_CBCLSEXTRA = (-20)
Private Const GCL_WNDPROC = (-24)
Private Const GCL_STYLE = (-26)
Private Const GCW_ATOM = (-32)

' Window Messages
Private Const WM_NULL = &H0
Private Const WM_CREATE = &H1
Private Const WM_DESTROY = &H2
Private Const WM_MOVE = &H3
Private Const WM_SIZE = &H5

Private Const WM_ACTIVATE = &H6
'
'  WM_ACTIVATE state values

Private Const WA_INACTIVE = 0
Private Const WA_ACTIVE = 1
Private Const WA_CLICKACTIVE = 2

Private Const WM_SETFOCUS = &H7
Private Const WM_KILLFOCUS = &H8
Private Const WM_ENABLE = &HA
Private Const WM_SETREDRAW = &HB
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_PAINT = &HF
Private Const WM_CLOSE = &H10
Private Const WM_QUERYENDSESSION = &H11
Private Const WM_QUIT = &H12
Private Const WM_QUERYOPEN = &H13
Private Const WM_ERASEBKGND = &H14
Private Const WM_SYSCOLORCHANGE = &H15
Private Const WM_ENDSESSION = &H16
Private Const WM_SHOWWINDOW = &H18
Private Const WM_WININICHANGE = &H1A
Private Const WM_DEVMODECHANGE = &H1B
Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_FONTCHANGE = &H1D
Private Const WM_TIMECHANGE = &H1E
Private Const WM_CANCELMODE = &H1F
Private Const WM_SETCURSOR = &H20
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_CHILDACTIVATE = &H22
Private Const WM_QUEUESYNC = &H23

Private Const WM_GETMINMAXINFO = &H24

Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

Private Const WM_PAINTICON = &H26
Private Const WM_ICONERASEBKGND = &H27
Private Const WM_NEXTDLGCTL = &H28
Private Const WM_SPOOLERSTATUS = &H2A
Private Const WM_DRAWITEM = &H2B
Private Const WM_MEASUREITEM = &H2C
Private Const WM_DELETEITEM = &H2D
Private Const WM_VKEYTOITEM = &H2E
Private Const WM_CHARTOITEM = &H2F
Private Const WM_SETFONT = &H30
Private Const WM_GETFONT = &H31
Private Const WM_SETHOTKEY = &H32
Private Const WM_GETHOTKEY = &H33
Private Const WM_QUERYDRAGICON = &H37
Private Const WM_COMPAREITEM = &H39
Private Const WM_COMPACTING = &H41
Private Const WM_OTHERWINDOWCREATED = &H42               '  no longer suported
Private Const WM_OTHERWINDOWDESTROYED = &H43             '  no longer suported
Private Const WM_COMMNOTIFY = &H44                       '  no longer suported

' notifications passed in low word of lParam on WM_COMMNOTIFY messages
Private Const CN_RECEIVE = &H1
Private Const CN_TRANSMIT = &H2
Private Const CN_EVENT = &H4

Private Const WM_WINDOWPOSCHANGING = &H46
Private Const WM_WINDOWPOSCHANGED = &H47

Private Const WM_POWER = &H48
'
'  wParam for WM_POWER window message and DRV_POWER driver notification

Private Const PWR_OK = 1
Private Const PWR_FAIL = (-1)
Private Const PWR_SUSPENDREQUEST = 1
Private Const PWR_SUSPENDRESUME = 2
Private Const PWR_CRITICALRESUME = 3

Private Const WM_COPYDATA = &H4A
Private Const WM_CANCELJOURNAL = &H4B

Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As Long
End Type

Private Const WM_NCCREATE = &H81
Private Const WM_NCDESTROY = &H82
Private Const WM_NCCALCSIZE = &H83
Private Const WM_NCHITTEST = &H84
Private Const WM_NCPAINT = &H85
Private Const WM_NCACTIVATE = &H86
Private Const WM_GETDLGCODE = &H87
Private Const WM_NCMOUSEMOVE = &HA0
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCLBUTTONUP = &HA2
Private Const WM_NCLBUTTONDBLCLK = &HA3
Private Const WM_NCRBUTTONDOWN = &HA4
Private Const WM_NCRBUTTONUP = &HA5
Private Const WM_NCRBUTTONDBLCLK = &HA6
Private Const WM_NCMBUTTONDOWN = &HA7
Private Const WM_NCMBUTTONUP = &HA8
Private Const WM_NCMBUTTONDBLCLK = &HA9

Private Const WM_KEYFIRST = &H100
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102
Private Const WM_DEADCHAR = &H103
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Const WM_SYSCHAR = &H106
Private Const WM_SYSDEADCHAR = &H107
Private Const WM_KEYLAST = &H108
Private Const WM_INITDIALOG = &H110
Private Const WM_COMMAND = &H111
Private Const WM_SYSCOMMAND = &H112
Private Const WM_TIMER = &H113
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const WM_INITMENU = &H116
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_MENUSELECT = &H11F
Private Const WM_MENUCHAR = &H120
Private Const WM_ENTERIDLE = &H121

Private Const WM_CTLCOLORMSGBOX = &H132
Private Const WM_CTLCOLOREDIT = &H133
Private Const WM_CTLCOLORLISTBOX = &H134
Private Const WM_CTLCOLORBTN = &H135
Private Const WM_CTLCOLORDLG = &H136
Private Const WM_CTLCOLORSCROLLBAR = &H137
Private Const WM_CTLCOLORSTATIC = &H138

Private Const WM_MOUSEFIRST = &H200
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSELAST = &H209

Private Const WM_PARENTNOTIFY = &H210
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_EXITMENULOOP = &H212
Private Const WM_MDICREATE = &H220
Private Const WM_MDIDESTROY = &H221
Private Const WM_MDIACTIVATE = &H222
Private Const WM_MDIRESTORE = &H223
Private Const WM_MDINEXT = &H224
Private Const WM_MDIMAXIMIZE = &H225
Private Const WM_MDITILE = &H226
Private Const WM_MDICASCADE = &H227
Private Const WM_MDIICONARRANGE = &H228
Private Const WM_MDIGETACTIVE = &H229
Private Const WM_MDISETMENU = &H230
Private Const WM_DROPFILES = &H233
Private Const WM_MDIREFRESHMENU = &H234


Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
Private Const WM_UNDO = &H304
Private Const WM_RENDERFORMAT = &H305
Private Const WM_RENDERALLFORMATS = &H306
Private Const WM_DESTROYCLIPBOARD = &H307
Private Const WM_DRAWCLIPBOARD = &H308
Private Const WM_PAINTCLIPBOARD = &H309
Private Const WM_VSCROLLCLIPBOARD = &H30A
Private Const WM_SIZECLIPBOARD = &H30B
Private Const WM_ASKCBFORMATNAME = &H30C
Private Const WM_CHANGECBCHAIN = &H30D
Private Const WM_HSCROLLCLIPBOARD = &H30E
Private Const WM_QUERYNEWPALETTE = &H30F
Private Const WM_PALETTEISCHANGING = &H310
Private Const WM_PALETTECHANGED = &H311
Private Const WM_HOTKEY = &H312

Private Const WM_PENWINFIRST = &H380
Private Const WM_PENWINLAST = &H38F

' NOTE: All Message Numbers below 0x0400 are RESERVED.

' Private Window Messages Start Here:
Private Const WM_USER = &H400

' WM_SYNCTASK Commands
Private Const ST_BEGINSWP = 0
Private Const ST_ENDSWP = 1

' WM_NCHITTEST and MOUSEHOOKSTRUCT Mouse Position Codes
Private Const HTERROR = (-2)
Private Const HTTRANSPARENT = (-1)
Private Const HTNOWHERE = 0
Private Const HTCLIENT = 1
Private Const HTCAPTION = 2
Private Const HTSYSMENU = 3
Private Const HTGROWBOX = 4
Private Const HTSIZE = HTGROWBOX
Private Const HTMENU = 5
Private Const HTHSCROLL = 6
Private Const HTVSCROLL = 7
Private Const HTMINBUTTON = 8
Private Const HTMAXBUTTON = 9
Private Const HTLEFT = 10
Private Const HTRIGHT = 11
Private Const HTTOP = 12
Private Const HTTOPLEFT = 13
Private Const HTTOPRIGHT = 14
Private Const HTBOTTOM = 15
Private Const HTBOTTOMLEFT = 16
Private Const HTBOTTOMRIGHT = 17
Private Const HTBORDER = 18
Private Const HTREDUCE = HTMINBUTTON
Private Const HTZOOM = HTMAXBUTTON
Private Const HTSIZEFIRST = HTLEFT
Private Const HTSIZELAST = HTBOTTOMRIGHT

'  SendMessageTimeout values
Private Const SMTO_NORMAL = &H0
Private Const SMTO_BLOCK = &H1
Private Const SMTO_ABORTIFHUNG = &H2

' WM_MOUSEACTIVATE Return Codes
Private Const MA_ACTIVATE = 1
Private Const MA_ACTIVATEANDEAT = 2
Private Const MA_NOACTIVATE = 3
Private Const MA_NOACTIVATEANDEAT = 4

Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

' WM_SIZE message wParam values
Private Const SIZE_RESTORED = 0
Private Const SIZE_MINIMIZED = 1
Private Const SIZE_MAXIMIZED = 2
Private Const SIZE_MAXSHOW = 3
Private Const SIZE_MAXHIDE = 4

' Obsolete constant names
Private Const SIZENORMAL = SIZE_RESTORED
Private Const SIZEICONIC = SIZE_MINIMIZED
Private Const SIZEFULLSCREEN = SIZE_MAXIMIZED
Private Const SIZEZOOMSHOW = SIZE_MAXSHOW
Private Const SIZEZOOMHIDE = SIZE_MAXHIDE

' WM_WINDOWPOSCHANGING/CHANGED struct pointed to by lParam
Type WINDOWPOS
        hWnd As Long
        hWndInsertAfter As Long
        X As Long
        Y As Long
        cx As Long
        cy As Long
        Flags As Long
End Type

' WM_NCCALCSIZE return flags
Private Const WVR_ALIGNTOP = &H10
Private Const WVR_ALIGNLEFT = &H20
Private Const WVR_ALIGNBOTTOM = &H40
Private Const WVR_ALIGNRIGHT = &H80
Private Const WVR_HREDRAW = &H100
Private Const WVR_VREDRAW = &H200
Private Const WVR_REDRAW = (WVR_HREDRAW Or WVR_VREDRAW)
Private Const WVR_VALIDRECTS = &H400

' Key State Masks for Mouse Messages
Private Const MK_LBUTTON = &H1
Private Const MK_RBUTTON = &H2
Private Const MK_SHIFT = &H4
Private Const MK_CONTROL = &H8
Private Const MK_MBUTTON = &H10

' Window Styles
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_CHILD = &H40000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_DISABLED = &H8000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_GROUP = &H20000
Private Const WS_TABSTOP = &H10000

Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000

Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

'
'   Common Window Styles
'  /


Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)

Private Const WS_CHILDWINDOW = (WS_CHILD)

' Extended Window Styles
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_TRANSPARENT = &H20&

' Class styles
Private Const CS_VREDRAW = &H1
Private Const CS_HREDRAW = &H2
Private Const CS_KEYCVTWINDOW = &H4
Private Const CS_DBLCLKS = &H8
Private Const CS_OWNDC = &H20
Private Const CS_CLASSDC = &H40
Private Const CS_PARENTDC = &H80
Private Const CS_NOKEYCVT = &H100
Private Const CS_NOCLOSE = &H200
Private Const CS_SAVEBITS = &H800
Private Const CS_BYTEALIGNCLIENT = &H1000
Private Const CS_BYTEALIGNWINDOW = &H2000
Private Const CS_PrivateCLASS = &H4000

' Predefined Clipboard Formats
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2
Private Const CF_METAFILEPICT = 3
Private Const CF_SYLK = 4
Private Const CF_DIF = 5
Private Const CF_TIFF = 6
Private Const CF_OEMTEXT = 7
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_RIFF = 11
Private Const CF_WAVE = 12
Private Const CF_UNICODETEXT = 13
Private Const CF_ENHMETAFILE = 14

Private Const CF_OWNERDISPLAY = &H80
Private Const CF_DSPTEXT = &H81
Private Const CF_DSPBITMAP = &H82
Private Const CF_DSPMETAFILEPICT = &H83
Private Const CF_DSPENHMETAFILE = &H8E

' "Private" formats don't get GlobalFree()'d
Private Const CF_PRIVATEFIRST = &H200
Private Const CF_PRIVATELAST = &H2FF

' "GDIOBJ" formats do get DeleteObject()'d
Private Const CF_GDIOBJFIRST = &H300
Private Const CF_GDIOBJLAST = &H3FF

'  Defines for the fVirt field of the Accelerator table structure.
Private Const FVIRTKEY = True          '  Assumed to be == TRUE
Private Const FNOINVERT = &H2
Private Const FSHIFT = &H4
Private Const FCONTROL = &H8
Private Const FALT = &H10

Type ACCEL
        fVirt As Byte
        key As Integer
        cmd As Integer
End Type

Type PAINTSTRUCT
        hdc As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved As Byte
End Type

Type CREATESTRUCT
        lpCreateParams As Long
        hInstance As Long
        hMenu As Long
        hwndParent As Long
        cy As Long
        cx As Long
        Y As Long
        X As Long
        style As Long
        lpszName As String
        lpszClass As String
        ExStyle As Long
End Type

' HCBT_CREATEWND parameters pointed to by lParam
Type CBT_CREATEWND
        lpcs As CREATESTRUCT
        hWndInsertAfter As Long
End Type

Type WINDOWPLACEMENT
        Length As Long
        Flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type

Private Const WPF_SETMINPOSITION = &H1
Private Const WPF_RESTORETOMAXIMIZED = &H2

' Owner draw control types
Private Const ODT_MENU = 1
Private Const ODT_LISTBOX = 2
Private Const ODT_COMBOBOX = 3
Private Const ODT_BUTTON = 4

' Owner draw actions
Private Const ODA_DRAWENTIRE = &H1
Private Const ODA_SELECT = &H2
Private Const ODA_FOCUS = &H4

' Owner draw state
Private Const ODS_SELECTED = &H1
Private Const ODS_GRAYED = &H2
Private Const ODS_DISABLED = &H4
Private Const ODS_CHECKED = &H8
Private Const ODS_FOCUS = &H10

' MEASUREITEMSTRUCT for ownerdraw
Type MEASUREITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemWidth As Long
        itemHeight As Long
        itemData As Long
End Type

' DRAWITEMSTRUCT for ownerdraw
Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hdc As Long
        rcItem As RECT
        itemData As Long
End Type

' DELETEITEMSTRUCT for ownerdraw
Type DELETEITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        hwndItem As Long
        itemData As Long
End Type

' COMPAREITEMSTRUCT for ownerdraw sorting
Type COMPAREITEMSTRUCT
        CtlType As Long
        CtlID As Long
        hwndItem As Long
        itemID1 As Long
        itemData1 As Long
        itemID2 As Long
        itemData2 As Long
End Type

' Message Function Templates
Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

' PeekMessage() Options
Private Const PM_NOREMOVE = &H0
Private Const PM_REMOVE = &H1
Private Const PM_NOYIELD = &H2

Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long

Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4

Private Const IDHOT_SNAPWINDOW = (-1)    '  SHIFT-PRINTSCRN
Private Const IDHOT_SNAPDESKTOP = (-2)    '  PRINTSCRN

Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4

Private Const READAPI = 0        '  Flags for _lopen
Private Const WRITEAPI = 1
Private Const READ_WRITE = 2

Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Declare Function GetMessagePos Lib "user32" () As Long
Declare Function GetMessageTime Lib "user32" () As Long
Declare Function GetMessageExtraInfo Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessageCallback Lib "user32" Alias "SendMessageCallbackA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lpResultCallBack As Long, ByVal dwData As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function PostThreadMessage Lib "user32" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Special HWND value for use with PostMessage and SendMessage
Private Const HWND_BROADCAST = &HFFFF&


Type WNDCLASS
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Declare Function ReplyMessage Lib "user32" (ByVal lReply As Long) As Long
Declare Function WaitMessage Lib "user32" () As Long
Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Declare Function InSendMessage Lib "user32" () As Long

Declare Function GetDoubleClickTime Lib "user32" () As Long
Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Declare Function RegisterClass Lib "user32" (Class As WNDCLASS) As Long
Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long

Private Const CW_USEDEFAULT = &H80000000
Private Const HWND_DESKTOP = 0

Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function IsChild Lib "user32" (ByVal hwndParent As Long, ByVal hWnd As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Declare Function ShowOwnedPopups Lib "user32" (ByVal hWnd As Long, ByVal fShow As Long) As Long

Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long

Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Declare Function AnyPopup Lib "user32" () As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long

' SetWindowPos Flags
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' SetWindowPos() hwndInsertAfter values
Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Type DLGTEMPLATE
        style As Long
        dwExtendedStyle As Long
        cdit As Integer
        X As Integer
        Y As Integer
        cx As Integer
        cy As Integer
End Type

Type DLGITEMTEMPLATE
        style As Long
        dwExtendedStyle As Long
        X As Integer
        Y As Integer
        cx As Integer
        cy As Integer
        id As Integer
End Type

Declare Function EndDialog Lib "user32" (ByVal hDlg As Long, ByVal nResult As Long) As Long
Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Declare Function SetDlgItemInt Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wValue As Long, ByVal bSigned As Long) As Long
Declare Function GetDlgItemInt Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpTranslated As Long, ByVal bSigned As Long) As Long
Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function CheckDlgButton Lib "user32" Alias "CheckDLGButtonA" (ByVal hDlg As Long, ByVal nIDButton As Long, ByVal wCheck As Long) As Long
Declare Function CheckRadioButton Lib "user32" Alias "CheckRadioButtonA" (ByVal hDlg As Long, ByVal nIDFirstButton As Long, ByVal nIDLastButton As Long, ByVal nIDCheckButton As Long) As Long
Declare Function IsDlgButtonChecked Lib "user32" (ByVal hDlg As Long, ByVal nIDButton As Long) As Long
Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetNextDlgGroupItem Lib "user32" (ByVal hDlg As Long, ByVal hCtl As Long, ByVal bPrevious As Long) As Long
Declare Function GetNextDlgTabItem Lib "user32" (ByVal hDlg As Long, ByVal hCtl As Long, ByVal bPrevious As Long) As Long
Declare Function GetDlgCtrlID Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDialogBaseUnits Lib "user32" () As Long
Declare Function DefDlgProc Lib "user32" Alias "DefDlgProcA" (ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const DLGWINDOWEXTRA = 30        '  Window extra bytes needed for private dialog classes

Declare Function CallMsgFilter Lib "user32" Alias "CallMsgFilterA" (lpMsg As Msg, ByVal nCode As Long) As Long

' Clipboard Manager Functions
Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function GetClipboardOwner Lib "user32" () As Long
Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetClipboardViewer Lib "user32" () As Long
Declare Function ChangeClipboardChain Lib "user32" (ByVal hWnd As Long, ByVal hWndNext As Long) As Long
Declare Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As Long
Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Declare Function CountClipboardFormats Lib "user32" () As Long
Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Declare Function GetPriorityClipboardFormat Lib "user32" (lpPriorityList As Long, ByVal nCount As Long) As Long
Declare Function GetOpenClipboardWindow Lib "user32" () As Long
Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Declare Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Declare Function CharUpperBuff Lib "user32" Alias "CharUpperBuffA" (ByVal lpsz As String, ByVal cchLength As Long) As Long
Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Declare Function CharLowerBuff Lib "user32" Alias "CharLowerBuffA" (ByVal lpsz As String, ByVal cchLength As Long) As Long
Declare Function CharNext Lib "user32" Alias "CharNextA" (ByVal lpsz As String) As String
Declare Function CharPrev Lib "user32" Alias "CharPrevA" (ByVal lpszStart As String, ByVal lpszCurrent As String) As String

' Language dependent Routines
Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long

Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetFocus Lib "user32" () As Long
Declare Function GetActiveWindow Lib "user32" () As Long

' Keyboard Information Routines
Declare Function GetKBCodePage Lib "user32" () As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
Declare Function ToUnicode Lib "user32" (ByVal wVirtKey As Long, ByVal wScanCode As Long, lpKeyState As Byte, ByVal pwszBuff As String, ByVal cchBuff As Long, ByVal wFlags As Long) As Long

Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Long) As Long
Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Private Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move

Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

Declare Function GetInputState Lib "user32" () As Long
Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long
Declare Function GetCapture Lib "user32" () As Long
Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long

Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long

' GetQueueStatus flags
Private Const QS_KEY = &H1
Private Const QS_MOUSEMOVE = &H2
Private Const QS_MOUSEBUTTON = &H4
Private Const QS_POSTMESSAGE = &H8
Private Const QS_TIMER = &H10
Private Const QS_PAINT = &H20
Private Const QS_SENDMESSAGE = &H40
Private Const QS_HOTKEY = &H80

Private Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)

Private Const QS_INPUT = (QS_MOUSE Or QS_KEY)

Private Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)

Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

' Windows Functions
Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long

Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long

Declare Function LoadAccelerators Lib "user32" Alias "LoadAcceleratorsA" (ByVal hInstance As Long, ByVal lpTableName As String) As Long
Declare Function CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableA" (lpaccl As ACCEL, ByVal cEntries As Long) As Long
Declare Function DestroyAcceleratorTable Lib "user32" (ByVal haccel As Long) As Long
Declare Function CopyAcceleratorTable Lib "user32" Alias "CopyAcceleratorTableA" (ByVal hAccelSrc As Long, lpAccelDst As ACCEL, ByVal cAccelEntries As Long) As Long
Declare Function TranslateAccelerator Lib "user32" Alias "TranslateAcceleratorA" (ByVal hWnd As Long, ByVal hAccTable As Long, lpMsg As Msg) As Long

' GetSystemMetrics() codes
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CYCAPTION = 4
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CXDLGFRAME = 7
Private Const SM_CYDLGFRAME = 8
Private Const SM_CYVTHUMB = 9
Private Const SM_CXHTHUMB = 10
Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
Private Const SM_CXCURSOR = 13
Private Const SM_CYCURSOR = 14
Private Const SM_CYMENU = 15
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17
Private Const SM_CYKANJIWINDOW = 18
Private Const SM_MOUSEPRESENT = 19
Private Const SM_CYVSCROLL = 20
Private Const SM_CXHSCROLL = 21
Private Const SM_DEBUG = 22
Private Const SM_SWAPBUTTON = 23
Private Const SM_RESERVED1 = 24
Private Const SM_RESERVED2 = 25
Private Const SM_RESERVED3 = 26
Private Const SM_RESERVED4 = 27
Private Const SM_CXMIN = 28
Private Const SM_CYMIN = 29
Private Const SM_CXSIZE = 30
Private Const SM_CYSIZE = 31
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33
Private Const SM_CXMINTRACK = 34
Private Const SM_CYMINTRACK = 35
Private Const SM_CXDOUBLECLK = 36
Private Const SM_CYDOUBLECLK = 37
Private Const SM_CXICONSPACING = 38
Private Const SM_CYICONSPACING = 39
Private Const SM_MENUDROPALIGNMENT = 40
Private Const SM_PENWINDOWS = 41
Private Const SM_DBCSENABLED = 42
Private Const SM_CMOUSEBUTTONS = 43
Private Const SM_CMETRICS = 44
Private Const SM_CXSIZEFRAME = SM_CXFRAME
Private Const SM_CYSIZEFRAME = SM_CYFRAME
Private Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Private Const SM_CYFIXEDFRAME = SM_CYDLGFRAME

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As String) As Long
Declare Function LoadMenuIndirect Lib "user32" Alias "LoadMenuIndirectA" (ByVal lpMenuTemplate As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Declare Function HiliteMenuItem Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function CreateMenu Lib "user32" () As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long

' Flags for TrackPopupMenu
Private Const TPM_LEFTBUTTON = &H0&
Private Const TPM_RIGHTBUTTON = &H2&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_CENTERALIGN = &H4&
Private Const TPM_RIGHTALIGN = &H8&

Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long

' DrawText() Format Flags
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const DT_NOCLIP = &H100
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000

Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function TabbedTextOut Lib "user32" Alias "TabbedTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long
Declare Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long

Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long

Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDCEx Lib "user32" (ByVal hWnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long

Private Const DCX_WINDOW = &H1&
Private Const DCX_CACHE = &H2&
Private Const DCX_NORESETATTRS = &H4&
Private Const DCX_CLIPCHILDREN = &H8&
Private Const DCX_CLIPSIBLINGS = &H10&
Private Const DCX_PARENTCLIP = &H20&

Private Const DCX_EXCLUDERGN = &H40&
Private Const DCX_INTERSECTRGN = &H80&

Private Const DCX_EXCLUDEUPDATE = &H100&
Private Const DCX_INTERSECTUPDATE = &H200&

Private Const DCX_LOCKWINDOWUPDATE = &H400&

Private Const DCX_NORECOMPUTE = &H100000
Private Const DCX_VALIDATE = &H200000

Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function GetUpdateRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal fErase As Long) As Long
Declare Function ExcludeUpdateRgn Lib "user32" (ByVal hdc As Long, ByVal hWnd As Long) As Long
Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function InvalidateRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bErase As Long) As Long
Declare Function ValidateRgn Lib "user32" Alias "ValidateRgn" (ByVal hwnd As Long, ByVal hRgn As Long) As Long

Declare Function RedrawWindow Lib "user32" Alias "RedrawWindow" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Const RDW_INVALIDATE = &H1
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_ERASE = &H4

Private Const RDW_VALIDATE = &H8
Private Const RDW_NOINTERNALPAINT = &H10
Private Const RDW_NOERASE = &H20

Private Const RDW_NOCHILDREN = &H40
Private Const RDW_ALLCHILDREN = &H80

Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASENOW = &H200

Private Const RDW_FRAME = &H400
Private Const RDW_NOFRAME = &H800

Declare Function LockWindowUpdate Lib "user32" Alias "LockWindowUpdate" (ByVal hwndLock As Long) As Long

Declare Function ScrollWindow Lib "user32" Alias "ScrollWindow" (ByVal hWnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long
Declare Function ScrollDC Lib "user32" Alias "ScrollDC" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Declare Function ScrollWindowEx Lib "user32" Alias "ScrollWindowEx" (ByVal hwnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT, ByVal fuScroll As Long) As Long

Private Const SW_SCROLLCHILDREN = &H1
Private Const SW_INVALIDATE = &H2
Private Const SW_ERASE = &H4

Declare Function SetScrollPos Lib "user32" Alias "SetScrollPos" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Declare Function GetScrollPos Lib "user32" Alias "GetScrollPos" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Declare Function SetScrollRange Lib "user32" Alias "SetScrollRange" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Declare Function GetScrollRange Lib "user32" Alias "GetScrollRange" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Declare Function ShowScrollBar Lib "user32" Alias "ShowScrollBar" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Declare Function EnableScrollBar Lib "user32" Alias "EnableScrollBar" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long

' EnableScrollBar() flags
Private Const ESB_ENABLE_BOTH = &H0
Private Const ESB_DISABLE_BOTH = &H3

Private Const ESB_DISABLE_LEFT = &H1
Private Const ESB_DISABLE_RIGHT = &H2

Private Const ESB_DISABLE_UP = &H1
Private Const ESB_DISABLE_DOWN = &H2

Private Const ESB_DISABLE_LTUP = ESB_DISABLE_LEFT
Private Const ESB_DISABLE_RTDN = ESB_DISABLE_RIGHT

Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Declare Function GetClientRect Lib "user32" Alias "GetClientRect" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function AdjustWindowRect Lib "user32" Alias "AdjustWindowRect" (lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
Declare Function AdjustWindowRectEx Lib "user32" Alias "AdjustWindowRectEx" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long

' MessageBox() Flags
Private Const MB_OK = &H0&
Private Const MB_OKCANCEL = &H1&
Private Const MB_ABORTRETRYIGNORE = &H2&
Private Const MB_YESNOCANCEL = &H3&
Private Const MB_YESNO = &H4&
Private Const MB_RETRYCANCEL = &H5&

Private Const MB_ICONHAND = &H10&
Private Const MB_ICONQUESTION = &H20&
Private Const MB_ICONEXCLAMATION = &H30&
Private Const MB_ICONASTERISK = &H40&

Private Const MB_ICONINFORMATION = MB_ICONASTERISK
Private Const MB_ICONSTOP = MB_ICONHAND

Private Const MB_DEFBUTTON1 = &H0&
Private Const MB_DEFBUTTON2 = &H100&
Private Const MB_DEFBUTTON3 = &H200&

Private Const MB_APPLMODAL = &H0&
Private Const MB_SYSTEMMODAL = &H1000&
Private Const MB_TASKMODAL = &H2000&

Private Const MB_NOFOCUS = &H8000&
Private Const MB_SETFOREGROUND = &H10000
Private Const MB_DEFAULT_DESKTOP_ONLY = &H20000

Private Const MB_TYPEMASK = &HF&
Private Const MB_ICONMASK = &HF0&
Private Const MB_DEFMASK = &HF00&
Private Const MB_MODEMASK = &H3000&
Private Const MB_MISCMASK = &HC000&

Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Declare Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
Declare Function MessageBeep Lib "user32" Alias "MessageBeep" (ByVal wType As Long) As Long

Declare Function ShowCursor Lib "user32" Alias "ShowCursor" (ByVal bShow As Long) As Long
Declare Function SetCursorPos Lib "user32" Alias "SetCursorPos" (ByVal x As Long, ByVal y As Long) As Long
Declare Function SetCursor Lib "user32" Alias "SetCursor" (ByVal hCursor As Long) As Long
Declare Function GetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
Declare Function ClipCursor Lib "user32" Alias "ClipCursor" (lpRect As Any) As Long
Declare Function GetCursor Lib "user32" Alias "GetCursor" () As Long
Declare Function GetClipCursor Lib "user32" Alias "GetClipCursor" (lprc As RECT) As Long

Declare Function CreateCaret Lib "user32" Alias "CreateCaret" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetCaretBlinkTime Lib "user32" Alias "GetCaretBlinkTime" () As Long
Declare Function SetCaretBlinkTime Lib "user32" Alias "SetCaretBlinkTime" (ByVal wMSeconds As Long) As Long
Declare Function DestroyCaret Lib "user32" Alias "DestroyCaret" () As Long
Declare Function HideCaret Lib "user32" Alias "HideCaret" (ByVal hwnd As Long) As Long
Declare Function ShowCaret Lib "user32" Alias "ShowCaret" (ByVal hwnd As Long) As Long
Declare Function SetCaretPos Lib "user32" Alias "SetCaretPos" (ByVal x As Long, ByVal y As Long) As Long
Declare Function GetCaretPos Lib "user32" Alias "GetCaretPos" (lpPoint As POINTAPI) As Long

Declare Function ClientToScreen Lib "user32" Alias "ClientToScreen" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function ScreenToClient Lib "user32" Alias "ScreenToClient" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Declare Function MapWindowPoints Lib "user32" Alias "MapWindowPoints" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Declare Function WindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function ChildWindowFromPoint Lib "user32" Alias "ChildWindowFromPoint" (ByVal hWnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long

' Color Types
Private Const CTLCOLOR_MSGBOX = 0
Private Const CTLCOLOR_EDIT = 1
Private Const CTLCOLOR_LISTBOX = 2
Private Const CTLCOLOR_BTN = 3
Private Const CTLCOLOR_DLG = 4
Private Const CTLCOLOR_SCROLLBAR = 5
Private Const CTLCOLOR_STATIC = 6
Private Const CTLCOLOR_MAX = 8   '  three bits max

Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_MENU = 4
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_WINDOWTEXT = 8
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const COLOR_BTNHIGHLIGHT = 20

Declare Function GetSysColor Lib "user32" Alias "GetSysColor" (ByVal nIndex As Long) As Long
Declare Function SetSysColors Lib "user32" Alias "SetSysColors" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long

Declare Function DrawFocusRect Lib "user32" Alias "DrawFocusRect" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function FillRect Lib "user32" Alias "FillRect" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function FrameRect Lib "user32" Alias "FrameRect" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function InvertRect Lib "user32" Alias "InvertRect" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" Alias "SetRect" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetRectEmpty Lib "user32" Alias "SetRectEmpty" (lpRect As RECT) As Long
Declare Function CopyRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Declare Function InflateRect Lib "user32" Alias "InflateRect" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Declare Function IntersectRect Lib "user32" Alias "IntersectRect" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Declare Function UnionRect Lib "user32" Alias "UnionRect" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Declare Function SubtractRect Lib "user32" Alias "SubtractRect" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Declare Function OffsetRect Lib "user32" Alias "OffsetRect" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Declare Function IsRectEmpty Lib "user32" Alias "IsRectEmpty" (lpRect As RECT) As Long
Declare Function EqualRect Lib "user32" Alias "EqualRect" (lpRect1 As RECT, lpRect2 As RECT) As Long
Declare Function PtInRect Lib "user32" Alias "PtInRect" (lpRect As RECT, pt As POINTAPI) As Long

Declare Function GetWindowWord Lib "user32" Alias "GetWindowWord" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Declare Function SetWindowWord Lib "user32" Alias "SetWindowWord" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetClassWord Lib "user32" Alias "GetClassWord" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetClassWord Lib "user32" Alias "SetClassWord" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long

Declare Function GetParent Lib "user32" Alias "GetParent" (ByVal hwnd As Long) As Long
Declare Function SetParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetTopWindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long

Declare Function GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Declare Function GetLastActivePopup Lib "user32" Alias "GetLastActivePopup" (ByVal hwndOwnder As Long) As Long

' GetWindow() Constants
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_OWNER = 4
Private Const GW_CHILD = 5
Private Const GW_MAX = 5

Declare Function GetWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" Alias "UnhookWindowsHookEx" (ByVal hHook As Long) As Long
Declare Function CallNextHookEx Lib "user32" Alias "CallNextHookEx" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long

' Menu flags for Add/Check/EnableMenuItem()
Private Const MF_INSERT = &H0&
Private Const MF_CHANGE = &H80&
Private Const MF_APPEND = &H100&
Private Const MF_DELETE = &H200&
Private Const MF_REMOVE = &H1000&

Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&

Private Const MF_SEPARATOR = &H800&

Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_DISABLED = &H2&

Private Const MF_UNCHECKED = &H0&
Private Const MF_CHECKED = &H8&
Private Const MF_USECHECKBITMAPS = &H200&

Private Const MF_STRING = &H0&
Private Const MF_BITMAP = &H4&
Private Const MF_OWNERDRAW = &H100&

Private Const MF_POPUP = &H10&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&

Private Const MF_UNHILITE = &H0&
Private Const MF_HILITE = &H80&

Private Const MF_SYSMENU = &H2000&
Private Const MF_HELP = &H4000&
Private Const MF_MOUSESELECT = &H8000&

' Menu item resource format
Type MENUITEMTEMPLATEHEADER
        versionNumber As Integer
        offset As Integer
End Type

Type MENUITEMTEMPLATE
        mtOption As Integer
        mtID As Integer
        mtString As Byte
End Type

Private Const MF_END = &H80

' System Menu Command Values
Private Const SC_SIZE = &HF000
Private Const SC_MOVE = &HF010
Private Const SC_MINIMIZE = &HF020
Private Const SC_MAXIMIZE = &HF030
Private Const SC_NEXTWINDOW = &HF040
Private Const SC_PREVWINDOW = &HF050
Private Const SC_CLOSE = &HF060
Private Const SC_VSCROLL = &HF070
Private Const SC_HSCROLL = &HF080
Private Const SC_MOUSEMENU = &HF090
Private Const SC_KEYMENU = &HF100
Private Const SC_ARRANGE = &HF110
Private Const SC_RESTORE = &HF120
Private Const SC_TASKLIST = &HF130
Private Const SC_SCREENSAVE = &HF140
Private Const SC_HOTKEY = &HF150

' Obsolete names
Private Const SC_ICON = SC_MINIMIZE
Private Const SC_ZOOM = SC_MAXIMIZE

' Resource Loading Routines
Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Declare Function CreateCursor Lib "user32" Alias "CreateCursor" (ByVal hInstance As Long, ByVal nXhotspot As Long, ByVal nYhotspot As Long, ByVal nWidth As Long, ByVal nHeight As Long, lpANDbitPlane As Any, lpXORbitPlane As Any) As Long
Declare Function DestroyCursor Lib "user32" Alias "DestroyCursor" (ByVal hCursor As Long) As Long
Declare Function CopyCursor Lib "user32" Alias "CopyCursor" (ByVal hcur As Long) As Long

' Standard Cursor IDs
Private Const IDC_ARROW = 32512&
Private Const IDC_IBEAM = 32513&
Private Const IDC_WAIT = 32514&
Private Const IDC_CROSS = 32515&
Private Const IDC_UPARROW = 32516&
Private Const IDC_SIZE = 32640&
Private Const IDC_ICON = 32641&
Private Const IDC_SIZENWSE = 32642&
Private Const IDC_SIZENESW = 32643&
Private Const IDC_SIZEWE = 32644&
Private Const IDC_SIZENS = 32645&
Private Const IDC_SIZEALL = 32646&
Private Const IDC_NO = 32648&
Private Const IDC_APPSTARTING = 32650&

Type ICONINFO
        fIcon As Long
        xHotspot As Long
        yHotspot As Long
        hbmMask As Long
        hbmColor As Long
End Type

Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Declare Function CreateIcon Lib "user32" Alias "CreateIcon" (ByVal hInstance As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Byte, ByVal nBitsPixel As Byte, lpANDbits As Byte, lpXORbits As Byte) As Long
Declare Function DestroyIcon Lib "user32" Alias "DestroyIcon" (ByVal hIcon As Long) As Long
Declare Function LookupIconIdFromDirectory Lib "user32" Alias "LookupIconIdFromDirectory" (presbits As Byte, ByVal fIcon As Long) As Long
Declare Function CreateIconFromResource Lib "user32" Alias "CreateIconFromResource" (presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As Long
Declare Function CreateIconIndirect Lib "user32" Alias "CreateIconIndirect" (piconinfo As ICONINFO) As Long
Declare Function CopyIcon Lib "user32" Alias "CopyIcon" (ByVal hIcon As Long) As Long
Declare Function GetIconInfo Lib "user32" Alias "GetIconInfo" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long

' OEM Resource Ordinal Numbers
Private Const OBM_CLOSE = 32754
Private Const OBM_UPARROW = 32753
Private Const OBM_DNARROW = 32752
Private Const OBM_RGARROW = 32751
Private Const OBM_LFARROW = 32750
Private Const OBM_REDUCE = 32749
Private Const OBM_ZOOM = 32748
Private Const OBM_RESTORE = 32747
Private Const OBM_REDUCED = 32746
Private Const OBM_ZOOMD = 32745
Private Const OBM_RESTORED = 32744
Private Const OBM_UPARROWD = 32743
Private Const OBM_DNARROWD = 32742
Private Const OBM_RGARROWD = 32741
Private Const OBM_LFARROWD = 32740
Private Const OBM_MNARROW = 32739
Private Const OBM_COMBO = 32738
Private Const OBM_UPARROWI = 32737
Private Const OBM_DNARROWI = 32736
Private Const OBM_RGARROWI = 32735
Private Const OBM_LFARROWI = 32734

Private Const OBM_OLD_CLOSE = 32767
Private Const OBM_SIZE = 32766
Private Const OBM_OLD_UPARROW = 32765
Private Const OBM_OLD_DNARROW = 32764
Private Const OBM_OLD_RGARROW = 32763
Private Const OBM_OLD_LFARROW = 32762
Private Const OBM_BTSIZE = 32761
Private Const OBM_CHECK = 32760
Private Const OBM_CHECKBOXES = 32759
Private Const OBM_BTNCORNERS = 32758
Private Const OBM_OLD_REDUCE = 32757
Private Const OBM_OLD_ZOOM = 32756
Private Const OBM_OLD_RESTORE = 32755

Private Const OCR_NORMAL = 32512
Private Const OCR_IBEAM = 32513
Private Const OCR_WAIT = 32514
Private Const OCR_CROSS = 32515
Private Const OCR_UP = 32516
Private Const OCR_SIZE = 32640
Private Const OCR_ICON = 32641
Private Const OCR_SIZENWSE = 32642
Private Const OCR_SIZENESW = 32643
Private Const OCR_SIZEWE = 32644
Private Const OCR_SIZENS = 32645
Private Const OCR_SIZEALL = 32646
Private Const OCR_ICOCUR = 32647
Private Const OCR_NO = 32648 ' not in win3.1

Private Const OIC_SAMPLE = 32512
Private Const OIC_HAND = 32513
Private Const OIC_QUES = 32514
Private Const OIC_BANG = 32515
Private Const OIC_NOTE = 32516

Private Const ORD_LANGDRIVER = 1 '  The ordinal number for the entry point of
                                '  language drivers.

' Standard Icon IDs
Private Const IDI_APPLICATION = 32512&
Private Const IDI_HAND = 32513&
Private Const IDI_QUESTION = 32514&
Private Const IDI_EXCLAMATION = 32515&
Private Const IDI_ASTERISK = 32516&

Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

' Dialog Box Command IDs
Private Const IDOK = 1
Private Const IDCANCEL = 2
Private Const IDABORT = 3
Private Const IDRETRY = 4
Private Const IDIGNORE = 5
Private Const IDYES = 6
Private Const IDNO = 7

' Control Manager Structures and Definitions

' Edit Control Styles
Private Const ES_LEFT = &H0&
Private Const ES_CENTER = &H1&
Private Const ES_RIGHT = &H2&
Private Const ES_MULTILINE = &H4&
Private Const ES_UPPERCASE = &H8&
Private Const ES_LOWERCASE = &H10&
Private Const ES_PASSWORD = &H20&
Private Const ES_AUTOVSCROLL = &H40&
Private Const ES_AUTOHSCROLL = &H80&
Private Const ES_NOHIDESEL = &H100&
Private Const ES_OEMCONVERT = &H400&
Private Const ES_READONLY = &H800&
Private Const ES_WANTRETURN = &H1000&

' Edit Control Notification Codes
Private Const EN_SETFOCUS = &H100
Private Const EN_KILLFOCUS = &H200
Private Const EN_CHANGE = &H300
Private Const EN_UPDATE = &H400
Private Const EN_ERRSPACE = &H500
Private Const EN_MAXTEXT = &H501
Private Const EN_HSCROLL = &H601
Private Const EN_VSCROLL = &H602

' Edit Control Messages
Private Const EM_GETSEL = &HB0
Private Const EM_SETSEL = &HB1
Private Const EM_GETRECT = &HB2
Private Const EM_SETRECT = &HB3
Private Const EM_SETRECTNP = &HB4
Private Const EM_SCROLL = &HB5
Private Const EM_LINESCROLL = &HB6
Private Const EM_SCROLLCARET = &HB7
Private Const EM_GETMODIFY = &HB8
Private Const EM_SETMODIFY = &HB9
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEINDEX = &HBB
Private Const EM_SETHANDLE = &HBC
Private Const EM_GETHANDLE = &HBD
Private Const EM_GETTHUMB = &HBE
Private Const EM_LINELENGTH = &HC1
Private Const EM_REPLACESEL = &HC2
Private Const EM_GETLINE = &HC4
Private Const EM_LIMITTEXT = &HC5
Private Const EM_CANUNDO = &HC6
Private Const EM_UNDO = &HC7
Private Const EM_FMTLINES = &HC8
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_SETTABSTOPS = &HCB
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const EM_EMPTYUNDOBUFFER = &HCD
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_SETREADONLY = &HCF
Private Const EM_SETWORDBREAKPROC = &HD0
Private Const EM_GETWORDBREAKPROC = &HD1
Private Const EM_GETPASSWORDCHAR = &HD2

' EDITWORDBREAKPROC code values
Private Const WB_LEFT = 0
Private Const WB_RIGHT = 1
Private Const WB_ISDELIMITER = 2

' Button Control Styles
Private Const BS_PUSHBUTTON = &H0&
Private Const BS_DEFPUSHBUTTON = &H1&
Private Const BS_CHECKBOX = &H2&
Private Const BS_AUTOCHECKBOX = &H3&
Private Const BS_RADIOBUTTON = &H4&
Private Const BS_3STATE = &H5&
Private Const BS_AUTO3STATE = &H6&
Private Const BS_GROUPBOX = &H7&
Private Const BS_USERBUTTON = &H8&
Private Const BS_AUTORADIOBUTTON = &H9&
Private Const BS_OWNERDRAW = &HB&
Private Const BS_LEFTTEXT = &H20&

' User Button Notification Codes
Private Const BN_CLICKED = 0
Private Const BN_PAINT = 1
Private Const BN_HILITE = 2
Private Const BN_UNHILITE = 3
Private Const BN_DISABLE = 4
Private Const BN_DOUBLECLICKED = 5

' Button Control Messages
Private Const BM_GETCHECK = &HF0
Private Const BM_SETCHECK = &HF1
Private Const BM_GETSTATE = &HF2
Private Const BM_SETSTATE = &HF3
Private Const BM_SETSTYLE = &HF4

' Static Control Constants
Private Const SS_LEFT = &H0&
Private Const SS_CENTER = &H1&
Private Const SS_RIGHT = &H2&
Private Const SS_ICON = &H3&
Private Const SS_BLACKRECT = &H4&
Private Const SS_GRAYRECT = &H5&
Private Const SS_WHITERECT = &H6&
Private Const SS_BLACKFRAME = &H7&
Private Const SS_GRAYFRAME = &H8&
Private Const SS_WHITEFRAME = &H9&
Private Const SS_USERITEM = &HA&
Private Const SS_SIMPLE = &HB&
Private Const SS_LEFTNOWORDWRAP = &HC&
Private Const SS_NOPREFIX = &H80           '  Don't do "&" character translation

' Static Control Mesages
Private Const STM_SETICON = &H170
Private Const STM_GETICON = &H171
Private Const STM_MSGMAX = &H172

Private Const WC_DIALOG = 8002&

'  Get/SetWindowWord/Long offsets for use with WC_DIALOG windows
Private Const DWL_MSGRESULT = 0
Private Const DWL_DLGPROC = 4
Private Const DWL_USER = 8

' Dialog Manager Routines
Declare Function IsDialogMessage Lib "user32" Alias "IsDialogMessageA" (ByVal hDlg As Long, lpMsg As MSG) As Long
Declare Function MapDialogRect Lib "user32" Alias "MapDialogRect" (ByVal hDlg As Long, lpRect As RECT) As Long
Declare Function DlgDirList Lib "user32" Alias "DlgDirListA" (ByVal hDlg As Long, ByVal lpPathSpec As String, ByVal nIDListBox As Long, ByVal nIDStaticPath As Long, ByVal wFileType As Long) As Long

' DlgDirList, DlgDirListComboBox flags values
Private Const DDL_READWRITE = &H0
Private Const DDL_READONLY = &H1
Private Const DDL_HIDDEN = &H2
Private Const DDL_SYSTEM = &H4
Private Const DDL_DIRECTORY = &H10
Private Const DDL_ARCHIVE = &H20

Private Const DDL_POSTMSGS = &H2000
Private Const DDL_DRIVES = &H4000
Private Const DDL_EXCLUSIVE = &H8000

Declare Function DlgDirSelectEx Lib "user32" Alias "DlgDirSelectExA" (ByVal hWndDlg As Long, ByVal lpszPath As String, ByVal cbPath As Long, ByVal idListBox As Long) As Long
Declare Function DlgDirListComboBox Lib "user32" Alias "DlgDirListComboBoxA" (ByVal hDlg As Long, ByVal lpPathSpec As String, ByVal nIDComboBox As Long, ByVal nIDStaticPath As Long, ByVal wFileType As Long) As Long
Declare Function DlgDirSelectComboBoxEx Lib "user32" Alias "DlgDirSelectComboBoxExA" (ByVal hWndDlg As Long, ByVal lpszPath As String, ByVal cbPath As Long, ByVal idComboBox As Long) As Long

' Dialog Styles
Private Const DS_ABSALIGN = &H1&
Private Const DS_SYSMODAL = &H2&
Private Const DS_LOCALEDIT = &H20          '  Edit items get Local storage.
Private Const DS_SETFONT = &H40            '  User specified font for Dlg controls
Private Const DS_MODALFRAME = &H80         '  Can be combined with WS_CAPTION
Private Const DS_NOIDLEMSG = &H100         '  WM_ENTERIDLE message will not be sent
Private Const DS_SETFOREGROUND = &H200     '  not in win3.1

Private Const DM_GETDEFID = WM_USER + 0
Private Const DM_SETDEFID = WM_USER + 1
Private Const DC_HASDEFID = &H534      '0x534B

' Dialog Codes
Private Const DLGC_WANTARROWS = &H1              '  Control wants arrow keys
Private Const DLGC_WANTTAB = &H2                 '  Control wants tab keys
Private Const DLGC_WANTALLKEYS = &H4             '  Control wants all keys
Private Const DLGC_WANTMESSAGE = &H4             '  Pass message to control
Private Const DLGC_HASSETSEL = &H8               '  Understands EM_SETSEL message
Private Const DLGC_DEFPUSHBUTTON = &H10          '  Default pushbutton
Private Const DLGC_UNDEFPUSHBUTTON = &H20        '  Non-default pushbutton
Private Const DLGC_RADIOBUTTON = &H40            '  Radio button
Private Const DLGC_WANTCHARS = &H80              '  Want WM_CHAR messages
Private Const DLGC_STATIC = &H100                '  Static item: don't include
Private Const DLGC_BUTTON = &H2000               '  Button item: can be checked

Private Const LB_CTLCODE = 0&

' Listbox Return Values
Private Const LB_OKAY = 0
Private Const LB_ERR = (-1)
Private Const LB_ERRSPACE = (-2)

' The idStaticPath parameter to DlgDirList can have the following values
' ORed if the list box should show other details of the files along with
' the name of the files;
' all other details also will be returned

' Listbox Notification Codes
Private Const LBN_ERRSPACE = (-2)
Private Const LBN_SELCHANGE = 1
Private Const LBN_DBLCLK = 2
Private Const LBN_SELCANCEL = 3
Private Const LBN_SETFOCUS = 4
Private Const LBN_KILLFOCUS = 5

' Listbox messages
Private Const LB_ADDSTRING = &H180
Private Const LB_INSERTSTRING = &H181
Private Const LB_DELETESTRING = &H182
Private Const LB_SELITEMRANGEEX = &H183
Private Const LB_RESETCONTENT = &H184
Private Const LB_SETSEL = &H185
Private Const LB_SETCURSEL = &H186
Private Const LB_GETSEL = &H187
Private Const LB_GETCURSEL = &H188
Private Const LB_GETTEXT = &H189
Private Const LB_GETTEXTLEN = &H18A
Private Const LB_GETCOUNT = &H18B
Private Const LB_SELECTSTRING = &H18C
Private Const LB_DIR = &H18D
Private Const LB_GETTOPINDEX = &H18E
Private Const LB_FINDSTRING = &H18F
Private Const LB_GETSELCOUNT = &H190
Private Const LB_GETSELITEMS = &H191
Private Const LB_SETTABSTOPS = &H192
Private Const LB_GETHORIZONTALEXTENT = &H193
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const LB_SETCOLUMNWIDTH = &H195
Private Const LB_ADDFILE = &H196
Private Const LB_SETTOPINDEX = &H197
Private Const LB_GETITEMRECT = &H198
Private Const LB_GETITEMDATA = &H199
Private Const LB_SETITEMDATA = &H19A
Private Const LB_SELITEMRANGE = &H19B
Private Const LB_SETANCHORINDEX = &H19C
Private Const LB_GETANCHORINDEX = &H19D
Private Const LB_SETCARETINDEX = &H19E
Private Const LB_GETCARETINDEX = &H19F
Private Const LB_SETITEMHEIGHT = &H1A0
Private Const LB_GETITEMHEIGHT = &H1A1
Private Const LB_FINDSTRINGEXACT = &H1A2
Private Const LB_SETLOCALE = &H1A5
Private Const LB_GETLOCALE = &H1A6
Private Const LB_SETCOUNT = &H1A7
Private Const LB_MSGMAX = &H1A8

' Listbox Styles
Private Const LBS_NOTIFY = &H1&
Private Const LBS_SORT = &H2&
Private Const LBS_NOREDRAW = &H4&
Private Const LBS_MULTIPLESEL = &H8&
Private Const LBS_OWNERDRAWFIXED = &H10&
Private Const LBS_OWNERDRAWVARIABLE = &H20&
Private Const LBS_HASSTRINGS = &H40&
Private Const LBS_USETABSTOPS = &H80&
Private Const LBS_NOINTEGRALHEIGHT = &H100&
Private Const LBS_MULTICOLUMN = &H200&
Private Const LBS_WANTKEYBOARDINPUT = &H400&
Private Const LBS_EXTENDEDSEL = &H800&
Private Const LBS_DISABLENOSCROLL = &H1000&
Private Const LBS_NODATA = &H2000&
Private Const LBS_STANDARD = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)

' Combo Box return Values
Private Const CB_OKAY = 0
Private Const CB_ERR = (-1)
Private Const CB_ERRSPACE = (-2)

' Combo Box Notification Codes
Private Const CBN_ERRSPACE = (-1)
Private Const CBN_SELCHANGE = 1
Private Const CBN_DBLCLK = 2
Private Const CBN_SETFOCUS = 3
Private Const CBN_KILLFOCUS = 4
Private Const CBN_EDITCHANGE = 5
Private Const CBN_EDITUPDATE = 6
Private Const CBN_DROPDOWN = 7
Private Const CBN_CLOSEUP = 8
Private Const CBN_SELENDOK = 9
Private Const CBN_SELENDCANCEL = 10

' Combo Box styles
Private Const CBS_SIMPLE = &H1&
Private Const CBS_DROPDOWN = &H2&
Private Const CBS_DROPDOWNLIST = &H3&
Private Const CBS_OWNERDRAWFIXED = &H10&
Private Const CBS_OWNERDRAWVARIABLE = &H20&
Private Const CBS_AUTOHSCROLL = &H40&
Private Const CBS_OEMCONVERT = &H80&
Private Const CBS_SORT = &H100&
Private Const CBS_HASSTRINGS = &H200&
Private Const CBS_NOINTEGRALHEIGHT = &H400&
Private Const CBS_DISABLENOSCROLL = &H800&

' Combo Box messages
Private Const CB_GETEDITSEL = &H140
Private Const CB_LIMITTEXT = &H141
Private Const CB_SETEDITSEL = &H142
Private Const CB_ADDSTRING = &H143
Private Const CB_DELETESTRING = &H144
Private Const CB_DIR = &H145
Private Const CB_GETCOUNT = &H146
Private Const CB_GETCURSEL = &H147
Private Const CB_GETLBTEXT = &H148
Private Const CB_GETLBTEXTLEN = &H149
Private Const CB_INSERTSTRING = &H14A
Private Const CB_RESETCONTENT = &H14B
Private Const CB_FINDSTRING = &H14C
Private Const CB_SELECTSTRING = &H14D
Private Const CB_SETCURSEL = &H14E
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETITEMDATA = &H150
Private Const CB_SETITEMDATA = &H151
Private Const CB_GETDROPPEDCONTROLRECT = &H152
Private Const CB_SETITEMHEIGHT = &H153
Private Const CB_GETITEMHEIGHT = &H154
Private Const CB_SETEXTENDEDUI = &H155
Private Const CB_GETEXTENDEDUI = &H156
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_SETLOCALE = &H159
Private Const CB_GETLOCALE = &H15A
Private Const CB_MSGMAX = &H15B

' Scroll Bar Styles
Private Const SBS_HORZ = &H0&
Private Const SBS_VERT = &H1&
Private Const SBS_TOPALIGN = &H2&
Private Const SBS_LEFTALIGN = &H2&
Private Const SBS_BOTTOMALIGN = &H4&
Private Const SBS_RIGHTALIGN = &H4&
Private Const SBS_SIZEBOXTOPLEFTALIGN = &H2&
Private Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
Private Const SBS_SIZEBOX = &H8&

'  Scroll bar messages
Private Const SBM_SETPOS = &HE0 ' not in win3.1
Private Const SBM_GETPOS = &HE1 ' not in win3.1
Private Const SBM_SETRANGE = &HE2 ' not in win3.1
Private Const SBM_SETRANGEREDRAW = &HE6 ' not in win3.1
Private Const SBM_GETRANGE = &HE3 ' not in win3.1
Private Const SBM_ENABLE_ARROWS = &HE4 ' not in win3.1

Private Const MDIS_ALLCHILDSTYLES = &H1

' wParam values for WM_MDITILE and WM_MDICASCADE messages.
Private Const MDITILE_VERTICAL = &H0
Private Const MDITILE_HORIZONTAL = &H1
Private Const MDITILE_SKIPDISABLED = &H2

Type MDICREATESTRUCT
        szClass As String
        szTitle As String
        hOwner As Long
        X As Long
        Y As Long
        cx As Long
        cy As Long
        style As Long
        lParam As Long
End Type

Type CLIENTCREATESTRUCT
        hWindowMenu As Long
        idFirstChild As Long
End Type

Declare Function DefFrameProc Lib "user32" Alias "DefFrameProcA" (ByVal hwnd As Long, ByVal hWndMDIClient As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DefMDIChildProc Lib "user32" Alias "DefMDIChildProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function TranslateMDISysAccel Lib "user32" Alias "TranslateMDISysAccel" (ByVal hWndClient As Long, lpMsg As MSG) As Long

Declare Function ArrangeIconicWindows Lib "user32" Alias "ArrangeIconicWindows" (ByVal hwnd As Long) As Long

Declare Function CreateMDIWindow Lib "user32" Alias "CreateMDIWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hInstance As Long, ByVal lParam As Long) As Long

'  Help engine section.

Type MULTIKEYHELP
        mkSize As Long
        mkKeylist As Byte
        szKeyphrase As String * 253 ' Array length is arbitrary; may be changed
End Type

Type HELPWININFO
        wStructSize As Long
        X As Long
        Y As Long
        dx As Long
        dy As Long
        wMax As Long
        rgchMember As String * 2
End Type

' Commands to pass WinHelp()
Private Const HELP_CONTEXT = &H1          '  Display topic in ulTopic
Private Const HELP_QUIT = &H2             '  Terminate help
Private Const HELP_INDEX = &H3            '  Display index
Private Const HELP_CONTENTS = &H3&
Private Const HELP_HELPONHELP = &H4       '  Display help on using help
Private Const HELP_SETINDEX = &H5         '  Set current Index for multi index help
Private Const HELP_SETCONTENTS = &H5&
Private Const HELP_CONTEXTPOPUP = &H8&
Private Const HELP_FORCEFILE = &H9&
Private Const HELP_KEY = &H101            '  Display topic for keyword in offabData
Private Const HELP_COMMAND = &H102&
Private Const HELP_PARTIALKEY = &H105&
Private Const HELP_MULTIKEY = &H201&
Private Const HELP_SETWINPOS = &H203&

Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

' Parameter for SystemParametersInfo()
Private Const SPI_GETBEEP = 1
Private Const SPI_SETBEEP = 2
Private Const SPI_GETMOUSE = 3
Private Const SPI_SETMOUSE = 4
Private Const SPI_GETBORDER = 5
Private Const SPI_SETBORDER = 6
Private Const SPI_GETKEYBOARDSPEED = 10
Private Const SPI_SETKEYBOARDSPEED = 11
Private Const SPI_LANGDRIVER = 12
Private Const SPI_ICONHORIZONTALSPACING = 13
Private Const SPI_GETSCREENSAVETIMEOUT = 14
Private Const SPI_SETSCREENSAVETIMEOUT = 15
Private Const SPI_GETSCREENSAVEACTIVE = 16
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPI_GETGRIDGRANULARITY = 18
Private Const SPI_SETGRIDGRANULARITY = 19
Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPI_SETDESKPATTERN = 21
Private Const SPI_GETKEYBOARDDELAY = 22
Private Const SPI_SETKEYBOARDDELAY = 23
Private Const SPI_ICONVERTICALSPACING = 24
Private Const SPI_GETICONTITLEWRAP = 25
Private Const SPI_SETICONTITLEWRAP = 26
Private Const SPI_GETMENUDROPALIGNMENT = 27
Private Const SPI_SETMENUDROPALIGNMENT = 28
Private Const SPI_SETDOUBLECLKWIDTH = 29
Private Const SPI_SETDOUBLECLKHEIGHT = 30
Private Const SPI_GETICONTITLELOGFONT = 31
Private Const SPI_SETDOUBLECLICKTIME = 32
Private Const SPI_SETMOUSEBUTTONSWAP = 33
Private Const SPI_SETICONTITLELOGFONT = 34
Private Const SPI_GETFASTTASKSWITCH = 35
Private Const SPI_SETFASTTASKSWITCH = 36
Private Const SPI_SETDRAGFULLWINDOWS = 37
Private Const SPI_GETDRAGFULLWINDOWS = 38
Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const SPI_SETNONCLIENTMETRICS = 42
Private Const SPI_GETMINIMIZEDMETRICS = 43
Private Const SPI_SETMINIMIZEDMETRICS = 44
Private Const SPI_GETICONMETRICS = 45
Private Const SPI_SETICONMETRICS = 46
Private Const SPI_SETWORKAREA = 47
Private Const SPI_GETWORKAREA = 48
Private Const SPI_SETPENWINDOWS = 49
Private Const SPI_GETFILTERKEYS = 50
Private Const SPI_SETFILTERKEYS = 51
Private Const SPI_GETTOGGLEKEYS = 52
Private Const SPI_SETTOGGLEKEYS = 53
Private Const SPI_GETMOUSEKEYS = 54
Private Const SPI_SETMOUSEKEYS = 55
Private Const SPI_GETSHOWSOUNDS = 56
Private Const SPI_SETSHOWSOUNDS = 57
Private Const SPI_GETSTICKYKEYS = 58
Private Const SPI_SETSTICKYKEYS = 59
Private Const SPI_GETACCESSTIMEOUT = 60
Private Const SPI_SETACCESSTIMEOUT = 61
Private Const SPI_GETSERIALKEYS = 62
Private Const SPI_SETSERIALKEYS = 63
Private Const SPI_GETSOUNDSENTRY = 64
Private Const SPI_SETSOUNDSENTRY = 65
Private Const SPI_GETHIGHCONTRAST = 66
Private Const SPI_SETHIGHCONTRAST = 67
Private Const SPI_GETKEYBOARDPREF = 68
Private Const SPI_SETKEYBOARDPREF = 69
Private Const SPI_GETSCREENREADER = 70
Private Const SPI_SETSCREENREADER = 71
Private Const SPI_GETANIMATION = 72
Private Const SPI_SETANIMATION = 73
Private Const SPI_GETFONTSMOOTHING = 74
Private Const SPI_SETFONTSMOOTHING = 75
Private Const SPI_SETDRAGWIDTH = 76
Private Const SPI_SETDRAGHEIGHT = 77
Private Const SPI_SETHANDHELD = 78
Private Const SPI_GETLOWPOWERTIMEOUT = 79
Private Const SPI_GETPOWEROFFTIMEOUT = 80
Private Const SPI_SETLOWPOWERTIMEOUT = 81
Private Const SPI_SETPOWEROFFTIMEOUT = 82
Private Const SPI_GETLOWPOWERACTIVE = 83
Private Const SPI_GETPOWEROFFACTIVE = 84
Private Const SPI_SETLOWPOWERACTIVE = 85
Private Const SPI_SETPOWEROFFACTIVE = 86
Private Const SPI_SETCURSORS = 87
Private Const SPI_SETICONS = 88
Private Const SPI_GETDEFAULTINPUTLANG = 89
Private Const SPI_SETDEFAULTINPUTLANG = 90
Private Const SPI_SETLANGTOGGLE = 91
Private Const SPI_GETWINDOWSEXTENSION = 92
Private Const SPI_SETMOUSETRAILS = 93
Private Const SPI_GETMOUSETRAILS = 94
Private Const SPI_SCREENSAVERRUNNING = 97

' SystemParametersInfo flags
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPIF_SENDWININICHANGE = &H2

Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

'  DDE window messages
Private Const WM_DDE_FIRST = &H3E0
Private Const WM_DDE_INITIATE = (WM_DDE_FIRST)
Private Const WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
Private Const WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
Private Const WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
Private Const WM_DDE_ACK = (WM_DDE_FIRST + 4)
Private Const WM_DDE_DATA = (WM_DDE_FIRST + 5)
Private Const WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
Private Const WM_DDE_POKE = (WM_DDE_FIRST + 7)
Private Const WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
Private Const WM_DDE_LAST = (WM_DDE_FIRST + 8)

' *****************************************************************************                                                                             *
' * dde.h -       Dynamic Data Exchange structures and definitions              *
' *                                                                             *
' * Copyright (c) 1993-1995, Microsoft Corp.        All rights reserved              *
' *                                                                             *
' \*****************************************************************************/


' ----------------------------------------------------------------------------
'        DDEACK structure
'
'         Structure of wStatus (LOWORD(lParam)) in WM_DDE_ACK message
'        sent in response to a WM_DDE_DATA, WM_DDE_REQUEST, WM_DDE_POKE,
'        WM_DDE_ADVISE, or WM_DDE_UNADVISE message.
'
' ----------------------------------------------------------------------------*/

Type DDEACK
        bAppReturnCode As Integer
        Reserved As Integer
        fbusy As Integer
        fAck As Integer
End Type

' ----------------------------------------------------------------------------
'        DDEADVISE structure
'
'         WM_DDE_ADVISE parameter structure for hOptions (LOWORD(lParam))
'
' ----------------------------------------------------------------------------*/

Type DDEADVISE
        Reserved As Integer
        fDeferUpd As Integer
        fAckReq As Integer
        cfFormat As Integer
End Type


' ----------------------------------------------------------------------------
'        DDEDATA structure
'
'        WM_DDE_DATA parameter structure for hData (LOWORD(lParam)).
'        The actual size of this structure depends on the size of
'        the Value array.
'
' ----------------------------------------------------------------------------*/

Type DDEDATA
        unused As Integer
        fresponse As Integer
        fRelease As Integer
        Reserved As Integer
        fAckReq As Integer
        cfFormat As Integer
        Value(1) As Byte
End Type


' ----------------------------------------------------------------------------
'         DDEPOKE structure
'
'         WM_DDE_POKE parameter structure for hData (LOWORD(lParam)).
'        The actual size of this structure depends on the size of
'        the Value array.
'
' ----------------------------------------------------------------------------*/

Type DDEPOKE
        unused As Integer
        fRelease As Integer
        fReserved As Integer
        cfFormat As Integer
        Value(1) As Byte
End Type

' ----------------------------------------------------------------------------
' The following typedef's were used in previous versions of the Windows SDK.
' They are still valid.  The above typedef's define exactly the same structures
' as those below.  The above typedef names are recommended, however, as they
' are more meaningful.

' Note that the DDEPOKE structure typedef'ed in earlier versions of DDE.H did
' not correctly define the bit positions.
' ----------------------------------------------------------------------------*/

Type DDELN
        unused As Integer
        fRelease As Integer
        fDeferUpd As Integer
        fAckReq As Integer
        cfFormat As Integer
End Type

Type DDEUP
        unused As Integer
        fAck As Integer
        fRelease As Integer
        fReserved As Integer
        fAckReq As Integer
        cfFormat As Integer
        rgb(1) As Byte
End Type

Declare Function DdeSetQualityOfService Lib "user32" Alias "DdeSetQualityOfService" (ByVal hWndClient As Long, pqosNew As SECURITY_QUALITY_OF_SERVICE, pqosPrev As SECURITY_QUALITY_OF_SERVICE) As Long
Declare Function ImpersonateDdeClientWindow Lib "user32" Alias "ImpersonateDdeClientWindow" (ByVal hWndClient As Long, ByVal hWndServer As Long) As Long
Declare Function PackDDElParam Lib "user32" Alias "PackDDElParam" (ByVal msg As Long, ByVal uiLo As Long, ByVal uiHi As Long) As Long
Declare Function UnpackDDElParam Lib "user32" Alias "UnpackDDElParam" (ByVal msg As Long, ByVal lParam As Long, puiLo As Long, puiHi As Long) As Long
Declare Function FreeDDElParam Lib "user32" Alias "FreeDDElParam" (ByVal msg As Long, ByVal lParam As Long) As Long
Declare Function ReuseDDElParam Lib "user32" Alias "ReuseDDElParam" (ByVal lParam As Long, ByVal msgIn As Long, ByVal msgOut As Long, ByVal uiLo As Long, ByVal uiHi As Long) As Long

Type HSZPAIR
        hszSvc As Long
        hszTopic As Long
End Type

'//
'// Quality Of Service
'//

Type SECURITY_QUALITY_OF_SERVICE
    Length As Long
    Impersonationlevel As Integer
    ContextTrackingMode As Integer
    EffectiveOnly As Long
End Type
    
Type CONVCONTEXT
        cb As Long
        wFlags As Long
        wCountryID As Long
        iCodePage As Long
        dwLangID As Long
        dwSecurity As Long
        qos As SECURITY_QUALITY_OF_SERVICE
End Type

Type CONVINFO
        cb As Long
        hUser As Long
        hConvPartner As Long
        hszSvcPartner As Long
        hszServiceReq As Long
        hszTopic As Long
        hszItem As Long
        wFmt As Long
        wType As Long
        wStatus As Long
        wConvst As Long
        wLastError As Long
        hConvList As Long
        ConvCtxt As CONVCONTEXT
        hWnd As Long
        hwndPartner As Long
End Type

'  conversation states (usState)
Private Const XST_NULL = 0  '  quiescent states
Private Const XST_INCOMPLETE = 1
Private Const XST_CONNECTED = 2
Private Const XST_INIT1 = 3  '  mid-initiation states
Private Const XST_INIT2 = 4
Private Const XST_REQSENT = 5  '  active conversation states
Private Const XST_DATARCVD = 6
Private Const XST_POKESENT = 7
Private Const XST_POKEACKRCVD = 8
Private Const XST_EXECSENT = 9
Private Const XST_EXECACKRCVD = 10
Private Const XST_ADVSENT = 11
Private Const XST_UNADVSENT = 12
Private Const XST_ADVACKRCVD = 13
Private Const XST_UNADVACKRCVD = 14
Private Const XST_ADVDATASENT = 15
Private Const XST_ADVDATAACKRCVD = 16

'  used in LOWORD(dwData1) of XTYP_ADVREQ callbacks...
Private Const CADV_LATEACK = &HFFFF

'  conversation status bits (fsStatus)
Private Const ST_CONNECTED = &H1
Private Const ST_ADVISE = &H2
Private Const ST_ISLOCAL = &H4
Private Const ST_BLOCKED = &H8
Private Const ST_CLIENT = &H10
Private Const ST_TERMINATED = &H20
Private Const ST_INLIST = &H40
Private Const ST_BLOCKNEXT = &H80
Private Const ST_ISSELF = &H100

'  DDE constants for wStatus field
Private Const DDE_FACK = &H8000
Private Const DDE_FBUSY = &H4000
Private Const DDE_FDEFERUPD = &H4000
Private Const DDE_FACKREQ = &H8000
Private Const DDE_FRELEASE = &H2000
Private Const DDE_FREQUESTED = &H1000
Private Const DDE_FAPPSTATUS = &HFF
Private Const DDE_FNOTPROCESSED = &H0

Private Const DDE_FACKRESERVED = (Not (DDE_FACK Or DDE_FBUSY Or DDE_FAPPSTATUS))
Private Const DDE_FADVRESERVED = (Not (DDE_FACKREQ Or DDE_FDEFERUPD))
Private Const DDE_FDATRESERVED = (Not (DDE_FACKREQ Or DDE_FRELEASE Or DDE_FREQUESTED))
Private Const DDE_FPOKRESERVED = (Not (DDE_FRELEASE))

'  message filter hook types
Private Const MSGF_DDEMGR = &H8001

'  codepage constants
Private Const CP_WINANSI = 1004  '  default codepage for windows old DDE convs.
Private Const CP_WINUNICODE = 1200

'  transaction types
Private Const XTYPF_NOBLOCK = &H2     '  CBR_BLOCK will not work
Private Const XTYPF_NODATA = &H4     '  DDE_FDEFERUPD
Private Const XTYPF_ACKREQ = &H8     '  DDE_FACKREQ

Private Const XCLASS_MASK = &HFC00
Private Const XCLASS_BOOL = &H1000
Private Const XCLASS_DATA = &H2000
Private Const XCLASS_FLAGS = &H4000
Private Const XCLASS_NOTIFICATION = &H8000

Private Const XTYP_ERROR = (&H0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Private Const XTYP_ADVDATA = (&H10 Or XCLASS_FLAGS)
Private Const XTYP_ADVREQ = (&H20 Or XCLASS_DATA Or XTYPF_NOBLOCK)
Private Const XTYP_ADVSTART = (&H30 Or XCLASS_BOOL)
Private Const XTYP_ADVSTOP = (&H40 Or XCLASS_NOTIFICATION)
Private Const XTYP_EXECUTE = (&H50 Or XCLASS_FLAGS)
Private Const XTYP_CONNECT = (&H60 Or XCLASS_BOOL Or XTYPF_NOBLOCK)
Private Const XTYP_CONNECT_CONFIRM = (&H70 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Private Const XTYP_XACT_COMPLETE = (&H80 Or XCLASS_NOTIFICATION)
Private Const XTYP_POKE = (&H90 Or XCLASS_FLAGS)
Private Const XTYP_REGISTER = (&HA0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Private Const XTYP_REQUEST = (&HB0 Or XCLASS_DATA)
Private Const XTYP_DISCONNECT = (&HC0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Private Const XTYP_UNREGISTER = (&HD0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Private Const XTYP_WILDCONNECT = (&HE0 Or XCLASS_DATA Or XTYPF_NOBLOCK)

Private Const XTYP_MASK = &HF0
Private Const XTYP_SHIFT = 4  '  shift to turn XTYP_ into an index

'  Timeout constants
Private Const TIMEOUT_ASYNC = &HFFFF

'  Transaction ID constants
Private Const QID_SYNC = &HFFFF

'  Private strings used in DDE
Private Const SZDDESYS_TOPIC = "System"
Private Const SZDDESYS_ITEM_TOPICS = "Topics"
Private Const SZDDESYS_ITEM_SYSITEMS = "SysItems"
Private Const SZDDESYS_ITEM_RTNMSG = "ReturnMessage"
Private Const SZDDESYS_ITEM_STATUS = "Status"
Private Const SZDDESYS_ITEM_FORMATS = "Formats"
Private Const SZDDESYS_ITEM_HELP = "Help"
Private Const SZDDE_ITEM_ITEMLIST = "TopicItemList"

Private Const CBR_BLOCK = &HFFFF

' Callback filter flags for use with standard apps.
Private Const CBF_FAIL_SELFCONNECTIONS = &H1000
Private Const CBF_FAIL_CONNECTIONS = &H2000
Private Const CBF_FAIL_ADVISES = &H4000
Private Const CBF_FAIL_EXECUTES = &H8000
Private Const CBF_FAIL_POKES = &H10000
Private Const CBF_FAIL_REQUESTS = &H20000
Private Const CBF_FAIL_ALLSVRXACTIONS = &H3F000

Private Const CBF_SKIP_CONNECT_CONFIRMS = &H40000
Private Const CBF_SKIP_REGISTRATIONS = &H80000
Private Const CBF_SKIP_UNREGISTRATIONS = &H100000
Private Const CBF_SKIP_DISCONNECTS = &H200000
Private Const CBF_SKIP_ALLNOTIFICATIONS = &H3C0000

' Application command flags
Private Const APPCMD_CLIENTONLY = &H10&
Private Const APPCMD_FILTERINITS = &H20&
Private Const APPCMD_MASK = &HFF0&

' Application classification flags
Private Const APPCLASS_STANDARD = &H0&
Private Const APPCLASS_MASK = &HF&

Declare Function DdeUninitialize Lib "user32" Alias "DdeUninitialize" (ByVal idInst As Long) As Long

' conversation enumeration functions
Declare Function DdeConnectList Lib "user32" Alias "DdeConnectList" (ByVal idInst As Long, ByVal hszService As Long, ByVal hszTopic As Long, ByVal hConvList As Long, pCC As CONVCONTEXT) As Long
Declare Function DdeQueryNextServer Lib "user32" Alias "DdeQueryNextServer" (ByVal hConvList As Long, ByVal hConvPrev As Long) As Long
Declare Function DdeDisconnectList Lib "user32" Alias "DdeDisconnectList" (ByVal hConvList As Long) As Long

' conversation control functions
Declare Function DdeConnect Lib "user32" Alias "DdeConnect" (ByVal idInst As Long, ByVal hszService As Long, ByVal hszTopic As Long, pCC As CONVCONTEXT) As Long
Declare Function DdeDisconnect Lib "user32" Alias "DdeDisconnect" (ByVal hConv As Long) As Long
Declare Function DdeReconnect Lib "user32" Alias "DdeReconnect" (ByVal hConv As Long) As Long
Declare Function DdeQueryConvInfo Lib "user32" Alias "DdeQueryConvInfo" (ByVal hConv As Long, ByVal idTransaction As Long, pConvInfo As CONVINFO) As Long
Declare Function DdeSetUserHandle Lib "user32" Alias "DdeSetUserHandle" (ByVal hConv As Long, ByVal id As Long, ByVal hUser As Long) As Long
Declare Function DdeAbandonTransaction Lib "user32" Alias "DdeAbandonTransaction" (ByVal idInst As Long, ByVal hConv As Long, ByVal idTransaction As Long) As Long

' app server interface functions
Declare Function DdePostAdvise Lib "user32" Alias "DdePostAdvise" (ByVal idInst As Long, ByVal hszTopic As Long, ByVal hszItem As Long) As Long
Declare Function DdeEnableCallback Lib "user32" Alias "DdeEnableCallback" (ByVal idInst As Long, ByVal hConv As Long, ByVal wCmd As Long) As Long
Declare Function DdeImpersonateClient Lib "user32" Alias "DdeImpersonateClient" (ByVal hConv As Long) As Long

Private Const EC_ENABLEALL = 0
Private Const EC_ENABLEONE = ST_BLOCKNEXT
Private Const EC_DISABLE = ST_BLOCKED
Private Const EC_QUERYWAITING = 2

Declare Function DdeNameService Lib "user32" Alias "DdeNameService" (ByVal idInst As Long, ByVal hsz1 As Long, ByVal hsz2 As Long, ByVal afCmd As Long) As Long

Private Const DNS_REGISTER = &H1
Private Const DNS_UNREGISTER = &H2
Private Const DNS_FILTERON = &H4
Private Const DNS_FILTEROFF = &H8

' app client interface functions
Declare Function DdeClientTransaction Lib "user32" Alias "DdeClientTransaction" (pData As Byte, ByVal cbData As Long, ByVal hConv As Long, ByVal hszItem As Long, ByVal wFmt As Long, ByVal wType As Long, ByVal dwTimeout As Long, pdwResult As Long) As Long

' data transfer functions
Declare Function DdeCreateDataHandle Lib "user32" Alias "DdeCreateDataHandle" (ByVal idInst As Long, pSrc As Byte, ByVal cb As Long, ByVal cbOff As Long, ByVal hszItem As Long, ByVal wFmt As Long, ByVal afCmd As Long) As Long
Declare Function DdeAddData Lib "user32" Alias "DdeAddDataA" (ByVal hData As Long, pSrc As Byte, ByVal cb As Long, ByVal cbOff As Long) As Long
Declare Function DdeGetData Lib "user32" Alias "DdeGetDataA" (ByVal hData As Long, pDst As Byte, ByVal cbMax As Long, ByVal cbOff As Long) As Long
Declare Function DdeAccessData Lib "user32" Alias "DdeAccessDataA" (ByVal hData As Long, pcbDataSize As Long) As Long
Declare Function DdeUnaccessData Lib "user32" Alias "DdeUnaccessDataA" (ByVal hData As Long) As Long
Declare Function DdeFreeDataHandle Lib "user32" Alias "DdeFreeDataHandle" (ByVal hData As Long) As Long

Private Const HDATA_APPOWNED = &H1

Declare Function DdeGetLastError Lib "user32" Alias "DdeGetLastError" (ByVal idInst As Long) As Long

Private Const DMLERR_NO_ERROR = 0                           '  must be 0

Private Const DMLERR_FIRST = &H4000

Private Const DMLERR_ADVACKTIMEOUT = &H4000
Private Const DMLERR_BUSY = &H4001
Private Const DMLERR_DATAACKTIMEOUT = &H4002
Private Const DMLERR_DLL_NOT_INITIALIZED = &H4003
Private Const DMLERR_DLL_USAGE = &H4004
Private Const DMLERR_EXECACKTIMEOUT = &H4005
Private Const DMLERR_INVALIDPARAMETER = &H4006
Private Const DMLERR_LOW_MEMORY = &H4007
Private Const DMLERR_MEMORY_ERROR = &H4008
Private Const DMLERR_NOTPROCESSED = &H4009
Private Const DMLERR_NO_CONV_ESTABLISHED = &H400A
Private Const DMLERR_POKEACKTIMEOUT = &H400B
Private Const DMLERR_POSTMSG_FAILED = &H400C
Private Const DMLERR_REENTRANCY = &H400D
Private Const DMLERR_SERVER_DIED = &H400E
Private Const DMLERR_SYS_ERROR = &H400F
Private Const DMLERR_UNADVACKTIMEOUT = &H4010
Private Const DMLERR_UNFOUND_QUEUE_ID = &H4011

Private Const DMLERR_LAST = &H4011

Declare Function DdeCreateStringHandle Lib "user32" Alias "DdeCreateStringHandleA" (ByVal idInst As Long, ByVal psz As String, ByVal iCodePage As Long) As Long

Declare Function DdeQueryString Lib "user32" Alias "DdeQueryStringA" (ByVal idInst As Long, ByVal hsz As Long, ByVal psz As String, ByVal cchMax As Long, ByVal iCodePage As Long) As Long

Declare Function DdeFreeStringHandle Lib "user32" Alias "DdeFreeStringHandle" (ByVal idInst As Long, ByVal hsz As Long) As Long
Declare Function DdeKeepStringHandle Lib "user32" Alias "DdeKeepStringHandle" (ByVal idInst As Long, ByVal hsz As Long) As Long
Declare Function DdeCmpStringHandles Lib "user32" Alias "DdeCmpStringHandles" (ByVal hsz1 As Long, ByVal hsz2 As Long) As Long

Type DDEML_MSG_HOOK_DATA    '  new for NT
        uiLo As Long  '  unpacked lo and hi parts of lParam
        uiHi As Long
        cbData As Long   '  amount of data in message, if any. May be > than 32 bytes.
        Data(8) As Long  '  data peeking by DDESPY is limited to 32 bytes.
End Type

Type MONMSGSTRUCT
        cb As Long
        hwndTo As Long
        dwTime As Long
        htask As Long
        wMsg As Long
        wParam As Long
        lParam As Long
        dmhd As DDEML_MSG_HOOK_DATA       '  new for NT
End Type

Type MONCBSTRUCT
        cb As Long
        dwTime As Long
        htask As Long
        dwRet As Long
        wType As Long
        wFmt As Long
        hConv As Long
        hsz1 As Long
        hsz2 As Long
        hData As Long
        dwData1 As Long
        dwData2 As Long
        cc As CONVCONTEXT                 '  new for NT for XTYP_CONNECT callbacks
        cbData As Long                  '  new for NT for data peeking
        Data(8) As Long                 '  new for NT for data peeking
End Type

Type MONHSZSTRUCT
        cb As Long
        fsAction As Long '  MH_ value
        dwTime As Long
        hsz As Long
        htask As Long
        str As Byte
End Type

Private Const MH_CREATE = 1
Private Const MH_KEEP = 2
Private Const MH_DELETE = 3
Private Const MH_CLEANUP = 4

Type MONERRSTRUCT
        cb As Long
        wLastError As Long
        dwTime As Long
        htask As Long
End Type

Type MONLINKSTRUCT
        cb As Long
        dwTime As Long
        htask As Long
        fEstablished As Long
        fNoData As Long
        hszSvc As Long
        hszTopic As Long
        hszItem As Long
        wFmt As Long
        fServer As Long
        hConvServer As Long
        hConvClient As Long
End Type

Type MONCONVSTRUCT
        cb As Long
        fConnect As Long
        dwTime As Long
        htask As Long
        hszSvc As Long
        hszTopic As Long
        hConvClient As Long        '  Globally unique value != apps local hConv
        hConvServer As Long        '  Globally unique value != apps local hConv
End Type

Private Const MAX_MONITORS = 4
Private Const APPCLASS_MONITOR = &H1&
Private Const XTYP_MONITOR = (&HF0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)

' Callback filter flags for use with MONITOR apps - 0 implies no monitor callbacks
Private Const MF_HSZ_INFO = &H1000000
Private Const MF_SENDMSGS = &H2000000
Private Const MF_POSTMSGS = &H4000000
Private Const MF_CALLBACKS = &H8000000
Private Const MF_ERRORS = &H10000000
Private Const MF_LINKS = &H20000000
Private Const MF_CONV = &H40000000

Private Const MF_MASK = &HFF000000

' -----------------------------------------
' Win32 API error code definitions
' -----------------------------------------
' This section contains the error code definitions for the Win32 API functions.

' NO_ERROR
Private Const NO_ERROR = 0 '  dderror

' The configuration registry database operation completed successfully.
Private Const ERROR_SUCCESS = 0&

'   Incorrect function.
Private Const ERROR_INVALID_FUNCTION = 1 '  dderror

'   The system cannot find the file specified.
Private Const ERROR_FILE_NOT_FOUND = 2&

'   The system cannot find the path specified.
Private Const ERROR_PATH_NOT_FOUND = 3&

'   The system cannot open the file.
Private Const ERROR_TOO_MANY_OPEN_FILES = 4&

'   Access is denied.
Private Const ERROR_ACCESS_DENIED = 5&

'   The handle is invalid.
Private Const ERROR_INVALID_HANDLE = 6&

'   The storage control blocks were destroyed.
Private Const ERROR_ARENA_TRASHED = 7&

'   Not enough storage is available to process this command.
Private Const ERROR_NOT_ENOUGH_MEMORY = 8 '  dderror

'   The storage control block address is invalid.
Private Const ERROR_INVALID_BLOCK = 9&

'   The environment is incorrect.
Private Const ERROR_BAD_ENVIRONMENT = 10&

'   An attempt was made to load a program with an
'   incorrect format.
Private Const ERROR_BAD_FORMAT = 11&

'   The access code is invalid.
Private Const ERROR_INVALID_ACCESS = 12&

'   The data is invalid.
Private Const ERROR_INVALID_DATA = 13&

'   Not enough storage is available to complete this operation.
Private Const ERROR_OUTOFMEMORY = 14&

'   The system cannot find the drive specified.
Private Const ERROR_INVALID_DRIVE = 15&

'   The directory cannot be removed.
Private Const ERROR_CURRENT_DIRECTORY = 16&

'   The system cannot move the file
'   to a different disk drive.
Private Const ERROR_NOT_SAME_DEVICE = 17&

'   There are no more files.
Private Const ERROR_NO_MORE_FILES = 18&

'   The media is write protected.
Private Const ERROR_WRITE_PROTECT = 19&

'   The system cannot find the device specified.
Private Const ERROR_BAD_UNIT = 20&

'   The device is not ready.
Private Const ERROR_NOT_READY = 21&

'   The device does not recognize the command.
Private Const ERROR_BAD_COMMAND = 22&

'   Data error (cyclic redundancy check)
Private Const ERROR_CRC = 23&

'   The program issued a command but the
'   command length is incorrect.
Private Const ERROR_BAD_LENGTH = 24&

'   The drive cannot locate a specific
'   area or track on the disk.
Private Const ERROR_SEEK = 25&

'   The specified disk or diskette cannot be accessed.
Private Const ERROR_NOT_DOS_DISK = 26&

'   The drive cannot find the sector requested.
Private Const ERROR_SECTOR_NOT_FOUND = 27&

'   The printer is out of paper.
Private Const ERROR_OUT_OF_PAPER = 28&

'   The system cannot write to the specified device.
Private Const ERROR_WRITE_FAULT = 29&

'   The system cannot read from the specified device.
Private Const ERROR_READ_FAULT = 30&

'   A device attached to the system is not functioning.
Private Const ERROR_GEN_FAILURE = 31&

'   The process cannot access the file because
'   it is being used by another process.
Private Const ERROR_SHARING_VIOLATION = 32&

'   The process cannot access the file because
'   another process has locked a portion of the file.
Private Const ERROR_LOCK_VIOLATION = 33&

'   The wrong diskette is in the drive.
'   Insert %2 (Volume Serial Number: %3)
'   into drive %1.
Private Const ERROR_WRONG_DISK = 34&

'   Too many files opened for sharing.
Private Const ERROR_SHARING_BUFFER_EXCEEDED = 36&

'   Reached end of file.
Private Const ERROR_HANDLE_EOF = 38&

'   The disk is full.
Private Const ERROR_HANDLE_DISK_FULL = 39&

'   The network request is not supported.
Private Const ERROR_NOT_SUPPORTED = 50&

'   The remote computer is not available.
Private Const ERROR_REM_NOT_LIST = 51&

'   A duplicate name exists on the network.
Private Const ERROR_DUP_NAME = 52&

'   The network path was not found.
Private Const ERROR_BAD_NETPATH = 53&

'   The network is busy.
Private Const ERROR_NETWORK_BUSY = 54&

'   The specified network resource or device is no longer
'   available.
Private Const ERROR_DEV_NOT_EXIST = 55 '  dderror

'   The network BIOS command limit has been reached.
Private Const ERROR_TOO_MANY_CMDS = 56&

'   A network adapter hardware error occurred.
Private Const ERROR_ADAP_HDW_ERR = 57&

'   The specified server cannot perform the requested
'   operation.
Private Const ERROR_BAD_NET_RESP = 58&

'   An unexpected network error occurred.
Private Const ERROR_UNEXP_NET_ERR = 59&

'   The remote adapter is not compatible.
Private Const ERROR_BAD_REM_ADAP = 60&

'   The printer queue is full.
Private Const ERROR_PRINTQ_FULL = 61&

'   Space to store the file waiting to be printed is
'   not available on the server.
Private Const ERROR_NO_SPOOL_SPACE = 62&

'   Your file waiting to be printed was deleted.
Private Const ERROR_PRINT_CANCELLED = 63&

'   The specified network name is no longer available.
Private Const ERROR_NETNAME_DELETED = 64&

'   Network access is denied.
Private Const ERROR_NETWORK_ACCESS_DENIED = 65&

'   The network resource type is not correct.
Private Const ERROR_BAD_DEV_TYPE = 66&

'   The network name cannot be found.
Private Const ERROR_BAD_NET_NAME = 67&

'   The name limit for the local computer network
'   adapter card was exceeded.
Private Const ERROR_TOO_MANY_NAMES = 68&

'   The network BIOS session limit was exceeded.
Private Const ERROR_TOO_MANY_SESS = 69&

'   The remote server has been paused or is in the
'   process of being started.
Private Const ERROR_SHARING_PAUSED = 70&

'   The network request was not accepted.
Private Const ERROR_REQ_NOT_ACCEP = 71&

'   The specified printer or disk device has been paused.
Private Const ERROR_REDIR_PAUSED = 72&

'   The file exists.
Private Const ERROR_FILE_EXISTS = 80&

'   The directory or file cannot be created.
Private Const ERROR_CANNOT_MAKE = 82&

'   Fail on INT 24
Private Const ERROR_FAIL_I24 = 83&

'   Storage to process this request is not available.
Private Const ERROR_OUT_OF_STRUCTURES = 84&

'   The local device name is already in use.
Private Const ERROR_ALREADY_ASSIGNED = 85&

'   The specified network password is not correct.
Private Const ERROR_INVALID_PASSWORD = 86&

'   The parameter is incorrect.
Private Const ERROR_INVALID_PARAMETER = 87 '  dderror

'   A write fault occurred on the network.
Private Const ERROR_NET_WRITE_FAULT = 88&

'   The system cannot start another process at
'   this time.
Private Const ERROR_NO_PROC_SLOTS = 89&

'   Cannot create another system semaphore.
Private Const ERROR_TOO_MANY_SEMAPHORES = 100&

'   The exclusive semaphore is owned by another process.
Private Const ERROR_EXCL_SEM_ALREADY_OWNED = 101&

'   The semaphore is set and cannot be closed.
Private Const ERROR_SEM_IS_SET = 102&

'   The semaphore cannot be set again.
Private Const ERROR_TOO_MANY_SEM_REQUESTS = 103&

'   Cannot request exclusive semaphores at interrupt time.
Private Const ERROR_INVALID_AT_INTERRUPT_TIME = 104&

'   The previous ownership of this semaphore has ended.
Private Const ERROR_SEM_OWNER_DIED = 105&

'   Insert the diskette for drive %1.
Private Const ERROR_SEM_USER_LIMIT = 106&

'   Program stopped because alternate diskette was not inserted.
Private Const ERROR_DISK_CHANGE = 107&

'   The disk is in use or locked by
'   another process.
Private Const ERROR_DRIVE_LOCKED = 108&

'   The pipe has been ended.
Private Const ERROR_BROKEN_PIPE = 109&

'   The system cannot open the
'   device or file specified.
Private Const ERROR_OPEN_FAILED = 110&

'   The file name is too long.
Private Const ERROR_BUFFER_OVERFLOW = 111&

'   There is not enough space on the disk.
Private Const ERROR_DISK_FULL = 112&

'   No more internal file identifiers available.
Private Const ERROR_NO_MORE_SEARCH_HANDLES = 113&

'   The target internal file identifier is incorrect.
Private Const ERROR_INVALID_TARGET_HANDLE = 114&

'   The IOCTL call made by the application program is
'   not correct.
Private Const ERROR_INVALID_CATEGORY = 117&

'   The verify-on-write switch parameter value is not
'   correct.
Private Const ERROR_INVALID_VERIFY_SWITCH = 118&

'   The system does not support the command requested.
Private Const ERROR_BAD_DRIVER_LEVEL = 119&

'   This function is only valid in Windows NT mode.
Private Const ERROR_CALL_NOT_IMPLEMENTED = 120&

'   The semaphore timeout period has expired.
Private Const ERROR_SEM_TIMEOUT = 121&

'   The data area passed to a system call is too
'   small.
Private Const ERROR_INSUFFICIENT_BUFFER = 122 '  dderror

'   The filename, directory name, or volume label syntax is incorrect.
Private Const ERROR_INVALID_NAME = 123&

'   The system call level is not correct.
Private Const ERROR_INVALID_LEVEL = 124&

'   The disk has no volume label.
Private Const ERROR_NO_VOLUME_LABEL = 125&

'   The specified module could not be found.
Private Const ERROR_MOD_NOT_FOUND = 126&

'   The specified procedure could not be found.
Private Const ERROR_PROC_NOT_FOUND = 127&

'   There are no child processes to wait for.
Private Const ERROR_WAIT_NO_CHILDREN = 128&

'   The %1 application cannot be run in Windows NT mode.
Private Const ERROR_CHILD_NOT_COMPLETE = 129&

'   Attempt to use a file handle to an open disk partition for an
'   operation other than raw disk I/O.
Private Const ERROR_DIRECT_ACCESS_HANDLE = 130&

'   An attempt was made to move the file pointer before the beginning of the file.
Private Const ERROR_NEGATIVE_SEEK = 131&

'   The file pointer cannot be set on the specified device or file.
Private Const ERROR_SEEK_ON_DEVICE = 132&

'   A JOIN or SUBST command
'   cannot be used for a drive that
'   contains previously joined drives.
Private Const ERROR_IS_JOIN_TARGET = 133&

'   An attempt was made to use a
'   JOIN or SUBST command on a drive that has
'   already been joined.
Private Const ERROR_IS_JOINED = 134&

'   An attempt was made to use a
'   JOIN or SUBST command on a drive that has
'   already been substituted.
Private Const ERROR_IS_SUBSTED = 135&

'   The system tried to delete
'   the JOIN of a drive that is not joined.
Private Const ERROR_NOT_JOINED = 136&

'   The system tried to delete the
'   substitution of a drive that is not substituted.
Private Const ERROR_NOT_SUBSTED = 137&

'   The system tried to join a drive
'   to a directory on a joined drive.
Private Const ERROR_JOIN_TO_JOIN = 138&

'   The system tried to substitute a
'   drive to a directory on a substituted drive.
Private Const ERROR_SUBST_TO_SUBST = 139&

'   The system tried to join a drive to
'   a directory on a substituted drive.
Private Const ERROR_JOIN_TO_SUBST = 140&

'   The system tried to SUBST a drive
'   to a directory on a joined drive.
Private Const ERROR_SUBST_TO_JOIN = 141&

'   The system cannot perform a JOIN or SUBST at this time.
Private Const ERROR_BUSY_DRIVE = 142&

'   The system cannot join or substitute a
'   drive to or for a directory on the same drive.
Private Const ERROR_SAME_DRIVE = 143&

'   The directory is not a subdirectory of the root directory.
Private Const ERROR_DIR_NOT_ROOT = 144&

'   The directory is not empty.
Private Const ERROR_DIR_NOT_EMPTY = 145&

'   The path specified is being used in
'   a substitute.
Private Const ERROR_IS_SUBST_PATH = 146&

'   Not enough resources are available to
'   process this command.
Private Const ERROR_IS_JOIN_PATH = 147&

'   The path specified cannot be used at this time.
Private Const ERROR_PATH_BUSY = 148&

'   An attempt was made to join
'   or substitute a drive for which a directory
'   on the drive is the target of a previous
'   substitute.
Private Const ERROR_IS_SUBST_TARGET = 149&

'   System trace information was not specified in your
'   CONFIG.SYS file, or tracing is disallowed.
Private Const ERROR_SYSTEM_TRACE = 150&

'   The number of specified semaphore events for
'   DosMuxSemWait is not correct.
Private Const ERROR_INVALID_EVENT_COUNT = 151&

'   DosMuxSemWait did not execute; too many semaphores
'   are already set.
Private Const ERROR_TOO_MANY_MUXWAITERS = 152&

'   The DosMuxSemWait list is not correct.
Private Const ERROR_INVALID_LIST_FORMAT = 153&

'   The volume label you entered exceeds the
'   11 character limit.  The first 11 characters were written
'   to disk.  Any characters that exceeded the 11 character limit
'   were automatically deleted.
Private Const ERROR_LABEL_TOO_LONG = 154&

'   Cannot create another thread.
Private Const ERROR_TOO_MANY_TCBS = 155&

'   The recipient process has refused the signal.
Private Const ERROR_SIGNAL_REFUSED = 156&

'   The segment is already discarded and cannot be locked.
Private Const ERROR_DISCARDED = 157&

'   The segment is already unlocked.
Private Const ERROR_NOT_LOCKED = 158&

'   The address for the thread ID is not correct.
Private Const ERROR_BAD_THREADID_ADDR = 159&

'   The argument string passed to DosExecPgm is not correct.
Private Const ERROR_BAD_ARGUMENTS = 160&

'   The specified path is invalid.
Private Const ERROR_BAD_PATHNAME = 161&

'   A signal is already pending.
Private Const ERROR_SIGNAL_PENDING = 162&

'   No more threads can be created in the system.
Private Const ERROR_MAX_THRDS_REACHED = 164&

'   Unable to lock a region of a file.
Private Const ERROR_LOCK_FAILED = 167&

'   The requested resource is in use.
Private Const ERROR_BUSY = 170&

'   A lock request was not outstanding for the supplied cancel region.
Private Const ERROR_CANCEL_VIOLATION = 173&

'   The file system does not support atomic changes to the lock type.
Private Const ERROR_ATOMIC_LOCKS_NOT_SUPPORTED = 174&

'   The system detected a segment number that was not correct.
Private Const ERROR_INVALID_SEGMENT_NUMBER = 180&

'   The operating system cannot run %1.
Private Const ERROR_INVALID_ORDINAL = 182&

'   Cannot create a file when that file already exists.
Private Const ERROR_ALREADY_EXISTS = 183&

'   The flag passed is not correct.
Private Const ERROR_INVALID_FLAG_NUMBER = 186&

'   The specified system semaphore name was not found.
Private Const ERROR_SEM_NOT_FOUND = 187&

'   The operating system cannot run %1.
Private Const ERROR_INVALID_STARTING_CODESEG = 188&

'   The operating system cannot run %1.
Private Const ERROR_INVALID_STACKSEG = 189&

'   The operating system cannot run %1.
Private Const ERROR_INVALID_MODULETYPE = 190&

'   Cannot run %1 in Windows NT mode.
Private Const ERROR_INVALID_EXE_SIGNATURE = 191&

'   The operating system cannot run %1.
Private Const ERROR_EXE_MARKED_INVALID = 192&

'   %1 is not a valid Windows NT application.
Private Const ERROR_BAD_EXE_FORMAT = 193&

'   The operating system cannot run %1.
Private Const ERROR_ITERATED_DATA_EXCEEDS_64k = 194&

'   The operating system cannot run %1.
Private Const ERROR_INVALID_MINALLOCSIZE = 195&

'   The operating system cannot run this
'   application program.
Private Const ERROR_DYNLINK_FROM_INVALID_RING = 196&

'   The operating system is not presently
'   configured to run this application.
Private Const ERROR_IOPL_NOT_ENABLED = 197&

'   The operating system cannot run %1.
Private Const ERROR_INVALID_SEGDPL = 198&

'   The operating system cannot run this
'   application program.
Private Const ERROR_AUTODATASEG_EXCEEDS_64k = 199&

'   The code segment cannot be greater than or equal to 64KB.
Private Const ERROR_RING2SEG_MUST_BE_MOVABLE = 200&

'   The operating system cannot run %1.
Private Const ERROR_RELOC_CHAIN_XEEDS_SEGLIM = 201&

'   The operating system cannot run %1.
Private Const ERROR_INFLOOP_IN_RELOC_CHAIN = 202&

'   The system could not find the environment
'   option that was entered.
Private Const ERROR_ENVVAR_NOT_FOUND = 203&

'   No process in the command subtree has a
'   signal handler.
Private Const ERROR_NO_SIGNAL_SENT = 205&

'   The filename or extension is too long.
Private Const ERROR_FILENAME_EXCED_RANGE = 206&

'   The ring 2 stack is in use.
Private Const ERROR_RING2_STACK_IN_USE = 207&

'   The Global filename characters,  or ?, are entered
'   incorrectly or too many Global filename characters are specified.
Private Const ERROR_META_EXPANSION_TOO_LONG = 208&

'   The signal being posted is not correct.
Private Const ERROR_INVALID_SIGNAL_NUMBER = 209&

'   The signal handler cannot be set.
Private Const ERROR_THREAD_1_INACTIVE = 210&

'   The segment is locked and cannot be reallocated.
Private Const ERROR_LOCKED = 212&

'   Too many dynamic link modules are attached to this
'   program or dynamic link module.
Private Const ERROR_TOO_MANY_MODULES = 214&

'   Can't nest calls to LoadModule.
Private Const ERROR_NESTING_NOT_ALLOWED = 215&

'   The pipe state is invalid.
Private Const ERROR_BAD_PIPE = 230&

'   All pipe instances are busy.
Private Const ERROR_PIPE_BUSY = 231&

'   The pipe is being closed.
Private Const ERROR_NO_DATA = 232&

'   No process is on the other end of the pipe.
Private Const ERROR_PIPE_NOT_CONNECTED = 233&

'   More data is available.
Private Const ERROR_MORE_DATA = 234 '  dderror

'   The session was cancelled.
Private Const ERROR_VC_DISCONNECTED = 240&

'   The specified extended attribute name was invalid.
Private Const ERROR_INVALID_EA_NAME = 254&

'   The extended attributes are inconsistent.
Private Const ERROR_EA_LIST_INCONSISTENT = 255&

'   No more data is available.
Private Const ERROR_NO_MORE_ITEMS = 259&

'   The Copy API cannot be used.
Private Const ERROR_CANNOT_COPY = 266&

'   The directory name is invalid.
Private Const ERROR_DIRECTORY = 267&

'   The extended attributes did not fit in the buffer.
Private Const ERROR_EAS_DIDNT_FIT = 275&

'   The extended attribute file on the mounted file system is corrupt.
Private Const ERROR_EA_FILE_CORRUPT = 276&

'   The extended attribute table file is full.
Private Const ERROR_EA_TABLE_FULL = 277&

'   The specified extended attribute handle is invalid.
Private Const ERROR_INVALID_EA_HANDLE = 278&

'   The mounted file system does not support extended attributes.
Private Const ERROR_EAS_NOT_SUPPORTED = 282&

'   Attempt to release mutex not owned by caller.
Private Const ERROR_NOT_OWNER = 288&

'   Too many posts were made to a semaphore.
Private Const ERROR_TOO_MANY_POSTS = 298&

'   The system cannot find message for message number 0x%1
'   in message file for %2.
Private Const ERROR_MR_MID_NOT_FOUND = 317&

'   Attempt to access invalid address.
Private Const ERROR_INVALID_ADDRESS = 487&

'   Arithmetic result exceeded 32 bits.
Private Const ERROR_ARITHMETIC_OVERFLOW = 534&

'   There is a process on other end of the pipe.
Private Const ERROR_PIPE_CONNECTED = 535&

'   Waiting for a process to open the other end of the pipe.
Private Const ERROR_PIPE_LISTENING = 536&

'   Access to the extended attribute was denied.
Private Const ERROR_EA_ACCESS_DENIED = 994&

'   The I/O operation has been aborted because of either a thread exit
'   or an application request.
Private Const ERROR_OPERATION_ABORTED = 995&

'   Overlapped I/O event is not in a signalled state.
Private Const ERROR_IO_INCOMPLETE = 996&

'   Overlapped I/O operation is in progress.
Private Const ERROR_IO_PENDING = 997 '  dderror

'   Invalid access to memory location.
Private Const ERROR_NOACCESS = 998&

'   Error performing inpage operation.
Private Const ERROR_SWAPERROR = 999&

'   Recursion too deep, stack overflowed.
Private Const ERROR_STACK_OVERFLOW = 1001&

'   The window cannot act on the sent message.
Private Const ERROR_INVALID_MESSAGE = 1002&

'   Cannot complete this function.
Private Const ERROR_CAN_NOT_COMPLETE = 1003&

'   Invalid flags.
Private Const ERROR_INVALID_FLAGS = 1004&

'   The volume does not contain a recognized file system.
'   Please make sure that all required file system drivers are loaded and that the
'   volume is not corrupt.
Private Const ERROR_UNRECOGNIZED_VOLUME = 1005&

'   The volume for a file has been externally altered such that the
'   opened file is no longer valid.
Private Const ERROR_FILE_INVALID = 1006&

'   The requested operation cannot be performed in full-screen mode.
Private Const ERROR_FULLSCREEN_MODE = 1007&

'   An attempt was made to reference a token that does not exist.
Private Const ERROR_NO_TOKEN = 1008&

'   The configuration registry database is corrupt.
Private Const ERROR_BADDB = 1009&

'   The configuration registry key is invalid.
Private Const ERROR_BADKEY = 1010&

'   The configuration registry key could not be opened.
Private Const ERROR_CANTOPEN = 1011&

'   The configuration registry key could not be read.
Private Const ERROR_CANTREAD = 1012&

'   The configuration registry key could not be written.
Private Const ERROR_CANTWRITE = 1013&

'   One of the files in the Registry database had to be recovered
'   by use of a log or alternate copy.  The recovery was successful.
Private Const ERROR_REGISTRY_RECOVERED = 1014&

'   The Registry is corrupt. The structure of one of the files that contains
'   Registry data is corrupt, or the system's image of the file in memory
'   is corrupt, or the file could not be recovered because the alternate
'   copy or log was absent or corrupt.
Private Const ERROR_REGISTRY_CORRUPT = 1015&

'   An I/O operation initiated by the Registry failed unrecoverably.
'   The Registry could not read in, or write out, or flush, one of the files
'   that contain the system's image of the Registry.
Private Const ERROR_REGISTRY_IO_FAILED = 1016&

'   The system has attempted to load or restore a file into the Registry, but the
'   specified file is not in a Registry file format.
Private Const ERROR_NOT_REGISTRY_FILE = 1017&

'   Illegal operation attempted on a Registry key which has been marked for deletion.
Private Const ERROR_KEY_DELETED = 1018&

'   System could not allocate the required space in a Registry log.
Private Const ERROR_NO_LOG_SPACE = 1019&

'   Cannot create a symbolic link in a Registry key that already
'   has subkeys or values.
Private Const ERROR_KEY_HAS_CHILDREN = 1020&

'   Cannot create a stable subkey under a volatile parent key.
Private Const ERROR_CHILD_MUST_BE_VOLATILE = 1021&

'   A notify change request is being completed and the information
'   is not being returned in the caller's buffer. The caller now
'   needs to enumerate the files to find the changes.
Private Const ERROR_NOTIFY_ENUM_DIR = 1022&

'   A stop control has been sent to a service which other running services
'   are dependent on.
Private Const ERROR_DEPENDENT_SERVICES_RUNNING = 1051&

'   The requested control is not valid for this service
Private Const ERROR_INVALID_SERVICE_CONTROL = 1052&

'   The service did not respond to the start or control request in a timely
'   fashion.
Private Const ERROR_SERVICE_REQUEST_TIMEOUT = 1053&

'   A thread could not be created for the service.
Private Const ERROR_SERVICE_NO_THREAD = 1054&

'   The service database is locked.
Private Const ERROR_SERVICE_DATABASE_LOCKED = 1055&

'   An instance of the service is already running.
Private Const ERROR_SERVICE_ALREADY_RUNNING = 1056&

'   The account name is invalid or does not exist.
Private Const ERROR_INVALID_SERVICE_ACCOUNT = 1057&

'   The specified service is disabled and cannot be started.
Private Const ERROR_SERVICE_DISABLED = 1058&

'   Circular service dependency was specified.
Private Const ERROR_CIRCULAR_DEPENDENCY = 1059&

'   The specified service does not exist as an installed service.
Private Const ERROR_SERVICE_DOES_NOT_EXIST = 1060&

'   The service cannot accept control messages at this time.
Private Const ERROR_SERVICE_CANNOT_ACCEPT_CTRL = 1061&

'   The service has not been started.
Private Const ERROR_SERVICE_NOT_ACTIVE = 1062&

'   The service process could not connect to the service controller.
Private Const ERROR_FAILED_SERVICE_CONTROLLER_CONNECT = 1063&

'   An exception occurred in the service when handling the control request.
Private Const ERROR_EXCEPTION_IN_SERVICE = 1064&

'   The database specified does not exist.
Private Const ERROR_DATABASE_DOES_NOT_EXIST = 1065&

'   The service has returned a service-specific error code.
Private Const ERROR_SERVICE_SPECIFIC_ERROR = 1066&

'   The process terminated unexpectedly.
Private Const ERROR_PROCESS_ABORTED = 1067&

'   The dependency service or group failed to start.
Private Const ERROR_SERVICE_DEPENDENCY_FAIL = 1068&

'   The service did not start due to a logon failure.
Private Const ERROR_SERVICE_LOGON_FAILED = 1069&

'   After starting, the service hung in a start-pending state.
Private Const ERROR_SERVICE_START_HANG = 1070&

'   The specified service database lock is invalid.
Private Const ERROR_INVALID_SERVICE_LOCK = 1071&

'   The specified service has been marked for deletion.
Private Const ERROR_SERVICE_MARKED_FOR_DELETE = 1072&

'   The specified service already exists.
Private Const ERROR_SERVICE_EXISTS = 1073&

'   The system is currently running with the last-known-good configuration.
Private Const ERROR_ALREADY_RUNNING_LKG = 1074&

'   The dependency service does not exist or has been marked for
'   deletion.
Private Const ERROR_SERVICE_DEPENDENCY_DELETED = 1075&

'   The current boot has already been accepted for use as the
'   last-known-good control set.
Private Const ERROR_BOOT_ALREADY_ACCEPTED = 1076&

'   No attempts to start the service have been made since the last boot.
Private Const ERROR_SERVICE_NEVER_STARTED = 1077&

'   The name is already in use as either a service name or a service display
'   name.
Private Const ERROR_DUPLICATE_SERVICE_NAME = 1078&

'   The physical end of the tape has been reached.
Private Const ERROR_END_OF_MEDIA = 1100&

'   A tape access reached a filemark.
Private Const ERROR_FILEMARK_DETECTED = 1101&

'   Beginning of tape or partition was encountered.
Private Const ERROR_BEGINNING_OF_MEDIA = 1102&

'   A tape access reached the end of a set of files.
Private Const ERROR_SETMARK_DETECTED = 1103&

'   No more data is on the tape.
Private Const ERROR_NO_DATA_DETECTED = 1104&

'   Tape could not be partitioned.
Private Const ERROR_PARTITION_FAILURE = 1105&

'   When accessing a new tape of a multivolume partition, the current
'   blocksize is incorrect.
Private Const ERROR_INVALID_BLOCK_LENGTH = 1106&

'   Tape partition information could not be found when loading a tape.
Private Const ERROR_DEVICE_NOT_PARTITIONED = 1107&

'   Unable to lock the media eject mechanism.
Private Const ERROR_UNABLE_TO_LOCK_MEDIA = 1108&

'   Unable to unload the media.
Private Const ERROR_UNABLE_TO_UNLOAD_MEDIA = 1109&

'   Media in drive may have changed.
Private Const ERROR_MEDIA_CHANGED = 1110&

'   The I/O bus was reset.
Private Const ERROR_BUS_RESET = 1111&

'   No media in drive.
Private Const ERROR_NO_MEDIA_IN_DRIVE = 1112&

'   No mapping for the Unicode character exists in the target multi-byte code page.
Private Const ERROR_NO_UNICODE_TRANSLATION = 1113&

'   A dynamic link library (DLL) initialization routine failed.
Private Const ERROR_DLL_INIT_FAILED = 1114&

'   A system shutdown is in progress.
Private Const ERROR_SHUTDOWN_IN_PROGRESS = 1115&

'   Unable to abort the system shutdown because no shutdown was in progress.
Private Const ERROR_NO_SHUTDOWN_IN_PROGRESS = 1116&

'   The request could not be performed because of an I/O device error.
Private Const ERROR_IO_DEVICE = 1117&

'   No serial device was successfully initialized.  The serial driver will unload.
Private Const ERROR_SERIAL_NO_DEVICE = 1118&

'   Unable to open a device that was sharing an interrupt request (IRQ)
'   with other devices. At least one other device that uses that IRQ
'   was already opened.
Private Const ERROR_IRQ_BUSY = 1119&

'   A serial I/O operation was completed by another write to the serial port.
'   (The IOCTL_SERIAL_XOFF_COUNTER reached zero.)
Private Const ERROR_MORE_WRITES = 1120&

'   A serial I/O operation completed because the time-out period expired.
'   (The IOCTL_SERIAL_XOFF_COUNTER did not reach zero.)
Private Const ERROR_COUNTER_TIMEOUT = 1121&

'   No ID address mark was found on the floppy disk.
Private Const ERROR_FLOPPY_ID_MARK_NOT_FOUND = 1122&

'   Mismatch between the floppy disk sector ID field and the floppy disk
'   controller track address.
Private Const ERROR_FLOPPY_WRONG_CYLINDER = 1123&

'   The floppy disk controller reported an error that is not recognized
'   by the floppy disk driver.
Private Const ERROR_FLOPPY_UNKNOWN_ERROR = 1124&

'   The floppy disk controller returned inconsistent results in its registers.
Private Const ERROR_FLOPPY_BAD_REGISTERS = 1125&

'   While accessing the hard disk, a recalibrate operation failed, even after retries.
Private Const ERROR_DISK_RECALIBRATE_FAILED = 1126&

'   While accessing the hard disk, a disk operation failed even after retries.
Private Const ERROR_DISK_OPERATION_FAILED = 1127&

'   While accessing the hard disk, a disk controller reset was needed, but
'   even that failed.
Private Const ERROR_DISK_RESET_FAILED = 1128&

'   Physical end of tape encountered.
Private Const ERROR_EOM_OVERFLOW = 1129&

'   Not enough server storage is available to process this command.
Private Const ERROR_NOT_ENOUGH_SERVER_MEMORY = 1130&

'   A potential deadlock condition has been detected.
Private Const ERROR_POSSIBLE_DEADLOCK = 1131&

'   The base address or the file offset specified does not have the proper
'   alignment.
Private Const ERROR_MAPPED_ALIGNMENT = 1132&

' NEW for Win32
Private Const ERROR_INVALID_PIXEL_FORMAT = 2000
Private Const ERROR_BAD_DRIVER = 2001
Private Const ERROR_INVALID_WINDOW_STYLE = 2002
Private Const ERROR_METAFILE_NOT_SUPPORTED = 2003
Private Const ERROR_TRANSFORM_NOT_SUPPORTED = 2004
Private Const ERROR_CLIPPING_NOT_SUPPORTED = 2005
Private Const ERROR_UNKNOWN_PRINT_MONITOR = 3000
Private Const ERROR_PRINTER_DRIVER_IN_USE = 3001
Private Const ERROR_SPOOL_FILE_NOT_FOUND = 3002
Private Const ERROR_SPL_NO_STARTDOC = 3003
Private Const ERROR_SPL_NO_ADDJOB = 3004
Private Const ERROR_PRINT_PROCESSOR_ALREADY_INSTALLED = 3005
Private Const ERROR_PRINT_MONITOR_ALREADY_INSTALLED = 3006
Private Const ERROR_WINS_INTERNAL = 4000
Private Const ERROR_CAN_NOT_DEL_LOCAL_WINS = 4001
Private Const ERROR_STATIC_INIT = 4002
Private Const ERROR_INC_BACKUP = 4003
Private Const ERROR_FULL_BACKUP = 4004
Private Const ERROR_REC_NON_EXISTENT = 4005
Private Const ERROR_RPL_NOT_ALLOWED = 4006
Private Const SEVERITY_SUCCESS = 0
Private Const SEVERITY_ERROR = 1
Private Const FACILITY_NT_BIT = &H10000000
Private Const NOERROR = 0
Private Const E_UNEXPECTED = &H8000FFFF
Private Const E_NOTIMPL = &H80004001
Private Const E_OUTOFMEMORY = &H8007000E
Private Const E_INVALIDARG = &H80070057
Private Const E_NOINTERFACE = &H80004002
Private Const E_POINTER = &H80004003
Private Const E_HANDLE = &H80070006
Private Const E_ABORT = &H80004004
Private Const E_FAIL = &H80004005
Private Const E_ACCESSDENIED = &H80070005
Private Const CO_E_INIT_TLS = &H80004006
Private Const CO_E_INIT_SHARED_ALLOCATOR = &H80004007
Private Const CO_E_INIT_MEMORY_ALLOCATOR = &H80004008
Private Const CO_E_INIT_CLASS_CACHE = &H80004009
Private Const CO_E_INIT_RPC_CHANNEL = &H8000400A
Private Const CO_E_INIT_TLS_SET_CHANNEL_CONTROL = &H8000400B
Private Const CO_E_INIT_TLS_CHANNEL_CONTROL = &H8000400C
Private Const CO_E_INIT_UNACCEPTED_USER_ALLOCATOR = &H8000400D
Private Const CO_E_INIT_SCM_MUTEX_EXISTS = &H8000400E
Private Const CO_E_INIT_SCM_FILE_MAPPING_EXISTS = &H8000400F
Private Const CO_E_INIT_SCM_MAP_VIEW_OF_FILE = &H80004010
Private Const CO_E_INIT_SCM_EXEC_FAILURE = &H80004011
Private Const CO_E_INIT_ONLY_SINGLE_THREADED = &H80004012
Private Const S_OK = &H0
Private Const S_FALSE = &H1
Private Const OLE_E_FIRST = &H80040000
Private Const OLE_E_LAST = &H800400FF
Private Const OLE_S_FIRST = &H40000
Private Const OLE_S_LAST = &H400FF
Private Const OLE_E_OLEVERB = &H80040000
Private Const OLE_E_ADVF = &H80040001
Private Const OLE_E_ENUM_NOMORE = &H80040002
Private Const OLE_E_ADVISENOTSUPPORTED = &H80040003
Private Const OLE_E_NOCONNECTION = &H80040004
Private Const OLE_E_NOTRUNNING = &H80040005
Private Const OLE_E_NOCACHE = &H80040006
Private Const OLE_E_BLANK = &H80040007
Private Const OLE_E_CLASSDIFF = &H80040008
Private Const OLE_E_CANT_GETMONIKER = &H80040009
Private Const OLE_E_CANT_BINDTOSOURCE = &H8004000A
Private Const OLE_E_STATIC = &H8004000B
Private Const OLE_E_PROMPTSAVECANCELLED = &H8004000C
Private Const OLE_E_INVALIDRECT = &H8004000D
Private Const OLE_E_WRONGCOMPOBJ = &H8004000E
Private Const OLE_E_INVALIDHWND = &H8004000F
Private Const OLE_E_NOT_INPLACEACTIVE = &H80040010
Private Const OLE_E_CANTCONVERT = &H80040011
Private Const OLE_E_NOSTORAGE = &H80040012
Private Const DV_E_FORMATETC = &H80040064
Private Const DV_E_DVTARGETDEVICE = &H80040065
Private Const DV_E_STGMEDIUM = &H80040066
Private Const DV_E_STATDATA = &H80040067
Private Const DV_E_LINDEX = &H80040068
Private Const DV_E_TYMED = &H80040069
Private Const DV_E_CLIPFORMAT = &H8004006A
Private Const DV_E_DVASPECT = &H8004006B
Private Const DV_E_DVTARGETDEVICE_SIZE = &H8004006C
Private Const DV_E_NOIVIEWOBJECT = &H8004006D
Private Const DRAGDROP_E_FIRST = &H80040100
Private Const DRAGDROP_E_LAST = &H8004010F
Private Const DRAGDROP_S_FIRST = &H40100
Private Const DRAGDROP_S_LAST = &H4010F
Private Const DRAGDROP_E_NOTREGISTERED = &H80040100
Private Const DRAGDROP_E_ALREADYREGISTERED = &H80040101
Private Const DRAGDROP_E_INVALIDHWND = &H80040102
Private Const CLASSFACTORY_E_FIRST = &H80040110
Private Const CLASSFACTORY_E_LAST = &H8004011F
Private Const CLASSFACTORY_S_FIRST = &H40110
Private Const CLASSFACTORY_S_LAST = &H4011F
Private Const CLASS_E_NOAGGREGATION = &H80040110
Private Const CLASS_E_CLASSNOTAVAILABLE = &H80040111
Private Const MARSHAL_E_FIRST = &H80040120
Private Const MARSHAL_E_LAST = &H8004012F
Private Const MARSHAL_S_FIRST = &H40120
Private Const MARSHAL_S_LAST = &H4012F
Private Const DATA_E_FIRST = &H80040130
Private Const DATA_E_LAST = &H8004013F
Private Const DATA_S_FIRST = &H40130
Private Const DATA_S_LAST = &H4013F
Private Const VIEW_E_FIRST = &H80040140
Private Const VIEW_E_LAST = &H8004014F
Private Const VIEW_S_FIRST = &H40140
Private Const VIEW_S_LAST = &H4014F
Private Const VIEW_E_DRAW = &H80040140
Private Const REGDB_E_FIRST = &H80040150
Private Const REGDB_E_LAST = &H8004015F
Private Const REGDB_S_FIRST = &H40150
Private Const REGDB_S_LAST = &H4015F
Private Const REGDB_E_READREGDB = &H80040150
Private Const REGDB_E_WRITEREGDB = &H80040151
Private Const REGDB_E_KEYMISSING = &H80040152
Private Const REGDB_E_INVALIDVALUE = &H80040153
Private Const REGDB_E_CLASSNOTREG = &H80040154
Private Const REGDB_E_IIDNOTREG = &H80040155
Private Const CACHE_E_FIRST = &H80040170
Private Const CACHE_E_LAST = &H8004017F
Private Const CACHE_S_FIRST = &H40170
Private Const CACHE_S_LAST = &H4017F
Private Const CACHE_E_NOCACHE_UPDATED = &H80040170
Private Const OLEOBJ_E_FIRST = &H80040180
Private Const OLEOBJ_E_LAST = &H8004018F
Private Const OLEOBJ_S_FIRST = &H40180
Private Const OLEOBJ_S_LAST = &H4018F
Private Const OLEOBJ_E_NOVERBS = &H80040180
Private Const OLEOBJ_E_INVALIDVERB = &H80040181
Private Const CLIENTSITE_E_FIRST = &H80040190
Private Const CLIENTSITE_E_LAST = &H8004019F
Private Const CLIENTSITE_S_FIRST = &H40190
Private Const CLIENTSITE_S_LAST = &H4019F
Private Const INPLACE_E_NOTUNDOABLE = &H800401A0
Private Const INPLACE_E_NOTOOLSPACE = &H800401A1
Private Const INPLACE_E_FIRST = &H800401A0
Private Const INPLACE_E_LAST = &H800401AF
Private Const INPLACE_S_FIRST = &H401A0
Private Const INPLACE_S_LAST = &H401AF
Private Const ENUM_E_FIRST = &H800401B0
Private Const ENUM_E_LAST = &H800401BF
Private Const ENUM_S_FIRST = &H401B0
Private Const ENUM_S_LAST = &H401BF
Private Const CONVERT10_E_FIRST = &H800401C0
Private Const CONVERT10_E_LAST = &H800401CF
Private Const CONVERT10_S_FIRST = &H401C0
Private Const CONVERT10_S_LAST = &H401CF
Private Const CONVERT10_E_OLESTREAM_GET = &H800401C0
Private Const CONVERT10_E_OLESTREAM_PUT = &H800401C1
Private Const CONVERT10_E_OLESTREAM_FMT = &H800401C2
Private Const CONVERT10_E_OLESTREAM_BITMAP_TO_DIB = &H800401C3
Private Const CONVERT10_E_STG_FMT = &H800401C4
Private Const CONVERT10_E_STG_NO_STD_STREAM = &H800401C5
Private Const CONVERT10_E_STG_DIB_TO_BITMAP = &H800401C6
Private Const CLIPBRD_E_FIRST = &H800401D0
Private Const CLIPBRD_E_LAST = &H800401DF
Private Const CLIPBRD_S_FIRST = &H401D0
Private Const CLIPBRD_S_LAST = &H401DF
Private Const CLIPBRD_E_CANT_OPEN = &H800401D0
Private Const CLIPBRD_E_CANT_EMPTY = &H800401D1
Private Const CLIPBRD_E_CANT_SET = &H800401D2
Private Const CLIPBRD_E_BAD_DATA = &H800401D3
Private Const CLIPBRD_E_CANT_CLOSE = &H800401D4
Private Const MK_E_FIRST = &H800401E0
Private Const MK_E_LAST = &H800401EF
Private Const MK_S_FIRST = &H401E0
Private Const MK_S_LAST = &H401EF
Private Const MK_E_CONNECTMANUALLY = &H800401E0
Private Const MK_E_EXCEEDEDDEADLINE = &H800401E1
Private Const MK_E_NEEDGENERIC = &H800401E2
Private Const MK_E_UNAVAILABLE = &H800401E3
Private Const MK_E_SYNTAX = &H800401E4
Private Const MK_E_NOOBJECT = &H800401E5
Private Const MK_E_INVALIDEXTENSION = &H800401E6
Private Const MK_E_INTERMEDIATEINTERFACENOTSUPPORTED = &H800401E7
Private Const MK_E_NOTBINDABLE = &H800401E8
Private Const MK_E_NOTBOUND = &H800401E9
Private Const MK_E_CANTOPENFILE = &H800401EA
Private Const MK_E_MUSTBOTHERUSER = &H800401EB
Private Const MK_E_NOINVERSE = &H800401EC
Private Const MK_E_NOSTORAGE = &H800401ED
Private Const MK_E_NOPREFIX = &H800401EE
Private Const MK_E_ENUMERATION_FAILED = &H800401EF
Private Const CO_E_FIRST = &H800401F0
Private Const CO_E_LAST = &H800401FF
Private Const CO_S_FIRST = &H401F0
Private Const CO_S_LAST = &H401FF
Private Const CO_E_NOTINITIALIZED = &H800401F0
Private Const CO_E_ALREADYINITIALIZED = &H800401F1
Private Const CO_E_CANTDETERMINECLASS = &H800401F2
Private Const CO_E_CLASSSTRING = &H800401F3
Private Const CO_E_IIDSTRING = &H800401F4
Private Const CO_E_APPNOTFOUND = &H800401F5
Private Const CO_E_APPSINGLEUSE = &H800401F6
Private Const CO_E_ERRORINAPP = &H800401F7
Private Const CO_E_DLLNOTFOUND = &H800401F8
Private Const CO_E_ERRORINDLL = &H800401F9
Private Const CO_E_WRONGOSFORAPP = &H800401FA
Private Const CO_E_OBJNOTREG = &H800401FB
Private Const CO_E_OBJISREG = &H800401FC
Private Const CO_E_OBJNOTCONNECTED = &H800401FD
Private Const CO_E_APPDIDNTREG = &H800401FE
Private Const CO_E_RELEASED = &H800401FF
Private Const OLE_S_USEREG = &H40000
Private Const OLE_S_STATIC = &H40001
Private Const OLE_S_MAC_CLIPFORMAT = &H40002
Private Const DRAGDROP_S_DROP = &H40100
Private Const DRAGDROP_S_CANCEL = &H40101
Private Const DRAGDROP_S_USEDEFAULTCURSORS = &H40102
Private Const DATA_S_SAMEFORMATETC = &H40130
Private Const VIEW_S_ALREADY_FROZEN = &H40140
Private Const CACHE_S_FORMATETC_NOTSUPPORTED = &H40170
Private Const CACHE_S_SAMECACHE = &H40171
Private Const CACHE_S_SOMECACHES_NOTUPDATED = &H40172
Private Const OLEOBJ_S_INVALIDVERB = &H40180
Private Const OLEOBJ_S_CANNOT_DOVERB_NOW = &H40181
Private Const OLEOBJ_S_INVALIDHWND = &H40182
Private Const INPLACE_S_TRUNCATED = &H401A0
Private Const CONVERT10_S_NO_PRESENTATION = &H401C0
Private Const MK_S_REDUCED_TO_SELF = &H401E2
Private Const MK_S_ME = &H401E4
Private Const MK_S_HIM = &H401E5
Private Const MK_S_US = &H401E6
Private Const MK_S_MONIKERALREADYREGISTERED = &H401E7
Private Const CO_E_CLASS_CREATE_FAILED = &H80080001
Private Const CO_E_SCM_ERROR = &H80080002
Private Const CO_E_SCM_RPC_FAILURE = &H80080003
Private Const CO_E_BAD_PATH = &H80080004
Private Const CO_E_SERVER_EXEC_FAILURE = &H80080005
Private Const CO_E_OBJSRV_RPC_FAILURE = &H80080006
Private Const MK_E_NO_NORMALIZED = &H80080007
Private Const CO_E_SERVER_STOPPING = &H80080008
Private Const MEM_E_INVALID_ROOT = &H80080009
Private Const MEM_E_INVALID_LINK = &H80080010
Private Const MEM_E_INVALID_SIZE = &H80080011
Private Const DISP_E_UNKNOWNINTERFACE = &H80020001
Private Const DISP_E_MEMBERNOTFOUND = &H80020003
Private Const DISP_E_PARAMNOTFOUND = &H80020004
Private Const DISP_E_TYPEMISMATCH = &H80020005
Private Const DISP_E_UNKNOWNNAME = &H80020006
Private Const DISP_E_NONAMEDARGS = &H80020007
Private Const DISP_E_BADVARTYPE = &H80020008
Private Const DISP_E_EXCEPTION = &H80020009
Private Const DISP_E_OVERFLOW = &H8002000A
Private Const DISP_E_BADINDEX = &H8002000B
Private Const DISP_E_UNKNOWNLCID = &H8002000C
Private Const DISP_E_ARRAYISLOCKED = &H8002000D
Private Const DISP_E_BADPARAMCOUNT = &H8002000E
Private Const DISP_E_PARAMNOTOPTIONAL = &H8002000F
Private Const DISP_E_BADCALLEE = &H80020010
Private Const DISP_E_NOTACOLLECTION = &H80020011
Private Const TYPE_E_BUFFERTOOSMALL = &H80028016
Private Const TYPE_E_INVDATAREAD = &H80028018
Private Const TYPE_E_UNSUPFORMAT = &H80028019
Private Const TYPE_E_REGISTRYACCESS = &H8002801C
Private Const TYPE_E_LIBNOTREGISTERED = &H8002801D
Private Const TYPE_E_UNDEFINEDTYPE = &H80028027
Private Const TYPE_E_QUALIFIEDNAMEDISALLOWED = &H80028028
Private Const TYPE_E_INVALIDSTATE = &H80028029
Private Const TYPE_E_WRONGTYPEKIND = &H8002802A
Private Const TYPE_E_ELEMENTNOTFOUND = &H8002802B
Private Const TYPE_E_AMBIGUOUSNAME = &H8002802C
Private Const TYPE_E_NAMECONFLICT = &H8002802D
Private Const TYPE_E_UNKNOWNLCID = &H8002802E
Private Const TYPE_E_DLLFUNCTIONNOTFOUND = &H8002802F
Private Const TYPE_E_BADMODULEKIND = &H800288BD
Private Const TYPE_E_SIZETOOBIG = &H800288C5
Private Const TYPE_E_DUPLICATEID = &H800288C6
Private Const TYPE_E_INVALIDID = &H800288CF
Private Const TYPE_E_TYPEMISMATCH = &H80028CA0
Private Const TYPE_E_OUTOFBOUNDS = &H80028CA1
Private Const TYPE_E_IOERROR = &H80028CA2
Private Const TYPE_E_CANTCREATETMPFILE = &H80028CA3
Private Const TYPE_E_CANTLOADLIBRARY = &H80029C4A
Private Const TYPE_E_INCONSISTENTPROPFUNCS = &H80029C83
Private Const TYPE_E_CIRCULARTYPE = &H80029C84
Private Const STG_E_INVALIDFUNCTION = &H80030001
Private Const STG_E_FILENOTFOUND = &H80030002
Private Const STG_E_PATHNOTFOUND = &H80030003
Private Const STG_E_TOOMANYOPENFILES = &H80030004
Private Const STG_E_ACCESSDENIED = &H80030005
Private Const STG_E_INVALIDHANDLE = &H80030006
Private Const STG_E_INSUFFICIENTMEMORY = &H80030008
Private Const STG_E_INVALIDPOINTER = &H80030009
Private Const STG_E_NOMOREFILES = &H80030012
Private Const STG_E_DISKISWRITEPROTECTED = &H80030013
Private Const STG_E_SEEKERROR = &H80030019
Private Const STG_E_WRITEFAULT = &H8003001D
Private Const STG_E_READFAULT = &H8003001E
Private Const STG_E_SHAREVIOLATION = &H80030020
Private Const STG_E_LOCKVIOLATION = &H80030021
Private Const STG_E_FILEALREADYEXISTS = &H80030050
Private Const STG_E_INVALIDPARAMETER = &H80030057
Private Const STG_E_MEDIUMFULL = &H80030070
Private Const STG_E_ABNORMALAPIEXIT = &H800300FA
Private Const STG_E_INVALIDHEADER = &H800300FB
Private Const STG_E_INVALIDNAME = &H800300FC
Private Const STG_E_UNKNOWN = &H800300FD
Private Const STG_E_UNIMPLEMENTEDFUNCTION = &H800300FE
Private Const STG_E_INVALIDFLAG = &H800300FF
Private Const STG_E_INUSE = &H80030100
Private Const STG_E_NOTCURRENT = &H80030101
Private Const STG_E_REVERTED = &H80030102
Private Const STG_E_CANTSAVE = &H80030103
Private Const STG_E_OLDFORMAT = &H80030104
Private Const STG_E_OLDDLL = &H80030105
Private Const STG_E_SHAREREQUIRED = &H80030106
Private Const STG_E_NOTFILEBASEDSTORAGE = &H80030107
Private Const STG_E_EXTANTMARSHALLINGS = &H80030108
Private Const STG_S_CONVERTED = &H30200
Private Const RPC_E_CALL_REJECTED = &H80010001
Private Const RPC_E_CALL_CANCELED = &H80010002
Private Const RPC_E_CANTPOST_INSENDCALL = &H80010003
Private Const RPC_E_CANTCALLOUT_INASYNCCALL = &H80010004
Private Const RPC_E_CANTCALLOUT_INEXTERNALCALL = &H80010005
Private Const RPC_E_CONNECTION_TERMINATED = &H80010006
Private Const RPC_E_SERVER_DIED = &H80010007
Private Const RPC_E_CLIENT_DIED = &H80010008
Private Const RPC_E_INVALID_DATAPACKET = &H80010009
Private Const RPC_E_CANTTRANSMIT_CALL = &H8001000A
Private Const RPC_E_CLIENT_CANTMARSHAL_DATA = &H8001000B
Private Const RPC_E_CLIENT_CANTUNMARSHAL_DATA = &H8001000C
Private Const RPC_E_SERVER_CANTMARSHAL_DATA = &H8001000D
Private Const RPC_E_SERVER_CANTUNMARSHAL_DATA = &H8001000E
Private Const RPC_E_INVALID_DATA = &H8001000F
Private Const RPC_E_INVALID_PARAMETER = &H80010010
Private Const RPC_E_CANTCALLOUT_AGAIN = &H80010011
Private Const RPC_E_SERVER_DIED_DNE = &H80010012
Private Const RPC_E_SYS_CALL_FAILED = &H80010100
Private Const RPC_E_OUT_OF_RESOURCES = &H80010101
Private Const RPC_E_ATTEMPTED_MULTITHREAD = &H80010102
Private Const RPC_E_NOT_REGISTERED = &H80010103
Private Const RPC_E_FAULT = &H80010104
Private Const RPC_E_SERVERFAULT = &H80010105
Private Const RPC_E_CHANGED_MODE = &H80010106
Private Const RPC_E_INVALIDMETHOD = &H80010107
Private Const RPC_E_DISCONNECTED = &H80010108
Private Const RPC_E_RETRY = &H80010109
Private Const RPC_E_SERVERCALL_RETRYLATER = &H8001010A
Private Const RPC_E_SERVERCALL_REJECTED = &H8001010B
Private Const RPC_E_INVALID_CALLDATA = &H8001010C
Private Const RPC_E_CANTCALLOUT_ININPUTSYNCCALL = &H8001010D
Private Const RPC_E_WRONG_THREAD = &H8001010E
Private Const RPC_E_THREAD_NOT_INIT = &H8001010F
Private Const RPC_E_UNEXPECTED = &H8001FFFF


' /////////////////////////
'                        //
'  Winnet32 Status Codes //
'                        //
' /////////////////////////

'   The specified username is invalid.
Private Const ERROR_BAD_USERNAME = 2202&

'   This network connection does not exist.
Private Const ERROR_NOT_CONNECTED = 2250&

'   This network connection has files open or requests pending.
Private Const ERROR_OPEN_FILES = 2401&

'   The device is in use by an active process and cannot be disconnected.
Private Const ERROR_DEVICE_IN_USE = 2404&

'   The specified device name is invalid.
Private Const ERROR_BAD_DEVICE = 1200&

'   The device is not currently connected but it is a remembered connection.
Private Const ERROR_CONNECTION_UNAVAIL = 1201&

'   An attempt was made to remember a device that had previously been remembered.
Private Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&

'   No network provider accepted the given network path.
Private Const ERROR_NO_NET_OR_BAD_PATH = 1203&

'   The specified network provider name is invalid.
Private Const ERROR_BAD_PROVIDER = 1204&

'   Unable to open the network connection profile.
Private Const ERROR_CANNOT_OPEN_PROFILE = 1205&

'   The network connection profile is corrupt.
Private Const ERROR_BAD_PROFILE = 1206&

'   Cannot enumerate a non-container.
Private Const ERROR_NOT_CONTAINER = 1207&

'   An extended error has occurred.
Private Const ERROR_EXTENDED_ERROR = 1208&

'   The format of the specified group name is invalid.
Private Const ERROR_INVALID_GROUPNAME = 1209&

'   The format of the specified computer name is invalid.
Private Const ERROR_INVALID_COMPUTERNAME = 1210&

'   The format of the specified event name is invalid.
Private Const ERROR_INVALID_EVENTNAME = 1211&

'   The format of the specified domain name is invalid.
Private Const ERROR_INVALID_DOMAINNAME = 1212&

'   The format of the specified service name is invalid.
Private Const ERROR_INVALID_SERVICENAME = 1213&

'   The format of the specified network name is invalid.
Private Const ERROR_INVALID_NETNAME = 1214&

'   The format of the specified share name is invalid.
Private Const ERROR_INVALID_SHARENAME = 1215&

'   The format of the specified password is invalid.
Private Const ERROR_INVALID_PASSWORDNAME = 1216&

'   The format of the specified message name is invalid.
Private Const ERROR_INVALID_MESSAGENAME = 1217&

'   The format of the specified message destination is invalid.
Private Const ERROR_INVALID_MESSAGEDEST = 1218&

'   The credentials supplied conflict with an existing set of credentials.
Private Const ERROR_SESSION_CREDENTIAL_CONFLICT = 1219&

'   An attempt was made to establish a session to a Lan Manager server, but there
'   are already too many sessions established to that server.
Private Const ERROR_REMOTE_SESSION_LIMIT_EXCEEDED = 1220&

'   The workgroup or domain name is already in use by another computer on the
'   network.
Private Const ERROR_DUP_DOMAINNAME = 1221&

'   The network is not present or not started.
Private Const ERROR_NO_NETWORK = 1222&


' /////////////////////////
'                        //
'  Security Status Codes //
'                        //
' /////////////////////////

'   Not all privileges referenced are assigned to the caller.
Private Const ERROR_NOT_ALL_ASSIGNED = 1300&

'   Some mapping between account names and security IDs was not done.
Private Const ERROR_SOME_NOT_MAPPED = 1301&

'   No system quota limits are specifically set for this account.
Private Const ERROR_NO_QUOTAS_FOR_ACCOUNT = 1302&

'   No encryption key is available.  A well-known encryption key was returned.
Private Const ERROR_LOCAL_USER_SESSION_KEY = 1303&

'   The NT password is too complex to be converted to a LAN Manager
'   password.  The LAN Manager password returned is a NULL string.
Private Const ERROR_NULL_LM_PASSWORD = 1304&

'   The revision level is unknown.
Private Const ERROR_UNKNOWN_REVISION = 1305&

'   Indicates two revision levels are incompatible.
Private Const ERROR_REVISION_MISMATCH = 1306&

'   This security ID may not be assigned as the owner of this object.
Private Const ERROR_INVALID_OWNER = 1307&

'   This security ID may not be assigned as the primary group of an object.
Private Const ERROR_INVALID_PRIMARY_GROUP = 1308&

'   An attempt has been made to operate on an impersonation token
'   by a thread that is not currently impersonating a client.
Private Const ERROR_NO_IMPERSONATION_TOKEN = 1309&

'   The group may not be disabled.
Private Const ERROR_CANT_DISABLE_MANDATORY = 1310&

'   There are currently no logon servers available to service the logon
'   request.
Private Const ERROR_NO_LOGON_SERVERS = 1311&

'    A specified logon session does not exist.  It may already have
'    been terminated.
Private Const ERROR_NO_SUCH_LOGON_SESSION = 1312&

'    A specified privilege does not exist.
Private Const ERROR_NO_SUCH_PRIVILEGE = 1313&

'    A required privilege is not held by the client.
Private Const ERROR_PRIVILEGE_NOT_HELD = 1314&

'   The name provided is not a properly formed account name.
Private Const ERROR_INVALID_ACCOUNT_NAME = 1315&

'   The specified user already exists.
Private Const ERROR_USER_EXISTS = 1316&

'   The specified user does not exist.
Private Const ERROR_NO_SUCH_USER = 1317&

'   The specified group already exists.
Private Const ERROR_GROUP_EXISTS = 1318&

'   The specified group does not exist.
Private Const ERROR_NO_SUCH_GROUP = 1319&

'   Either the specified user account is already a member of the specified
'   group, or the specified group cannot be deleted because it contains
'   a member.
Private Const ERROR_MEMBER_IN_GROUP = 1320&

'   The specified user account is not a member of the specified group account.
Private Const ERROR_MEMBER_NOT_IN_GROUP = 1321&

'   The last remaining administration account cannot be disabled
'   or deleted.
Private Const ERROR_LAST_ADMIN = 1322&

'   Unable to update the password.  The value provided as the current
'   password is incorrect.
Private Const ERROR_WRONG_PASSWORD = 1323&

'   Unable to update the password.  The value provided for the new password
'   contains values that are not allowed in passwords.
Private Const ERROR_ILL_FORMED_PASSWORD = 1324&

'   Unable to update the password because a password update rule has been
'   violated.
Private Const ERROR_PASSWORD_RESTRICTION = 1325&

'   Logon failure: unknown user name or bad password.
Private Const ERROR_LOGON_FAILURE = 1326&

'   Logon failure: user account restriction.
Private Const ERROR_ACCOUNT_RESTRICTION = 1327&

'   Logon failure: account logon time restriction violation.
Private Const ERROR_INVALID_LOGON_HOURS = 1328&

'   Logon failure: user not allowed to log on to this computer.
Private Const ERROR_INVALID_WORKSTATION = 1329&

'   Logon failure: the specified account password has expired.
Private Const ERROR_PASSWORD_EXPIRED = 1330&

'   Logon failure: account currently disabled.
Private Const ERROR_ACCOUNT_DISABLED = 1331&

'   No mapping between account names and security IDs was done.
Private Const ERROR_NONE_MAPPED = 1332&

'   Too many local user identifiers (LUIDs) were requested at one time.
Private Const ERROR_TOO_MANY_LUIDS_REQUESTED = 1333&

'   No more local user identifiers (LUIDs) are available.
Private Const ERROR_LUIDS_EXHAUSTED = 1334&

'   The subauthority part of a security ID is invalid for this particular use.
Private Const ERROR_INVALID_SUB_AUTHORITY = 1335&

'   The access control list (ACL) structure is invalid.
Private Const ERROR_INVALID_ACL = 1336&

'   The security ID structure is invalid.
Private Const ERROR_INVALID_SID = 1337&

'   The security descriptor structure is invalid.
Private Const ERROR_INVALID_SECURITY_DESCR = 1338&

'   The inherited access control list (ACL) or access control entry (ACE)
'   could not be built.
Private Const ERROR_BAD_INHERITANCE_ACL = 1340&

'   The server is currently disabled.
Private Const ERROR_SERVER_DISABLED = 1341&

'   The server is currently enabled.
Private Const ERROR_SERVER_NOT_DISABLED = 1342&

'   The value provided was an invalid value for an identifier authority.
Private Const ERROR_INVALID_ID_AUTHORITY = 1343&

'   No more memory is available for security information updates.
Private Const ERROR_ALLOTTED_SPACE_EXCEEDED = 1344&

'   The specified attributes are invalid, or incompatible with the
'   attributes for the group as a whole.
Private Const ERROR_INVALID_GROUP_ATTRIBUTES = 1345&

'   Either a required impersonation level was not provided, or the
'   provided impersonation level is invalid.
Private Const ERROR_BAD_IMPERSONATION_LEVEL = 1346&

'   Cannot open an anonymous level security token.
Private Const ERROR_CANT_OPEN_ANONYMOUS = 1347&

'   The validation information class requested was invalid.
Private Const ERROR_BAD_VALIDATION_CLASS = 1348&

'   The type of the token is inappropriate for its attempted use.
Private Const ERROR_BAD_TOKEN_TYPE = 1349&

'   Unable to perform a security operation on an object
'   which has no associated security.
Private Const ERROR_NO_SECURITY_ON_OBJECT = 1350&

'   Indicates a Windows NT Advanced Server could not be contacted or that
'   objects within the domain are protected such that necessary
'   information could not be retrieved.
Private Const ERROR_CANT_ACCESS_DOMAIN_INFO = 1351&

'   The security account manager (SAM) or local security
'   authority (LSA) server was in the wrong state to perform
'   the security operation.
Private Const ERROR_INVALID_SERVER_STATE = 1352&

'   The domain was in the wrong state to perform the security operation.
Private Const ERROR_INVALID_DOMAIN_STATE = 1353&

'   This operation is only allowed for the Primary Domain Controller of the domain.
Private Const ERROR_INVALID_DOMAIN_ROLE = 1354&

'   The specified domain did not exist.
Private Const ERROR_NO_SUCH_DOMAIN = 1355&

'   The specified domain already exists.
Private Const ERROR_DOMAIN_EXISTS = 1356&

'   An attempt was made to exceed the limit on the number of domains per server.
Private Const ERROR_DOMAIN_LIMIT_EXCEEDED = 1357&

'   Unable to complete the requested operation because of either a
'   catastrophic media failure or a data structure corruption on the disk.
Private Const ERROR_INTERNAL_DB_CORRUPTION = 1358&

'   The security account database contains an internal inconsistency.
Private Const ERROR_INTERNAL_ERROR = 1359&

'   Generic access types were contained in an access mask which should
'   already be mapped to non-generic types.
Private Const ERROR_GENERIC_NOT_MAPPED = 1360&

'   A security descriptor is not in the right format (absolute or self-relative).
Private Const ERROR_BAD_DESCRIPTOR_FORMAT = 1361&

'   The requested action is restricted for use by logon processes
'   only.  The calling process has not registered as a logon process.
Private Const ERROR_NOT_LOGON_PROCESS = 1362&

'   Cannot start a new logon session with an ID that is already in use.
Private Const ERROR_LOGON_SESSION_EXISTS = 1363&

'   A specified authentication package is unknown.
Private Const ERROR_NO_SUCH_PACKAGE = 1364&

'   The logon session is not in a state that is consistent with the
'   requested operation.
Private Const ERROR_BAD_LOGON_SESSION_STATE = 1365&

'   The logon session ID is already in use.
Private Const ERROR_LOGON_SESSION_COLLISION = 1366&

'   A logon request contained an invalid logon type value.
Private Const ERROR_INVALID_LOGON_TYPE = 1367&

'   Unable to impersonate via a named pipe until data has been read
'   from that pipe.
Private Const ERROR_CANNOT_IMPERSONATE = 1368&

'   The transaction state of a Registry subtree is incompatible with the
'   requested operation.
Private Const ERROR_RXACT_INVALID_STATE = 1369&

'   An internal security database corruption has been encountered.
Private Const ERROR_RXACT_COMMIT_FAILURE = 1370&

'   Cannot perform this operation on built-in accounts.
Private Const ERROR_SPECIAL_ACCOUNT = 1371&

'   Cannot perform this operation on this built-in special group.
Private Const ERROR_SPECIAL_GROUP = 1372&

'   Cannot perform this operation on this built-in special user.
Private Const ERROR_SPECIAL_USER = 1373&

'   The user cannot be removed from a group because the group
'   is currently the user's primary group.
Private Const ERROR_MEMBERS_PRIMARY_GROUP = 1374&

'   The token is already in use as a primary token.
Private Const ERROR_TOKEN_ALREADY_IN_USE = 1375&

'   The specified local group does not exist.
Private Const ERROR_NO_SUCH_ALIAS = 1376&

'   The specified account name is not a member of the local group.
Private Const ERROR_MEMBER_NOT_IN_ALIAS = 1377&

'   The specified account name is already a member of the local group.
Private Const ERROR_MEMBER_IN_ALIAS = 1378&

'   The specified local group already exists.
Private Const ERROR_ALIAS_EXISTS = 1379&

'   Logon failure: the user has not been granted the requested
'   logon type at this computer.
Private Const ERROR_LOGON_NOT_GRANTED = 1380&

'   The maximum number of secrets that may be stored in a single system has been
'   exceeded.
Private Const ERROR_TOO_MANY_SECRETS = 1381&

'   The length of a secret exceeds the maximum length allowed.
Private Const ERROR_SECRET_TOO_LONG = 1382&

'   The local security authority database contains an internal inconsistency.
Private Const ERROR_INTERNAL_DB_ERROR = 1383&

'   During a logon attempt, the user's security context accumulated too many
'   security IDs.
Private Const ERROR_TOO_MANY_CONTEXT_IDS = 1384&

'   Logon failure: the user has not been granted the requested logon type
'   at this computer.
Private Const ERROR_LOGON_TYPE_NOT_GRANTED = 1385&

'   A cross-encrypted password is necessary to change a user password.
Private Const ERROR_NT_CROSS_ENCRYPTION_REQUIRED = 1386&

'   A new member could not be added to a local group because the member does
'   not exist.
Private Const ERROR_NO_SUCH_MEMBER = 1387&

'   A new member could not be added to a local group because the member has the
'   wrong account type.
Private Const ERROR_INVALID_MEMBER = 1388&

'   Too many security IDs have been specified.
Private Const ERROR_TOO_MANY_SIDS = 1389&

'   A cross-encrypted password is necessary to change this user password.
Private Const ERROR_LM_CROSS_ENCRYPTION_REQUIRED = 1390&

'   Indicates an ACL contains no inheritable components
Private Const ERROR_NO_INHERITANCE = 1391&

'   The file or directory is corrupt and non-readable.
Private Const ERROR_FILE_CORRUPT = 1392&

'   The disk structure is corrupt and non-readable.
Private Const ERROR_DISK_CORRUPT = 1393&

'   There is no user session key for the specified logon session.
Private Const ERROR_NO_USER_SESSION_KEY = 1394&

'  End of security error codes


' /////////////////////////
'                        //
'  WinUser Error Codes   //
'                        //
' /////////////////////////

'   Invalid window handle.
Private Const ERROR_INVALID_WINDOW_HANDLE = 1400&

'   Invalid menu handle.
Private Const ERROR_INVALID_MENU_HANDLE = 1401&

'   Invalid cursor handle.
Private Const ERROR_INVALID_CURSOR_HANDLE = 1402&

'   Invalid accelerator table handle.
Private Const ERROR_INVALID_ACCEL_HANDLE = 1403&

'   Invalid hook handle.
Private Const ERROR_INVALID_HOOK_HANDLE = 1404&

'   Invalid handle to a multiple-window position structure.
Private Const ERROR_INVALID_DWP_HANDLE = 1405&

'   Cannot create a top-level child window.
Private Const ERROR_TLW_WITH_WSCHILD = 1406&

'   Cannot find window class.
Private Const ERROR_CANNOT_FIND_WND_CLASS = 1407&

'   Invalid window, belongs to other thread.
Private Const ERROR_WINDOW_OF_OTHER_THREAD = 1408&

'   Hot key is already registered.
Private Const ERROR_HOTKEY_ALREADY_REGISTERED = 1409&

'   Class already exists.
Private Const ERROR_CLASS_ALREADY_EXISTS = 1410&

'   Class does not exist.
Private Const ERROR_CLASS_DOES_NOT_EXIST = 1411&

'   Class still has open windows.
Private Const ERROR_CLASS_HAS_WINDOWS = 1412&

'   Invalid index.
Private Const ERROR_INVALID_INDEX = 1413&

'   Invalid icon handle.
Private Const ERROR_INVALID_ICON_HANDLE = 1414&

'   Using private DIALOG window words.
Private Const ERROR_PRIVATE_DIALOG_INDEX = 1415&

'   The listbox identifier was not found.
Private Const ERROR_LISTBOX_ID_NOT_FOUND = 1416&

'   No wildcards were found.
Private Const ERROR_NO_WILDCARD_CHARACTERS = 1417&

'   Thread does not have a clipboard open.
Private Const ERROR_CLIPBOARD_NOT_OPEN = 1418&

'   Hot key is not registered.
Private Const ERROR_HOTKEY_NOT_REGISTERED = 1419&

'   The window is not a valid dialog window.
Private Const ERROR_WINDOW_NOT_DIALOG = 1420&

'   Control ID not found.
Private Const ERROR_CONTROL_ID_NOT_FOUND = 1421&

'   Invalid message for a combo box because it does not have an edit control.
Private Const ERROR_INVALID_COMBOBOX_MESSAGE = 1422&

'   The window is not a combo box.
Private Const ERROR_WINDOW_NOT_COMBOBOX = 1423&

'   Height must be less than 256.
Private Const ERROR_INVALID_EDIT_HEIGHT = 1424&

'   Invalid device context (DC) handle.
Private Const ERROR_DC_NOT_FOUND = 1425&

'   Invalid hook procedure type.
Private Const ERROR_INVALID_HOOK_FILTER = 1426&

'   Invalid hook procedure.
Private Const ERROR_INVALID_FILTER_PROC = 1427&

'   Cannot set non-local hook without a module handle.
Private Const ERROR_HOOK_NEEDS_HMOD = 1428&

'   This hook procedure can only be set Globally.
'
Private Const ERROR_Private_ONLY_HOOK = 1429&

'   The journal hook procedure is already installed.
Private Const ERROR_JOURNAL_HOOK_SET = 1430&

'   The hook procedure is not installed.
Private Const ERROR_HOOK_NOT_INSTALLED = 1431&

'   Invalid message for single-selection listbox.
Private Const ERROR_INVALID_LB_MESSAGE = 1432&

'   LB_SETCOUNT sent to non-lazy listbox.
Private Const ERROR_SETCOUNT_ON_BAD_LB = 1433&

'   This list box does not support tab stops.
Private Const ERROR_LB_WITHOUT_TABSTOPS = 1434&

'   Cannot destroy object created by another thread.
Private Const ERROR_DESTROY_OBJECT_OF_OTHER_THREAD = 1435&

'   Child windows cannot have menus.
Private Const ERROR_CHILD_WINDOW_MENU = 1436&

'   The window does not have a system menu.
Private Const ERROR_NO_SYSTEM_MENU = 1437&

'   Invalid message box style.
Private Const ERROR_INVALID_MSGBOX_STYLE = 1438&

'   Invalid system-wide (SPI_) parameter.
Private Const ERROR_INVALID_SPI_VALUE = 1439&

'   Screen already locked.
Private Const ERROR_SCREEN_ALREADY_LOCKED = 1440&

'   All handles to windows in a multiple-window position structure must
'   have the same parent.
Private Const ERROR_HWNDS_HAVE_DIFF_PARENT = 1441&

'   The window is not a child window.
Private Const ERROR_NOT_CHILD_WINDOW = 1442&

'   Invalid GW_ command.
Private Const ERROR_INVALID_GW_COMMAND = 1443&

'   Invalid thread identifier.
Private Const ERROR_INVALID_THREAD_ID = 1444&

'   Cannot process a message from a window that is not a multiple document
'   interface (MDI) window.
Private Const ERROR_NON_MDICHILD_WINDOW = 1445&

'   Popup menu already active.
Private Const ERROR_POPUP_ALREADY_ACTIVE = 1446&

'   The window does not have scroll bars.
Private Const ERROR_NO_SCROLLBARS = 1447&

'   Scroll bar range cannot be greater than 0x7FFF.
Private Const ERROR_INVALID_SCROLLBAR_RANGE = 1448&

'   Cannot show or remove the window in the way specified.
Private Const ERROR_INVALID_SHOWWIN_COMMAND = 1449&

'  End of WinUser error codes


' /////////////////////////
'                        //
'  Eventlog Status Codes //
'                        //
' /////////////////////////

'   The event log file is corrupt.
Private Const ERROR_EVENTLOG_FILE_CORRUPT = 1500&

'   No event log file could be opened, so the event logging service did not start.
Private Const ERROR_EVENTLOG_CANT_START = 1501&

'   The event log file is full.
Private Const ERROR_LOG_FILE_FULL = 1502&

'   The event log file has changed between reads.
Private Const ERROR_EVENTLOG_FILE_CHANGED = 1503&

'  End of eventlog error codes


' /////////////////////////
'                        //
'    RPC Status Codes    //
'                        //
' /////////////////////////

'   The string binding is invalid.
Private Const RPC_S_INVALID_STRING_BINDING = 1700&

'   The binding handle is not the correct type.
Private Const RPC_S_WRONG_KIND_OF_BINDING = 1701&

'   The binding handle is invalid.
Private Const RPC_S_INVALID_BINDING = 1702&

'   The RPC protocol sequence is not supported.
Private Const RPC_S_PROTSEQ_NOT_SUPPORTED = 1703&

'   The RPC protocol sequence is invalid.
Private Const RPC_S_INVALID_RPC_PROTSEQ = 1704&

'   The string universal unique identifier (UUID) is invalid.
Private Const RPC_S_INVALID_STRING_UUID = 1705&

'   The endpoint format is invalid.
Private Const RPC_S_INVALID_ENDPOINT_FORMAT = 1706&

'   The network address is invalid.
Private Const RPC_S_INVALID_NET_ADDR = 1707&

'   No endpoint was found.
Private Const RPC_S_NO_ENDPOINT_FOUND = 1708&

'   The timeout value is invalid.
Private Const RPC_S_INVALID_TIMEOUT = 1709&

'   The object universal unique identifier (UUID) was not found.
Private Const RPC_S_OBJECT_NOT_FOUND = 1710&

'   The object universal unique identifier (UUID) has already been registered.
Private Const RPC_S_ALREADY_REGISTERED = 1711&

'   The type universal unique identifier (UUID) has already been registered.
Private Const RPC_S_TYPE_ALREADY_REGISTERED = 1712&

'   The RPC server is already listening.
Private Const RPC_S_ALREADY_LISTENING = 1713&

'   No protocol sequences have been registered.
Private Const RPC_S_NO_PROTSEQS_REGISTERED = 1714&

'   The RPC server is not listening.
Private Const RPC_S_NOT_LISTENING = 1715&

'   The manager type is unknown.
Private Const RPC_S_UNKNOWN_MGR_TYPE = 1716&

'   The interface is unknown.
Private Const RPC_S_UNKNOWN_IF = 1717&

'   There are no bindings.
Private Const RPC_S_NO_BINDINGS = 1718&

'   There are no protocol sequences.
Private Const RPC_S_NO_PROTSEQS = 1719&

'   The endpoint cannot be created.
Private Const RPC_S_CANT_CREATE_ENDPOINT = 1720&

'   Not enough resources are available to complete this operation.
Private Const RPC_S_OUT_OF_RESOURCES = 1721&

'   The RPC server is unavailable.
Private Const RPC_S_SERVER_UNAVAILABLE = 1722&

'   The RPC server is too busy to complete this operation.
Private Const RPC_S_SERVER_TOO_BUSY = 1723&

'   The network options are invalid.
Private Const RPC_S_INVALID_NETWORK_OPTIONS = 1724&

'   There is not a remote procedure call active in this thread.
Private Const RPC_S_NO_CALL_ACTIVE = 1725&

'   The remote procedure call failed.
Private Const RPC_S_CALL_FAILED = 1726&

'   The remote procedure call failed and did not execute.
Private Const RPC_S_CALL_FAILED_DNE = 1727&

'   A remote procedure call (RPC) protocol error occurred.
Private Const RPC_S_PROTOCOL_ERROR = 1728&

'   The transfer syntax is not supported by the RPC server.
Private Const RPC_S_UNSUPPORTED_TRANS_SYN = 1730&

'   The universal unique identifier (UUID) type is not supported.
Private Const RPC_S_UNSUPPORTED_TYPE = 1732&

'   The tag is invalid.
Private Const RPC_S_INVALID_TAG = 1733&

'   The array bounds are invalid.
Private Const RPC_S_INVALID_BOUND = 1734&

'   The binding does not contain an entry name.
Private Const RPC_S_NO_ENTRY_NAME = 1735&

'   The name syntax is invalid.
Private Const RPC_S_INVALID_NAME_SYNTAX = 1736&

'   The name syntax is not supported.
Private Const RPC_S_UNSUPPORTED_NAME_SYNTAX = 1737&

'   No network address is available to use to construct a universal
'   unique identifier (UUID).
Private Const RPC_S_UUID_NO_ADDRESS = 1739&

'   The endpoint is a duplicate.
Private Const RPC_S_DUPLICATE_ENDPOINT = 1740&

'   The authentication type is unknown.
Private Const RPC_S_UNKNOWN_AUTHN_TYPE = 1741&

'   The maximum number of calls is too small.
Private Const RPC_S_MAX_CALLS_TOO_SMALL = 1742&

'   The string is too long.
Private Const RPC_S_STRING_TOO_LONG = 1743&

'   The RPC protocol sequence was not found.
Private Const RPC_S_PROTSEQ_NOT_FOUND = 1744&

'   The procedure number is out of range.
Private Const RPC_S_PROCNUM_OUT_OF_RANGE = 1745&

'   The binding does not contain any authentication information.
Private Const RPC_S_BINDING_HAS_NO_AUTH = 1746&

'   The authentication service is unknown.
Private Const RPC_S_UNKNOWN_AUTHN_SERVICE = 1747&

'   The authentication level is unknown.
Private Const RPC_S_UNKNOWN_AUTHN_LEVEL = 1748&

'   The security context is invalid.
Private Const RPC_S_INVALID_AUTH_IDENTITY = 1749&

'   The authorization service is unknown.
Private Const RPC_S_UNKNOWN_AUTHZ_SERVICE = 1750&

'   The entry is invalid.
Private Const EPT_S_INVALID_ENTRY = 1751&

'   The server endpoint cannot perform the operation.
Private Const EPT_S_CANT_PERFORM_OP = 1752&

'   There are no more endpoints available from the endpoint mapper.
Private Const EPT_S_NOT_REGISTERED = 1753&

'   No interfaces have been exported.
Private Const RPC_S_NOTHING_TO_EXPORT = 1754&

'   The entry name is incomplete.
Private Const RPC_S_INCOMPLETE_NAME = 1755&

'   The version option is invalid.
Private Const RPC_S_INVALID_VERS_OPTION = 1756&

'   There are no more members.
Private Const RPC_S_NO_MORE_MEMBERS = 1757&

'   There is nothing to unexport.
Private Const RPC_S_NOT_ALL_OBJS_UNEXPORTED = 1758&

'   The interface was not found.
Private Const RPC_S_INTERFACE_NOT_FOUND = 1759&

'   The entry already exists.
Private Const RPC_S_ENTRY_ALREADY_EXISTS = 1760&

'   The entry is not found.
Private Const RPC_S_ENTRY_NOT_FOUND = 1761&

'   The name service is unavailable.
Private Const RPC_S_NAME_SERVICE_UNAVAILABLE = 1762&

'   The network address family is invalid.
Private Const RPC_S_INVALID_NAF_ID = 1763&

'   The requested operation is not supported.
Private Const RPC_S_CANNOT_SUPPORT = 1764&

'   No security context is available to allow impersonation.
Private Const RPC_S_NO_CONTEXT_AVAILABLE = 1765&

'   An internal error occurred in a remote procedure call (RPC).
Private Const RPC_S_INTERNAL_ERROR = 1766&

'   The RPC server attempted an integer division by zero.'
Private Const RPC_S_ZERO_DIVIDE = 1767&

'   An addressing error occurred in the RPC server.
Private Const RPC_S_ADDRESS_ERROR = 1768&

'   A floating-point operation at the RPC server caused a division by zero.
Private Const RPC_S_FP_DIV_ZERO = 1769&

'   A floating-point underflow occurred at the RPC server.
Private Const RPC_S_FP_UNDERFLOW = 1770&

'   A floating-point overflow occurred at the RPC server.
Private Const RPC_S_FP_OVERFLOW = 1771&

'   The list of RPC servers available for the binding of auto handles
'   has been exhausted.
Private Const RPC_X_NO_MORE_ENTRIES = 1772&

'   Unable to open the character translation table file.
Private Const RPC_X_SS_CHAR_TRANS_OPEN_FAIL = 1773&

'   The file containing the character translation table has fewer than
'   512 bytes.
Private Const RPC_X_SS_CHAR_TRANS_SHORT_FILE = 1774&

'   A null context handle was passed from the client to the host during
'   a remote procedure call.
Private Const RPC_X_SS_IN_NULL_CONTEXT = 1775&

'   The context handle changed during a remote procedure call.
Private Const RPC_X_SS_CONTEXT_DAMAGED = 1777&

'   The binding handles passed to a remote procedure call do not match.
Private Const RPC_X_SS_HANDLES_MISMATCH = 1778&

'   The stub is unable to get the remote procedure call handle.
Private Const RPC_X_SS_CANNOT_GET_CALL_HANDLE = 1779&

'   A null reference pointer was passed to the stub.
Private Const RPC_X_NULL_REF_POINTER = 1780&

'   The enumeration value is out of range.
Private Const RPC_X_ENUM_VALUE_OUT_OF_RANGE = 1781&

'   The byte count is too small.
Private Const RPC_X_BYTE_COUNT_TOO_SMALL = 1782&

'   The stub received bad data.
Private Const RPC_X_BAD_STUB_DATA = 1783&

'   The supplied user buffer is not valid for the requested operation.
Private Const ERROR_INVALID_USER_BUFFER = 1784&

'   The disk media is not recognized.  It may not be formatted.
Private Const ERROR_UNRECOGNIZED_MEDIA = 1785&

'   The workstation does not have a trust secret.
Private Const ERROR_NO_TRUST_LSA_SECRET = 1786&

'   The SAM database on the Windows NT Advanced Server does not have a computer
'   account for this workstation trust relationship.
Private Const ERROR_NO_TRUST_SAM_ACCOUNT = 1787&

'   The trust relationship between the primary domain and the trusted
'   domain failed.
Private Const ERROR_TRUSTED_DOMAIN_FAILURE = 1788&

'   The trust relationship between this workstation and the primary
'   domain failed.
Private Const ERROR_TRUSTED_RELATIONSHIP_FAILURE = 1789&

'   The network logon failed.
Private Const ERROR_TRUST_FAILURE = 1790&

'   A remote procedure call is already in progress for this thread.
Private Const RPC_S_CALL_IN_PROGRESS = 1791&

'   An attempt was made to logon, but the network logon service was not started.
Private Const ERROR_NETLOGON_NOT_STARTED = 1792&

'   The user's account has expired.
Private Const ERROR_ACCOUNT_EXPIRED = 1793&

'   The redirector is in use and cannot be unloaded.
Private Const ERROR_REDIRECTOR_HAS_OPEN_HANDLES = 1794&

'   The specified printer driver is already installed.
Private Const ERROR_PRINTER_DRIVER_ALREADY_INSTALLED = 1795&

'   The specified port is unknown.
Private Const ERROR_UNKNOWN_PORT = 1796&

'   The printer driver is unknown.
Private Const ERROR_UNKNOWN_PRINTER_DRIVER = 1797&

'   The print processor is unknown.
'
Private Const ERROR_UNKNOWN_PRINTPROCESSOR = 1798&

'   The specified separator file is invalid.
Private Const ERROR_INVALID_SEPARATOR_FILE = 1799&

'   The specified priority is invalid.
Private Const ERROR_INVALID_PRIORITY = 1800&

'   The printer name is invalid.
Private Const ERROR_INVALID_PRINTER_NAME = 1801&

'   The printer already exists.
Private Const ERROR_PRINTER_ALREADY_EXISTS = 1802&

'   The printer command is invalid.
Private Const ERROR_INVALID_PRINTER_COMMAND = 1803&

'   The specified datatype is invalid.
Private Const ERROR_INVALID_DATATYPE = 1804&

'   The Environment specified is invalid.
Private Const ERROR_INVALID_ENVIRONMENT = 1805&

'   There are no more bindings.
Private Const RPC_S_NO_MORE_BINDINGS = 1806&

'   The account used is an interdomain trust account.  Use your Global user account or local user account to access this server.
Private Const ERROR_NOLOGON_INTERDOMAIN_TRUST_ACCOUNT = 1807&

'   The account used is a Computer Account.  Use your Global user account or local user account to access this server.
Private Const ERROR_NOLOGON_WORKSTATION_TRUST_ACCOUNT = 1808&

'   The account used is an server trust account.  Use your Global user account or local user account to access this server.
Private Const ERROR_NOLOGON_SERVER_TRUST_ACCOUNT = 1809&

'   The name or security ID (SID) of the domain specified is inconsistent
'   with the trust information for that domain.
Private Const ERROR_DOMAIN_TRUST_INCONSISTENT = 1810&

'   The server is in use and cannot be unloaded.
Private Const ERROR_SERVER_HAS_OPEN_HANDLES = 1811&

'   The specified image file did not contain a resource section.
Private Const ERROR_RESOURCE_DATA_NOT_FOUND = 1812&

'   The specified resource type can not be found in the image file.
Private Const ERROR_RESOURCE_TYPE_NOT_FOUND = 1813&

'   The specified resource name can not be found in the image file.
Private Const ERROR_RESOURCE_NAME_NOT_FOUND = 1814&

'   The specified resource language ID cannot be found in the image file.
Private Const ERROR_RESOURCE_LANG_NOT_FOUND = 1815&

'   Not enough quota is available to process this command.
Private Const ERROR_NOT_ENOUGH_QUOTA = 1816&

'   The group member was not found.
Private Const RPC_S_GROUP_MEMBER_NOT_FOUND = 1898&

'   The endpoint mapper database could not be created.
Private Const EPT_S_CANT_CREATE = 1899&

'   The object universal unique identifier (UUID) is the nil UUID.
Private Const RPC_S_INVALID_OBJECT = 1900&

'   The specified time is invalid.
Private Const ERROR_INVALID_TIME = 1901&

'   The specified Form name is invalid.
Private Const ERROR_INVALID_FORM_NAME = 1902&

'   The specified Form size is invalid
Private Const ERROR_INVALID_FORM_SIZE = 1903&

'   The specified Printer handle is already being waited on
Private Const ERROR_ALREADY_WAITING = 1904&

'   The specified Printer has been deleted
Private Const ERROR_PRINTER_DELETED = 1905&

'   The state of the Printer is invalid
Private Const ERROR_INVALID_PRINTER_STATE = 1906&

'   The list of servers for this workgroup is not currently available
Private Const ERROR_NO_BROWSER_SERVERS_FOUND = 6118&


' -------------------------
'  MMSystem Section
' -------------------------

' This section defines all the support for Multimedia applications

'  general constants
Private Const MAXPNAMELEN = 32  '  max product name length (including NULL)
Private Const MAXERRORLENGTH = 128  '  max error text length (including final NULL)

Type smpte
        hour As Byte
        min As Byte
        sec As Byte
        frame As Byte
        fps As Byte
        dummy As Byte
        pad(2) As Byte
End Type

Type midi
        songptrpos As Long
End Type

Type MMTIME
        wType As Long
        u As Long
End Type

'  values for wType field in MMTIME struct
Private Const TIME_MS = &H1     '  time in Milliseconds
Private Const TIME_SAMPLES = &H2     '  number of wave samples
Private Const TIME_BYTES = &H4     '  current byte offset
Private Const TIME_SMPTE = &H8     '  SMPTE time
Private Const TIME_MIDI = &H10    '  MIDI time

'  Multimedia Window Messages
Private Const MM_JOY1MOVE = &H3A0  '  joystick
Private Const MM_JOY2MOVE = &H3A1
Private Const MM_JOY1ZMOVE = &H3A2
Private Const MM_JOY2ZMOVE = &H3A3
Private Const MM_JOY1BUTTONDOWN = &H3B5
Private Const MM_JOY2BUTTONDOWN = &H3B6
Private Const MM_JOY1BUTTONUP = &H3B7
Private Const MM_JOY2BUTTONUP = &H3B8

Private Const MM_MCINOTIFY = &H3B9  '  MCI
Private Const MM_MCISYSTEM_STRING = &H3CA

Private Const MM_WOM_OPEN = &H3BB  '  waveform output
Private Const MM_WOM_CLOSE = &H3BC
Private Const MM_WOM_DONE = &H3BD

Private Const MM_WIM_OPEN = &H3BE  '  waveform input
Private Const MM_WIM_CLOSE = &H3BF
Private Const MM_WIM_DATA = &H3C0

Private Const MM_MIM_OPEN = &H3C1  '  MIDI input
Private Const MM_MIM_CLOSE = &H3C2
Private Const MM_MIM_DATA = &H3C3
Private Const MM_MIM_LONGDATA = &H3C4
Private Const MM_MIM_ERROR = &H3C5
Private Const MM_MIM_LONGERROR = &H3C6

Private Const MM_MOM_OPEN = &H3C7  '  MIDI output
Private Const MM_MOM_CLOSE = &H3C8
Private Const MM_MOM_DONE = &H3C9

' String resource number bases (internal use)

Private Const MMSYSERR_BASE = 0
Private Const WAVERR_BASE = 32
Private Const MIDIERR_BASE = 64
Private Const TIMERR_BASE = 96   '  was 128, changed to match Win 31 Sonic
Private Const JOYERR_BASE = 160
Private Const MCIERR_BASE = 256

Private Const MCI_STRING_OFFSET = 512  '  if this number is changed you MUST
                                    '  alter the MCI_DEVTYPE_... list below
Private Const MCI_VD_OFFSET = 1024
Private Const MCI_CD_OFFSET = 1088
Private Const MCI_WAVE_OFFSET = 1152
Private Const MCI_SEQ_OFFSET = 1216

' General error return values
Private Const MMSYSERR_NOERROR = 0  '  no error
Private Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)  '  unspecified error
Private Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)  '  device ID out of range
Private Const MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)  '  driver failed enable
Private Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)  '  device already allocated
Private Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)  '  device handle is invalid
Private Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)  '  no device driver present
Private Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)  '  memory allocation error
Private Const MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)  '  function isn't supported
Private Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)  '  error value out of range
Private Const MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10) '  invalid flag passed
Private Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11) '  invalid parameter passed
Private Const MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12) '  handle being used
                                                   '  simultaneously on another
                                                   '  thread (eg callback)
Private Const MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13) '  "Specified alias not found in WIN.INI
Private Const MMSYSERR_LASTERROR = (MMSYSERR_BASE + 13) '  last error in range
Private Const MM_MOM_POSITIONCB = &H3CA              '  Callback for MEVT_POSITIONCB
Private Const MM_MCISIGNAL = &H3CB
Private Const MM_MIM_MOREDATA = &H3CC                '  MIM_DONE w/ pending events
Private Const MIDICAPS_STREAM = &H8               '  driver supports midiStreamOut directly
 

Type MIDIEVENT
        dwDeltaTime As Long          '  Ticks since last event
        dwStreamID As Long           '  Reserved; must be zero
        dwEvent As Long              '  Event type and parameters
        dwParms(1) As Long           '  Parameters if this is a long event
End Type

Type MIDISTRMBUFFVER
        dwVersion As Long                  '  Stream buffer format version
        dwMid As Long                      '  Manufacturer ID as defined in MMREG.H
        dwOEMVersion As Long               '  Manufacturer version for custom ext
End Type

'  Type codes which go in the high byte of the event DWORD of a stream buffer

'  Type codes 00-7F contain parameters within the low 24 bits
'  Type codes 80-FF contain a length of their parameter in the low 24
'  bits, followed by their parameter data in the buffer. The event
'  DWORD contains the exact byte length; the parm data itself must be
'  padded to be an even multiple of 4 Byte long.
'

Private Const MEVT_F_SHORT = &H0&
Private Const MEVT_F_LONG = &H80000000
Private Const MEVT_F_CALLBACK = &H40000000
Private Const MIDISTRM_ERROR = -2

'
'  Structures and defines for midiStreamProperty
'
Private Const MIDIPROP_SET = &H80000000
Private Const MIDIPROP_GET = &H40000000

'  These are intentionally both non-zero so the app cannot accidentally
'  leave the operation off and happen to appear to work due to default
'  action.

Private Const MIDIPROP_TIMEDIV = &H1&
Private Const MIDIPROP_TEMPO = &H2&

Type MIDIPROPTIMEDIV
        cbStruct As Long
        dwTimeDiv As Long
End Type

Type MIDIPROPTEMPO
        cbStruct As Long
        dwTempo As Long
End Type


'  MIDI function prototypes *

' ***************************************************************************

'                             Mixer Support

' **************************************************************************

Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXER_LONG_NAME_CHARS = 64

'
'   MMRESULT error return values specific to the mixer API
'
'
Private Const MIXERR_BASE = 1024
Private Const MIXERR_INVALLINE = (MIXERR_BASE + 0)
Private Const MIXERR_INVALCONTROL = (MIXERR_BASE + 1)
Private Const MIXERR_INVALVALUE = (MIXERR_BASE + 2)
Private Const MIXERR_LASTERROR = (MIXERR_BASE + 2)


Private Const MIXER_OBJECTF_HANDLE = &H80000000
Private Const MIXER_OBJECTF_MIXER = &H0&
Private Const MIXER_OBJECTF_HMIXER = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)
Private Const MIXER_OBJECTF_WAVEOUT = &H10000000
Private Const MIXER_OBJECTF_HWAVEOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)
Private Const MIXER_OBJECTF_WAVEIN = &H20000000
Private Const MIXER_OBJECTF_HWAVEIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)
Private Const MIXER_OBJECTF_MIDIOUT = &H30000000
Private Const MIXER_OBJECTF_HMIDIOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)
Private Const MIXER_OBJECTF_MIDIIN = &H40000000
Private Const MIXER_OBJECTF_HMIDIIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIIN)
Private Const MIXER_OBJECTF_AUX = &H50000000

Declare Function mixerGetNumDevs Lib "winmm.dll" Alias "mixerGetNumDevs" () As Long

Type MIXERCAPS
        wMid As Integer                   '  manufacturer id
        wPid As Integer                   '  product id
        vDriverVersion As Long            '  version of the driver
        szPname As String * MAXPNAMELEN   '  product name
        fdwSupport As Long             '  misc. support bits
        cDestinations As Long          '  count of destinations
End Type

Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As Long, ByVal pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Declare Function mixerOpen Lib "winmm.dll" Alias "mixerOpen" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function mixerClose Lib "winmm.dll" Alias "mixerClose" (ByVal hmx As Long) As Long
Declare Function mixerMessage Lib "winmm.dll" Alias "mixerMessage" (ByVal hmx As Long, ByVal uMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long

Type Target    ' for use in MIXERLINE and others (embedded structure)
        
        dwType As Long                 '  MIXERLINE_TARGETTYPE_xxxx
        dwDeviceID As Long             '  target device ID of device type
        wMid As Integer                   '  of target device
        wPid As Integer                   '       "
        vDriverVersion As Long            '       "
        szPname As String * MAXPNAMELEN
End Type

Type MIXERLINE
        cbStruct As Long               '  size of MIXERLINE structure
        dwDestination As Long          '  zero based destination index
        dwSource As Long               '  zero based source index (if source)
        dwLineID As Long               '  unique line id for mixer device
        fdwLine As Long                '  state/information about line
        dwUser As Long                 '  driver specific information
        dwComponentType As Long        '  component type line connects to
        cChannels As Long              '  number of channels line supports
        cConnections As Long           '  number of connections (possible)
        cControls As Long              '  number of controls at this line
        szShortName As String * MIXER_SHORT_NAME_CHARS
        szName As String * MIXER_LONG_NAME_CHARS
        lpTarget As Target
End Type

'   MIXERLINE.fdwLine

Private Const MIXERLINE_LINEF_ACTIVE = &H1&
Private Const MIXERLINE_LINEF_DISCONNECTED = &H8000&
Private Const MIXERLINE_LINEF_SOURCE = &H80000000

'   MIXERLINE.dwComponentType

'   component types for destinations and sources

Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 0)
Private Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 1)
Private Const MIXERLINE_COMPONENTTYPE_DST_LINE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 2)
Private Const MIXERLINE_COMPONENTTYPE_DST_MONITOR = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 3)
Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Private Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 5)
Private Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 6)
Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Private Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Private Const MIXERLINE_COMPONENTTYPE_DST_LAST = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)

Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Private Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 0)
Private Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)
Private Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Private Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Private Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Private Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Private Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
Private Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Private Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)
Private Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
Private Const MIXERLINE_COMPONENTTYPE_SRC_LAST = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)


'
'   MIXERLINE.Target.dwType
'
'
Private Const MIXERLINE_TARGETTYPE_UNDEFINED = 0
Private Const MIXERLINE_TARGETTYPE_WAVEOUT = 1
Private Const MIXERLINE_TARGETTYPE_WAVEIN = 2
Private Const MIXERLINE_TARGETTYPE_MIDIOUT = 3
Private Const MIXERLINE_TARGETTYPE_MIDIIN = 4
Private Const MIXERLINE_TARGETTYPE_AUX = 5

Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Private Const MIXER_GETLINEINFOF_DESTINATION = &H0&
Private Const MIXER_GETLINEINFOF_SOURCE = &H1&
Private Const MIXER_GETLINEINFOF_LINEID = &H2&
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_GETLINEINFOF_TARGETTYPE = &H4&
Private Const MIXER_GETLINEINFOF_QUERYMASK = &HF&

Declare Function mixerGetID Lib "winmm.dll" Alias "mixerGetID" (ByVal hmxobj As Long, pumxID As Long, ByVal fdwId As Long) As Long

'   MIXERCONTROL

Type MIXERCONTROL
        cbStruct As Long           '  size in Byte of MIXERCONTROL
        dwControlID As Long        '  unique control id for mixer device
        dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
        fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
        cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
        szShortName As String * MIXER_SHORT_NAME_CHARS
        szName As String * MIXER_LONG_NAME_CHARS
        Bounds As Double
        Metrics As Long
End Type
'
'   MIXERCONTROL.fdwControl

Private Const MIXERCONTROL_CONTROLF_UNIFORM = &H1&
Private Const MIXERCONTROL_CONTROLF_MULTIPLE = &H2&
Private Const MIXERCONTROL_CONTROLF_DISABLED = &H80000000

'   MIXERCONTROL_CONTROLTYPE_xxx building block defines

Private Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Private Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Private Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Private Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_CLASS_TIME = &H60000000
Private Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Private Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000
Private Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Private Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Private Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Private Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Private Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Private Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Private Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Private Const MIXERCONTROL_CT_UNITS_MASK = &HFF0000
Private Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Private Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Private Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000 '  in 10ths
Private Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000 '  in 10ths
'
'   Commonly used control types for specifying MIXERCONTROL.dwControlType
'
Private Const MIXERCONTROL_CONTROLTYPE_CUSTOM = (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
Private Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Private Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_ONOFF = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Private Const MIXERCONTROL_CONTROLTYPE_MONO = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
Private Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
Private Const MIXERCONTROL_CONTROLTYPE_STEREOENH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
Private Const MIXERCONTROL_CONTROLTYPE_BUTTON = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_DECIBELS = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
Private Const MIXERCONTROL_CONTROLTYPE_SIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_PERCENT = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
Private Const MIXERCONTROL_CONTROLTYPE_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_PAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
Private Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
Private Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Private Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Private Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Private Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
Private Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
Private Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_MIXER = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Private Const MIXERCONTROL_CONTROLTYPE_MICROTIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_MILLITIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
'
'   MIXERLINECONTROLS
'
Type MIXERLINECONTROLS
        cbStruct As Long       '  size in Byte of MIXERLINECONTROLS
        dwLineID As Long       '  line id (from MIXERLINE.dwLineID)
                                             '  MIXER_GETLINECONTROLSF_ONEBYID or
        dwControl As Long  '  MIXER_GETLINECONTROLSF_ONEBYTYPE
        cControls As Long      '  count of controls pmxctrl points to
        cbmxctrl As Long       '  size in Byte of _one_ MIXERCONTROL
        pamxctrl As MIXERCONTROL       '  pointer to first MIXERCONTROL array
End Type

Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long

Private Const MIXER_GETLINECONTROLSF_ALL = &H0&
Private Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Private Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&

Type MIXERCONTROLDETAILS
        cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
        dwControlID As Long    '  control id to get/set details on
        cChannels As Long      '  number of channels in paDetails array
        item As Long                           ' hwndOwner or cMultipleItems
        cbDetails As Long      '  size of _one_ details_XX struct
        paDetails As Long      '  pointer to array of details_XX structs
End Type

'   MIXER_GETCONTROLDETAILSF_LISTTEXT

Type MIXERCONTROLDETAILS_LISTTEXT
        dwParam1 As Long
        dwParam2 As Long
        szName As String * MIXER_LONG_NAME_CHARS
End Type

'   MIXER_GETCONTROLDETAILSF_VALUE

Type MIXERCONTROLDETAILS_BOOLEAN
        fValue As Long
End Type

Type MIXERCONTROLDETAILS_SIGNED
        lValue As Long
End Type

Type MIXERCONTROLDETAILS_UNSIGNED
        dwValue As Long
End Type

Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long

Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_GETCONTROLDETAILSF_LISTTEXT = &H1&
Private Const MIXER_GETCONTROLDETAILSF_QUERYMASK = &HF&

Declare Function mixerSetControlDetails Lib "winmm.dll" Alias "mixerSetControlDetails" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long

Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_SETCONTROLDETAILSF_CUSTOM = &H1&
Private Const MIXER_SETCONTROLDETAILSF_QUERYMASK = &HF&

'  constants used with JOYINFOEX
Private Const JOY_BUTTON5 = &H10&
Private Const JOY_BUTTON6 = &H20&
Private Const JOY_BUTTON7 = &H40&
Private Const JOY_BUTTON8 = &H80&
Private Const JOY_BUTTON9 = &H100&
Private Const JOY_BUTTON10 = &H200&
Private Const JOY_BUTTON11 = &H400&
Private Const JOY_BUTTON12 = &H800&
Private Const JOY_BUTTON13 = &H1000&
Private Const JOY_BUTTON14 = &H2000&
Private Const JOY_BUTTON15 = &H4000&
Private Const JOY_BUTTON16 = &H8000&
Private Const JOY_BUTTON17 = &H10000
Private Const JOY_BUTTON18 = &H20000
Private Const JOY_BUTTON19 = &H40000
Private Const JOY_BUTTON20 = &H80000
Private Const JOY_BUTTON21 = &H100000
Private Const JOY_BUTTON22 = &H200000
Private Const JOY_BUTTON23 = &H400000
Private Const JOY_BUTTON24 = &H800000
Private Const JOY_BUTTON25 = &H1000000
Private Const JOY_BUTTON26 = &H2000000
Private Const JOY_BUTTON27 = &H4000000
Private Const JOY_BUTTON28 = &H8000000
Private Const JOY_BUTTON29 = &H10000000
Private Const JOY_BUTTON30 = &H20000000
Private Const JOY_BUTTON31 = &H40000000
Private Const JOY_BUTTON32 = &H80000000

'  constants used with JOYINFOEX structure
Private Const JOY_POVCENTERED = -1
Private Const JOY_POVFORWARD = 0
Private Const JOY_POVRIGHT = 9000
Private Const JOY_POVBACKWARD = 18000
Private Const JOY_POVLEFT = 27000
Private Const JOY_RETURNX = &H1&
Private Const JOY_RETURNY = &H2&
Private Const JOY_RETURNZ = &H4&
Private Const JOY_RETURNR = &H8&
Private Const JOY_RETURNU = &H10                             '  axis 5
Private Const JOY_RETURNV = &H20                             '  axis 6
Private Const JOY_RETURNPOV = &H40&
Private Const JOY_RETURNBUTTONS = &H80&
Private Const JOY_RETURNRAWDATA = &H100&
Private Const JOY_RETURNPOVCTS = &H200&
Private Const JOY_RETURNCENTERED = &H400&
Private Const JOY_USEDEADZONE = &H800&
Private Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)
Private Const JOY_CAL_READALWAYS = &H10000
Private Const JOY_CAL_READXYONLY = &H20000
Private Const JOY_CAL_READ3 = &H40000
Private Const JOY_CAL_READ4 = &H80000
Private Const JOY_CAL_READXONLY = &H100000
Private Const JOY_CAL_READYONLY = &H200000
Private Const JOY_CAL_READ5 = &H400000
Private Const JOY_CAL_READ6 = &H800000
Private Const JOY_CAL_READZONLY = &H1000000
Private Const JOY_CAL_READRONLY = &H2000000
Private Const JOY_CAL_READUONLY = &H4000000
Private Const JOY_CAL_READVONLY = &H8000000

Declare Function joyGetPos Lib "winmm.dll" Alias "joyGetPos" (ByVal uJoyID As Long, pji As JOYINFO) As Long
Declare Function joyGetPosEx Lib "winmm.dll" Alias "joyGetPosEx" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Private Const WAVE_FORMAT_QUERY = &H1
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_APPLICATION = &H80         '  look for application specific association
Private Const WAVE_MAPPED = &H4
Private Const WAVE_FORMAT_DIRECT = &H8
Private Const WAVE_FORMAT_DIRECT_QUERY = (WAVE_FORMAT_QUERY Or WAVE_FORMAT_DIRECT)
Private Const MIM_MOREDATA = MM_MIM_MOREDATA
Private Const MOM_POSITIONCB = MM_MOM_POSITIONCB

'  flags for dwFlags parm of midiInOpen()
Private Const MIDI_IO_STATUS = &H20&

Declare Function midiStreamOpen Lib "winmm.dll" Alias "midiStreamOpen" (phms As Long, puDeviceID As Long, ByVal cMidi As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function midiStreamClose Lib "winmm.dll" Alias "midiStreamClose" (ByVal hms As Long) As Long

Declare Function midiStreamProperty Lib "winmm.dll" Alias "midiStreamProperty" (ByVal hms As Long, lppropdata As Byte, ByVal dwProperty As Long) As Long
Declare Function midiStreamPosition Lib "winmm.dll" Alias "midiStreamPosition" (ByVal hms As Long, lpmmt As MMTIME, ByVal cbmmt As Long) As Long
Declare Function midiStreamOut Lib "winmm.dll" Alias "midiStreamOut" (ByVal hms As Long, pmh As MIDIHDR, ByVal cbmh As Long) As Long
Declare Function midiStreamPause Lib "winmm.dll" Alias "midiStreamPause" (ByVal hms As Long) As Long
Declare Function midiStreamRestart Lib "winmm.dll" Alias "midiStreamRestart" (ByVal hms As Long) As Long
Declare Function midiStreamStop Lib "winmm.dll" Alias "midiStreamStop" (ByVal hms As Long) As Long
Declare Function midiConnect Lib "winmm.dll" Alias "midiConnect" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
Declare Function midiDisconnect Lib "winmm.dll" Alias "midiDisconnect" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long

Type JOYINFOEX
        dwSize As Long                 '  size of structure
        dwFlags As Long                 '  flags to indicate what to return
        dwXpos As Long                '  x position
        dwYpos As Long                '  y position
        dwZpos As Long                '  z position
        dwRpos As Long                 '  rudder/4th axis position
        dwUpos As Long                 '  5th axis position
        dwVpos As Long                 '  6th axis position
        dwButtons As Long             '  button states
        dwButtonNumber As Long        '  current button number pressed
        dwPOV As Long                 '  point of view state
        dwReserved1 As Long                 '  reserved for communication between winmm driver
        dwReserved2 As Long                 '  reserved for future expansion
End Type
' Installable driver support

' Driver messages
Private Const DRV_LOAD = &H1
Private Const DRV_ENABLE = &H2
Private Const DRV_OPEN = &H3
Private Const DRV_CLOSE = &H4
Private Const DRV_DISABLE = &H5
Private Const DRV_FREE = &H6
Private Const DRV_CONFIGURE = &H7
Private Const DRV_QUERYCONFIGURE = &H8
Private Const DRV_INSTALL = &H9
Private Const DRV_REMOVE = &HA
Private Const DRV_EXITSESSION = &HB
Private Const DRV_POWER = &HF
Private Const DRV_RESERVED = &H800
Private Const DRV_USER = &H4000

Type DRVCONFIGINFO
        dwDCISize As Long
        lpszDCISectionName As String
        lpszDCIAliasName As String
        dnDevNode As Long
End Type

' Supported return values for DRV_CONFIGURE message
Private Const DRVCNF_CANCEL = &H0
Private Const DRVCNF_OK = &H1
Private Const DRVCNF_RESTART = &H2

'  return values from DriverProc() function
Private Const DRV_CANCEL = DRVCNF_CANCEL
Private Const DRV_OK = DRVCNF_OK
Private Const DRV_RESTART = DRVCNF_RESTART

Declare Function CloseDriver Lib "winmm.dll" Alias "CloseDriver" (ByVal hDriver As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Declare Function OpenDriver Lib "winmm.dll" Alias "OpenDriver" (ByVal szDriverName As String, ByVal szSectionName As String, ByVal lParam2 As Long) As Long
Declare Function SendDriverMessage Lib "winmm.dll" Alias "SendDriverMessage" (ByVal hDriver As Long, ByVal message As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Declare Function DrvGetModuleHandle Lib "winmm.dll" Alias "DrvGetModuleHandle" (ByVal hDriver As Long) As Long
Declare Function GetDriverModuleHandle Lib "winmm.dll" Alias "GetDriverModuleHandle" (ByVal hDriver As Long) As Long
Declare Function DefDriverProc Lib "winmm.dll" Alias "DefDriverProc" (ByVal dwDriverIdentifier As Long, ByVal hdrvr As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long

Private Const DRV_MCI_FIRST = DRV_RESERVED
Private Const DRV_MCI_LAST = DRV_RESERVED + &HFFF

' Driver callback support

'  flags used with waveOutOpen(), waveInOpen(), midiInOpen(), and
'  midiOutOpen() to specify the type of the dwCallback parameter.
Private Const CALLBACK_TYPEMASK = &H70000      '  callback type mask
Private Const CALLBACK_NULL = &H0        '  no callback
Private Const CALLBACK_WINDOW = &H10000      '  dwCallback is a HWND
Private Const CALLBACK_TASK = &H20000      '  dwCallback is a HTASK
Private Const CALLBACK_FUNCTION = &H30000      '  dwCallback is a FARPROC

'  manufacturer IDs
Private Const MM_MICROSOFT = 1  '  Microsoft Corp.

'  product IDs
Private Const MM_MIDI_MAPPER = 1  '  MIDI Mapper
Private Const MM_WAVE_MAPPER = 2  '  Wave Mapper

Private Const MM_SNDBLST_MIDIOUT = 3  '  Sound Blaster MIDI output port
Private Const MM_SNDBLST_MIDIIN = 4  '  Sound Blaster MIDI input port
Private Const MM_SNDBLST_SYNTH = 5  '  Sound Blaster internal synthesizer
Private Const MM_SNDBLST_WAVEOUT = 6  '  Sound Blaster waveform output
Private Const MM_SNDBLST_WAVEIN = 7  '  Sound Blaster waveform input

Private Const MM_ADLIB = 9  '  Ad Lib-compatible synthesizer

Private Const MM_MPU401_MIDIOUT = 10  '  MPU401-compatible MIDI output port
Private Const MM_MPU401_MIDIIN = 11  '  MPU401-compatible MIDI input port

Private Const MM_PC_JOYSTICK = 12  '  Joystick adapter

Declare Function mmsystemGetVersion Lib "winmm.dll" Alias "mmsystemGetVersion" () As Long
Declare Sub OutputDebugStr Lib "winmm.dll" Alias "OutputDebugStr" (ByVal lpszOutputString As String)

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'  flag values for uFlags parameter
Private Const SND_SYNC = &H0         '  play synchronously (default)
Private Const SND_ASYNC = &H1         '  play asynchronously

Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found

Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier

Private Const SND_ALIAS_START = 0  '  must be > 4096 to keep strings in same section of resource file

Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_VALID = &H1F        '  valid flags          / ;Internal /

Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy

Private Const SND_VALIDFLAGS = &H17201F    '  Set of valid flag bits.  Anything outside
                                    '  this range will raise an error
Private Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved

Private Const SND_TYPE_MASK = &H170007

Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'  waveform audio error return values
Private Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)    '  unsupported wave format
Private Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)    '  still something playing
Private Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)    '  header not prepared
Private Const WAVERR_SYNC = (WAVERR_BASE + 3)    '  device is synchronous
Private Const WAVERR_LASTERROR = (WAVERR_BASE + 3)    '  last error in range

'  wave callback messages
Private Const WOM_OPEN = MM_WOM_OPEN
Private Const WOM_CLOSE = MM_WOM_CLOSE
Private Const WOM_DONE = MM_WOM_DONE
Private Const WIM_OPEN = MM_WIM_OPEN
Private Const WIM_CLOSE = MM_WIM_CLOSE
Private Const WIM_DATA = MM_WIM_DATA

'  device ID for wave device mapper
Private Const WAVE_MAPPER = -1&

'  flags for dwFlags parameter in waveOutOpen() and waveInOpen()

Private Const WAVE_ALLOWSYNC = &H2
Private Const WAVE_VALID = &H3         '  ;Internal

Type WAVEHDR
        lpData As String
        dwBufferLength As Long
        dwBytesRecorded As Long
        dwUser As Long
        dwFlags As Long
        dwLoops As Long
        lpNext As Long
        Reserved As Long
End Type

'  flags for dwFlags field of WAVEHDR
Private Const WHDR_DONE = &H1         '  done bit
Private Const WHDR_PREPARED = &H2         '  set if this header has been prepared
Private Const WHDR_BEGINLOOP = &H4         '  loop start block
Private Const WHDR_ENDLOOP = &H8         '  loop end block
Private Const WHDR_INQUEUE = &H10        '  reserved for driver
Private Const WHDR_VALID = &H1F        '  valid flags      / ;Internal /

Type WAVEOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        dwFormats As Long
        wChannels As Integer
        dwSupport As Long
End Type

'  flags for dwSupport field of WAVEOUTCAPS
Private Const WAVECAPS_PITCH = &H1         '  supports pitch control
Private Const WAVECAPS_PLAYBACKRATE = &H2         '  supports playback rate control
Private Const WAVECAPS_VOLUME = &H4         '  supports volume control
Private Const WAVECAPS_LRVOLUME = &H8         '  separate left-right volume control
Private Const WAVECAPS_SYNC = &H10

Type WAVEINCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        dwFormats As Long
        wChannels As Integer
End Type

'  defines for dwFormat field of WAVEINCAPS and WAVEOUTCAPS
Private Const WAVE_INVALIDFORMAT = &H0              '  invalid format
Private Const WAVE_FORMAT_1M08 = &H1              '  11.025 kHz, Mono,   8-bit
Private Const WAVE_FORMAT_1S08 = &H2              '  11.025 kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_1M16 = &H4              '  11.025 kHz, Mono,   16-bit
Private Const WAVE_FORMAT_1S16 = &H8              '  11.025 kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_2M08 = &H10             '  22.05  kHz, Mono,   8-bit
Private Const WAVE_FORMAT_2S08 = &H20             '  22.05  kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_2M16 = &H40             '  22.05  kHz, Mono,   16-bit
Private Const WAVE_FORMAT_2S16 = &H80             '  22.05  kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_4M08 = &H100            '  44.1   kHz, Mono,   8-bit
Private Const WAVE_FORMAT_4S08 = &H200            '  44.1   kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_4M16 = &H400            '  44.1   kHz, Mono,   16-bit
Private Const WAVE_FORMAT_4S16 = &H800            '  44.1   kHz, Stereo, 16-bit

'  flags for wFormatTag field of WAVEFORMAT
Private Const WAVE_FORMAT_PCM = 1  '  Needed in resource files so outside #ifndef RC_INVOKED

Type WAVEFORMAT
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
End Type

Type PCMWAVEFORMAT
        wf As WAVEFORMAT
        wBitsPerSample As Integer
End Type

Declare Function waveOutGetNumDevs Lib "winmm.dll" Alias "waveOutGetNumDevs" () As Long
Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long

Declare Function waveOutGetVolume Lib "winmm.dll" Alias "waveOutGetVolume" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Declare Function waveOutSetVolume Lib "winmm.dll" Alias "waveOutSetVolume" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long

Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long

Declare Function waveOutOpen Lib "winmm.dll" Alias "waveOutOpen" (lphWaveOut As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveOutClose Lib "winmm.dll" Alias "waveOutClose" (ByVal hWaveOut As Long) As Long
Declare Function waveOutPrepareHeader Lib "winmm.dll" Alias "waveOutPrepareHeader" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutUnprepareHeader Lib "winmm.dll" Alias "waveOutUnprepareHeader" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutWrite Lib "winmm.dll" Alias "waveOutWrite" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutPause Lib "winmm.dll" Alias "waveOutPause" (ByVal hWaveOut As Long) As Long
Declare Function waveOutRestart Lib "winmm.dll" Alias "waveOutRestart" (ByVal hWaveOut As Long) As Long
Declare Function waveOutReset Lib "winmm.dll" Alias "waveOutReset" (ByVal hWaveOut As Long) As Long
Declare Function waveOutBreakLoop Lib "winmm.dll" Alias "waveOutBreakLoop" (ByVal hWaveOut As Long) As Long
Declare Function waveOutGetPosition Lib "winmm.dll" Alias "waveOutGetPosition" (ByVal hWaveOut As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
Declare Function waveOutGetPitch Lib "winmm.dll" Alias "waveOutGetPitch" (ByVal hWaveOut As Long, lpdwPitch As Long) As Long
Declare Function waveOutSetPitch Lib "winmm.dll" Alias "waveOutSetPitch" (ByVal hWaveOut As Long, ByVal dwPitch As Long) As Long
Declare Function waveOutGetPlaybackRate Lib "winmm.dll" Alias "waveOutGetPlaybackRate" (ByVal hWaveOut As Long, lpdwRate As Long) As Long
Declare Function waveOutSetPlaybackRate Lib "winmm.dll" Alias "waveOutSetPlaybackRate" (ByVal hWaveOut As Long, ByVal dwRate As Long) As Long
Declare Function waveOutGetID Lib "winmm.dll" Alias "waveOutGetID" (ByVal hWaveOut As Long, lpuDeviceID As Long) As Long
Declare Function waveOutMessage Lib "winmm.dll" Alias "waveOutMessage" (ByVal hWaveOut As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Declare Function waveInGetNumDevs Lib "winmm.dll" Alias "waveInGetNumDevs" () As Long

Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long

Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long

Declare Function waveInOpen Lib "winmm.dll" Alias "waveInOpen" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveInClose Lib "winmm.dll" Alias "waveInClose" (ByVal hWaveIn As Long) As Long
Declare Function waveInPrepareHeader Lib "winmm.dll" Alias "waveInPrepareHeader" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInUnprepareHeader Lib "winmm.dll" Alias "waveInUnprepareHeader" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInAddBuffer Lib "winmm.dll" Alias "waveInAddBuffer" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInStart Lib "winmm.dll" Alias "waveInStart" (ByVal hWaveIn As Long) As Long
Declare Function waveInStop Lib "winmm.dll" Alias "waveInStop" (ByVal hWaveIn As Long) As Long
Declare Function waveInReset Lib "winmm.dll" Alias "waveInReset" (ByVal hWaveIn As Long) As Long
Declare Function waveInGetPosition Lib "winmm.dll" Alias "waveInGetPosition" (ByVal hWaveIn As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
Declare Function waveInGetID Lib "winmm.dll" Alias "waveInGetID" (ByVal hWaveIn As Long, lpuDeviceID As Long) As Long
Declare Function waveInMessage Lib "winmm.dll" Alias "waveInMessage" (ByVal hWaveIn As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

'  MIDI error return values
Private Const MIDIERR_UNPREPARED = (MIDIERR_BASE + 0)   '  header not prepared
Private Const MIDIERR_STILLPLAYING = (MIDIERR_BASE + 1)   '  still something playing
Private Const MIDIERR_NOMAP = (MIDIERR_BASE + 2)   '  no current map
Private Const MIDIERR_NOTREADY = (MIDIERR_BASE + 3)   '  hardware is still busy
Private Const MIDIERR_NODEVICE = (MIDIERR_BASE + 4)   '  port no longer connected
Private Const MIDIERR_INVALIDSETUP = (MIDIERR_BASE + 5)   '  invalid setup
Private Const MIDIERR_LASTERROR = (MIDIERR_BASE + 5)   '  last error in range

'  MIDI callback messages
Private Const MIM_OPEN = MM_MIM_OPEN
Private Const MIM_CLOSE = MM_MIM_CLOSE
Private Const MIM_DATA = MM_MIM_DATA
Private Const MIM_LONGDATA = MM_MIM_LONGDATA
Private Const MIM_ERROR = MM_MIM_ERROR
Private Const MIM_LONGERROR = MM_MIM_LONGERROR
Private Const MOM_OPEN = MM_MOM_OPEN
Private Const MOM_CLOSE = MM_MOM_CLOSE
Private Const MOM_DONE = MM_MOM_DONE

'  device ID for MIDI mapper
Private Const MIDIMAPPER = (-1)  '  Cannot be cast to DWORD as RC complains
Private Const MIDI_MAPPER = -1&

'  flags for wFlags parm of midiOutCachePatches(), midiOutCacheDrumPatches()
Private Const MIDI_CACHE_ALL = 1
Private Const MIDI_CACHE_BESTFIT = 2
Private Const MIDI_CACHE_QUERY = 3
Private Const MIDI_UNCACHE = 4
Private Const MIDI_CACHE_VALID = (MIDI_CACHE_ALL Or MIDI_CACHE_BESTFIT Or MIDI_CACHE_QUERY Or MIDI_UNCACHE)  '  ;Internal

Type MIDIOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        wTechnology As Integer
        wVoices As Integer
        wNotes As Integer
        wChannelMask As Integer
        dwSupport As Long
End Type

'  flags for wTechnology field of MIDIOUTCAPS structure
Private Const MOD_MIDIPORT = 1  '  output port
Private Const MOD_SYNTH = 2  '  generic internal synth
Private Const MOD_SQSYNTH = 3  '  square wave internal synth
Private Const MOD_FMSYNTH = 4  '  FM internal synth
Private Const MOD_MAPPER = 5  '  MIDI mapper

'  flags for dwSupport field of MIDIOUTCAPS
Private Const MIDICAPS_VOLUME = &H1         '  supports volume control
Private Const MIDICAPS_LRVOLUME = &H2         '  separate left-right volume control
Private Const MIDICAPS_CACHE = &H4

Type MIDIINCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
End Type

Type MIDIHDR
        lpData As String
        dwBufferLength As Long
        dwBytesRecorded As Long
        dwUser As Long
        dwFlags As Long
        lpNext As Long
        Reserved As Long
End Type

'  flags for dwFlags field of MIDIHDR structure
Private Const MHDR_DONE = &H1         '  done bit
Private Const MHDR_PREPARED = &H2         '  set if header prepared
Private Const MHDR_INQUEUE = &H4         '  reserved for driver
Private Const MHDR_VALID = &H7         '  valid flags / ;Internal /

Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long

Declare Function midiOutGetVolume Lib "winmm.dll" Alias "midiOutGetVolume" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Declare Function midiOutSetVolume Lib "winmm.dll" Alias "midiOutSetVolume" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long

Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long

Declare Function midiOutOpen Lib "winmm.dll" Alias "midiOutOpen" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiOutClose Lib "winmm.dll" Alias "midiOutClose" (ByVal hMidiOut As Long) As Long
Declare Function midiOutPrepareHeader Lib "winmm.dll" Alias "midiOutPrepareHeader" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiOutUnprepareHeader Lib "winmm.dll" Alias "midiOutUnprepareHeader" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiOutShortMsg Lib "winmm.dll" Alias "midiOutShortMsg" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Declare Function midiOutLongMsg Lib "winmm.dll" Alias "midiOutLongMsg" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiOutReset Lib "winmm.dll" Alias "midiOutReset" (ByVal hMidiOut As Long) As Long
Declare Function midiOutCachePatches Lib "winmm.dll" Alias "midiOutCachePatches" (ByVal hMidiOut As Long, ByVal uBank As Long, lpPatchArray As Long, ByVal uFlags As Long) As Long
Declare Function midiOutCacheDrumPatches Lib "winmm.dll" Alias "midiOutCacheDrumPatches" (ByVal hMidiOut As Long, ByVal uPatch As Long, lpKeyArray As Long, ByVal uFlags As Long) As Long
Declare Function midiOutGetID Lib "winmm.dll" Alias "midiOutGetID" (ByVal hMidiOut As Long, lpuDeviceID As Long) As Long
Declare Function midiOutMessage Lib "winmm.dll" Alias "midiOutMessage" (ByVal hMidiOut As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Declare Function midiInGetNumDevs Lib "winmm.dll" Alias "midiInGetNumDevs" () As Long

Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIINCAPS, ByVal uSize As Long) As Long

Declare Function midiInGetErrorText Lib "winmm.dll" Alias "midiInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long

Declare Function midiInOpen Lib "winmm.dll" Alias "midiInOpen" (lphMidiIn As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiInClose Lib "winmm.dll" Alias "midiInClose" (ByVal hMidiIn As Long) As Long
Declare Function midiInPrepareHeader Lib "winmm.dll" Alias "midiInPrepareHeader" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiInUnprepareHeader Lib "winmm.dll" Alias "midiInUnprepareHeader" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiInAddBuffer Lib "winmm.dll" Alias "midiInAddBuffer" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiInStart Lib "winmm.dll" Alias "midiInStart" (ByVal hMidiIn As Long) As Long
Declare Function midiInStop Lib "winmm.dll" Alias "midiInStop" (ByVal hMidiIn As Long) As Long
Declare Function midiInReset Lib "winmm.dll" Alias "midiInReset" (ByVal hMidiIn As Long) As Long
Declare Function midiInGetID Lib "winmm.dll" Alias "midiInGetID" (ByVal hMidiIn As Long, lpuDeviceID As Long) As Long
Declare Function midiInMessage Lib "winmm.dll" Alias "midiInMessage" (ByVal hMidiIn As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

'  device ID for aux device mapper
Private Const AUX_MAPPER = -1&

Type AUXCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        wTechnology As Integer
        dwSupport As Long
End Type

'  flags for wTechnology field in AUXCAPS structure
Private Const AUXCAPS_CDAUDIO = 1  '  audio from internal CD-ROM drive
Private Const AUXCAPS_AUXIN = 2  '  audio from auxiliary input jacks

'  flags for dwSupport field in AUXCAPS structure
Private Const AUXCAPS_VOLUME = &H1         '  supports volume control
Private Const AUXCAPS_LRVOLUME = &H2         '  separate left-right volume control

Declare Function auxGetNumDevs Lib "winmm.dll" Alias "auxGetNumDevs" () As Long
Declare Function auxGetDevCaps Lib "winmm.dll" Alias "auxGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As AUXCAPS, ByVal uSize As Long) As Long

Declare Function auxSetVolume Lib "winmm.dll" Alias "auxSetVolume" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function auxGetVolume Lib "winmm.dll" Alias "auxGetVolume" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Declare Function auxOutMessage Lib "winmm.dll" Alias "auxOutMessage" (ByVal uDeviceID As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

'  timer error return values
Private Const TIMERR_NOERROR = (0)  '  no error
Private Const TIMERR_NOCANDO = (TIMERR_BASE + 1) '  request not completed
Private Const TIMERR_STRUCT = (TIMERR_BASE + 33) '  time struct size

'  flags for wFlags parameter of timeSetEvent() function
Private Const TIME_ONESHOT = 0  '  program timer for single event
Private Const TIME_PERIODIC = 1  '  program for continuous periodic event

Type TIMECAPS
        wPeriodMin As Long
        wPeriodMax As Long
End Type

Declare Function timeGetSystemTime Lib "winmm.dll" Alias "timeGetSystemTime" (lpTime As MMTIME, ByVal uSize As Long) As Long
Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Long
Declare Function timeSetEvent Lib "winmm.dll" Alias "timeSetEvent" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
Declare Function timeKillEvent Lib "winmm.dll" Alias "timeKillEvent" (ByVal uID As Long) As Long
Declare Function timeGetDevCaps Lib "winmm.dll" Alias "timeGetDevCaps" (lpTimeCaps As TIMECAPS, ByVal uSize As Long) As Long
Declare Function timeBeginPeriod Lib "winmm.dll" Alias "timeBeginPeriod" (ByVal uPeriod As Long) As Long
Declare Function timeEndPeriod Lib "winmm.dll" Alias "timeEndPeriod" (ByVal uPeriod As Long) As Long

'  joystick error return values
Private Const JOYERR_NOERROR = (0)  '  no error
Private Const JOYERR_PARMS = (JOYERR_BASE + 5) '  bad parameters
Private Const JOYERR_NOCANDO = (JOYERR_BASE + 6) '  request not completed
Private Const JOYERR_UNPLUGGED = (JOYERR_BASE + 7) '  joystick is unplugged

'  constants used with JOYINFO structure and MM_JOY messages
Private Const JOY_BUTTON1 = &H1
Private Const JOY_BUTTON2 = &H2
Private Const JOY_BUTTON3 = &H4
Private Const JOY_BUTTON4 = &H8
Private Const JOY_BUTTON1CHG = &H100
Private Const JOY_BUTTON2CHG = &H200
Private Const JOY_BUTTON3CHG = &H400
Private Const JOY_BUTTON4CHG = &H800

'  joystick ID constants
Private Const JOYSTICKID1 = 0
Private Const JOYSTICKID2 = 1

Type JOYCAPS
        wMid As Integer
        wPid As Integer
        szPname As String * MAXPNAMELEN
        wXmin As Integer
        wXmax As Integer
        wYmin As Integer
        wYmax As Integer
        wZmin As Integer
        wZmax As Integer
        wNumButtons As Integer
        wPeriodMin As Integer
        wPeriodMax As Integer
End Type

Type JOYINFO
        wXpos As Integer
        wYpos As Integer
        wZpos As Integer
        wButtons As Integer
End Type

Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long

Declare Function joyGetNumDevs Lib "winmm.dll" Alias "joyGetNumDev" () As Long
Declare Function joyGetThreshold Lib "winmm.dll" Alias "joyGetThreshold" (ByVal id As Long, lpuThreshold As Long) As Long
Declare Function joyReleaseCapture Lib "winmm.dll" Alias "joyReleaseCapture" (ByVal id As Long) As Long
Declare Function joySetCapture Lib "winmm.dll" Alias "joySetCapture" (ByVal hwnd As Long, ByVal uID As Long, ByVal uPeriod As Long, ByVal bChanged As Long) As Long
Declare Function joySetThreshold Lib "winmm.dll" Alias "joySetThreshold" (ByVal id As Long, ByVal uThreshold As Long) As Long

'  MMIO error return values
Private Const MMIOERR_BASE = 256
Private Const MMIOERR_FILENOTFOUND = (MMIOERR_BASE + 1)  '  file not found
Private Const MMIOERR_OUTOFMEMORY = (MMIOERR_BASE + 2)  '  out of memory
Private Const MMIOERR_CANNOTOPEN = (MMIOERR_BASE + 3)  '  cannot open
Private Const MMIOERR_CANNOTCLOSE = (MMIOERR_BASE + 4)  '  cannot close
Private Const MMIOERR_CANNOTREAD = (MMIOERR_BASE + 5)  '  cannot read
Private Const MMIOERR_CANNOTWRITE = (MMIOERR_BASE + 6) '  cannot write
Private Const MMIOERR_CANNOTSEEK = (MMIOERR_BASE + 7)  '  cannot seek
Private Const MMIOERR_CANNOTEXPAND = (MMIOERR_BASE + 8)  '  cannot expand file
Private Const MMIOERR_CHUNKNOTFOUND = (MMIOERR_BASE + 9)  '  chunk not found
Private Const MMIOERR_UNBUFFERED = (MMIOERR_BASE + 10) '  file is unbuffered

'  MMIO constants
Private Const CFSEPCHAR = "+"  '  compound file name separator char.

Type MMIOINFO
        dwFlags As Long
        fccIOProc As Long
        pIOProc As Long
        wErrorRet As Long
        htask As Long
        cchBuffer As Long
        pchBuffer As String
        pchNext As String
        pchEndRead As String
        pchEndWrite As String
        lBufOffset As Long
        lDiskOffset As Long
        adwInfo(4) As Long
        dwReserved1 As Long
        dwReserved2 As Long
        hmmio As Long
End Type

Private Const MMIO_RWMODE = &H3         '  mask to get bits used for opening
                                        '  file for reading/writing/both
Private Const MMIO_SHAREMODE = &H70        '  file sharing mode number

'  constants for dwFlags field of MMIOINFO
Private Const MMIO_CREATE = &H1000      '  create new file (or truncate file)
Private Const MMIO_PARSE = &H100       '  parse new file returning path
Private Const MMIO_DELETE = &H200       '  create new file (or truncate file)
Private Const MMIO_EXIST = &H4000      '  checks for existence of file
Private Const MMIO_ALLOCBUF = &H10000     '  mmioOpen() should allocate a buffer
Private Const MMIO_GETTEMP = &H20000     '  mmioOpen() should retrieve temp name

Private Const MMIO_DIRTY = &H10000000  '  I/O buffer is dirty
'  MMIO_DIRTY is also used in the <dwFlags> field of MMCKINFO structure

Private Const MMIO_OPEN_VALID = &H3FFFF     '  valid flags for mmioOpen / ;Internal /

'  read/write mode numbers (bit field MMIO_RWMODE)
Private Const MMIO_READ = &H0         '  open file for reading only
Private Const MMIO_WRITE = &H1         '  open file for writing only
Private Const MMIO_READWRITE = &H2         '  open file for reading and writing

'  share mode numbers (bit field MMIO_SHAREMODE)
Private Const MMIO_COMPAT = &H0         '  compatibility mode
Private Const MMIO_EXCLUSIVE = &H10        '  exclusive-access mode
Private Const MMIO_DENYWRITE = &H20        '  deny writing to other processes
Private Const MMIO_DENYREAD = &H30        '  deny reading to other processes
Private Const MMIO_DENYNONE = &H40        '  deny nothing to other processes

'  flags for other functions
Private Const MMIO_FHOPEN = &H10    '  mmioClose(): keep file handle open
Private Const MMIO_EMPTYBUF = &H10    '  mmioFlush(): empty the I/O buffer
Private Const MMIO_TOUPPER = &H10    '  mmioStringToFOURCC(): cvt. to u-case
Private Const MMIO_INSTALLPROC = &H10000     '  mmioInstallIOProc(): install MMIOProc
Private Const MMIO_PrivatePROC = &H10000000  '  mmioInstallIOProc: install Globally
Private Const MMIO_UNICODEPROC = &H1000000   '  mmioInstallIOProc(): Unicode MMIOProc
Private Const MMIO_REMOVEPROC = &H20000     '  mmioInstallIOProc(): remove MMIOProc
Private Const MMIO_FINDPROC = &H40000     '  mmioInstallIOProc(): find an MMIOProc
Private Const MMIO_FINDCHUNK = &H10    '  mmioDescend(): find a chunk by ID
Private Const MMIO_FINDRIFF = &H20    '  mmioDescend(): find a LIST chunk
Private Const MMIO_FINDLIST = &H40    '  mmioDescend(): find a RIFF chunk
Private Const MMIO_CREATERIFF = &H20    '  mmioCreateChunk(): make a LIST chunk
Private Const MMIO_CREATELIST = &H40    '  mmioCreateChunk(): make a RIFF chunk

Private Const MMIO_VALIDPROC = &H11070000  '  valid for mmioInstallIOProc / ;Internal /

'  message numbers for MMIOPROC I/O procedure functions
Private Const MMIOM_READ = MMIO_READ  '  read (must equal MMIO_READ!)
Private Const MMIOM_WRITE = MMIO_WRITE  '  write (must equal MMIO_WRITE!)
Private Const MMIOM_SEEK = 2  '  seek to a new position in file
Private Const MMIOM_OPEN = 3  '  open file
Private Const MMIOM_CLOSE = 4  '  close file
Private Const MMIOM_WRITEFLUSH = 5  '  write and flush
Private Const MMIOM_RENAME = 6  '  rename specified file
Private Const MMIOM_USER = &H8000  '  beginning of user-defined messages

'  flags for mmioSeek()
Private Const SEEK_SET = 0  '  seek to an absolute position
Private Const SEEK_CUR = 1  '  seek relative to current position
Private Const SEEK_END = 2  '  seek relative to end of file

'  other constants
Private Const MMIO_DEFAULTBUFFER = 8192  '  default buffer size

Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type

Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long

Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As MMIOINFO, ByVal dwOpenFlags As Long) As Long

Declare Function mmioRename Lib "winmm.dll" Alias "mmioRenameA" (ByVal szFileName As String, ByVal SzNewFileName As String, lpmmioinfo As MMIOINFO, ByVal dwRenameFlags As Long) As Long

Declare Function mmioClose Lib "winmm.dll" Alias "mmioClose" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioRead Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Declare Function mmioWrite Lib "winmm.dll" Alias "mmioWrite" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Declare Function mmioSeek Lib "winmm.dll" Alias "mmioSeek" (ByVal hmmio As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function mmioGetInfo Lib "winmm.dll" Alias "mmioGetInfo" (ByVal hmmio As Long, lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Declare Function mmioSetInfo Lib "winmm.dll" Alias "mmioSetInfo" (ByVal hmmio As Long, lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Declare Function mmioSetBuffer Lib "winmm.dll" Alias "mmioSetBuffer" (ByVal hmmio As Long, ByVal pchBuffer As String, ByVal cchBuffer As Long, ByVal uFlags As Long) As Long
Declare Function mmioFlush Lib "winmm.dll" Alias "mmioFlush" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioAdvance Lib "winmm.dll" Alias "mmioAdvance" (ByVal hmmio As Long, lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Declare Function mmioSendMessage Lib "winmm.dll" Alias "mmioSendMessage" (ByVal hmmio As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long

Declare Function mmioDescend Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioAscend Lib "winmm.dll" Alias "mmioAscend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioCreateChunk Lib "winmm.dll" Alias "mmioCreateChunk" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long

' MCI functions

Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Declare Function mciGetCreatorTask Lib "winmm.dll" Alias "mciGetCreatorTask" (ByVal wDeviceID As Long) As Long

Declare Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long

Declare Function mciGetDeviceIDFromElementID Lib "winmm.dll" Alias "mciGetDeviceIDFromElementIDA" (ByVal dwElementID As Long, ByVal lpstrType As String) As Long

Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Declare Function mciExecute Lib "winmm.dll" Alias "mciExecute" (ByVal lpstrCommand As String) As Long

'   MCI error return values
Private Const MCIERR_INVALID_DEVICE_ID = (MCIERR_BASE + 1)
Private Const MCIERR_UNRECOGNIZED_KEYWORD = (MCIERR_BASE + 3)
Private Const MCIERR_UNRECOGNIZED_COMMAND = (MCIERR_BASE + 5)
Private Const MCIERR_HARDWARE = (MCIERR_BASE + 6)
Private Const MCIERR_INVALID_DEVICE_NAME = (MCIERR_BASE + 7)
Private Const MCIERR_OUT_OF_MEMORY = (MCIERR_BASE + 8)
Private Const MCIERR_DEVICE_OPEN = (MCIERR_BASE + 9)
Private Const MCIERR_CANNOT_LOAD_DRIVER = (MCIERR_BASE + 10)
Private Const MCIERR_MISSING_COMMAND_STRING = (MCIERR_BASE + 11)
Private Const MCIERR_PARAM_OVERFLOW = (MCIERR_BASE + 12)
Private Const MCIERR_MISSING_STRING_ARGUMENT = (MCIERR_BASE + 13)
Private Const MCIERR_BAD_INTEGER = (MCIERR_BASE + 14)
Private Const MCIERR_PARSER_INTERNAL = (MCIERR_BASE + 15)
Private Const MCIERR_DRIVER_INTERNAL = (MCIERR_BASE + 16)
Private Const MCIERR_MISSING_PARAMETER = (MCIERR_BASE + 17)
Private Const MCIERR_UNSUPPORTED_FUNCTION = (MCIERR_BASE + 18)
Private Const MCIERR_FILE_NOT_FOUND = (MCIERR_BASE + 19)
Private Const MCIERR_DEVICE_NOT_READY = (MCIERR_BASE + 20)
Private Const MCIERR_INTERNAL = (MCIERR_BASE + 21)
Private Const MCIERR_DRIVER = (MCIERR_BASE + 22)
Private Const MCIERR_CANNOT_USE_ALL = (MCIERR_BASE + 23)
Private Const MCIERR_MULTIPLE = (MCIERR_BASE + 24)
Private Const MCIERR_EXTENSION_NOT_FOUND = (MCIERR_BASE + 25)
Private Const MCIERR_OUTOFRANGE = (MCIERR_BASE + 26)
Private Const MCIERR_FLAGS_NOT_COMPATIBLE = (MCIERR_BASE + 28)
Private Const MCIERR_FILE_NOT_SAVED = (MCIERR_BASE + 30)
Private Const MCIERR_DEVICE_TYPE_REQUIRED = (MCIERR_BASE + 31)
Private Const MCIERR_DEVICE_LOCKED = (MCIERR_BASE + 32)
Private Const MCIERR_DUPLICATE_ALIAS = (MCIERR_BASE + 33)
Private Const MCIERR_BAD_CONSTANT = (MCIERR_BASE + 34)
Private Const MCIERR_MUST_USE_SHAREABLE = (MCIERR_BASE + 35)
Private Const MCIERR_MISSING_DEVICE_NAME = (MCIERR_BASE + 36)
Private Const MCIERR_BAD_TIME_FORMAT = (MCIERR_BASE + 37)
Private Const MCIERR_NO_CLOSING_QUOTE = (MCIERR_BASE + 38)
Private Const MCIERR_DUPLICATE_FLAGS = (MCIERR_BASE + 39)
Private Const MCIERR_INVALID_FILE = (MCIERR_BASE + 40)
Private Const MCIERR_NULL_PARAMETER_BLOCK = (MCIERR_BASE + 41)
Private Const MCIERR_UNNAMED_RESOURCE = (MCIERR_BASE + 42)
Private Const MCIERR_NEW_REQUIRES_ALIAS = (MCIERR_BASE + 43)
Private Const MCIERR_NOTIFY_ON_AUTO_OPEN = (MCIERR_BASE + 44)
Private Const MCIERR_NO_ELEMENT_ALLOWED = (MCIERR_BASE + 45)
Private Const MCIERR_NONAPPLICABLE_FUNCTION = (MCIERR_BASE + 46)
Private Const MCIERR_ILLEGAL_FOR_AUTO_OPEN = (MCIERR_BASE + 47)
Private Const MCIERR_FILENAME_REQUIRED = (MCIERR_BASE + 48)
Private Const MCIERR_EXTRA_CHARACTERS = (MCIERR_BASE + 49)
Private Const MCIERR_DEVICE_NOT_INSTALLED = (MCIERR_BASE + 50)
Private Const MCIERR_GET_CD = (MCIERR_BASE + 51)
Private Const MCIERR_SET_CD = (MCIERR_BASE + 52)
Private Const MCIERR_SET_DRIVE = (MCIERR_BASE + 53)
Private Const MCIERR_DEVICE_LENGTH = (MCIERR_BASE + 54)
Private Const MCIERR_DEVICE_ORD_LENGTH = (MCIERR_BASE + 55)
Private Const MCIERR_NO_INTEGER = (MCIERR_BASE + 56)

Private Const MCIERR_WAVE_OUTPUTSINUSE = (MCIERR_BASE + 64)
Private Const MCIERR_WAVE_SETOUTPUTINUSE = (MCIERR_BASE + 65)
Private Const MCIERR_WAVE_INPUTSINUSE = (MCIERR_BASE + 66)
Private Const MCIERR_WAVE_SETINPUTINUSE = (MCIERR_BASE + 67)
Private Const MCIERR_WAVE_OUTPUTUNSPECIFIED = (MCIERR_BASE + 68)
Private Const MCIERR_WAVE_INPUTUNSPECIFIED = (MCIERR_BASE + 69)
Private Const MCIERR_WAVE_OUTPUTSUNSUITABLE = (MCIERR_BASE + 70)
Private Const MCIERR_WAVE_SETOUTPUTUNSUITABLE = (MCIERR_BASE + 71)
Private Const MCIERR_WAVE_INPUTSUNSUITABLE = (MCIERR_BASE + 72)
Private Const MCIERR_WAVE_SETINPUTUNSUITABLE = (MCIERR_BASE + 73)

Private Const MCIERR_SEQ_DIV_INCOMPATIBLE = (MCIERR_BASE + 80)
Private Const MCIERR_SEQ_PORT_INUSE = (MCIERR_BASE + 81)
Private Const MCIERR_SEQ_PORT_NONEXISTENT = (MCIERR_BASE + 82)
Private Const MCIERR_SEQ_PORT_MAPNODEVICE = (MCIERR_BASE + 83)
Private Const MCIERR_SEQ_PORT_MISCERROR = (MCIERR_BASE + 84)
Private Const MCIERR_SEQ_TIMER = (MCIERR_BASE + 85)
Private Const MCIERR_SEQ_PORTUNSPECIFIED = (MCIERR_BASE + 86)
Private Const MCIERR_SEQ_NOMIDIPRESENT = (MCIERR_BASE + 87)

Private Const MCIERR_NO_WINDOW = (MCIERR_BASE + 90)
Private Const MCIERR_CREATEWINDOW = (MCIERR_BASE + 91)
Private Const MCIERR_FILE_READ = (MCIERR_BASE + 92)
Private Const MCIERR_FILE_WRITE = (MCIERR_BASE + 93)

'  All custom device driver errors must be >= this value
Private Const MCIERR_CUSTOM_DRIVER_BASE = (MCIERR_BASE + 256)

'  Message numbers must be in the range between MCI_FIRST and MCI_LAST

Private Const MCI_FIRST = &H800
'  Messages 0x801 and 0x802 are reserved
Private Const MCI_OPEN = &H803
Private Const MCI_CLOSE = &H804
Private Const MCI_ESCAPE = &H805
Private Const MCI_PLAY = &H806
Private Const MCI_SEEK = &H807
Private Const MCI_STOP = &H808
Private Const MCI_PAUSE = &H809
Private Const MCI_INFO = &H80A
Private Const MCI_GETDEVCAPS = &H80B
Private Const MCI_SPIN = &H80C
Private Const MCI_SET = &H80D
Private Const MCI_STEP = &H80E
Private Const MCI_RECORD = &H80F
Private Const MCI_SYSINFO = &H810
Private Const MCI_BREAK = &H811
Private Const MCI_SOUND = &H812
Private Const MCI_SAVE = &H813
Private Const MCI_STATUS = &H814

Private Const MCI_CUE = &H830

Private Const MCI_REALIZE = &H840
Private Const MCI_WINDOW = &H841
Private Const MCI_PUT = &H842
Private Const MCI_WHERE = &H843
Private Const MCI_FREEZE = &H844
Private Const MCI_UNFREEZE = &H845

Private Const MCI_LOAD = &H850
Private Const MCI_CUT = &H851
Private Const MCI_COPY = &H852
Private Const MCI_PASTE = &H853
Private Const MCI_UPDATE = &H854
Private Const MCI_RESUME = &H855
Private Const MCI_DELETE = &H856

Private Const MCI_LAST = &HFFF

'  the next 0x400 message ID's are reserved for custom drivers
'  all custom MCI command messages must be >= than this value
Private Const MCI_USER_MESSAGES = (&H400 + MCI_FIRST)
Private Const MCI_ALL_DEVICE_ID = -1   '  Matches all MCI devices

'  constants for predefined MCI device types
Private Const MCI_DEVTYPE_VCR = 513
Private Const MCI_DEVTYPE_VIDEODISC = 514
Private Const MCI_DEVTYPE_OVERLAY = 515
Private Const MCI_DEVTYPE_CD_AUDIO = 516
Private Const MCI_DEVTYPE_DAT = 517
Private Const MCI_DEVTYPE_SCANNER = 518
Private Const MCI_DEVTYPE_ANIMATION = 519
Private Const MCI_DEVTYPE_DIGITAL_VIDEO = 520
Private Const MCI_DEVTYPE_OTHER = 521
Private Const MCI_DEVTYPE_WAVEFORM_AUDIO = 522
Private Const MCI_DEVTYPE_SEQUENCER = 523

Private Const MCI_DEVTYPE_FIRST = MCI_DEVTYPE_VCR
Private Const MCI_DEVTYPE_LAST = MCI_DEVTYPE_SEQUENCER

Private Const MCI_DEVTYPE_FIRST_USER = &H1000

'  return values for 'status mode' command
Private Const MCI_MODE_NOT_READY = (MCI_STRING_OFFSET + 12)
Private Const MCI_MODE_STOP = (MCI_STRING_OFFSET + 13)
Private Const MCI_MODE_PLAY = (MCI_STRING_OFFSET + 14)
Private Const MCI_MODE_RECORD = (MCI_STRING_OFFSET + 15)
Private Const MCI_MODE_SEEK = (MCI_STRING_OFFSET + 16)
Private Const MCI_MODE_PAUSE = (MCI_STRING_OFFSET + 17)
Private Const MCI_MODE_OPEN = (MCI_STRING_OFFSET + 18)

'  constants used in 'set time format' and 'status time format' commands
Private Const MCI_FORMAT_MILLISECONDS = 0
Private Const MCI_FORMAT_HMS = 1
Private Const MCI_FORMAT_MSF = 2
Private Const MCI_FORMAT_FRAMES = 3
Private Const MCI_FORMAT_SMPTE_24 = 4
Private Const MCI_FORMAT_SMPTE_25 = 5
Private Const MCI_FORMAT_SMPTE_30 = 6
Private Const MCI_FORMAT_SMPTE_30DROP = 7
Private Const MCI_FORMAT_BYTES = 8
Private Const MCI_FORMAT_SAMPLES = 9
Private Const MCI_FORMAT_TMSF = 10

'  Flags for wParam of the MM_MCINOTIFY message
Private Const MCI_NOTIFY_SUCCESSFUL = &H1
Private Const MCI_NOTIFY_SUPERSEDED = &H2
Private Const MCI_NOTIFY_ABORTED = &H4
Private Const MCI_NOTIFY_FAILURE = &H8

'  common flags for dwFlags parameter of MCI command messages
Private Const MCI_NOTIFY = &H1&
Private Const MCI_WAIT = &H2&
Private Const MCI_FROM = &H4&
Private Const MCI_TO = &H8&
Private Const MCI_TRACK = &H10&

'  flags for dwFlags parameter of MCI_OPEN command message
Private Const MCI_OPEN_SHAREABLE = &H100&
Private Const MCI_OPEN_ELEMENT = &H200&
Private Const MCI_OPEN_ALIAS = &H400&
Private Const MCI_OPEN_ELEMENT_ID = &H800&
Private Const MCI_OPEN_TYPE_ID = &H1000&
Private Const MCI_OPEN_TYPE = &H2000&

'  flags for dwFlags parameter of MCI_SEEK command message
Private Const MCI_SEEK_TO_START = &H100&
Private Const MCI_SEEK_TO_END = &H200&

'  flags for dwFlags parameter of MCI_STATUS command message
Private Const MCI_STATUS_ITEM = &H100&
Private Const MCI_STATUS_START = &H200&

'  flags for dwItem field of the MCI_STATUS_PARMS parameter block
Private Const MCI_STATUS_LENGTH = &H1&
Private Const MCI_STATUS_POSITION = &H2&
Private Const MCI_STATUS_NUMBER_OF_TRACKS = &H3&
Private Const MCI_STATUS_MODE = &H4&
Private Const MCI_STATUS_MEDIA_PRESENT = &H5&
Private Const MCI_STATUS_TIME_FORMAT = &H6&
Private Const MCI_STATUS_READY = &H7&
Private Const MCI_STATUS_CURRENT_TRACK = &H8&

'  flags for dwFlags parameter of MCI_INFO command message
Private Const MCI_INFO_PRODUCT = &H100&
Private Const MCI_INFO_FILE = &H200&

'  flags for dwFlags parameter of MCI_GETDEVCAPS command message
Private Const MCI_GETDEVCAPS_ITEM = &H100&

'  flags for dwItem field of the MCI_GETDEVCAPS_PARMS parameter block
Private Const MCI_GETDEVCAPS_CAN_RECORD = &H1&
Private Const MCI_GETDEVCAPS_HAS_AUDIO = &H2&
Private Const MCI_GETDEVCAPS_HAS_VIDEO = &H3&
Private Const MCI_GETDEVCAPS_DEVICE_TYPE = &H4&
Private Const MCI_GETDEVCAPS_USES_FILES = &H5&
Private Const MCI_GETDEVCAPS_COMPOUND_DEVICE = &H6&
Private Const MCI_GETDEVCAPS_CAN_EJECT = &H7&
Private Const MCI_GETDEVCAPS_CAN_PLAY = &H8&
Private Const MCI_GETDEVCAPS_CAN_SAVE = &H9&

'  flags for dwFlags parameter of MCI_SYSINFO command message
Private Const MCI_SYSINFO_QUANTITY = &H100&
Private Const MCI_SYSINFO_OPEN = &H200&
Private Const MCI_SYSINFO_NAME = &H400&
Private Const MCI_SYSINFO_INSTALLNAME = &H800&

'  flags for dwFlags parameter of MCI_SET command message
Private Const MCI_SET_DOOR_OPEN = &H100&
Private Const MCI_SET_DOOR_CLOSED = &H200&
Private Const MCI_SET_TIME_FORMAT = &H400&
Private Const MCI_SET_AUDIO = &H800&
Private Const MCI_SET_VIDEO = &H1000&
Private Const MCI_SET_ON = &H2000&
Private Const MCI_SET_OFF = &H4000&

'  flags for dwAudio field of MCI_SET_PARMS or MCI_SEQ_SET_PARMS
Private Const MCI_SET_AUDIO_ALL = &H4001&
Private Const MCI_SET_AUDIO_LEFT = &H4002&
Private Const MCI_SET_AUDIO_RIGHT = &H4003&

'  flags for dwFlags parameter of MCI_BREAK command message
Private Const MCI_BREAK_KEY = &H100&
Private Const MCI_BREAK_HWND = &H200&
Private Const MCI_BREAK_OFF = &H400&

'  flags for dwFlags parameter of MCI_RECORD command message
Private Const MCI_RECORD_INSERT = &H100&
Private Const MCI_RECORD_OVERWRITE = &H200&

'  flags for dwFlags parameter of MCI_SOUND command message
Private Const MCI_SOUND_NAME = &H100&

'  flags for dwFlags parameter of MCI_SAVE command message
Private Const MCI_SAVE_FILE = &H100&

'  flags for dwFlags parameter of MCI_LOAD command message
Private Const MCI_LOAD_FILE = &H100&

Type MCI_GENERIC_PARMS
        dwCallback As Long
End Type

Type MCI_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDeviceType As String
        lpstrElementName As String
        lpstrAlias As String
End Type

Type MCI_PLAY_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
End Type

Type MCI_SEEK_PARMS
        dwCallback As Long
        dwTo As Long
End Type

Type MCI_STATUS_PARMS
        dwCallback As Long
        dwReturn As Long
        dwItem As Long
        dwTrack As Integer
End Type

Type MCI_INFO_PARMS
        dwCallback As Long
        lpstrReturn As String
        dwRetSize As Long
End Type

Type MCI_GETDEVCAPS_PARMS
        dwCallback As Long
        dwReturn As Long
        dwIten As Long
End Type

Type MCI_SYSINFO_PARMS
        dwCallback As Long
        lpstrReturn As String
        dwRetSize As Long
        dwNumber As Long
        wDeviceType As Long
End Type

Type MCI_SET_PARMS
        dwCallback As Long
        dwTimeFormat As Long
        dwAudio As Long
End Type

Type MCI_BREAK_PARMS
        dwCallback As Long
        nVirtKey As Long
        hwndBreak As Long
End Type

Type MCI_SOUND_PARMS
        dwCallback As Long
        lpstrSoundName As String
End Type

Type MCI_SAVE_PARMS
        dwCallback As Long
        lpFileName As String
End Type

Type MCI_LOAD_PARMS
        dwCallback As Long
        lpFileName As String
End Type

Type MCI_RECORD_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
End Type

Private Const MCI_VD_MODE_PARK = (MCI_VD_OFFSET + 1)

'  return ID's for videodisc MCI_GETDEVCAPS command

'  flag for dwReturn field of MCI_STATUS_PARMS
'  MCI_STATUS command, (dwItem == MCI_VD_STATUS_MEDIA_TYPE)
Private Const MCI_VD_MEDIA_CLV = (MCI_VD_OFFSET + 2)
Private Const MCI_VD_MEDIA_CAV = (MCI_VD_OFFSET + 3)
Private Const MCI_VD_MEDIA_OTHER = (MCI_VD_OFFSET + 4)

Private Const MCI_VD_FORMAT_TRACK = &H4001

'  flags for dwFlags parameter of MCI_PLAY command message
Private Const MCI_VD_PLAY_REVERSE = &H10000
Private Const MCI_VD_PLAY_FAST = &H20000
Private Const MCI_VD_PLAY_SPEED = &H40000
Private Const MCI_VD_PLAY_SCAN = &H80000
Private Const MCI_VD_PLAY_SLOW = &H100000

'  flag for dwFlags parameter of MCI_SEEK command message
Private Const MCI_VD_SEEK_REVERSE = &H10000

'  flags for dwItem field of MCI_STATUS_PARMS parameter block
Private Const MCI_VD_STATUS_SPEED = &H4002&
Private Const MCI_VD_STATUS_FORWARD = &H4003&
Private Const MCI_VD_STATUS_MEDIA_TYPE = &H4004&
Private Const MCI_VD_STATUS_SIDE = &H4005&
Private Const MCI_VD_STATUS_DISC_SIZE = &H4006&

'  flags for dwFlags parameter of MCI_GETDEVCAPS command message
Private Const MCI_VD_GETDEVCAPS_CLV = &H10000
Private Const MCI_VD_GETDEVCAPS_CAV = &H20000

Private Const MCI_VD_SPIN_UP = &H10000
Private Const MCI_VD_SPIN_DOWN = &H20000

'  flags for dwItem field of MCI_GETDEVCAPS_PARMS parameter block
Private Const MCI_VD_GETDEVCAPS_CAN_REVERSE = &H4002&
Private Const MCI_VD_GETDEVCAPS_FAST_RATE = &H4003&
Private Const MCI_VD_GETDEVCAPS_SLOW_RATE = &H4004&
Private Const MCI_VD_GETDEVCAPS_NORMAL_RATE = &H4005&

'  flags for the dwFlags parameter of MCI_STEP command message
Private Const MCI_VD_STEP_FRAMES = &H10000
Private Const MCI_VD_STEP_REVERSE = &H20000

'  flag for the MCI_ESCAPE command message
Private Const MCI_VD_ESCAPE_STRING = &H100&

Type MCI_VD_PLAY_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
        dwSpeed As Long
End Type

Type MCI_VD_STEP_PARMS
        dwCallback As Long
        dwFrames As Long
End Type

Type MCI_VD_ESCAPE_PARMS
        dwCallback As Long
        lpstrCommand As String
End Type

Private Const MCI_WAVE_PCM = (MCI_WAVE_OFFSET + 0)
Private Const MCI_WAVE_MAPPER = (MCI_WAVE_OFFSET + 1)

'  flags for the dwFlags parameter of MCI_OPEN command message
Private Const MCI_WAVE_OPEN_BUFFER = &H10000

'  flags for the dwFlags parameter of MCI_SET command message
Private Const MCI_WAVE_SET_FORMATTAG = &H10000
Private Const MCI_WAVE_SET_CHANNELS = &H20000
Private Const MCI_WAVE_SET_SAMPLESPERSEC = &H40000
Private Const MCI_WAVE_SET_AVGBYTESPERSEC = &H80000
Private Const MCI_WAVE_SET_BLOCKALIGN = &H100000
Private Const MCI_WAVE_SET_BITSPERSAMPLE = &H200000

'  flags for the dwFlags parameter of MCI_STATUS, MCI_SET command messages
Private Const MCI_WAVE_INPUT = &H400000
Private Const MCI_WAVE_OUTPUT = &H800000

'  flags for the dwItem field of MCI_STATUS_PARMS parameter block
Private Const MCI_WAVE_STATUS_FORMATTAG = &H4001&
Private Const MCI_WAVE_STATUS_CHANNELS = &H4002&
Private Const MCI_WAVE_STATUS_SAMPLESPERSEC = &H4003&
Private Const MCI_WAVE_STATUS_AVGBYTESPERSEC = &H4004&
Private Const MCI_WAVE_STATUS_BLOCKALIGN = &H4005&
Private Const MCI_WAVE_STATUS_BITSPERSAMPLE = &H4006&
Private Const MCI_WAVE_STATUS_LEVEL = &H4007&

'  flags for the dwFlags parameter of MCI_SET command message
Private Const MCI_WAVE_SET_ANYINPUT = &H4000000
Private Const MCI_WAVE_SET_ANYOUTPUT = &H8000000

'  flags for the dwFlags parameter of MCI_GETDEVCAPS command message
Private Const MCI_WAVE_GETDEVCAPS_INPUTS = &H4001&
Private Const MCI_WAVE_GETDEVCAPS_OUTPUTS = &H4002&

Type MCI_WAVE_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDeviceType As String
        lpstrElementName As String
        lpstrAlias As String
        dwBufferSeconds As Long
End Type

Type MCI_WAVE_DELETE_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
End Type

Type MCI_WAVE_SET_PARMS
        dwCallback As Long
        dwTimeFormat As Long
        dwAudio As Long
        wInput As Long
        wOutput As Long
        wFormatTag As Integer
        wReserved2 As Integer
        nChannels As Integer
        wReserved3 As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
        wReserved4 As Integer
        wBitsPerSample As Integer
        wReserved5 As Integer
End Type

'  flags for the dwReturn field of MCI_STATUS_PARMS parameter block
'  MCI_STATUS command, (dwItem == MCI_SEQ_STATUS_DIVTYPE)
Private Const MCI_SEQ_DIV_PPQN = (0 + MCI_SEQ_OFFSET)
Private Const MCI_SEQ_DIV_SMPTE_24 = (1 + MCI_SEQ_OFFSET)
Private Const MCI_SEQ_DIV_SMPTE_25 = (2 + MCI_SEQ_OFFSET)
Private Const MCI_SEQ_DIV_SMPTE_30DROP = (3 + MCI_SEQ_OFFSET)
Private Const MCI_SEQ_DIV_SMPTE_30 = (4 + MCI_SEQ_OFFSET)

'  flags for the dwMaster field of MCI_SEQ_SET_PARMS parameter block
'  MCI_SET command, (dwFlags == MCI_SEQ_SET_MASTER)
Private Const MCI_SEQ_FORMAT_SONGPTR = &H4001
Private Const MCI_SEQ_FILE = &H4002
Private Const MCI_SEQ_MIDI = &H4003
Private Const MCI_SEQ_SMPTE = &H4004
Private Const MCI_SEQ_NONE = 65533

Private Const MCI_SEQ_MAPPER = 65535

'  flags for the dwItem field of MCI_STATUS_PARMS parameter block
Private Const MCI_SEQ_STATUS_TEMPO = &H4002&
Private Const MCI_SEQ_STATUS_PORT = &H4003&
Private Const MCI_SEQ_STATUS_SLAVE = &H4007&
Private Const MCI_SEQ_STATUS_MASTER = &H4008&
Private Const MCI_SEQ_STATUS_OFFSET = &H4009&
Private Const MCI_SEQ_STATUS_DIVTYPE = &H400A&

'  flags for the dwFlags parameter of MCI_SET command message
Private Const MCI_SEQ_SET_TEMPO = &H10000
Private Const MCI_SEQ_SET_PORT = &H20000
Private Const MCI_SEQ_SET_SLAVE = &H40000
Private Const MCI_SEQ_SET_MASTER = &H80000
Private Const MCI_SEQ_SET_OFFSET = &H1000000

Type MCI_SEQ_SET_PARMS
        dwCallback As Long
        dwTimeFormat As Long
        dwAudio As Long
        dwTempo As Long
        dwPort As Long
        dwSlave As Long
        dwMaster As Long
        dwOffset As Long
End Type

'  flags for dwFlags parameter of MCI_OPEN command message
Private Const MCI_ANIM_OPEN_WS = &H10000
Private Const MCI_ANIM_OPEN_PARENT = &H20000
Private Const MCI_ANIM_OPEN_NOSTATIC = &H40000

'  flags for dwFlags parameter of MCI_PLAY command message
Private Const MCI_ANIM_PLAY_SPEED = &H10000
Private Const MCI_ANIM_PLAY_REVERSE = &H20000
Private Const MCI_ANIM_PLAY_FAST = &H40000
Private Const MCI_ANIM_PLAY_SLOW = &H80000
Private Const MCI_ANIM_PLAY_SCAN = &H100000

'  flags for dwFlags parameter of MCI_STEP command message
Private Const MCI_ANIM_STEP_REVERSE = &H10000
Private Const MCI_ANIM_STEP_FRAMES = &H20000

'  flags for dwItem field of MCI_STATUS_PARMS parameter block
Private Const MCI_ANIM_STATUS_SPEED = &H4001&
Private Const MCI_ANIM_STATUS_FORWARD = &H4002&
Private Const MCI_ANIM_STATUS_HWND = &H4003&
Private Const MCI_ANIM_STATUS_HPAL = &H4004&
Private Const MCI_ANIM_STATUS_STRETCH = &H4005&

'  flags for the dwFlags parameter of MCI_INFO command message
Private Const MCI_ANIM_INFO_TEXT = &H10000

'  flags for dwItem field of MCI_GETDEVCAPS_PARMS parameter block
Private Const MCI_ANIM_GETDEVCAPS_CAN_REVERSE = &H4001&
Private Const MCI_ANIM_GETDEVCAPS_FAST_RATE = &H4002&
Private Const MCI_ANIM_GETDEVCAPS_SLOW_RATE = &H4003&
Private Const MCI_ANIM_GETDEVCAPS_NORMAL_RATE = &H4004&
Private Const MCI_ANIM_GETDEVCAPS_PALETTES = &H4006&
Private Const MCI_ANIM_GETDEVCAPS_CAN_STRETCH = &H4007&
Private Const MCI_ANIM_GETDEVCAPS_MAX_WINDOWS = &H4008&

'  flags for the MCI_REALIZE command message
Private Const MCI_ANIM_REALIZE_NORM = &H10000
Private Const MCI_ANIM_REALIZE_BKGD = &H20000

'  flags for dwFlags parameter of MCI_WINDOW command message
Private Const MCI_ANIM_WINDOW_HWND = &H10000
Private Const MCI_ANIM_WINDOW_STATE = &H40000
Private Const MCI_ANIM_WINDOW_TEXT = &H80000
Private Const MCI_ANIM_WINDOW_ENABLE_STRETCH = &H100000
Private Const MCI_ANIM_WINDOW_DISABLE_STRETCH = &H200000

'  flags for hWnd field of MCI_ANIM_WINDOW_PARMS parameter block
'  MCI_WINDOW command message, (dwFlags == MCI_ANIM_WINDOW_HWND)
Private Const MCI_ANIM_WINDOW_DEFAULT = &H0&

'  flags for dwFlags parameter of MCI_PUT command message
Private Const MCI_ANIM_RECT = &H10000
Private Const MCI_ANIM_PUT_SOURCE = &H20000      '  also  MCI_WHERE
Private Const MCI_ANIM_PUT_DESTINATION = &H40000      '  also  MCI_WHERE

'  flags for dwFlags parameter of MCI_WHERE command message
Private Const MCI_ANIM_WHERE_SOURCE = &H20000
Private Const MCI_ANIM_WHERE_DESTINATION = &H40000

'  flags for dwFlags parameter of MCI_UPDATE command message
Private Const MCI_ANIM_UPDATE_HDC = &H20000

Type MCI_ANIM_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDeviceType As String
        lpstrElementName As String
        lpstrAlias As String
        dwStyle As Long
        hwndParent As Long
End Type

Type MCI_ANIM_PLAY_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
        dwSpeed As Long
End Type

Type MCI_ANIM_STEP_PARMS
        dwCallback As Long
        dwFrames As Long
End Type

Type MCI_ANIM_WINDOW_PARMS
        dwCallback As Long
        hWnd As Long
        nCmdShow As Long
        lpstrText As String
End Type

Type MCI_ANIM_RECT_PARMS
        dwCallback As Long
        rc As RECT
End Type

Type MCI_ANIM_UPDATE_PARMS
        dwCallback As Long
        rc As RECT
        hdc As Long
End Type

'  flags for dwFlags parameter of MCI_OPEN command message
Private Const MCI_OVLY_OPEN_WS = &H10000
Private Const MCI_OVLY_OPEN_PARENT = &H20000

'  flags for dwFlags parameter of MCI_STATUS command message
Private Const MCI_OVLY_STATUS_HWND = &H4001&
Private Const MCI_OVLY_STATUS_STRETCH = &H4002&

'  flags for dwFlags parameter of MCI_INFO command message
Private Const MCI_OVLY_INFO_TEXT = &H10000

'  flags for dwItem field of MCI_GETDEVCAPS_PARMS parameter block
Private Const MCI_OVLY_GETDEVCAPS_CAN_STRETCH = &H4001&
Private Const MCI_OVLY_GETDEVCAPS_CAN_FREEZE = &H4002&
Private Const MCI_OVLY_GETDEVCAPS_MAX_WINDOWS = &H4003&

'  flags for dwFlags parameter of MCI_WINDOW command message
Private Const MCI_OVLY_WINDOW_HWND = &H10000
Private Const MCI_OVLY_WINDOW_STATE = &H40000
Private Const MCI_OVLY_WINDOW_TEXT = &H80000
Private Const MCI_OVLY_WINDOW_ENABLE_STRETCH = &H100000
Private Const MCI_OVLY_WINDOW_DISABLE_STRETCH = &H200000

'  flags for hWnd parameter of MCI_OVLY_WINDOW_PARMS parameter block
Private Const MCI_OVLY_WINDOW_DEFAULT = &H0&

'  flags for dwFlags parameter of MCI_PUT command message
Private Const MCI_OVLY_RECT = &H10000
Private Const MCI_OVLY_PUT_SOURCE = &H20000
Private Const MCI_OVLY_PUT_DESTINATION = &H40000
Private Const MCI_OVLY_PUT_FRAME = &H80000
Private Const MCI_OVLY_PUT_VIDEO = &H100000

'  flags for dwFlags parameter of MCI_WHERE command message
Private Const MCI_OVLY_WHERE_SOURCE = &H20000
Private Const MCI_OVLY_WHERE_DESTINATION = &H40000
Private Const MCI_OVLY_WHERE_FRAME = &H80000
Private Const MCI_OVLY_WHERE_VIDEO = &H100000

Type MCI_OVLY_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDeviceType As String
        lpstrElementName As String
        lpstrAlias As String
        dwStyle As Long
        hwndParent As Long
End Type

Type MCI_OVLY_WINDOW_PARMS
        dwCallback As Long
        hWnd As Long
        nCmdShow As Long
        lpstrText As String
End Type

Type MCI_OVLY_RECT_PARMS
        dwCallback As Long
        rc As RECT
End Type

Type MCI_OVLY_SAVE_PARMS
        dwCallback As Long
        lpFileName As String
        rc As RECT
End Type

Type MCI_OVLY_LOAD_PARMS
        dwCallback As Long
        lpFileName As String
        rc As RECT
End Type

Private Const CAPS1 = 94              '  other caps
Private Const C1_TRANSPARENT = &H1     '  new raster cap
Private Const NEWTRANSPARENT = 3  '  use with SetBkMode()

Private Const QUERYROPSUPPORT = 40  '  use to determine ROP support

Private Const SELECTDIB = 41  '  DIB.DRV select dib escape

' ----------------
' shell association database management functions
' -----------------

'  error values for ShellExecute() beyond the regular WinExec() codes
Private Const SE_ERR_SHARE = 26
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_NOASSOC = 31

' -------------
' Print APIs
' -------------

Type PRINTER_INFO_1
        Flags As Long
        pDescription As String
        pName As String
        pComment As String
End Type

Type PRINTER_INFO_2
        pServerName As String
        pPrinterName As String
        pShareName As String
        pPortName As String
        pDriverName As String
        pComment As String
        pLocation As String
        pDevmode As DEVMODE
        pSepFile As String
        pPrintProcessor As String
        pDatatype As String
        pParameters As String
        pSecurityDescriptor As SECURITY_DESCRIPTOR
        Attributes As Long
        Priority As Long
        DefaultPriority As Long
        StartTime As Long
        UntilTime As Long
        Status As Long
        cJobs As Long
        AveragePPM As Long
End Type

Type PRINTER_INFO_3
        pSecurityDescriptor As SECURITY_DESCRIPTOR
End Type

Private Const PRINTER_CONTROL_PAUSE = 1
Private Const PRINTER_CONTROL_RESUME = 2
Private Const PRINTER_CONTROL_PURGE = 3

Private Const PRINTER_STATUS_PAUSED = &H1
Private Const PRINTER_STATUS_ERROR = &H2
Private Const PRINTER_STATUS_PENDING_DELETION = &H4
Private Const PRINTER_STATUS_PAPER_JAM = &H8
Private Const PRINTER_STATUS_PAPER_OUT = &H10
Private Const PRINTER_STATUS_MANUAL_FEED = &H20
Private Const PRINTER_STATUS_PAPER_PROBLEM = &H40
Private Const PRINTER_STATUS_OFFLINE = &H80
Private Const PRINTER_STATUS_IO_ACTIVE = &H100
Private Const PRINTER_STATUS_BUSY = &H200
Private Const PRINTER_STATUS_PRINTING = &H400
Private Const PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
Private Const PRINTER_STATUS_NOT_AVAILABLE = &H1000
Private Const PRINTER_STATUS_WAITING = &H2000
Private Const PRINTER_STATUS_PROCESSING = &H4000
Private Const PRINTER_STATUS_INITIALIZING = &H8000
Private Const PRINTER_STATUS_WARMING_UP = &H10000
Private Const PRINTER_STATUS_TONER_LOW = &H20000
Private Const PRINTER_STATUS_NO_TONER = &H40000
Private Const PRINTER_STATUS_PAGE_PUNT = &H80000
Private Const PRINTER_STATUS_USER_INTERVENTION = &H100000
Private Const PRINTER_STATUS_OUT_OF_MEMORY = &H200000
Private Const PRINTER_STATUS_DOOR_OPEN = &H400000

Private Const PRINTER_ATTRIBUTE_QUEUED = &H1
Private Const PRINTER_ATTRIBUTE_DIRECT = &H2
Private Const PRINTER_ATTRIBUTE_DEFAULT = &H4
Private Const PRINTER_ATTRIBUTE_SHARED = &H8
Private Const PRINTER_ATTRIBUTE_NETWORK = &H10
Private Const PRINTER_ATTRIBUTE_HIDDEN = &H20
Private Const PRINTER_ATTRIBUTE_LOCAL = &H40

Private Const NO_PRIORITY = 0
Private Const MAX_PRIORITY = 99
Private Const MIN_PRIORITY = 1
Private Const DEF_PRIORITY = 1

Type JOB_INFO_1
        JobId As Long
        pPrinterName As String
        pMachineName As String
        pUserName As String
        pDocument As String
        pDatatype As String
        pStatus As String
        Status As Long
        Priority As Long
        Position As Long
        TotalPages As Long
        PagesPrinted As Long
        Submitted As SystemTime
End Type

Type JOB_INFO_2
        JobId As Long
        pPrinterName As String
        pMachineName As String
        pUserName As String
        pDocument As String
        pNotifyName As String
        pDatatype As String
        pPrintProcessor As String
        pParameters As String
        pDriverName As String
        pDevmode As DEVMODE
        pStatus As String
        pSecurityDescriptor As SECURITY_DESCRIPTOR
        Status As Long
        Priority As Long
        Position As Long
        StartTime As Long
        UntilTime As Long
        TotalPages As Long
        Size As Long
        Submitted As SystemTime
        time As Long
        PagesPrinted As Long
End Type

Private Const JOB_CONTROL_PAUSE = 1
Private Const JOB_CONTROL_RESUME = 2
Private Const JOB_CONTROL_CANCEL = 3
Private Const JOB_CONTROL_RESTART = 4

Private Const JOB_STATUS_PAUSED = &H1
Private Const JOB_STATUS_ERROR = &H2
Private Const JOB_STATUS_DELETING = &H4
Private Const JOB_STATUS_SPOOLING = &H8
Private Const JOB_STATUS_PRINTING = &H10
Private Const JOB_STATUS_OFFLINE = &H20
Private Const JOB_STATUS_PAPEROUT = &H40
Private Const JOB_STATUS_PRINTED = &H80

Private Const JOB_POSITION_UNSPECIFIED = 0

Type ADDJOB_INFO_1
        Path As String
        JobId As Long
End Type

Type DRIVER_INFO_1
        pName As String
End Type

Type DRIVER_INFO_2
        cVersion As Long
        pName As String
        pEnvironment As String
        pDriverPath As String
        pDataFile As String
        pConfigFile As String
End Type

Type DOC_INFO_1
        pDocName As String
        pOutputFile As String
        pDatatype As String
End Type

Type FORM_INFO_1
        pName As String
        Size As SIZEL
        ImageableArea As RECTL
End Type

Private Const FORM_BUILTIN = &H1

Type PRINTPROCESSOR_INFO_1
        pName As String
End Type

Type PORT_INFO_1
        pName As String
End Type

Type MONITOR_INFO_1
        pName As String
End Type

Type MONITOR_INFO_2
        pName As String
        pEnvironment As String
        pDLLName As String
End Type

Type DATATYPES_INFO_1
        pName As String
End Type

Type PRINTER_DEFAULTS
        pDatatype As String
        pDevmode As DEVMODE
        DesiredAccess As Long
End Type

Type PRINTER_INFO_4
    pPrinterName As String
    pServerName As String
    Attributes As Long
End Type

Type PRINTER_INFO_5
        pPrinterName As String
        pPortName As String
        Attributes As Long
        DeviceNotSelectedTimeout As Long
        TransmissionRetryTimeout As Long
End Type

Private Const PRINTER_CONTROL_SET_STATUS = 4
Private Const PRINTER_ATTRIBUTE_WORK_OFFLINE = &H400
Private Const PRINTER_ATTRIBUTE_ENABLE_BIDI = &H800
Private Const JOB_CONTROL_DELETE = 5
Private Const JOB_STATUS_USER_INTERVENTION = &H10000

Type DRIVER_INFO_3
        cVersion As Long
        pName As String                    '  QMS 810
        pEnvironment As String             '  Win32 x86
        pDriverPath As String              '  c:\drivers\pscript.dll
        pDataFile As String                '  c:\drivers\QMS810.PPD
        pConfigFile As String              '  c:\drivers\PSCRPTUI.DLL
        pHelpFile As String                '  c:\drivers\PSCRPTUI.HLP
        pDependentFiles As String          '
        pMonitorName As String             '  "PJL monitor"
        pDefaultDataType As String         '  "EMF"
End Type

Type DOC_INFO_2
        pDocName As String
        pOutputFile As String
        pDatatype As String
        dwMode As Long
        JobId As Long
End Type

Private Const DI_CHANNEL = 1                  '  start direct read/write channel,
Private Const DI_READ_SPOOL_JOB = 3

Type PORT_INFO_2
        pPortName As String
        pMonitorName As String
        pDescription As String
        fPortType As Long
        Reserved As Long
End Type

Private Const PORT_TYPE_WRITE = &H1
Private Const PORT_TYPE_READ = &H2
Private Const PORT_TYPE_REDIRECTED = &H4
Private Const PORT_TYPE_NET_ATTACHED = &H8

Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long

Declare Function EnumPrinterPropertySheets Lib "winspool.drv" Alias "EnumPrinterPropertySheets" (hPrinter As Long, hwnd As Long, lpfnAdd As Long, ByVal lParam As Long) As Long

Private Const PRINTER_ENUM_DEFAULT = &H1
Private Const PRINTER_ENUM_LOCAL = &H2
Private Const PRINTER_ENUM_CONNECTIONS = &H4
Private Const PRINTER_ENUM_FAVORITE = &H4
Private Const PRINTER_ENUM_NAME = &H8
Private Const PRINTER_ENUM_REMOTE = &H10
Private Const PRINTER_ENUM_SHARED = &H20
Private Const PRINTER_ENUM_NETWORK = &H40

Private Const PRINTER_ENUM_EXPAND = &H4000
Private Const PRINTER_ENUM_CONTAINER = &H8000

Private Const PRINTER_ENUM_ICONMASK = &HFF0000
Private Const PRINTER_ENUM_ICON1 = &H10000
Private Const PRINTER_ENUM_ICON2 = &H20000
Private Const PRINTER_ENUM_ICON3 = &H40000
Private Const PRINTER_ENUM_ICON4 = &H80000
Private Const PRINTER_ENUM_ICON5 = &H100000
Private Const PRINTER_ENUM_ICON6 = &H200000
Private Const PRINTER_ENUM_ICON7 = &H400000
Private Const PRINTER_ENUM_ICON8 = &H800000

Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Declare Function ResetPrinter Lib "winspool.drv" Alias "ResetPrinterA" (ByVal hPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Declare Function SetJob Lib "winspool.drv" Alias "SetJobA" (ByVal hPrinter As Long, ByVal JobId As Long, ByVal Level As Long, pJob As Byte, ByVal Command As Long) As Long
Declare Function GetJob Lib "winspool.drv" Alias "GetJobA" (ByVal hPrinter As Long, ByVal JobId As Long, ByVal Level As Long, pJob As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function AddPrinter Lib "winspool.drv" Alias "AddPrinterA" (ByVal pName As String, ByVal Level As Long, pPrinter As Any) As Long

Declare Function AddPrinterDriver Lib "winspool.drv" Alias "AddPrinterDriverA" (ByVal pName As String, ByVal Level As Long, pDriverInfo As Any) As Long

Declare Function EnumPrinterDrivers Lib "winspool.drv" Alias "EnumPrinterDriversA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverInfo As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcRetruned As Long) As Long
Declare Function GetPrinterDriver Lib "winspool.drv" Alias "GetPrinterDriverA" (ByVal hPrinter As Long, ByVal pEnvironment As String, ByVal Level As Long, pDriverInfo As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Declare Function GetPrinterDriverDirectory Lib "winspool.drv" Alias "GetPrinterDriverDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverDirectory As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Declare Function DeletePrinterDriver Lib "winspool.drv" Alias "DeletePrinterDriverA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pDriverName As String) As Long

Declare Function AddPrintProcessor Lib "winspool.drv" Alias "AddPrintProcessorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pPathName As String, ByVal pPrintProcessorName As String) As Long
Declare Function EnumPrintProcessors Lib "winspool.drv" Alias "EnumPrintProcessorsA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pPrintProcessorInfo As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function GetPrintProcessorDirectory Lib "winspool.drv" Alias "GetPrintProcessorDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, ByVal pPrintProcessorInfo As String, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Declare Function EnumPrintProcessorDatatypes Lib "winspool.drv" Alias "EnumPrintProcessorDatatypesA" (ByVal pName As String, ByVal pPrintProcessorName As String, ByVal Level As Long, pDatatypes As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcRetruned As Long) As Long
Declare Function DeletePrintProcessor Lib "winspool.drv" Alias "DeletePrintProcessorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pPrintProcessorName As String) As Long

Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As Byte) As Long
Declare Function StartPagePrinter Lib "winspool.drv" Alias "StartPagePrinter" (ByVal hPrinter As Long) As Long
Declare Function WritePrinter Lib "winspool.drv" Alias "WritePrinter" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Declare Function EndPagePrinter Lib "winspool.drv" Alias "EndPagePrinter" (ByVal hPrinter As Long) As Long
Declare Function AbortPrinter Lib "winspool.drv" Alias "AbortPrinter" (ByVal hPrinter As Long) As Long
Declare Function ReadPrinter Lib "winspool.drv" Alias "ReadPrinter" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pNoBytesRead As Long) As Long
Declare Function EndDocPrinter Lib "winspool.drv" Alias "EndDocPrinter" (ByVal hPrinter As Long) As Long

Declare Function AddJob Lib "winspool.drv" Alias "AddJobA" (ByVal hPrinter As Long, ByVal Level As Long, pData As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Declare Function ScheduleJob Lib "winspool.drv" Alias "ScheduleJob" (ByVal hPrinter As Long, ByVal JobId As Long) As Long
Declare Function PrinterProperties Lib "winspool.drv" Alias "PrinterProperties" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, pDevModeInput As DEVMODE, ByVal fMode As Long) As Long
Declare Function AdvancedDocumentProperties Lib "winspool.drv" Alias "AdvancedDocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, pDevModeInput As DEVMODE) As Long

Declare Function GetPrinterData Lib "winspool.drv" Alias "GetPrinterDataA" (ByVal hPrinter As Long, ByVal pValueName As String, pType As Long, pData As Byte, ByVal nSize As Long, pcbNeeded As Long) As Long
Declare Function SetPrinterData Lib "winspool.drv" Alias "SetPrinterDataA" (ByVal hPrinter As Long, ByVal pValueName As String, ByVal dwType As Long, pData As Byte, ByVal cbData As Long) As Long
Declare Function WaitForPrinterChange Lib "winspool.drv" Alias "WaitForPrinterChange" (ByVal hPrinter As Long, ByVal flags As Long) As Long

Private Const PRINTER_CHANGE_ADD_PRINTER = &H1
Private Const PRINTER_CHANGE_SET_PRINTER = &H2
Private Const PRINTER_CHANGE_DELETE_PRINTER = &H4
Private Const PRINTER_CHANGE_PRINTER = &HFF
Private Const PRINTER_CHANGE_ADD_JOB = &H100
Private Const PRINTER_CHANGE_SET_JOB = &H200
Private Const PRINTER_CHANGE_DELETE_JOB = &H400
Private Const PRINTER_CHANGE_WRITE_JOB = &H800
Private Const PRINTER_CHANGE_JOB = &HFF00
Private Const PRINTER_CHANGE_ADD_FORM = &H10000
Private Const PRINTER_CHANGE_SET_FORM = &H20000
Private Const PRINTER_CHANGE_DELETE_FORM = &H40000
Private Const PRINTER_CHANGE_FORM = &H70000
Private Const PRINTER_CHANGE_ADD_PORT = &H100000
Private Const PRINTER_CHANGE_CONFIGURE_PORT = &H200000
Private Const PRINTER_CHANGE_DELETE_PORT = &H400000
Private Const PRINTER_CHANGE_PORT = &H700000
Private Const PRINTER_CHANGE_ADD_PRINT_PROCESSOR = &H1000000
Private Const PRINTER_CHANGE_DELETE_PRINT_PROCESSOR = &H4000000
Private Const PRINTER_CHANGE_PRINT_PROCESSOR = &H7000000
Private Const PRINTER_CHANGE_ADD_PRINTER_DRIVER = &H10000000
Private Const PRINTER_CHANGE_DELETE_PRINTER_DRIVER = &H40000000
Private Const PRINTER_CHANGE_PRINTER_DRIVER = &H70000000
Private Const PRINTER_CHANGE_TIMEOUT = &H80000000
Private Const PRINTER_CHANGE_ALL = &H7777FFFF

Declare Function PrinterMessageBox Lib "winspool.drv" Alias "PrinterMessageBoxA" (ByVal hPrinter As Long, ByVal error As Long, ByVal hwnd As Long, ByVal pText As String, ByVal pCaption As String, ByVal dwType As Long) As Long

Private Const PRINTER_ERROR_INFORMATION = &H80000000
Private Const PRINTER_ERROR_WARNING = &H40000000
Private Const PRINTER_ERROR_SEVERE = &H20000000

Private Const PRINTER_ERROR_OUTOFPAPER = &H1
Private Const PRINTER_ERROR_JAM = &H2
Private Const PRINTER_ERROR_OUTOFTONER = &H4

Declare Function ClosePrinter Lib "winspool.drv" Alias "ClosePrinter" (ByVal hPrinter As Long) As Long
Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte) As Long
Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As Long, ByVal pFormName As String) As Long
Declare Function GetForm Lib "winspool.drv" Alias "GetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Declare Function SetForm Lib "winspool.drv" Alias "SetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte) As Long
Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long

Declare Function EnumMonitors Lib "winspool.drv" Alias "EnumMonitorsA" (ByVal pName As String, ByVal Level As Long, pMonitors As Byte, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function AddMonitor Lib "winspool.drv" Alias "AddMonitorA" (ByVal pName As String, ByVal Level As Long, pMonitors As Byte) As Long
Declare Function DeleteMonitor Lib "winspool.drv" Alias "DeleteMonitorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pMonitorName As String) As Long

Declare Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, ByVal lpbPorts As Long, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function AddPort Lib "winspool.drv" Alias "AddPortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pMonitorName As String) As Long
Declare Function ConfigurePort Lib "winspool.drv" Alias "ConfigurePortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pPortName As String) As Long
Declare Function DeletePort Lib "winspool.drv" Alias "DeletePortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pPortName As String) As Long

Declare Function AddPrinterConnection Lib "winspool.drv" Alias "AddPrinterConnectionA" (ByVal pName As String) As Long

Declare Function DeletePrinterConnection Lib "winspool.drv" Alias "DeletePrinterConnectionA" (ByVal pName As String) As Long
Declare Function ConnectToPrinterDlg Lib "winspool.drv" Alias "ConnectToPrinterDlg" (ByVal hwnd As Long, ByVal flags As Long) As Long

Type PROVIDOR_INFO_1
        pName As String
        pEnvironment As String
        pDLLName As String
End Type

Declare Function AddPrintProvidor Lib "winspool.drv" Alias "AddPrintProvidorA" (ByVal pName As String, ByVal Level As Long, pProvidorInfo As Byte) As Long
Declare Function DeletePrintProvidor Lib "winspool.drv" Alias "DeletePrintProvidorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pPrintProvidorName As String) As Long

Private Const SERVER_ACCESS_ADMINISTER = &H1
Private Const SERVER_ACCESS_ENUMERATE = &H2

Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8

Private Const JOB_ACCESS_ADMINISTER = &H10

' Access rights for print servers

Private Const SERVER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVER_ACCESS_ADMINISTER Or SERVER_ACCESS_ENUMERATE)
Private Const SERVER_READ = (STANDARD_RIGHTS_READ Or SERVER_ACCESS_ENUMERATE)
Private Const SERVER_WRITE = (STANDARD_RIGHTS_WRITE Or SERVER_ACCESS_ADMINISTER Or SERVER_ACCESS_ENUMERATE)
Private Const SERVER_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or SERVER_ACCESS_ENUMERATE)

' Access rights for printers
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
Private Const PRINTER_READ = (STANDARD_RIGHTS_READ Or PRINTER_ACCESS_USE)
Private Const PRINTER_WRITE = (STANDARD_RIGHTS_WRITE Or PRINTER_ACCESS_USE)
Private Const PRINTER_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or PRINTER_ACCESS_USE)

' Access rights for jobs
Private Const JOB_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or JOB_ACCESS_ADMINISTER)
Private Const JOB_READ = (STANDARD_RIGHTS_READ Or JOB_ACCESS_ADMINISTER)
Private Const JOB_WRITE = (STANDARD_RIGHTS_WRITE Or JOB_ACCESS_ADMINISTER)
Private Const JOB_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or JOB_ACCESS_ADMINISTER)

'  Windows Network support

'  RESOURCE ENUMERATION

Private Const RESOURCE_CONNECTED = &H1
Private Const RESOURCE_PrivateNET = &H2
Private Const RESOURCE_REMEMBERED = &H3

Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const RESOURCETYPE_UNKNOWN = &HFFFF

Private Const RESOURCEUSAGE_CONNECTABLE = &H1
Private Const RESOURCEUSAGE_CONTAINER = &H2
Private Const RESOURCEUSAGE_RESERVED = &H80000000

Private Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Private Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Private Const RESOURCEDISPLAYTYPE_SERVER = &H2
Private Const RESOURCEDISPLAYTYPE_SHARE = &H3
Private Const RESOURCEDISPLAYTYPE_FILE = &H4
Private Const RESOURCEDISPLAYTYPE_GROUP = &H5

Type NETRESOURCE
        dwScope As Long
        dwType As Long
        dwDisplayType As Long
        dwUsage As Long
        lpLocalName As String
        lpRemoteName As String
        lpComment As String
        lpProvider As String
End Type

Private Const CONNECT_UPDATE_PROFILE = &H1

Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long
Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As NETRESOURCE, lphEnum As Long) As Long
Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As Any, lpBufferSize As Long) As Long
Declare Function WNetCloseEnum Lib "mpr.dll" Alias "WNetCloseEnum" (ByVal hEnum As Long) As Long

Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

Declare Function WNetConnectionDialog Lib "mpr.dll" Alias "WNetConnectionDialog" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Declare Function WNetDisconnectDialog Lib "mpr.dll" Alias "WNetDisconnectDialog" (ByVal hwnd As Long, ByVal dwType As Long) As Long

Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long

' Status Codes

' This section is provided for backward compatibility.  Use of the ERROR_
' codes is preferred.  The WN_ error codes may not be available in future
' releases.

' General

Private Const WN_SUCCESS = NO_ERROR
Private Const WN_NOT_SUPPORTED = ERROR_NOT_SUPPORTED
Private Const WN_NET_ERROR = ERROR_UNEXP_NET_ERR
Private Const WN_MORE_DATA = ERROR_MORE_DATA
Private Const WN_BAD_POINTER = ERROR_INVALID_ADDRESS
Private Const WN_BAD_VALUE = ERROR_INVALID_PARAMETER
Private Const WN_BAD_PASSWORD = ERROR_INVALID_PASSWORD
Private Const WN_ACCESS_DENIED = ERROR_ACCESS_DENIED
Private Const WN_FUNCTION_BUSY = ERROR_BUSY
Private Const WN_WINDOWS_ERROR = ERROR_UNEXP_NET_ERR
Private Const WN_BAD_USER = ERROR_BAD_USERNAME
Private Const WN_OUT_OF_MEMORY = ERROR_NOT_ENOUGH_MEMORY
Private Const WN_NO_NETWORK = ERROR_NO_NETWORK
Private Const WN_EXTENDED_ERROR = ERROR_EXTENDED_ERROR

' Connection

Private Const WN_NOT_CONNECTED = ERROR_NOT_CONNECTED
Private Const WN_OPEN_FILES = ERROR_OPEN_FILES
Private Const WN_DEVICE_IN_USE = ERROR_DEVICE_IN_USE
Private Const WN_BAD_NETNAME = ERROR_BAD_NET_NAME
Private Const WN_BAD_LOCALNAME = ERROR_BAD_DEVICE
Private Const WN_ALREADY_CONNECTED = ERROR_ALREADY_ASSIGNED
Private Const WN_DEVICE_ERROR = ERROR_GEN_FAILURE
Private Const WN_CONNECTION_CLOSED = ERROR_CONNECTION_UNAVAIL
Private Const WN_NO_NET_OR_BAD_PATH = ERROR_NO_NET_OR_BAD_PATH
Private Const WN_BAD_PROVIDER = ERROR_BAD_PROVIDER
Private Const WN_CANNOT_OPEN_PROFILE = ERROR_CANNOT_OPEN_PROFILE
Private Const WN_BAD_PROFILE = ERROR_BAD_PROFILE

' Enumeration

Private Const WN_BAD_HANDLE = ERROR_INVALID_HANDLE
Private Const WN_NO_MORE_ENTRIES = ERROR_NO_MORE_ITEMS
Private Const WN_NOT_CONTAINER = ERROR_NOT_CONTAINER

Private Const WN_NO_ERROR = NO_ERROR

' This section contains the definitions
' for portable NetBIOS 3.0 support.

Private Const NCBNAMSZ = 16  '  absolute length of a net name
Private Const MAX_LANA = 254  '  lana's in range 0 to MAX_LANA

Type NCB
        ncb_command As Integer
        ncb_retcode As Integer
        ncb_lsn As Integer
        ncb_num As Integer
        ncb_buffer As String
        ncb_length As Integer
        ncb_callname As String * NCBNAMSZ
        ncb_name As String * NCBNAMSZ
        ncb_rto As Integer
        ncb_sto As Integer
        ncb_post As Long
        ncb_lana_num As Integer
        ncb_cmd_cplt As Integer
        ncb_reserve(10) As Byte ' Reserved, must be 0
        ncb_event As Long
End Type

Type ADAPTER_STATUS
        adapter_address As String * 6
        rev_major As Integer
        reserved0 As Integer
        adapter_type As Integer
        rev_minor As Integer
        duration As Integer
        frmr_recv As Integer
        frmr_xmit As Integer
        iframe_recv_err As Integer
        xmit_aborts As Integer
        xmit_success As Long
        recv_success As Long
        iframe_xmit_err As Integer
        recv_buff_unavail As Integer
        t1_timeouts As Integer
        ti_timeouts As Integer
        Reserved1 As Long
        free_ncbs As Integer
        max_cfg_ncbs As Integer
        max_ncbs As Integer
        xmit_buf_unavail As Integer
        max_dgram_size As Integer
        pending_sess As Integer
        max_cfg_sess As Integer
        max_sess As Integer
        max_sess_pkt_size As Integer
        name_count As Integer
End Type

Type NAME_BUFFER
        name  As String * NCBNAMSZ
        name_num As Integer
        name_flags As Integer
End Type

' values for name_flags bits.

Private Const NAME_FLAGS_MASK = &H87

Private Const GROUP_NAME = &H80
Private Const UNIQUE_NAME = &H0

Private Const REGISTERING = &H0
Private Const REGISTERED = &H4
Private Const DEREGISTERED = &H5
Private Const DUPLICATE = &H6
Private Const DUPLICATE_DEREG = &H7

Type SESSION_HEADER
        sess_name As Integer
        num_sess As Integer
        rcv_dg_outstanding As Integer
        rcv_any_outstanding As Integer
End Type

Type SESSION_BUFFER
        lsn As Integer
        State As Integer
        local_name As String * NCBNAMSZ
        remote_name As String * NCBNAMSZ
        rcvs_outstanding As Integer
        sends_outstanding As Integer
End Type

' Values for state
Private Const LISTEN_OUTSTANDING = &H1
Private Const CALL_PENDING = &H2
Private Const SESSION_ESTABLISHED = &H3
Private Const HANGUP_PENDING = &H4
Private Const HANGUP_COMPLETE = &H5
Private Const SESSION_ABORTED = &H6

Type LANA_ENUM
        Length As Integer
        lana(MAX_LANA) As Integer
End Type

Type FIND_NAME_HEADER
        node_count As Integer
        Reserved As Integer
        unique_group As Integer
End Type

Type FIND_NAME_BUFFER
        Length As Integer
        access_control As Integer
        frame_control As Integer
        destination_addr(6) As Integer
        source_addr(6) As Integer
        routing_info(18) As Integer
End Type

Type ACTION_HEADER
        transport_id As Long
        action_code As Integer
        Reserved As Integer
End Type

' Values for transport_id
Private Const ALL_TRANSPORTS = "M\0\0\0"
Private Const MS_NBF = "MNBF"

' NCB Command codes
Private Const NCBCALL = &H10  '  NCB CALL
Private Const NCBLISTEN = &H11  '  NCB LISTEN
Private Const NCBHANGUP = &H12  '  NCB HANG UP
Private Const NCBSEND = &H14  '  NCB SEND
Private Const NCBRECV = &H15  '  NCB RECEIVE
Private Const NCBRECVANY = &H16  '  NCB RECEIVE ANY
Private Const NCBCHAINSEND = &H17  '  NCB CHAIN SEND
Private Const NCBDGSEND = &H20  '  NCB SEND DATAGRAM
Private Const NCBDGRECV = &H21  '  NCB RECEIVE DATAGRAM
Private Const NCBDGSENDBC = &H22  '  NCB SEND BROADCAST DATAGRAM
Private Const NCBDGRECVBC = &H23  '  NCB RECEIVE BROADCAST DATAGRAM
Private Const NCBADDNAME = &H30  '  NCB ADD NAME
Private Const NCBDELNAME = &H31  '  NCB DELETE NAME
Private Const NCBRESET = &H32  '  NCB RESET
Private Const NCBASTAT = &H33  '  NCB ADAPTER STATUS
Private Const NCBSSTAT = &H34  '  NCB SESSION STATUS
Private Const NCBCANCEL = &H35  '  NCB CANCEL
Private Const NCBADDGRNAME = &H36  '  NCB ADD GROUP NAME
Private Const NCBENUM = &H37  '  NCB ENUMERATE LANA NUMBERS
Private Const NCBUNLINK = &H70  '  NCB UNLINK
Private Const NCBSENDNA = &H71  '  NCB SEND NO ACK
Private Const NCBCHAINSENDNA = &H72  '  NCB CHAIN SEND NO ACK
Private Const NCBLANSTALERT = &H73  '  NCB LAN STATUS ALERT
Private Const NCBACTION = &H77  '  NCB ACTION
Private Const NCBFINDNAME = &H78  '  NCB FIND NAME
Private Const NCBTRACE = &H79  '  NCB TRACE

Private Const ASYNCH = &H80  '  high bit set == asynchronous

' NCB Return codes
Private Const NRC_GOODRET = &H0   '  good return
                                '  also returned when ASYNCH request accepted
Private Const NRC_BUFLEN = &H1   '  illegal buffer length
Private Const NRC_ILLCMD = &H3   '  illegal command
Private Const NRC_CMDTMO = &H5   '  command timed out
Private Const NRC_INCOMP = &H6   '  message incomplete, issue another command
Private Const NRC_BADDR = &H7   '  illegal buffer address
Private Const NRC_SNUMOUT = &H8   '  session number out of range
Private Const NRC_NORES = &H9   '  no resource available
Private Const NRC_SCLOSED = &HA   '  session closed
Private Const NRC_CMDCAN = &HB   '  command cancelled
Private Const NRC_DUPNAME = &HD   '  duplicate name
Private Const NRC_NAMTFUL = &HE   '  name table full
Private Const NRC_ACTSES = &HF   '  no deletions, name has active sessions
Private Const NRC_LOCTFUL = &H11  '  local session table full
Private Const NRC_REMTFUL = &H12  '  remote session table full
Private Const NRC_ILLNN = &H13  '  illegal name number
Private Const NRC_NOCALL = &H14  '  no callname
Private Const NRC_NOWILD = &H15  '  cannot put  in NCB_NAME
Private Const NRC_INUSE = &H16  '  name in use on remote adapter
Private Const NRC_NAMERR = &H17  '  name deleted
Private Const NRC_SABORT = &H18  '  session ended abnormally
Private Const NRC_NAMCONF = &H19  '  name conflict detected
Private Const NRC_IFBUSY = &H21  '  interface busy, IRET before retrying
Private Const NRC_TOOMANY = &H22  '  too many commands outstanding, retry later
Private Const NRC_BRIDGE = &H23  '  ncb_lana_num field invalid
Private Const NRC_CANOCCR = &H24  '  command completed while cancel occurring
Private Const NRC_CANCEL = &H26  '  command not valid to cancel
Private Const NRC_DUPENV = &H30  '  name defined by anther local process
Private Const NRC_ENVNOTDEF = &H34  '  environment undefined. RESET required
Private Const NRC_OSRESNOTAV = &H35  '  required OS resources exhausted
Private Const NRC_MAXAPPS = &H36  '  max number of applications exceeded
Private Const NRC_NOSAPS = &H37  '  no saps available for netbios
Private Const NRC_NORESOURCES = &H38  '  requested resources are not available
Private Const NRC_INVADDRESS = &H39  '  invalid ncb address or length > segment
Private Const NRC_INVDDID = &H3B  '  invalid NCB DDID
Private Const NRC_LOCKFAIL = &H3C  '  lock of user area failed
Private Const NRC_OPENERR = &H3F  '  NETBIOS not loaded
Private Const NRC_SYSTEM = &H40  '  system error

Private Const NRC_PENDING = &HFF  '  asynchronous command is not yet finished

Declare Function Netbios Lib "netapi32.dll" Alias "Netbios" (pncb As NCB) As Byte

' Legal values for expression in except().
Private Const EXCEPTION_EXECUTE_HANDLER = 1
Private Const EXCEPTION_CONTINUE_SEARCH = 0
Private Const EXCEPTION_CONTINUE_EXECUTION = -1

' UI dialog constants and types

' ----Constants--------------------------------------------------------------
Private Const ctlFirst = &H400
Private Const ctlLast = &H4FF
    '  Push buttons
Private Const psh1 = &H400
Private Const psh2 = &H401
Private Const psh3 = &H402
Private Const psh4 = &H403
Private Const psh5 = &H404
Private Const psh6 = &H405
Private Const psh7 = &H406
Private Const psh8 = &H407
Private Const psh9 = &H408
Private Const psh10 = &H409
Private Const psh11 = &H40A
Private Const psh12 = &H40B
Private Const psh13 = &H40C
Private Const psh14 = &H40D
Private Const psh15 = &H40E
Private Const pshHelp = psh15
Private Const psh16 = &H40F
    '  Checkboxes
Private Const chx1 = &H410
Private Const chx2 = &H411
Private Const chx3 = &H412
Private Const chx4 = &H413
Private Const chx5 = &H414
Private Const chx6 = &H415
Private Const chx7 = &H416
Private Const chx8 = &H417
Private Const chx9 = &H418
Private Const chx10 = &H419
Private Const chx11 = &H41A
Private Const chx12 = &H41B
Private Const chx13 = &H41C
Private Const chx14 = &H41D
Private Const chx15 = &H41E
Private Const chx16 = &H41D
    '  Radio buttons
Private Const rad1 = &H420
Private Const rad2 = &H421
Private Const rad3 = &H422
Private Const rad4 = &H423
Private Const rad5 = &H424
Private Const rad6 = &H425
Private Const rad7 = &H426
Private Const rad8 = &H427
Private Const rad9 = &H428
Private Const rad10 = &H429
Private Const rad11 = &H42A
Private Const rad12 = &H42B
Private Const rad13 = &H42C
Private Const rad14 = &H42D
Private Const rad15 = &H42E
Private Const rad16 = &H42F
    '  Groups, frames, rectangles, and icons
Private Const grp1 = &H430
Private Const grp2 = &H431
Private Const grp3 = &H432
Private Const grp4 = &H433
Private Const frm1 = &H434
Private Const frm2 = &H435
Private Const frm3 = &H436
Private Const frm4 = &H437
Private Const rct1 = &H438
Private Const rct2 = &H439
Private Const rct3 = &H43A
Private Const rct4 = &H43B
Private Const ico1 = &H43C
Private Const ico2 = &H43D
Private Const ico3 = &H43E
Private Const ico4 = &H43F
    '  Static text
Private Const stc1 = &H440
Private Const stc2 = &H441
Private Const stc3 = &H442
Private Const stc4 = &H443
Private Const stc5 = &H444
Private Const stc6 = &H445
Private Const stc7 = &H446
Private Const stc8 = &H447
Private Const stc9 = &H448
Private Const stc10 = &H449
Private Const stc11 = &H44A
Private Const stc12 = &H44B
Private Const stc13 = &H44C
Private Const stc14 = &H44D
Private Const stc15 = &H44E
Private Const stc16 = &H44F
Private Const stc17 = &H450
Private Const stc18 = &H451
Private Const stc19 = &H452
Private Const stc20 = &H453
Private Const stc21 = &H454
Private Const stc22 = &H455
Private Const stc23 = &H456
Private Const stc24 = &H457
Private Const stc25 = &H458
Private Const stc26 = &H459
Private Const stc27 = &H45A
Private Const stc28 = &H45B
Private Const stc29 = &H45C
Private Const stc30 = &H45D
Private Const stc31 = &H45E
Private Const stc32 = &H45F
    '  Listboxes
Private Const lst1 = &H460
Private Const lst2 = &H461
Private Const lst3 = &H462
Private Const lst4 = &H463
Private Const lst5 = &H464
Private Const lst6 = &H465
Private Const lst7 = &H466
Private Const lst8 = &H467
Private Const lst9 = &H468
Private Const lst10 = &H469
Private Const lst11 = &H46A
Private Const lst12 = &H46B
Private Const lst13 = &H46C
Private Const lst14 = &H46D
Private Const lst15 = &H46E
Private Const lst16 = &H46F
    '  Combo boxes
Private Const cmb1 = &H470
Private Const cmb2 = &H471
Private Const cmb3 = &H472
Private Const cmb4 = &H473
Private Const cmb5 = &H474
Private Const cmb6 = &H475
Private Const cmb7 = &H476
Private Const cmb8 = &H477
Private Const cmb9 = &H478
Private Const cmb10 = &H479
Private Const cmb11 = &H47A
Private Const cmb12 = &H47B
Private Const cmb13 = &H47C
Private Const cmb14 = &H47D
Private Const cmb15 = &H47E
Private Const cmb16 = &H47F
    '  Edit controls
Private Const edt1 = &H480
Private Const edt2 = &H481
Private Const edt3 = &H482
Private Const edt4 = &H483
Private Const edt5 = &H484
Private Const edt6 = &H485
Private Const edt7 = &H486
Private Const edt8 = &H487
Private Const edt9 = &H488
Private Const edt10 = &H489
Private Const edt11 = &H48A
Private Const edt12 = &H48B
Private Const edt13 = &H48C
Private Const edt14 = &H48D
Private Const edt15 = &H48E
Private Const edt16 = &H48F
    '  Scroll bars
Private Const scr1 = &H490
Private Const scr2 = &H491
Private Const scr3 = &H492
Private Const scr4 = &H493
Private Const scr5 = &H494
Private Const scr6 = &H495
Private Const scr7 = &H496
Private Const scr8 = &H497

Private Const FILEOPENORD = 1536
Private Const MULTIFILEOPENORD = 1537
Private Const PRINTDLGORD = 1538
Private Const PRNSETUPDLGORD = 1539
Private Const FINDDLGORD = 1540
Private Const REPLACEDLGORD = 1541
Private Const FONTDLGORD = 1542
Private Const FORMATDLGORD31 = 1543
Private Const FORMATDLGORD30 = 1544

Type CRGB
        bRed As Byte
        bGreen As Byte
        bBlue As Byte
        bExtra As Byte
End Type

' -----------------
' ADVAPI32
' -----------------

' function prototypes, constants, and type definitions
' for Windows 32-bit Registry API

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

' Registry API prototypes

Declare Function RegCloseKey Lib "advapi32.dll" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Declare Function RegFlushKey Lib "advapi32.dll" Alias "RegFlushKey" (ByVal hKey As Long) As Long
Declare Function RegGetKeySecurity Lib "advapi32.dll" Alias "RegGetKeySecurity" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long
Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" Alias "RegNotifyChangeKeyValue" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function RegSetKeySecurity Lib "advapi32.dll" Alias "RegSetKeySecurity" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long

' Service database names
Private Const SERVICES_ACTIVE_DATABASE = "ServicesActive"
Private Const SERVICES_FAILED_DATABASE = "ServicesFailed"

' Value to indicate no change to an optional parameter
Private Const SERVICE_NO_CHANGE = &HFFFF

' Service State -- for Enum Requests (Bit Mask)
Private Const SERVICE_ACTIVE = &H1
Private Const SERVICE_INACTIVE = &H2
Private Const SERVICE_STATE_ALL = (SERVICE_ACTIVE Or SERVICE_INACTIVE)

' Controls
Private Const SERVICE_CONTROL_STOP = &H1
Private Const SERVICE_CONTROL_PAUSE = &H2
Private Const SERVICE_CONTROL_CONTINUE = &H3
Private Const SERVICE_CONTROL_INTERROGATE = &H4
Private Const SERVICE_CONTROL_SHUTDOWN = &H5

' Service State -- for CurrentState
Private Const SERVICE_STOPPED = &H1
Private Const SERVICE_START_PENDING = &H2
Private Const SERVICE_STOP_PENDING = &H3
Private Const SERVICE_RUNNING = &H4
Private Const SERVICE_CONTINUE_PENDING = &H5
Private Const SERVICE_PAUSE_PENDING = &H6
Private Const SERVICE_PAUSED = &H7

' Controls Accepted  (Bit Mask)
Private Const SERVICE_ACCEPT_STOP = &H1
Private Const SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
Private Const SERVICE_ACCEPT_SHUTDOWN = &H4

' Service Control Manager object specific access types
Private Const SC_MANAGER_CONNECT = &H1
Private Const SC_MANAGER_CREATE_SERVICE = &H2
Private Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Private Const SC_MANAGER_LOCK = &H8
Private Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Private Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20

Private Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)

' Service object specific access type
Private Const SERVICE_QUERY_CONFIG = &H1
Private Const SERVICE_CHANGE_CONFIG = &H2
Private Const SERVICE_QUERY_STATUS = &H4
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Private Const SERVICE_START = &H10
Private Const SERVICE_STOP = &H20
Private Const SERVICE_PAUSE_CONTINUE = &H40
Private Const SERVICE_INTERROGATE = &H80
Private Const SERVICE_USER_DEFINED_CONTROL = &H100

Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)


Type SERVICE_STATUS
        dwServiceType As Long
        dwCurrentState As Long
        dwControlsAccepted As Long
        dwWin32ExitCode As Long
        dwServiceSpecificExitCode As Long
        dwCheckPoint As Long
        dwWaitHint As Long
End Type

Type ENUM_SERVICE_STATUS
        lpServiceName As String
        lpDisplayName As String
        ServiceStatus As SERVICE_STATUS
End Type

Type QUERY_SERVICE_LOCK_STATUS
        fIsLocked As Long
        lpLockOwner As String
        dwLockDuration As Long
End Type

Type QUERY_SERVICE_CONFIG
        dwServiceType As Long
        dwStartType As Long
        dwErrorControl As Long
        lpBinaryPathName As String
        lpLoadOrderGroup As String
        dwTagId As Long
        lpDependencies As String
        lpServiceStartName As String
        lpDisplayName As String
End Type

Type SERVICE_TABLE_ENTRY
        lpServiceName As String
        lpServiceProc As Long
End Type


' ++ BUILD Version: 0010    '  Increment this if a change has global effects
' Copyright (c) 1995  Microsoft Corporation
' Module Name:
'     winsvc.h
' Abstract:
'     Header file for the Service Control Manager
' Environment:
'     User Mode - Win32
' --*/
'
'  Constants

'  Character to designate that a name is a group
'
Private Const SC_GROUP_IDENTIFIER = "+"

'  Prototype for the Service Control Handler Function

' /////////////////////////////////////////////////////////////////////////
'  API Function Prototypes
' /////////////////////////////////////////////////////////////////////////

Declare Function ChangeServiceConfig Lib "advapi32.dll" Alias "ChangeServiceConfigA" (ByVal hService As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lpServiceStartName As String, ByVal lpPassword As String, ByVal lpDisplayName As String) As Long
Declare Function CloseServiceHandle Lib "advapi32.dll" Alias "CloseServiceHandle" (ByVal hSCObject As Long) As Long
Declare Function ControlService Lib "advapi32.dll" Alias "ControlService" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Declare Function CreateService Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
Declare Function DeleteService Lib "advapi32.dll" Alias "DeleteService" (ByVal hService As Long) As Long
Declare Function EnumDependentServices Lib "advapi32.dll" Alias "EnumDependentServicesA" (ByVal hService As Long, ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long) As Long
Declare Function EnumServicesStatus Lib "advapi32.dll" Alias "EnumServicesStatusA" (ByVal hSCManager As Long, ByVal dwServiceType As Long, ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long, lpResumeHandle As Long) As Long
Declare Function GetServiceKeyName Lib "advapi32.dll" Alias "GetServiceKeyNameA" (ByVal hSCManager As Long, ByVal lpDisplayName As String, ByVal lpServiceName As String, lpcchBuffer As Long) As Long
Declare Function GetServiceDisplayName Lib "advapi32.dll" Alias "GetServiceDisplayNameA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, lpcchBuffer As Long) As Long
Declare Function LockServiceDatabase Lib "advapi32.dll" Alias "LockServiceDatabase" (ByVal hSCManager As Long) As Long
Declare Function NotifyBootConfigStatus Lib "advapi32.dll" Alias "NotifyBootConfigStatus" (ByVal BootAcceptable As Long) As Long
Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Declare Function QueryServiceConfig Lib "advapi32.dll" Alias "QueryServiceConfigA" (ByVal hService As Long, lpServiceConfig As QUERY_SERVICE_CONFIG, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Declare Function QueryServiceLockStatus Lib "advapi32.dll" Alias "QueryServiceLockStatusA" (ByVal hSCManager As Long, lpLockStatus As QUERY_SERVICE_LOCK_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Declare Function QueryServiceObjectSecurity Lib "advapi32.dll" Alias "QueryServiceObjectSecurity" (ByVal hService As Long, ByVal dwSecurityInformation As Long, lpSecurityDescriptor As Any, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Declare Function QueryServiceStatus Lib "advapi32.dll" Alias "QueryServiceStatus" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Declare Function RegisterServiceCtrlHandler Lib "advapi32.dll" Alias "RegisterServiceCtrlHandlerA" (ByVal lpServiceName As String, ByVal lpHandlerProc As Long) As Long
Declare Function SetServiceObjectSecurity Lib "advapi32.dll" Alias "SetServiceObjectSecurity" (ByVal hService As Long, ByVal dwSecurityInformation As Long, lpSecurityDescriptor As Any) As Long
Declare Function SetServiceStatus Lib "advapi32.dll" Alias "SetServiceStatus" (ByVal hServiceStatus As Long, lpServiceStatus As SERVICE_STATUS) As Long
Declare Function StartServiceCtrlDispatcher Lib "advapi32.dll" Alias "StartServiceCtrlDispatcherA" (lpServiceStartTable As SERVICE_TABLE_ENTRY) As Long
Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Declare Function UnlockServiceDatabase Lib "advapi32.dll" Alias "UnlockServiceDatabase" (ScLock As Any) As Long

Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
    

' Section for Performance Monitor data

Private Const PERF_DATA_VERSION = 1
Private Const PERF_DATA_REVISION = 1

Type PERF_DATA_BLOCK
        Signature As String * 4
        LittleEndian As Long
        Version As Long
        Revision As Long
        TotalByteLength As Long
        HeaderLength As Long
        NumObjectTypes As Long
        DefaultObject As Long
        SystemTime As SystemTime
        PerfTime As LARGE_INTEGER
        PerfFreq As LARGE_INTEGER
        PerTime100nSec As LARGE_INTEGER
        SystemNameLength As Long
        SystemNameOffset As Long
End Type

Type PERF_OBJECT_TYPE
        TotalByteLength As Long
        DefinitionLength As Long
        HeaderLength As Long
        ObjectNameTitleIndex As Long
        ObjectNameTitle As String
        ObjectHelpTitleIndex As Long
        ObjectHelpTitle As String
        DetailLevel As Long
        NumCounters As Long
        DefaultCounter As Long
        NumInstances As Long
        CodePage As Long
        PerfTime As LARGE_INTEGER
        PerfFreq As LARGE_INTEGER
End Type

Private Const PERF_NO_INSTANCES = -1  '  no instances

' The counter type is the "or" of the following values as described below
'
' select one of the following to indicate the counter's data size
Private Const PERF_SIZE_DWORD = &H0
Private Const PERF_SIZE_LARGE = &H100
Private Const PERF_SIZE_ZERO = &H200       '  for Zero Length fields
Private Const PERF_SIZE_VARIABLE_LEN = &H300       '  length is in CounterLength field of Counter Definition struct

' select one of the following values to indicate the counter field usage
Private Const PERF_TYPE_NUMBER = &H0         '  a number (not a counter)
Private Const PERF_TYPE_COUNTER = &H400       '  an increasing numeric value
Private Const PERF_TYPE_TEXT = &H800       '  a text field
Private Const PERF_TYPE_ZERO = &HC00       '  displays a zero

' If the PERF_TYPE_NUMBER field was selected, then select one of the
' following to describe the Number
Private Const PERF_NUMBER_HEX = &H0         '  display as HEX value
Private Const PERF_NUMBER_DECIMAL = &H10000     '  display as a decimal integer
Private Const PERF_NUMBER_DEC_1000 = &H20000     '  display as a decimal/1000
'
' If the PERF_TYPE_COUNTER value was selected then select one of the
' following to indicate the type of counter
Private Const PERF_COUNTER_VALUE = &H0         '  display counter value
Private Const PERF_COUNTER_RATE = &H10000     '  divide ctr / delta time
Private Const PERF_COUNTER_FRACTION = &H20000     '  divide ctr / base
Private Const PERF_COUNTER_BASE = &H30000     '  base value used in fractions
Private Const PERF_COUNTER_ELAPSED = &H40000     '  subtract counter from current time
Private Const PERF_COUNTER_QUEUELEN = &H50000     '  Use Queuelen processing func.
Private Const PERF_COUNTER_HISTOGRAM = &H60000     '  Counter begins or ends a histogram

' If the PERF_TYPE_TEXT value was selected, then select one of the
' following to indicate the type of TEXT data.
Private Const PERF_TEXT_UNICODE = &H0         '  type of text in text field
Private Const PERF_TEXT_ASCII = &H10000     '  ASCII using the CodePage field

' Timer SubTypes
Private Const PERF_TIMER_TICK = &H0         '  use system perf. freq for base
Private Const PERF_TIMER_100NS = &H100000    '  use 100 NS timer time base units
Private Const PERF_OBJECT_TIMER = &H200000    '  use the object timer freq

' Any types that have calculations performed can use one or more of
' the following calculation modification flags listed here
Private Const PERF_DELTA_COUNTER = &H400000    '  compute difference first
Private Const PERF_DELTA_BASE = &H800000    '  compute base diff as well
Private Const PERF_INVERSE_COUNTER = &H1000000   '  show as 1.00-value (assumes:
Private Const PERF_MULTI_COUNTER = &H2000000   '  sum of multiple instances

' Select one of the following values to indicate the display suffix (if any)
Private Const PERF_DISPLAY_NO_SUFFIX = &H0         '  no suffix
Private Const PERF_DISPLAY_PER_SEC = &H10000000  '  "/sec"
Private Const PERF_DISPLAY_PERCENT = &H20000000  '  "%"
Private Const PERF_DISPLAY_SECONDS = &H30000000  '  "secs"
Private Const PERF_DISPLAY_NOSHOW = &H40000000  '  value is not displayed

' Predefined counter types

' 32-bit Counter.  Divide delta by delta time.  Display suffix: "/sec"
Private Const PERF_COUNTER_COUNTER = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_PER_SEC)

' 64-bit Timer.  Divide delta by delta time.  Display suffix: "%"
Private Const PERF_COUNTER_TIMER = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_PERCENT)

' Queue Length Space-Time Product. Divide delta by delta time. No Display Suffix.
Private Const PERF_COUNTER_QUEUELEN_TYPE = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_QUEUELEN Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_NO_SUFFIX)

' 64-bit Counter.  Divide delta by delta time. Display Suffix: "/sec"
Private Const PERF_COUNTER_BULK_COUNT = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_PER_SEC)

' Indicates the counter is not a  counter but rather Unicode text Display as text.
Private Const PERF_COUNTER_TEXT = (PERF_SIZE_VARIABLE_LEN Or PERF_TYPE_TEXT Or PERF_TEXT_UNICODE Or PERF_DISPLAY_NO_SUFFIX)

' Indicates the data is a counter  which should not be
' time averaged on display (such as an error counter on a serial line)
' Display as is.  No Display Suffix.
Private Const PERF_COUNTER_RAWCOUNT = (PERF_SIZE_DWORD Or PERF_TYPE_NUMBER Or PERF_NUMBER_DECIMAL Or PERF_DISPLAY_NO_SUFFIX)

' A count which is either 1 or 0 on each sampling interrupt (% busy)
' Divide delta by delta base. Display Suffix: "%"
Private Const PERF_SAMPLE_FRACTION = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_FRACTION Or PERF_DELTA_COUNTER Or PERF_DELTA_BASE Or PERF_DISPLAY_PERCENT)

' A count which is sampled on each sampling interrupt (queue length)
' Divide delta by delta time. No Display Suffix.
Private Const PERF_SAMPLE_COUNTER = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_NO_SUFFIX)

' A label: no data is associated with this counter (it has 0 length)
' Do not display.
Private Const PERF_COUNTER_NODATA = (PERF_SIZE_ZERO Or PERF_DISPLAY_NOSHOW)

' 64-bit Timer inverse (e.g., idle is measured, but display busy  As Integer)
' Display 100 - delta divided by delta time.  Display suffix: "%"
Private Const PERF_COUNTER_TIMER_INV = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_INVERSE_COUNTER Or PERF_DISPLAY_PERCENT)

' The divisor for a sample, used with the previous counter to form a
' sampled %.  You must check for >0 before dividing by this!  This
' counter will directly follow the  numerator counter.  It should not
' be displayed to the user.
Private Const PERF_SAMPLE_BASE = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_BASE Or PERF_DISPLAY_NOSHOW Or &H1)         '  for compatibility with pre-beta versions

' A timer which, when divided by an average base, produces a time
' in seconds which is the average time of some operation.  This
' timer times total operations, and  the base is the number of opera-
' tions.  Display Suffix: "sec"
Private Const PERF_AVERAGE_TIMER = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_FRACTION Or PERF_DISPLAY_SECONDS)

' Used as the denominator in the computation of time or count
' averages.  Must directly follow the numerator counter.  Not dis-
' played to the user.
Private Const PERF_AVERAGE_BASE = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_BASE Or PERF_DISPLAY_NOSHOW Or &H2)         '  for compatibility with pre-beta versions

' A bulk count which, when divided (typically) by the number of
' operations, gives (typically) the number of bytes per operation.
' No Display Suffix.
Private Const PERF_AVERAGE_BULK = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_FRACTION Or PERF_DISPLAY_NOSHOW)

' 64-bit Timer in 100 nsec units. Display delta divided by
' delta time.  Display suffix: "%"
Private Const PERF_100NSEC_TIMER = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_100NS Or PERF_DELTA_COUNTER Or PERF_DISPLAY_PERCENT)

' 64-bit Timer inverse (e.g., idle is measured, but display busy  As Integer)
' Display 100 - delta divided by delta time.  Display suffix: "%"
Private Const PERF_100NSEC_TIMER_INV = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_100NS Or PERF_DELTA_COUNTER Or PERF_INVERSE_COUNTER Or PERF_DISPLAY_PERCENT)

' 64-bit Timer.  Divide delta by delta time.  Display suffix: "%"
' Timer for multiple instances, so result can exceed 100%.
Private Const PERF_COUNTER_MULTI_TIMER = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_DELTA_COUNTER Or PERF_TIMER_TICK Or PERF_MULTI_COUNTER Or PERF_DISPLAY_PERCENT)

' 64-bit Timer inverse (e.g., idle is measured, but display busy  As Integer)
' Display 100  _MULTI_BASE - delta divided by delta time.
' Display suffix: "%" Timer for multiple instances, so result
' can exceed 100%.  Followed by a counter of type _MULTI_BASE.
Private Const PERF_COUNTER_MULTI_TIMER_INV = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_DELTA_COUNTER Or PERF_MULTI_COUNTER Or PERF_TIMER_TICK Or PERF_INVERSE_COUNTER Or PERF_DISPLAY_PERCENT)

' Number of instances to which the preceding _MULTI_..._INV counter
' applies.  Used as a factor to get the percentage.
Private Const PERF_COUNTER_MULTI_BASE = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_BASE Or PERF_MULTI_COUNTER Or PERF_DISPLAY_NOSHOW)

' 64-bit Timer in 100 nsec units. Display delta divided by delta time.
' Display suffix: "%" Timer for multiple instances, so result can exceed 100%.
Private Const PERF_100NSEC_MULTI_TIMER = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_DELTA_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_100NS Or PERF_MULTI_COUNTER Or PERF_DISPLAY_PERCENT)

' 64-bit Timer inverse (e.g., idle is measured, but display busy  As Integer)
' Display 100  _MULTI_BASE - delta divided by delta time.
' Display suffix: "%" Timer for multiple instances, so result
' can exceed 100%.  Followed by a counter of type _MULTI_BASE.
Private Const PERF_100NSEC_MULTI_TIMER_INV = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_DELTA_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_100NS Or PERF_MULTI_COUNTER Or PERF_INVERSE_COUNTER Or PERF_DISPLAY_PERCENT)

' Indicates the data is a fraction of the following counter  which
' should not be time averaged on display (such as free space over
' total space.) Display as is.  Display the quotient as "%".
Private Const PERF_RAW_FRACTION = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_FRACTION Or PERF_DISPLAY_PERCENT)

' Indicates the data is a base for the preceding counter which should
' not be time averaged on display (such as free space over total space.)
Private Const PERF_RAW_BASE = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_BASE Or PERF_DISPLAY_NOSHOW Or &H3)         '  for compatibility with pre-beta versions

' The data collected in this counter is actually the start time of the
' item being measured. For display, this data is subtracted from the
' sample time to yield the elapsed time as the difference between the two.
' In the definition below, the PerfTime field of the Object contains
' the sample time as indicated by the PERF_OBJECT_TIMER bit and the
' difference is scaled by the PerfFreq of the Object to convert the time
' units into seconds.
Private Const PERF_ELAPSED_TIME = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_ELAPSED Or PERF_OBJECT_TIMER Or PERF_DISPLAY_SECONDS)

' The following counter type can be used with the preceding types to
' define a range of values to be displayed in a histogram.
Private Const PERF_COUNTER_HISTOGRAM_TYPE = &H80000000  ' Counter begins or ends a histogram

' The following are used to determine the level of detail associated
' with the counter.  The user will be setting the level of detail
' that should be displayed at any given time.
Private Const PERF_DETAIL_NOVICE = 100 '  The uninformed can understand it
Private Const PERF_DETAIL_ADVANCED = 200 '  For the advanced user
Private Const PERF_DETAIL_EXPERT = 300 '  For the expert user
Private Const PERF_DETAIL_WIZARD = 400 '  For the system designer

Type PERF_COUNTER_DEFINITION
        ByteLength As Long
        CounterNameTitleIndex As Long
        CounterNameTitle As String
        CounterHelpTitleIndex As Long
        CounterHelpTitle As String
        DefaultScale As Long
        DetailLevel As Long
        CounterType As Long
        CounterSize As Long
        CounterOffset As Long
End Type

Private Const PERF_NO_UNIQUE_ID = -1

Type PERF_INSTANCE_DEFINITION
        ByteLength As Long
        ParentObjectTitleIndex As Long
        ParentObjectInstance As Long
        UniqueID As Long
        NameOffset As Long
        NameLength As Long
End Type

Type PERF_COUNTER_BLOCK
        ByteLength As Long
End Type

Private Const CDERR_DIALOGFAILURE = &HFFFF

Private Const CDERR_GENERALCODES = &H0
Private Const CDERR_STRUCTSIZE = &H1
Private Const CDERR_INITIALIZATION = &H2
Private Const CDERR_NOTEMPLATE = &H3
Private Const CDERR_NOHINSTANCE = &H4
Private Const CDERR_LOADSTRFAILURE = &H5
Private Const CDERR_FINDRESFAILURE = &H6
Private Const CDERR_LOADRESFAILURE = &H7
Private Const CDERR_LOCKRESFAILURE = &H8
Private Const CDERR_MEMALLOCFAILURE = &H9
Private Const CDERR_MEMLOCKFAILURE = &HA
Private Const CDERR_NOHOOK = &HB
Private Const CDERR_REGISTERMSGFAIL = &HC

Private Const PDERR_PRINTERCODES = &H1000
Private Const PDERR_SETUPFAILURE = &H1001
Private Const PDERR_PARSEFAILURE = &H1002
Private Const PDERR_RETDEFFAILURE = &H1003
Private Const PDERR_LOADDRVFAILURE = &H1004
Private Const PDERR_GETDEVMODEFAIL = &H1005
Private Const PDERR_INITFAILURE = &H1006
Private Const PDERR_NODEVICES = &H1007
Private Const PDERR_NODEFAULTPRN = &H1008
Private Const PDERR_DNDMMISMATCH = &H1009
Private Const PDERR_CREATEICFAILURE = &H100A
Private Const PDERR_PRINTERNOTFOUND = &H100B
Private Const PDERR_DEFAULTDIFFERENT = &H100C

Private Const CFERR_CHOOSEFONTCODES = &H2000
Private Const CFERR_NOFONTS = &H2001
Private Const CFERR_MAXLESSTHANMIN = &H2002

Private Const FNERR_FILENAMECODES = &H3000
Private Const FNERR_SUBCLASSFAILURE = &H3001
Private Const FNERR_INVALIDFILENAME = &H3002
Private Const FNERR_BUFFERTOOSMALL = &H3003

Private Const FRERR_FINDREPLACECODES = &H4000
Private Const FRERR_BUFFERLENGTHZERO = &H4001

Private Const CCERR_CHOOSECOLORCODES = &H5000


' Private interface to LZEXP?.LIB

'  LZEXPAND error return codes
Private Const LZERROR_BADINHANDLE = (-1)  '  invalid input handle
Private Const LZERROR_BADOUTHANDLE = (-2) '  invalid output handle
Private Const LZERROR_READ = (-3)         '  corrupt compressed file format
Private Const LZERROR_WRITE = (-4)        '  out of space for output file
Private Const LZERROR_PrivateLOC = (-5)    '  insufficient memory for LZFile struct
Private Const LZERROR_GLOBLOCK = (-6)     '  bad Global handle
Private Const LZERROR_BADVALUE = (-7)     '  input parameter out of range
Private Const LZERROR_UNKNOWNALG = (-8)   '  compression algorithm not recognized

Declare Function LZCopy Lib "lz32.dll" Alias "LZCopy" (ByVal hfSource As Long, ByVal hfDest As Long) As Long
Declare Function LZInit Lib "lz32.dll" Alias "LZInit" (ByVal hfSrc As Long) As Long
Declare Function GetExpandedName Lib "lz32.dll" Alias "GetExpandedNameA" (ByVal lpszSource As String, ByVal lpszBuffer As String) As Long
Declare Function LZOpenFile Lib "lz32.dll" Alias "LZOpenFileA" (ByVal lpszFile As String, lpOf As OFSTRUCT, ByVal style As Long) As Long
Declare Function LZSeek Lib "lz32.dll" Alias "LZSeek" (ByVal hfFile As Long, ByVal lOffset As Long, ByVal nOrigin As Long) As Long
Declare Function LZRead Lib "lz32.dll" Alias "LZRead" (ByVal hfFile As Long, ByVal lpvBuf As String, ByVal cbread As Long) As Long
Declare Sub LZClose Lib "lz32.dll" Alias "LZClose" (ByVal hfFile As Long)

' ********************************************************************
'       IMM.H - Input Method Manager definitions
'
'       Copyright (c) 1993-1995  Microsoft Corporation
' ********************************************************************

Private Const VK_PROCESSKEY = &HE5

Type COMPOSITIONFORM
        dwStyle As Long
        ptCurrentPos As POINTAPI
        rcArea As RECT
End Type

Type CANDIDATEFORM
        dwIndex As Long
        dwStyle As Long
        ptCurrentPos As POINTAPI
        rcArea As RECT
End Type

Type CANDIDATELIST
        dwSize As Long
        dwStyle As Long
        dwCount As Long
        dwSelection As Long
        dwPageStart As Long
        dwPageSize As Long
        dwOffset(1) As Long
End Type

Private Const STYLE_DESCRIPTION_SIZE = 32

Type STYLEBUF
        dwStyle As Long
        szDescription As String * STYLE_DESCRIPTION_SIZE
End Type
'  prototype of IMM API

Declare Function ImmInstallIME Lib "imm32.dll" Alias "ImmInstallIMEA" (ByVal lpszIMEFileName As String, ByVal lpszLayoutText As String) As Long
Declare Function ImmGetDefaultIMEWnd Lib "imm32.dll" Alias "ImmGetDefaultIMEWnd" (ByVal hwnd As Long) As Long
Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Declare Function ImmGetIMEFileName Lib "imm32.dll" Alias "ImmGetIMEFileNameA" (ByVal hkl As Long, ByVal lpStr As String, ByVal uBufLen As Long) As Long
Declare Function ImmGetProperty Lib "imm32.dll" Alias "ImmGetProperty" (ByVal hkl As Long, ByVal dw As Long) As Long
Declare Function ImmIsIME Lib "imm32.dll" Alias "ImmIsIME" (ByVal hkl As Long) As Long
Declare Function ImmSimulateHotKey Lib "imm32.dll" Alias "ImmSimulateHotKey" (ByVal hwnd As Long, ByVal dw As Long) As Long
Declare Function ImmCreateContext Lib "imm32.dll" Alias "ImmCreateContext" () As Long
Declare Function ImmDestroyContext Lib "imm32.dll" Alias "ImmDestroyContext" (ByVal himc As Long) As Long
Declare Function ImmGetContext Lib "imm32.dll" Alias "ImmGetContext" (ByVal hwnd As Long) As Long
Declare Function ImmReleaseContext Lib "imm32.dll" Alias "ImmReleaseContext" (ByVal hwnd As Long, ByVal himc As Long) As Long
Declare Function ImmAssociateContext Lib "imm32.dll" Alias "ImmAssociateContext" (ByVal hwnd As Long, ByVal himc As Long) As Long
Declare Function ImmGetCompositionString Lib "imm32.dll" Alias "ImmGetCompositionStringA" (ByVal himc As Long, ByVal dw As Long, lpv As Any, ByVal dw2 As Long) As Long
Declare Function ImmSetCompositionString Lib "imm32.dll" Alias "ImmSetCompositionStringA" (ByVal himc As Long, ByVal dwIndex As Long, lpComp As Any, ByVal dw As Long, lpRead As Any, ByVal dw2 As Long) As Long
Declare Function ImmGetCandidateListCount Lib "imm32.dll" Alias "ImmGetCandidateListCountA" (ByVal himc As Long, lpdwListCount As Long) As Long
Declare Function ImmGetCandidateList Lib "imm32.dll" Alias "ImmGetCandidateListA" (ByVal himc As Long, ByVal deIndex As Long, lpCandidateList As CANDIDATELIST, ByVal dwBufLen As Long) As Long
Declare Function ImmGetGuideLine Lib "imm32.dll" Alias " ImmGetGuideLineA" (ByVal himc As Long, ByVal dwIndex As Long, ByVal lpStr As String, ByVal dwBufLen As Long) As Long
Declare Function ImmGetConversionStatus Lib "imm32.dll" Alias "ImmGetConversionStatus" (ByVal himc As Long, lpdw As Long, lpdw2 As Long) As Long
Declare Function ImmSetConversionStatus Lib "imm32.dll" Alias "ImmSetConversionStatus" (ByVal himc As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Declare Function ImmGetOpenStatus Lib "imm32.dll" Alias "ImmGetOpenStatus" (ByVal himc As Long) As Long
Declare Function ImmSetOpenStatus Lib "imm32.dll" Alias "ImmSetOpenStatus" (ByVal himc As Long, ByVal b As Long) As Long
Declare Function ImmGetCompositionFont Lib "imm32.dll" Alias "ImmGetCompositionFontA" (ByVal himc As Long, lpLogFont As LOGFONT) As Long
Declare Function ImmSetCompositionFont Lib "imm32.dll" Alias "ImmSetCompositionFontA" (ByVal himc As Long, lpLogFont As LOGFONT) As Long
Declare Function ImmConfigureIME Lib "imm32.dll" Alias "ImmConfigureIME" (ByVal hkl As Long, ByVal hwnd As Long, ByVal dw As Long) As Long
Declare Function ImmEscape Lib "imm32.dll" Alias "ImmEscapeA" (ByVal hkl As Long, ByVal himc As Long, ByVal un As Long, lpv As Any) As Long
Declare Function ImmGetConversionList Lib "imm32.dll" Alias "ImmGetConversionListA" (ByVal hkl As Long, ByVal himc As Long, ByVal lpsz As String, lpCandidateList As CANDIDATELIST, ByVal dwBufLen As Long, ByVal uFlag As Long) As Long
Declare Function ImmNotifyIME Lib "imm32.dll" Alias "ImmNotifyIME" (ByVal himc As Long, ByVal dwAction As Long, ByVal dwIndex As Long, ByVal dwValue As Long) As Long
Declare Function ImmGetStatusWindowPos Lib "imm32.dll" Alias "ImmGetStatusWindowPos" (ByVal himc As Long, lpPoint As POINTAPI) As Long
Declare Function ImmSetStatusWindowPos Lib "imm32.dll" Alias "ImmSetStatusWindowPos" (ByVal himc As Long, lpPoint As POINTAPI) As Long
Declare Function ImmGetCompositionWindow Lib "imm32.dll" Alias "ImmGetCompositionWindow" (ByVal himc As Long, lpCompositionForm As COMPOSITIONFORM) As Long
Declare Function ImmSetCompositionWindow Lib "imm32.dll" Alias "ImmSetCompositionWindow" (ByVal himc As Long, lpCompositionForm As COMPOSITIONFORM) As Long
Declare Function ImmGetCandidateWindow Lib "imm32.dll" Alias "ImmGetCandidateWindow" (ByVal himc As Long, ByVal dw As Long, lpCandidateForm As CANDIDATEFORM) As Long
Declare Function ImmSetCandidateWindow Lib "imm32.dll" Alias "ImmSetCandidateWindow" (ByVal himc As Long, lpCandidateForm As CANDIDATEFORM) As Long
Declare Function ImmIsUIMessage Lib "imm32.dll" Alias "ImmIsUIMessageA" (ByVal hwnd As Long, ByVal un As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function ImmGetVirtualKey Lib "imm32.dll" Alias "ImmGetVirtualKey" (ByVal hwnd As Long) As Long
Declare Function ImmRegisterWord Lib "imm32.dll" Alias "ImmRegisterWordA" (ByVal hkl As Long, ByVal lpszReading As String, ByVal dw As Long, ByVal lpszRegister As String) As Long
Declare Function ImmUnregisterWord Lib "imm32.dll" Alias "ImmUnregisterWordA" (ByVal hkl As Long, ByVal lpszReading As String, ByVal dw As Long, ByVal lpszUnregister As String) As Long
Declare Function ImmGetRegisterWordStyle Lib "imm32.dll" Alias " ImmGetRegisterWordStyleA" (ByVal hkl As Long, ByVal nItem As Long, lpStyleBuf As STYLEBUF) As Long
Declare Function ImmEnumRegisterWord Lib "imm32.dll" Alias "ImmEnumRegisterWordA" (ByVal hkl As Long, ByVal RegisterWordEnumProc As Long, ByVal lpszReading As String, ByVal dw As Long, ByVal lpszRegister As String, lpv As Any) As Long
'  the IME related messages
Private Const WM_CONVERTREQUESTEX = &H108
Private Const WM_IME_STARTCOMPOSITION = &H10D
Private Const WM_IME_ENDCOMPOSITION = &H10E
Private Const WM_IME_COMPOSITION = &H10F
Private Const WM_IME_KEYLAST = &H10F

Private Const WM_IME_SETCONTEXT = &H281
Private Const WM_IME_NOTIFY = &H282
Private Const WM_IME_CONTROL = &H283
Private Const WM_IME_COMPOSITIONFULL = &H284
Private Const WM_IME_SELECT = &H285
Private Const WM_IME_CHAR = &H286

Private Const WM_IME_KEYDOWN = &H290
Private Const WM_IME_KEYUP = &H291

'  wParam for WM_IME_CONTROL
Private Const IMC_GETCANDIDATEPOS = &H7
Private Const IMC_SETCANDIDATEPOS = &H8
Private Const IMC_GETCOMPOSITIONFONT = &H9
Private Const IMC_SETCOMPOSITIONFONT = &HA
Private Const IMC_GETCOMPOSITIONWINDOW = &HB
Private Const IMC_SETCOMPOSITIONWINDOW = &HC
Private Const IMC_GETSTATUSWINDOWPOS = &HF
Private Const IMC_SETSTATUSWINDOWPOS = &H10
Private Const IMC_CLOSESTATUSWINDOW = &H21
Private Const IMC_OPENSTATUSWINDOW = &H22


'  wParam for WM_IME_CONTROL to the soft keyboard
'  dwAction for ImmNotifyIME
Private Const NI_OPENCANDIDATE = &H10
Private Const NI_CLOSECANDIDATE = &H11
Private Const NI_SELECTCANDIDATESTR = &H12
Private Const NI_CHANGECANDIDATELIST = &H13
Private Const NI_FINALIZECONVERSIONRESULT = &H14
Private Const NI_COMPOSITIONSTR = &H15
Private Const NI_SETCANDIDATE_PAGESTART = &H16
Private Const NI_SETCANDIDATE_PAGESIZE = &H17

'  lParam for WM_IME_SETCONTEXT
Private Const ISC_SHOWUICANDIDATEWINDOW = &H1
Private Const ISC_SHOWUICOMPOSITIONWINDOW = &H80000000
Private Const ISC_SHOWUIGUIDELINE = &H40000000
Private Const ISC_SHOWUIALLCANDIDATEWINDOW = &HF
Private Const ISC_SHOWUIALL = &HC000000F

'  dwIndex for ImmNotifyIME/NI_COMPOSITIONSTR
Private Const CPS_COMPLETE = &H1
Private Const CPS_CONVERT = &H2
Private Const CPS_REVERT = &H3
Private Const CPS_CANCEL = &H4

'  Windows for Simplified Chinese Edition hot key ID from 0x10 - 0x2F
Private Const IME_CHOTKEY_IME_NONIME_TOGGLE = &H10
Private Const IME_CHOTKEY_SHAPE_TOGGLE = &H11
Private Const IME_CHOTKEY_SYMBOL_TOGGLE = &H12

'  Windows for Japanese Edition hot key ID from 0x30 - 0x4F
Private Const IME_JHOTKEY_CLOSE_OPEN = &H30

'  Windows for Korean Edition hot key ID from 0x50 - 0x6F
Private Const IME_KHOTKEY_SHAPE_TOGGLE = &H50
Private Const IME_KHOTKEY_HANJACONVERT = &H51
Private Const IME_KHOTKEY_ENGLISH = &H52

'  Windows for Tranditional Chinese Edition hot key ID from 0x70 - 0x8F
Private Const IME_THOTKEY_IME_NONIME_TOGGLE = &H70
Private Const IME_THOTKEY_SHAPE_TOGGLE = &H71
Private Const IME_THOTKEY_SYMBOL_TOGGLE = &H72

'  direct switch hot key ID from 0x100 - 0x11F
Private Const IME_HOTKEY_DSWITCH_FIRST = &H100
Private Const IME_HOTKEY_DSWITCH_LAST = &H11F

'  IME private hot key from 0x200 - 0x21F
Private Const IME_ITHOTKEY_RESEND_RESULTSTR = &H200
Private Const IME_ITHOTKEY_PREVIOUS_COMPOSITION = &H201
Private Const IME_ITHOTKEY_UISTYLE_TOGGLE = &H202

'  parameter of ImmGetCompositionString
Private Const GCS_COMPREADSTR = &H1
Private Const GCS_COMPREADATTR = &H2
Private Const GCS_COMPREADCLAUSE = &H4
Private Const GCS_COMPSTR = &H8
Private Const GCS_COMPATTR = &H10
Private Const GCS_COMPCLAUSE = &H20
Private Const GCS_CURSORPOS = &H80
Private Const GCS_DELTASTART = &H100
Private Const GCS_RESULTREADSTR = &H200
Private Const GCS_RESULTREADCLAUSE = &H400
Private Const GCS_RESULTSTR = &H800
Private Const GCS_RESULTCLAUSE = &H1000

'  style bit flags for WM_IME_COMPOSITION
Private Const CS_INSERTCHAR = &H2000
Private Const CS_NOMOVECARET = &H4000

'  bits of fdwInit of INPUTCONTEXT
'  IME property bits
Private Const IME_PROP_AT_CARET = &H10000
Private Const IME_PROP_SPECIAL_UI = &H20000
Private Const IME_PROP_CANDLIST_START_FROM_1 = &H40000
Private Const IME_PROP_UNICODE = &H80000

'  IME UICapability bits
Private Const UI_CAP_2700 = &H1
Private Const UI_CAP_ROT90 = &H2
Private Const UI_CAP_ROTANY = &H4

'  ImmSetCompositionString Capability bits
Private Const SCS_CAP_COMPSTR = &H1
Private Const SCS_CAP_MAKEREAD = &H2

'  IME WM_IME_SELECT inheritance Capability bits
Private Const SELECT_CAP_CONVERSION = &H1
Private Const SELECT_CAP_SENTENCE = &H2

'  ID for deIndex of ImmGetGuideLine
Private Const GGL_LEVEL = &H1
Private Const GGL_INDEX = &H2
Private Const GGL_STRING = &H3
Private Const GGL_PRIVATE = &H4

'  ID for dwLevel of GUIDELINE Structure
Private Const GL_LEVEL_NOGUIDELINE = &H0
Private Const GL_LEVEL_FATAL = &H1
Private Const GL_LEVEL_ERROR = &H2
Private Const GL_LEVEL_WARNING = &H3
Private Const GL_LEVEL_INFORMATION = &H4

'  ID for dwIndex of GUIDELINE Structure
Private Const GL_ID_UNKNOWN = &H0
Private Const GL_ID_NOMODULE = &H1
Private Const GL_ID_NODICTIONARY = &H10
Private Const GL_ID_CANNOTSAVE = &H11
Private Const GL_ID_NOCONVERT = &H20
Private Const GL_ID_TYPINGERROR = &H21
Private Const GL_ID_TOOMANYSTROKE = &H22
Private Const GL_ID_READINGCONFLICT = &H23
Private Const GL_ID_INPUTREADING = &H24
Private Const GL_ID_INPUTRADICAL = &H25
Private Const GL_ID_INPUTCODE = &H26
Private Const GL_ID_INPUTSYMBOL = &H27
Private Const GL_ID_CHOOSECANDIDATE = &H28
Private Const GL_ID_REVERSECONVERSION = &H29
Private Const GL_ID_PRIVATE_FIRST = &H8000
Private Const GL_ID_PRIVATE_LAST = &HFFFF

'  ID for dwIndex of ImmGetProperty
Private Const IGP_PROPERTY = &H4
Private Const IGP_CONVERSION = &H8
Private Const IGP_SENTENCE = &HC
Private Const IGP_UI = &H10
Private Const IGP_SETCOMPSTR = &H14
Private Const IGP_SELECT = &H18

'  dwIndex for ImmSetCompositionString API
Private Const SCS_SETSTR = (GCS_COMPREADSTR Or GCS_COMPSTR)
Private Const SCS_CHANGEATTR = (GCS_COMPREADATTR Or GCS_COMPATTR)
Private Const SCS_CHANGECLAUSE = (GCS_COMPREADCLAUSE Or GCS_COMPCLAUSE)

'  attribute for COMPOSITIONSTRING Structure
Private Const ATTR_INPUT = &H0
Private Const ATTR_TARGET_CONVERTED = &H1
Private Const ATTR_CONVERTED = &H2
Private Const ATTR_TARGET_NOTCONVERTED = &H3
Private Const ATTR_INPUT_ERROR = &H4

'  bit field for IMC_SETCOMPOSITIONWINDOW, IMC_SETCANDIDATEWINDOW
Private Const CFS_DEFAULT = &H0
Private Const CFS_RECT = &H1
Private Const CFS_POINT = &H2
Private Const CFS_SCREEN = &H4
Private Const CFS_FORCE_POSITION = &H20
Private Const CFS_CANDIDATEPOS = &H40
Private Const CFS_EXCLUDE = &H80

'  conversion direction for ImmGetConversionList
Private Const GCL_CONVERSION = &H1
Private Const GCL_REVERSECONVERSION = &H2
Private Const GCL_REVERSE_LENGTH = &H3

'  bit field for conversion mode
Private Const IME_CMODE_ALPHANUMERIC = &H0
Private Const IME_CMODE_NATIVE = &H1
Private Const IME_CMODE_CHINESE = IME_CMODE_NATIVE
Private Const IME_CMODE_HANGEUL = IME_CMODE_NATIVE
Private Const IME_CMODE_JAPANESE = IME_CMODE_NATIVE
Private Const IME_CMODE_KATAKANA = &H2                   '  only effect under IME_CMODE_NATIVE
Private Const IME_CMODE_LANGUAGE = &H3
Private Const IME_CMODE_FULLSHAPE = &H8
Private Const IME_CMODE_ROMAN = &H10
Private Const IME_CMODE_CHARCODE = &H20
Private Const IME_CMODE_HANJACONVERT = &H40
Private Const IME_CMODE_SOFTKBD = &H80
Private Const IME_CMODE_NOCONVERSION = &H100
Private Const IME_CMODE_EUDC = &H200
Private Const IME_CMODE_SYMBOL = &H400

Private Const IME_SMODE_NONE = &H0
Private Const IME_SMODE_PLAURALCLAUSE = &H1
Private Const IME_SMODE_SINGLECONVERT = &H2
Private Const IME_SMODE_AUTOMATIC = &H4
Private Const IME_SMODE_PHRASEPREDICT = &H8

'  style of candidate
Private Const IME_CAND_UNKNOWN = &H0
Private Const IME_CAND_READ = &H1
Private Const IME_CAND_CODE = &H2
Private Const IME_CAND_MEANING = &H3
Private Const IME_CAND_RADICAL = &H4
Private Const IME_CAND_STROKE = &H5

'  wParam of report message WM_IME_NOTIFY
Private Const IMN_CLOSESTATUSWINDOW = &H1
Private Const IMN_OPENSTATUSWINDOW = &H2
Private Const IMN_CHANGECANDIDATE = &H3
Private Const IMN_CLOSECANDIDATE = &H4
Private Const IMN_OPENCANDIDATE = &H5
Private Const IMN_SETCONVERSIONMODE = &H6
Private Const IMN_SETSENTENCEMODE = &H7
Private Const IMN_SETOPENSTATUS = &H8
Private Const IMN_SETCANDIDATEPOS = &H9
Private Const IMN_SETCOMPOSITIONFONT = &HA
Private Const IMN_SETCOMPOSITIONWINDOW = &HB
Private Const IMN_SETSTATUSWINDOWPOS = &HC
Private Const IMN_GUIDELINE = &HD
Private Const IMN_PRIVATE = &HE

'  error code of ImmGetCompositionString
Private Const IMM_ERROR_NODATA = (-1)
Private Const IMM_ERROR_GENERAL = (-2)

'  dialog mode of ImmConfigureIME
Private Const IME_CONFIG_GENERAL = 1
Private Const IME_CONFIG_REGISTERWORD = 2
Private Const IME_CONFIG_SELECTDICTIONARY = 3

'  dialog mode of ImmEscape
Private Const IME_ESC_QUERY_SUPPORT = &H3
Private Const IME_ESC_RESERVED_FIRST = &H4
Private Const IME_ESC_RESERVED_LAST = &H7FF
Private Const IME_ESC_PRIVATE_FIRST = &H800
Private Const IME_ESC_PRIVATE_LAST = &HFFF
Private Const IME_ESC_SEQUENCE_TO_INTERNAL = &H1001
Private Const IME_ESC_GET_EUDC_DICTIONARY = &H1003
Private Const IME_ESC_SET_EUDC_DICTIONARY = &H1004
Private Const IME_ESC_MAX_KEY = &H1005
Private Const IME_ESC_IME_NAME = &H1006
Private Const IME_ESC_SYNC_HOTKEY = &H1007
Private Const IME_ESC_HANJA_MODE = &H1008

'  style of word registration
Private Const IME_REGWORD_STYLE_EUDC = &H1
Private Const IME_REGWORD_STYLE_USER_FIRST = &H80000000
Private Const IME_REGWORD_STYLE_USER_LAST = &HFFFF

'  type of soft keyboard
'  for Windows Tranditional Chinese Edition
Private Const SOFTKEYBOARD_TYPE_T1 = &H1
'  for Windows Simplified Chinese Edition
Private Const SOFTKEYBOARD_TYPE_C1 = &H2



' ***********************************************************************
' *                                                                       *
' *   mcx.h -- This module defines the 32-Bit Windows MCX APIs            *
' *                                                                       *
' *   Copyright (c) 1990-1995, Microsoft Corp. All rights reserved.       *
' *                                                                       *
' ************************************************************************/


Type MODEMDEVCAPS
        dwActualSize As Long
        dwRequiredSize As Long
        dwDevSpecificOffset As Long
        dwDevSpecificSize As Long

    '  product and version identification
        dwModemProviderVersion As Long
        dwModemManufacturerOffset As Long
        dwModemManufacturerSize As Long
        dwModemModelOffset As Long
        dwModemModelSize As Long
        dwModemVersionOffset As Long
        dwModemVersionSize As Long

    '  local option capabilities
        dwDialOptions As Long          '  bitmap of supported values
        dwCallSetupFailTimer As Long   '  maximum in seconds
        dwInactivityTimeout As Long    '  maximum in seconds
        dwSpeakerVolume As Long        '  bitmap of supported values
        dwSpeakerMode As Long          '  bitmap of supported values
        dwModemOptions As Long         '  bitmap of supported values
        dwMaxDTERate As Long           '  maximum value in bit/s
        dwMaxDCERate As Long           '  maximum value in bit/s

    '  Variable portion for proprietary expansion
        abVariablePortion(1) As Byte
End Type

Type MODEMSETTINGS
        dwActualSize As Long
        dwRequiredSize As Long
        dwDevSpecificOffset As Long
        dwDevSpecificSize As Long

    '  static local options (read/write)
        dwCallSetupFailTimer As Long       '  seconds
        dwInactivityTimeout As Long        '  seconds
        dwSpeakerVolume As Long            '  level
        dwSpeakerMode As Long              '  mode
        dwPreferredModemOptions As Long    '  bitmap
    
    '  negotiated options (read only) for current or last call
        dwNegotiatedModemOptions As Long   '  bitmap
        dwNegotiatedDCERate As Long        '  bit/s

    '  Variable portion for proprietary expansion
        abVariablePortion(1) As Byte
End Type

'  Dial Options
Private Const DIALOPTION_BILLING = &H40          '  Supports wait for bong "$"
Private Const DIALOPTION_QUIET = &H80            '  Supports wait for quiet "@"
Private Const DIALOPTION_DIALTONE = &H100        '  Supports wait for dial tone "W"

'  SpeakerVolume for MODEMDEVCAPS
Private Const MDMVOLFLAG_LOW = &H1
Private Const MDMVOLFLAG_MEDIUM = &H2
Private Const MDMVOLFLAG_HIGH = &H4

'  SpeakerVolume for MODEMSETTINGS
Private Const MDMVOL_LOW = &H0
Private Const MDMVOL_MEDIUM = &H1
Private Const MDMVOL_HIGH = &H2

'  SpeakerMode for MODEMDEVCAPS
Private Const MDMSPKRFLAG_OFF = &H1
Private Const MDMSPKRFLAG_DIAL = &H2
Private Const MDMSPKRFLAG_ON = &H4
Private Const MDMSPKRFLAG_CALLSETUP = &H8

'  SpeakerMode for MODEMSETTINGS
Private Const MDMSPKR_OFF = &H0
Private Const MDMSPKR_DIAL = &H1
Private Const MDMSPKR_ON = &H2
Private Const MDMSPKR_CALLSETUP = &H3
 
'  Modem Options
Private Const MDM_COMPRESSION = &H1
Private Const MDM_ERROR_CONTROL = &H2
Private Const MDM_FORCED_EC = &H4
Private Const MDM_CELLULAR = &H8
Private Const MDM_FLOWCONTROL_HARD = &H10
Private Const MDM_FLOWCONTROL_SOFT = &H20
Private Const MDM_CCITT_OVERRIDE = &H40
Private Const MDM_SPEED_ADJUST = &H80
Private Const MDM_TONE_DIAL = &H100
Private Const MDM_BLIND_DIAL = &H200
Private Const MDM_V23_OVERRIDE = &H400



' *****************************************************************************                                                                             *
' * shellapi.h -  SHELL.DLL functions, types, and definitions                   *
' *                                                                             *
' * Copyright (c) 1992-1995, Microsoft Corp.  All rights reserved               *
' *                                                                             *
' \*****************************************************************************/

Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Declare Function DragQueryPoint Lib "shell32.dll" Alias "DragQueryPoint" (ByVal HDROP As Long, lpPoint As POINTAPI) As Long
Declare Sub DragFinish Lib "shell32.dll" Alias "DragFinish" (ByVal hDrop As Long)
Declare Sub DragAcceptFiles Lib "shell32.dll" Alias "DragAcceptFiles" (ByVal hwnd As Long, ByVal fAccept As Long)
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function DuplicateIcon Lib "shell32.dll" Alias "DuplicateIcon" (ByVal hInst As Long, ByVal hIcon As Long) As Long
Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociateIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Type DRAGINFO
        uSize As Long                 '  init with sizeof(DRAGINFO)
        pt As POINTAPI
        fNC As Long
        lpFileList As String
        grfKeyState As Long
End Type

' // AppBar stuff

Private Const ABM_NEW = &H0
Private Const ABM_REMOVE = &H1
Private Const ABM_QUERYPOS = &H2
Private Const ABM_SETPOS = &H3
Private Const ABM_GETSTATE = &H4
Private Const ABM_GETTASKBARPOS = &H5
Private Const ABM_ACTIVATE = &H6               '  lParam == TRUE/FALSE means activate/deactivate
Private Const ABM_GETAUTOHIDEBAR = &H7
Private Const ABM_SETAUTOHIDEBAR = &H8          '  this can fail at any time.  MUST check the result
                                        '  lParam = TRUE/FALSE  Set/Unset
                                        '  uEdge = what edge
Private Const ABM_WINDOWPOSCHANGED = &H9


'  these are put in the wparam of callback messages
Private Const ABN_STATECHANGE = &H0
Private Const ABN_POSCHANGED = &H1
Private Const ABN_FULLSCREENAPP = &H2
Private Const ABN_WINDOWARRANGE = &H3 '  lParam == TRUE means hide

'  flags for get state

Private Const ABS_AUTOHIDE = &H1
Private Const ABS_ALWAYSONTOP = &H2
Private Const ABE_LEFT = 0
Private Const ABE_TOP = 1
Private Const ABE_RIGHT = 2
Private Const ABE_BOTTOM = 3

Type APPBARDATA
        cbSize As Long
        hWnd As Long
        uCallbackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long '  message specific
End Type

Declare Function SHAppBarMessage Lib "shell32.dll" Alias "SHAppBarMessage" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

' //  EndAppBar

Declare Function DoEnvironmentSubst Lib "shell32.dll" Alias "DoEnvironmentSubstA" (ByVal szString As String, ByVal cbString As Long) As Long
Declare Function FindEnvironmentString Lib "shell32.dll" Alias "FindEnvironmentStringA" (ByVal szEnvVar As String) As String

Private Const EIRESID = -1

Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long

' // Shell File Operations

Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4
Private Const FOF_MULTIDESTFILES = &H1
Private Const FOF_CONFIRMMOUSE = &H2
Private Const FOF_SILENT = &H4                      '  don't create progress/report
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Private Const FOF_WANTMAPPINGHANDLE = &H20          '  Fill in SHFILEOPSTRUCT.hNameMappings
                                      '  Must be freed using SHFreeNameMappings
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_FILESONLY = &H80                  '  on *.*, do only files
Private Const FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files
Private Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs

Private Const PO_DELETE = &H13           '  printer is being deleted
Private Const PO_RENAME = &H14           '  printer is being renamed
Private Const PO_PORTCHANGE = &H20       '  port this printer connected to is being changed
                                '  if this id is set, the strings received by
                                '  the copyhook are a doubly-null terminated
                                '  list of strings.  The first is the printer
                                '  name and the second is the printer port.
Private Const PO_REN_PORT = &H34         '  PO_RENAME and PO_PORTCHANGE at same time.

'  no POF_ flags currently defined

'  implicit parameters are:
'       if pFrom or pTo are unqualified names the current directories are
'       taken from the global current drive/directory settings managed
'       by Get/SetCurrentDrive/Directory
'
'       the global confirmation settings

Type SHFILEOPSTRUCT
        hWnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type

Declare Function SHFileOperation Lib "shell32.dll" Alias " SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Declare Sub SHFreeNameMappings Lib "shell32.dll" Alias "SHFreeNameMappings" (ByVal hNameMappings As Long)

Type SHNAMEMAPPING
        pszOldPath As String
        pszNewPath As String
        cchOldPath As Long
        cchNewPath As Long
End Type

' // End Shell File Operations

' //  Begin ShellExecuteEx and family

'  ShellExecute() and ShellExecuteEx() error codes

'  regular WinExec() codes
Private Const SE_ERR_FNF = 2                     '  file not found
Private Const SE_ERR_PNF = 3                     '  path not found
Private Const SE_ERR_ACCESSDENIED = 5            '  access denied
Private Const SE_ERR_OOM = 8                     '  out of memory
Private Const SE_ERR_DLLNOTFOUND = 32


'  Note CLASSKEY overrides CLASSNAME
Private Const SEE_MASK_CLASSNAME = &H1
Private Const SEE_MASK_CLASSKEY = &H3
'  Note INVOKEIDLIST overrides IDLIST
Private Const SEE_MASK_IDLIST = &H4
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_ICON = &H10
Private Const SEE_MASK_HOTKEY = &H20
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_CONNECTNETDRV = &H80
Private Const SEE_MASK_FLAG_DDEWAIT = &H100
Private Const SEE_MASK_DOENVSUBST = &H200
Private Const SEE_MASK_FLAG_NO_UI = &H400

Type SHELLEXECUTEINFO
        cbSize As Long
        fMask As Long
        hWnd As Long
        lpVerb As String
        lpFile As String
        lpParameters As String
        lpDirectory As String
        nShow As Long
        hInstApp As Long
        '  Optional fields
        lpIDList As Long
        lpClass As String
        hkeyClass As Long
        dwHotKey As Long
        hIcon As Long
        hProcess As Long
End Type

Declare Sub WinExecError Lib "shell32.dll" Alias "WinExecErrorA" (ByVal hwnd As Long, ByVal error As Long, ByVal lpstrFileName As String, ByVal lpstrTitle As String)

' //  End ShellExecuteEx and family

' // Tray notification definitions

Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias " Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

' // End Tray Notification Icons

' // Begin SHGetFileInfo

'  * The SHGetFileInfo API provides an easy way to get attributes
'  * for a file given a pathname.
'  *
'  *   PARAMETERS
'  *
'  *     pszPath              file name to get info about
'  *     dwFileAttributes     file attribs, only used with SHGFI_USEFILEATTRIBUTES
'  *     psfi                 place to return file info
'  *     cbFileInfo           size of structure
'  *     uFlags               flags
'  *
'  *   RETURN
'  *     TRUE if things worked
'  */

Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long          '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80         '  out: type name
End Type

Private Const SHGFI_ICON = &H100                         '  get icon
Private Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Private Const SHGFI_TYPENAME = &H400                     '  get type name
Private Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Private Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Private Const SHGFI_EXETYPE = &H2000                     '  return exe type
Private Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Private Const SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
Private Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Private Const SHGFI_LARGEICON = &H0                      '  get large icon
Private Const SHGFI_SMALLICON = &H1                      '  get small icon
Private Const SHGFI_OPENICON = &H2                       '  get open icon
Private Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Private Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
Private Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute

Declare Function SHGetFileInfo Lib "shell32.dll" Alias " SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Declare Function SHGetNewLinkInfo Lib "shell32.dll" Alias "SHGetNewLinkInfoA" (ByVal pszLinkto As String, ByVal pszDir As String, ByVal pszName As String, pfMustCopy As Long, ByVal uFlags As Long) As Long

Private Const SHGNLI_PIDL = &H1                          '  pszLinkTo is a pidl
Private Const SHGNLI_PREFIXNAME = &H2                    '  Make name "Shortcut to xxx"

' // End SHGetFileInfo


' Copyright (C) 1993 - 1995 Microsoft Corporation

' Module Name:

'     winperf.h

' Abstract:

'     Header file for the Performance Monitor data.

'     This file contains the definitions of the data structures returned
'     by the Configuration Registry in response to a request for
'     performance data.  This file is used by both the Configuration
'     Registry and the Performance Monitor to define their interface.
'     The complete interface is described here, except for the name
'     of the node to query in the registry.  It is

'                    HKEY_PERFORMANCE_DATA.

'     By querying that node with a subkey of "Global" the caller will
'     retrieve the structures described here.

'     There is no need to RegOpenKey() the reserved handle HKEY_PERFORMANCE_DATA,
'     but the caller should RegCloseKey() the handle so that network transports
'     and drivers can be removed or installed (which cannot happen while
'     they are open for monitoring.)  Remote requests must first
'     RegConnectRegistry().

' --*/

'   Data structure definitions.

'   In order for data to be returned through the Configuration Registry
'   in a system-independent fashion, it must be self-describing.

'   In the following, all offsets are in bytes.

'
'   Data is returned through the Configuration Registry in a
'   a data block which begins with a _PERF_DATA_BLOCK structure.
'

'   The _PERF_DATA_BLOCK structure is followed by NumObjectTypes of
'   data sections, one for each type of object measured.  Each object
'   type section begins with a _PERF_OBJECT_TYPE structure.


' *****************************************************************************                                                                             *
' * winver.h -    Version management functions, types, and definitions          *
' *                                                                             *
' *               Include file for VER.DLL.  This library is                    *
' *               designed to allow version stamping of Windows executable files*
' *               and of special .VER files for DOS executable files.           *
' *                                                                             *
' *               Copyright (c) 1993, Microsoft Corp.  All rights reserved      *
' *                                                                             *
' \*****************************************************************************/

'  ----- Symbols -----
Private Const VS_VERSION_INFO = 1
Private Const VS_USER_DEFINED = 100

'  ----- VS_VERSION.dwFileFlags -----
Private Const VS_FFI_SIGNATURE = &HFEEF04BD
Private Const VS_FFI_STRUCVERSION = &H10000
Private Const VS_FFI_FILEFLAGSMASK = &H3F&

'  ----- VS_VERSION.dwFileFlags -----
Private Const VS_FF_DEBUG = &H1&
Private Const VS_FF_PRERELEASE = &H2&
Private Const VS_FF_PATCHED = &H4&
Private Const VS_FF_PRIVATEBUILD = &H8&
Private Const VS_FF_INFOINFERRED = &H10&
Private Const VS_FF_SPECIALBUILD = &H20&

'  ----- VS_VERSION.dwFileOS -----
Private Const VOS_UNKNOWN = &H0&
Private Const VOS_DOS = &H10000
Private Const VOS_OS216 = &H20000
Private Const VOS_OS232 = &H30000
Private Const VOS_NT = &H40000

Private Const VOS__BASE = &H0&
Private Const VOS__WINDOWS16 = &H1&
Private Const VOS__PM16 = &H2&
Private Const VOS__PM32 = &H3&
Private Const VOS__WINDOWS32 = &H4&

Private Const VOS_DOS_WINDOWS16 = &H10001
Private Const VOS_DOS_WINDOWS32 = &H10004
Private Const VOS_OS216_PM16 = &H20002
Private Const VOS_OS232_PM32 = &H30003
Private Const VOS_NT_WINDOWS32 = &H40004

'  ----- VS_VERSION.dwFileType -----
Private Const VFT_UNKNOWN = &H0&
Private Const VFT_APP = &H1&
Private Const VFT_DLL = &H2&
Private Const VFT_DRV = &H3&
Private Const VFT_FONT = &H4&
Private Const VFT_VXD = &H5&
Private Const VFT_STATIC_LIB = &H7&

'  ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_DRV -----
Private Const VFT2_UNKNOWN = &H0&
Private Const VFT2_DRV_PRINTER = &H1&
Private Const VFT2_DRV_KEYBOARD = &H2&
Private Const VFT2_DRV_LANGUAGE = &H3&
Private Const VFT2_DRV_DISPLAY = &H4&
Private Const VFT2_DRV_MOUSE = &H5&
Private Const VFT2_DRV_NETWORK = &H6&
Private Const VFT2_DRV_SYSTEM = &H7&
Private Const VFT2_DRV_INSTALLABLE = &H8&
Private Const VFT2_DRV_SOUND = &H9&
Private Const VFT2_DRV_COMM = &HA&
Private Const VFT2_DRV_INPUTMETHOD = &HB&

'  ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_FONT -----
Private Const VFT2_FONT_RASTER = &H1&
Private Const VFT2_FONT_VECTOR = &H2&
Private Const VFT2_FONT_TRUETYPE = &H3&

'  ----- VerFindFile() flags -----
Private Const VFFF_ISSHAREDFILE = &H1

Private Const VFF_CURNEDEST = &H1
Private Const VFF_FILEINUSE = &H2
Private Const VFF_BUFFTOOSMALL = &H4

'  ----- VerInstallFile() flags -----
Private Const VIFF_FORCEINSTALL = &H1
Private Const VIFF_DONTDELETEOLD = &H2

Private Const VIF_TEMPFILE = &H1&
Private Const VIF_MISMATCH = &H2&
Private Const VIF_SRCOLD = &H4&

Private Const VIF_DIFFLANG = &H8&
Private Const VIF_DIFFCODEPG = &H10&
Private Const VIF_DIFFTYPE = &H20&

Private Const VIF_WRITEPROT = &H40&
Private Const VIF_FILEINUSE = &H80&
Private Const VIF_OUTOFSPACE = &H100&
Private Const VIF_ACCESSVIOLATION = &H200&
Private Const VIF_SHARINGVIOLATION = &H400&
Private Const VIF_CANNOTCREATE = &H800&
Private Const VIF_CANNOTDELETE = &H1000&
Private Const VIF_CANNOTRENAME = &H2000&
Private Const VIF_CANNOTDELETECUR = &H4000&
Private Const VIF_OUTOFMEMORY = &H8000&

Private Const VIF_CANNOTREADSRC = &H10000
Private Const VIF_CANNOTREADDST = &H20000
Private Const VIF_BUFFTOOSMALL = &H40000

'  ----- Types and structures -----

Type VS_FIXEDFILEINFO
        dwSignature As Long
        dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
        dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
        dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
        dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
        dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
        dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
        dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
        dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
        dwFileType As Long             '  e.g. VFT_DRIVER
        dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
        dwFileDateMS As Long           '  e.g. 0
        dwFileDateLS As Long           '  e.g. 0
End Type

'  ----- Function prototypes -----

Declare Function VerFindFile Lib "version.dll" Alias "VerFindFileA" (ByVal uFlags As Long, ByVal szFileName As String, ByVal szWinDir As String, ByVal szAppDir As String, ByVal szCurDir As String, lpuCurDirLen As Long, ByVal szDestDir As String, lpuDestDirLen As Long) As Long
Declare Function VerInstallFile Lib "version.dll" Alias " VerInstallFileA" (ByVal uFlags As Long, ByVal szSrcFileName As String, ByVal szDestFileName As String, ByVal szSrcDir As String, ByVal szDestDir As String, ByVal szCurDir As String, ByVal szTmpFile As String, lpuTmpFileLen As Long) As Long

'  Returns size of version info in Bytes
Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

'  Read version info into buffer
' /* Length of buffer for info *
' /* Information from GetFileVersionSize *
' /* Filename of version stamped file *
Declare Function GetFileVersionInfo Lib "version.dll"  Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValue" (pBlock As Any, ByVal lpSubBlock As String, ByVal lplpBuffer As Long, puLen As Long) As Long

' ***********************************************************************
' *                                                                       *
' *   winbase.h -- This module defines the 32-Bit Windows Base APIs       *
' *                                                                       *
' *   Copyright (c) 1990-1995, Microsoft Corp. All rights reserved.       *
' *                                                                       *
' ************************************************************************/
Type ICONMETRICS
    cbSize As Long
    iHorzSpacing As Long
    iVertSpacing As Long
    iTitleWrap As Long
    lfFont As LOGFONT
End Type

Type HELPINFO
    cbSize As Long
    iContextType As Long
    iCtrlId As Long
    hItemHandle As Long
    dwContextId As Long
    MousePos As POINTAPI
End Type

Type ANIMATIONINFO
    cbSize As Long
    iMinAnimate As Long
End Type

Type MINIMIZEDMETRICS
    cbSize As Long
    iWidth As Long
    iHorzGap As Long
    iVertGap As Long
    iArrange As Long
    lfFont As LOGFONT
End Type

'  Define API decoration for direct importing of DLL references.
Declare Function HeapValidate Lib "kernel32" Alias "HeapValidate" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Declare Function HeapCompact Lib "kernel32" Alias "HeapCompact" (ByVal hHeap As Long, ByVal dwFlags As Long) As Long

Private Const PROCESS_HEAP_REGION = &H1
Private Const PROCESS_HEAP_UNCOMMITTED_RANGE = &H2
Private Const PROCESS_HEAP_ENTRY_BUSY = &H4
Private Const PROCESS_HEAP_ENTRY_MOVEABLE = &H10
Private Const PROCESS_HEAP_ENTRY_DDESHARE = &H20

Declare Function HeapLock Lib "kernel32" Alias "HeapLock" (ByVal hHeap As Long) As Long
Declare Function HeapUnlock Lib "kernel32" Alias "HeapUnlock" (ByVal hHeap As Long) As Long

'  GetBinaryType return values.

Private Const SCS_32BIT_BINARY = 0
Private Const SCS_DOS_BINARY = 1
Private Const SCS_WOW_BINARY = 2
Private Const SCS_PIF_BINARY = 3
Private Const SCS_POSIX_BINARY = 4
Private Const SCS_OS216_BINARY = 5

Declare Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeA" (ByVal lpApplicationName As String, lpBinaryType As Long) As Long

Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathName" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function GetProcessAffinityMask Lib "kernel32" Alias "GetProcessAffinityMask" (ByVal hProcess As Long, lpProcessAffinityMask As Long, SystemAffinityMask As Long) As Long

'  Logon Support APIs

Private Const LOGON32_LOGON_INTERACTIVE = 2
Private Const LOGON32_LOGON_BATCH = 4
Private Const LOGON32_LOGON_SERVICE = 5

Private Const LOGON32_PROVIDER_DEFAULT = 0
Private Const LOGON32_PROVIDER_WINNT35 = 1

Declare Function LogonUser Lib "ADVAPI32.DLL" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long

Declare Function ImpersonateLoggedOnUser Lib "kernel32" Alias "ImpersonateLoggedOnUser" (ByVal hToken As Long) As Long
Declare Function CreateProcessAsUser Lib "kernel32" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As SECURITY_ATTRIBUTES, ByVal lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, ByVal lpStartupInfo As STARTUPINFO, ByVal lpProcessInformation As PROCESS_INFORMATION) As Long

'  Performance counter API's

Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

'  dwPlatformId defines:
'
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByVal lpVersionInformation As OSVERSIONINFO) As Long

'  Power Management APIs

Private Const AC_LINE_OFFLINE = &H0
Private Const AC_LINE_ONLINE = &H1
Private Const AC_LINE_BACKUP_POWER = &H2
Private Const AC_LINE_UNKNOWN = &HFF
Private Const BATTERY_FLAG_HIGH = &H1
Private Const BATTERY_FLAG_LOW = &H2
Private Const BATTERY_FLAG_CRITICAL = &H4
Private Const BATTERY_FLAG_CHARGING = &H8
Private Const BATTERY_FLAG_NO_BATTERY = &H80
Private Const BATTERY_FLAG_UNKNOWN = &HFF
Private Const BATTERY_PERCENTAGE_UNKNOWN = &HFF
Private Const BATTERY_LIFE_UNKNOWN = &HFFFF

Type SYSTEM_POWER_STATUS
        ACLineStatus As Byte
        BatteryFlag As Byte
        BatteryLifePercent As Byte
        Reserved1 As Byte
        BatteryLifeTime As Long
        BatteryFullLifeTime As Long
End Type

Declare Function GetSystemPowerStatus Lib "kernel32" Alias "GetSystemPowerStatus" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Declare Function SetSystemPowerState Lib "kernel32" Alias "SetSystemPowerState" (ByVal fSuspend As Long, ByVal fForce As Long) As Long

' *   commdlg.h -- This module defines the 32-Bit Common Dialog APIs      *

Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
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
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer

Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Private Const OFN_EXPLORER = &H80000                         '  new look commdlg
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules

Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Type OFNOTIFY
        hdr As NMHDR
        lpOFN As OPENFILENAME
        pszFile As String        '  May be NULL
End Type

Private Const CDM_FIRST = (WM_USER + 100)
Private Const CDM_LAST = (WM_USER + 200)
Private Const CDM_GETSPEC = (CDM_FIRST + &H0)
Private Const CDM_GETFILEPATH = (CDM_FIRST + &H1)
Private Const CDM_GETFOLDERPATH = (CDM_FIRST + &H2)
Private Const CDM_GETFOLDERIDLIST = (CDM_FIRST + &H3)
Private Const CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)
Private Const CDM_HIDECONTROL = (CDM_FIRST + &H5)
Private Const CDM_SETDEFEXT = (CDM_FIRST + &H6)

Type ChooseColor
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        Flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Private Const CC_RGBINIT = &H1
Private Const CC_FULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_SHOWHELP = &H8
Private Const CC_ENABLEHOOK = &H10
Private Const CC_ENABLETEMPLATE = &H20
Private Const CC_ENABLETEMPLATEHANDLE = &H40
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_ANYCOLOR = &H100

Type FINDREPLACE
        lStructSize As Long        '  size of this struct 0x20
        hwndOwner As Long          '  handle to owner's window
        hInstance As Long          '  instance handle of.EXE that
                                    '    contains cust. dlg. template
        Flags As Long              '  one or more of the FR_??
        lpstrFindWhat As String      '  ptr. to search string
        lpstrReplaceWith As String   '  ptr. to replace string
        wFindWhatLen As Integer       '  size of find buffer
        wReplaceWithLen As Integer    '  size of replace buffer
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long            '  ptr. to hook fn. or NULL
        lpTemplateName As String     '  custom template name
End Type

Private Const FR_DOWN = &H1
Private Const FR_WHOLEWORD = &H2
Private Const FR_MATCHCASE = &H4
Private Const FR_FINDNEXT = &H8
Private Const FR_REPLACE = &H10
Private Const FR_REPLACEALL = &H20
Private Const FR_DIALOGTERM = &H40
Private Const FR_SHOWHELP = &H80
Private Const FR_ENABLEHOOK = &H100
Private Const FR_ENABLETEMPLATE = &H200
Private Const FR_NOUPDOWN = &H400
Private Const FR_NOMATCHCASE = &H800
Private Const FR_NOWHOLEWORD = &H1000
Private Const FR_ENABLETEMPLATEHANDLE = &H2000
Private Const FR_HIDEUPDOWN = &H4000
Private Const FR_HIDEMATCHCASE = &H8000
Private Const FR_HIDEWHOLEWORD = &H10000

Declare Function FindText Lib "comdlg32.dll" Alias "FindTextA " (pFindreplace As FINDREPLACE) As Long

Declare Function ReplaceText Lib "comdlg32.dll" Alias "ReplaceTextA" (pFindreplace As FINDREPLACE) As Long

Type ChooseFont
        lStructSize As Long
        hwndOwner As Long          '  caller's window handle
        hdc As Long                '  printer DC/IC or NULL
        lpLogFont As LOGFONT          '  ptr. to a LOGFONT struct
        iPointSize As Long         '  10 * size in points of selected font
        Flags As Long              '  enum. type flags
        rgbColors As Long          '  returned text color
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long           '  ptr. to hook function
        lpTemplateName As String     '  custom template name
        hInstance As Long          '  instance handle of.EXE that
                                       '    contains cust. dlg. template
        lpszStyle As String          '  return the style field here
                                       '  must be LF_FACESIZE or bigger
        nFontType As Integer          '  same value reported to the EnumFonts
                                       '    call back with the extra FONTTYPE_
                                       '    bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           '  minimum pt size allowed &
        nSizeMax As Long           '  max pt size allowed if
                                       '    CF_LIMITSIZE is used
End Type

Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long

Private Const CF_SCREENFONTS = &H1
Private Const CF_PRINTERFONTS = &H2
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_SHOWHELP = &H4&
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_USESTYLE = &H80&
Private Const CF_EFFECTS = &H100&
Private Const CF_APPLY = &H200&
Private Const CF_ANSIONLY = &H400&
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_TTONLY = &H40000
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_NOSIZESEL = &H200000
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOVERTFONTS = &H1000000

Private Const SIMULATED_FONTTYPE = &H8000
Private Const PRINTER_FONTTYPE = &H4000
Private Const SCREEN_FONTTYPE = &H2000
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Private Const REGULAR_FONTTYPE = &H400

Private Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
Private Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
Private Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)

Private Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Private Const SHAREVISTRING = "commdlg_ShareViolation"
Private Const FILEOKSTRING = "commdlg_FileNameOK"
Private Const COLOROKSTRING = "commdlg_ColorOK"
Private Const SETRGBSTRING = "commdlg_SetRGBColor"
Private Const HELPMSGSTRING = "commdlg_help"
Private Const FINDMSGSTRING = "commdlg_FindReplace"

Private Const CD_LBSELNOITEMS = -1
Private Const CD_LBSELCHANGE = 0
Private Const CD_LBSELSUB = 1
Private Const CD_LBSELADD = 2

Type PrintDlg
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        Flags As Long
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

Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG) As Long

Private Const PD_ALLPAGES = &H0
Private Const PD_SELECTION = &H1
Private Const PD_PAGENUMS = &H2
Private Const PD_NOSELECTION = &H4
Private Const PD_NOPAGENUMS = &H8
Private Const PD_COLLATE = &H10
Private Const PD_PRINTTOFILE = &H20
Private Const PD_PRINTSETUP = &H40
Private Const PD_NOWARNING = &H80
Private Const PD_RETURNDC = &H100
Private Const PD_RETURNIC = &H200
Private Const PD_RETURNDEFAULT = &H400
Private Const PD_SHOWHELP = &H800
Private Const PD_ENABLEPRINTHOOK = &H1000
Private Const PD_ENABLESETUPHOOK = &H2000
Private Const PD_ENABLEPRINTTEMPLATE = &H4000
Private Const PD_ENABLESETUPTEMPLATE = &H8000
Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Private Const PD_USEDEVMODECOPIES = &H40000
Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Private Const PD_DISABLEPRINTTOFILE = &H80000
Private Const PD_HIDEPRINTTOFILE = &H100000
Private Const PD_NONETWORKBUTTON = &H200000

Type DEVNAMES
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
End Type

Private Const DN_DEFAULTPRN = &H1

Declare Function CommDlgExtendedError Lib "comdlg32.dll" Alias "CommDlgExtendedError" () As Long

Private Const WM_PSD_PAGESETUPDLG = (WM_USER)
Private Const WM_PSD_FULLPAGERECT = (WM_USER + 1)
Private Const WM_PSD_MINMARGINRECT = (WM_USER + 2)
Private Const WM_PSD_MARGINRECT = (WM_USER + 3)
Private Const WM_PSD_GREEKTEXTRECT = (WM_USER + 4)
Private Const WM_PSD_ENVSTAMPRECT = (WM_USER + 5)
Private Const WM_PSD_YAFULLPAGERECT = (WM_USER + 6)

Type PageSetupDlg
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        Flags As Long
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

Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long

Private Const PSD_DEFAULTMINMARGINS = &H0 '  default (printer's)
Private Const PSD_INWININIINTLMEASURE = &H0 '  1st of 4 possible

Private Const PSD_MINMARGINS = &H1 '  use caller's
Private Const PSD_MARGINS = &H2 '  use caller's
Private Const PSD_INTHOUSANDTHSOFINCHES = &H4 '  2nd of 4 possible
Private Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8 '  3rd of 4 possible
Private Const PSD_DISABLEMARGINS = &H10
Private Const PSD_DISABLEPRINTER = &H20
Private Const PSD_NOWARNING = &H80 '  must be same as PD_*
Private Const PSD_DISABLEORIENTATION = &H100
Private Const PSD_RETURNDEFAULT = &H400 '  must be same as PD_*
Private Const PSD_DISABLEPAPER = &H200
Private Const PSD_SHOWHELP = &H800 '  must be same as PD_*
Private Const PSD_ENABLEPAGESETUPHOOK = &H2000 '  must be same as PD_*
Private Const PSD_ENABLEPAGESETUPTEMPLATE = &H8000 '  must be same as PD_*
Private Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE = &H20000 '  must be same as PD_*
Private Const PSD_ENABLEPAGEPAINTHOOK = &H40000
Private Const PSD_DISABLEPAGEPAINTING = &H80000

Declare Function DdeInitialize Lib "user32" Alias "DdeInitializeA" (pidInst As Long, ByVal pfnCallback As Long, ByVal afCmd As Long, ByVal ulRes As Long) As Integer
Declare Function SetServiceBits Lib "advapi32" Alias "SetServiceBits" (ByVal hServiceStatus As Long, ByVal dwServiceBits As Long, ByVal bSetBitsOn As Boolean, ByVal bUpdateImmediately As Boolean) As Long
Declare Function CopyLZFile Lib "lz32" Alias "CopyLZFile" (ByVal n1 As Long, ByVal n2 As Long) As Long
Declare Function LZStart Lib "lz32" Alias "LZStart" () As Long
Declare Sub LZDone Lib "lz32" Alias "LZDone" ()
Declare Function mciGetYieldProc Lib "winmm" Alias "mciGetYieldProc" (ByVal mciId As Long, pdwYieldData As Long) As Long
Declare Function mciSetYieldProc Lib "winmm" Alias "mciSetYieldProc" (ByVal mciId As Long, ByVal fpYieldProc As Long, ByVal dwYieldData As Long) As Long
Declare Function midiOutGetNumDevs Lib "winmm" Alias "midiOutGetNumDevs" () As Integer
Declare Function mmioInstallIOProcA Lib "winmm" Alias "mmioInstallIOProcA" (ByVal fccIOProc As String * 4, ByVal pIOProc As Long, ByVal dwFlags As Long) As Long
Declare Function CommandLineToArgv Lib "shell32" Alias "CommandLineToArgvW" (ByVal lpCmdLine As String, pNumArgs As Integer) As Long
Declare Function IsTextUnicode Lib "advapi32" Alias "IsTextUnicode" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Declare Function NotifyChangeEventLog Lib "advapi32" Alias "NotifyChangeEventLog" (ByVal hEventLog As Long, ByVal hEvent As Long) As Long
Declare Function ObjectOpenAuditAlarm Lib "advapi32" Alias "ObjectOpenAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As Long, ByVal DesiredAccess As Long, ByVal GrantedAccess As Long, Privileges As PRIVILEGE_SET, ByVal ObjectCreation As Boolean, ByVal AccessGranted As Boolean, GenerateOnClose As Boolean) As Long
Declare Function SetThreadToken Lib "advapi32" Alias "SetThreadToken" (Thread As Long, ByVal Token As Long) As Long

Type COMMCONFIG
    dwSize As Long
    wVersion As Integer
    wReserved As Integer
    dcbx As DCB
    dwProviderSubType As Long
    dwProviderOffset As Long
    dwProviderSize As Long
    wcProviderData As Byte
End Type

Declare Function CommConfigDialog Lib "kernel32" Alias "CommConfigDialogA" (ByVal lpszName As String, ByVal hWnd As Long, lpCC As COMMCONFIG) As Long
Declare Function CreateIoCompletionPort Lib "kernel32" Alias "CreateIoCompletionPort" (ByVal FileHandle As Long, ByVal ExistingCompletionPort As Long, ByVal CompletionKey As Long, ByVal NumberOfConcurrentThreads As Long) As Long
Declare Function DisableThreadLibraryCalls Lib "kernel32" Alias "DisableThreadLibraryCalls" (ByVal hLibModule As Long) As Long
Declare Function EnumResourceLanguages Lib "kernel32" Alias "EnumResourceLanguagesA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function FreeEnvironmentStrings Lib "kernel32" Alias "FreeEnvironmentStringsA" (ByVal lpsz As String) As Long

Declare Sub FreeLibraryAndExitThread Lib "kernel32" Alias "FreeLibraryAndExitThread" (ByVal hLibModule As Long, ByVal dwExitCode As Long)
Declare Function FreeResource Lib "kernel32" Alias "FreeResource" (ByVal hResData As Long) As Long
Declare Function GetCommConfig Lib "kernel32" Alias "GetCommConfig" (ByVal hCommDev As Long, lpCC As COMMCONFIG, lpdwSize As Long) As Long
Declare Function GetCompressedFileSize Lib "kernel32" Alias "GetCompressedFileSizeA" (ByVal lpFileName As String, lpFileSizeHigh As Long) As Long
Declare Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, lpdwSize As Long) As Long
Declare Function GetHandleInformation Lib "kernel32" Alias "GetHandleInformation" (ByVal hObject As Long, lpdwFlags As Long) As Long
Declare Function GetProcessHeaps Lib "kernel32" Alias "GetProcessHeaps" (ByVal NumberOfHeaps As Long, ProcessHeaps As Long) As Long
Declare Function GetProcessWorkingSetSize Lib "kernel32" Alias "GetProcessWorkingSetSize" (ByVal hProcess As Long, lpMinimumWorkingSetSize As Long, lpMaximumWorkingSetSize As Long) As Long
Declare Function GetQueuedCompletionStatus Lib "kernel32" Alias "GetQueuedCompletionStatus" (ByVal CompletionPort As Long, lpNumberOfBytesTransferred As Long, lpCompletionKey As Long, lpOverlapped As Long, ByVal dwMilliseconds As Long) As Long
Declare Function GetSystemTimeAdjustment Lib "kernel32" Alias "GetSystemTimeAdjustment" (lpTimeAdjustment As Long, lpTimeIncrement As Long, lpTimeAdjustmentDisabled As Boolean) As Long

Declare Function GlobalCompact Lib "kernel32" Alias "GlobalCompact" (ByVal dwMinFree As Long) As Long
Declare Sub GlobalFix Lib "kernel32" Alias "GlobalFix" (ByVal hMem As Long)
Declare Sub GlobalUnfix Lib "kernel32" Alias "GlobalUnfix" (ByVal hMem As Long)
Declare Function GlobalWire Lib "kernel32" Alias "GlobalWire" (ByVal hMem As Long) As Long
Declare Function GlobalUnWire Lib "kernel32" Alias "GlobalUnWire" (ByVal hMem As Long) As Long

Declare Function IsBadCodePtr Lib "kernel32" Alias "IsBadCodePtr" (ByVal lpfn As Long) As Long
Declare Function LocalCompact Lib "kernel32" Alias "LocalCompact" (ByVal uMinFree As Long) As Long
Declare Function LocalShrink Lib "kernel32" Alias "LocalShrink" (ByVal hMem As Long, ByVal cbNewSize As Long) As Long
Declare Function MapViewOfFile Lib "kernel32" Alias "MapViewOfFile" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Declare Function ReadFileEx Lib "kernel32" Alias "ReadFileEx" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long

Declare Function SetCommConfig Lib "kernel32" Alias "SetCommConfig" (ByVal hCommDev As Long, lpCC As COMMCONFIG, ByVal dwSize As Long) As Long
Declare Function SetDefaultCommConfig Lib "kernel32" Alias "SetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, ByVal dwSize As Long) As Long
Declare Sub SetFileApisToANSI Lib "kernel32" Alias "SetFileApisToANSI" ()
Declare Function SetHandleInformation Lib "kernel32" Alias "SetHandleInformation" (ByVal hObject As Long, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Declare Function SetProcessWorkingSetSize Lib "kernel32" Alias "SetProcessWorkingSetSize" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long

Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function SetSystemTimeAdjustment Lib "kernel32" Alias "SetSystemTimeAdjustment" (ByVal dwTimeAdjustment As Long, ByVal bTimeAdjustmentDisabled As Boolean) As Long
Declare Function SetThreadAffinityMask Lib "kernel32" Alias "SetThreadAffinityMask" (ByVal hThread As Long, ByVal dwThreadAffinityMask As Long) As Long
Declare Function SetUnhandledExceptionFilter Lib "kernel32" Alias "SetUnhandledExceptionFilter" (ByVal lpTopLevelExceptionFilter As Long) As Long
Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
Declare Function WriteFileEx Lib "kernel32" Alias "WriteFileEx" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long

Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlphaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type

Declare Function ChoosePixelFormat Lib "gdi32" Alias "ChoosePixelFormat" (ByVal hDC As Long, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Declare Function CreateDIBSection Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Declare Function DescribePixelFormat Lib "gdi32" Alias "DescribePixelFormat" (ByVal hDC As Long, ByVal n As Long, ByVal un As Long, lpPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Declare Function EndDoc Lib "gdi32" Alias "EndDoc" (ByVal hDC As Long) As Long
Declare Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hDC As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long
Declare Function EnumMetaFile Lib "gdi32" Alias "EnumMetaFile" (ByVal hDC As Long, ByVal hMetafile As Long, ByVal lpMFEnumProc As Long, ByVal lParam As Long) As Long
Declare Function EnumObjects Lib "gdi32" Alias "EnumObjects" (ByVal hDC As Long, ByVal n As Long, ByVal lpGOBJEnumProc As Long, lpVoid As Any) As Long

Declare Function FixBrushOrgEx Lib "gdi32" Alias "FixBrushOrgEx" (ByVal hDC As Long, ByVal n1 As Long, ByVal n2 As Long, lpPoint As POINTAPI) As Long
Declare Function GetBrushOrgEx Lib "gdi32" Alias "GetBrushOrgEx" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
Declare Function GetCharWidth Lib "gdi32" Alias "GetCharWidthA" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, lpn As Long) As Long
Declare Function GetDIBColorTable Lib "gdi32" Alias "GetDIBColorTable" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQUAD) As Long
Declare Function GetPixelFormat Lib "gdi32" Alias "GetPixelFormat" (ByVal hDC As Long) As Long
Declare Function LineDDA Lib "gdi32" Alias "LineDDA" (ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal lpLineDDAProc As Long, ByVal lParam As Long) As Long

Declare Function SetAbortProc Lib "gdi32" Alias "SetAbortProc" (ByVal hDC As Long, ByVal lpAbortProc As Long) As Long
Declare Function SetDIBColorTable Lib "gdi32" Alias "SetDIBColorTable" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
Declare Function SetPixelFormat Lib "gdi32" Alias "SetPixelFormat" (ByVal hDC As Long, ByVal n As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Declare Function SwapBuffers Lib "gdi32" Alias "SwapBuffers" (ByVal hDC As Long) As Long
Declare Function EnumCalendarInfo Lib "kernel32" Alias "EnumCalendarInfoA" (ByVal lpCalInfoEnumProc As Long, ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long) As Long
Declare Function GetCurrencyFormat Lib "kernel32" Alias "GetCurrencyFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As String, lpFormat As CURRENCYFMT, ByVal lpCurrencyStr As String, ByVal cchCurrency As Long) As Long
Declare Function GetNumberFormat Lib "kernel32" Alias "GetNumberFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As String, lpFormat As NUMBERFMT, ByVal lpNumberStr As String, ByVal cchNumber As Long) As Long
Declare Function GetStringTypeEx Lib "kernel32" Alias "GetStringTypeExA" (ByVal Locale As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Integer) As Long
Declare Function GetStringTypeW Lib "kernel32" Alias "GetStringTypeW" (ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Integer) As Long
Declare Function IsDBCSLeadByte Lib "kernel32" Alias "IsDBCSLeadByte" (ByVal TestChar As Byte) As Long
Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long

Declare Function DeletePrinter Lib "winspool.drv" Alias "DeletePrinter" (ByVal hPrinter As Long) As Long

Declare Function FindClosePrinterChangeNotification Lib "winspool.drv" Alias "FindClosePrinterChangeNotification" (ByVal hChange As Long) As Long
Declare Function FindFirstPrinterChangeNotification Lib "winspool.drv" Alias "FindFirstPrinterChangeNotification" (ByVal hPrinter As Long, ByVal fdwFlags As Long, ByVal fdwOptions As Long, ByVal pPrinterNotifyOptions As String) As Long
Declare Function FindNextPrinterChangeNotification Lib "winspool.drv" Alias "FindNextPrinterChangeNotification" (ByVal hChange As Long, pdwChange As Long, ByVal pvReserved As String, ByVal ppPrinterNotifyInfo As Long) As Long
Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Byte, ByVal Command As Long) As Long

Declare Function BroadcastSystemMessage Lib "user32" Alias "BroadcastSystemMessage" (ByVal dw As Long, pdw As Long, ByVal un As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function CascadeWindows Lib "user32" Alias "CascadeWindows" (ByVal hwndParent As Long, ByVal wHow As Long, ByVal lpRect As RECT, ByVal cKids As Long, lpkids As Long) As Integer
Declare Function ChangeMenu Lib "user32" Alias "ChangeMenuA" (ByVal hMenu As Long, ByVal cmd As Long, ByVal lpszNewItem As String, ByVal cmdInsert As Long, ByVal flags As Long) As Long
Declare Function CheckMenuRadioItem Lib "user32" Alias "CheckMenuRadioItem" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Declare Function ChildWindowFromPoint Lib "user32" Alias "ChildWindowFromPoint" (ByVal hWndParent As Long, ByVal pt As POINTAPI) As Long
Declare Function ChildWindowFromPointEx Lib "user32" Alias "ChildWindowFromPointEx" (ByVal hWnd As Long, ByVal pt As POINTAPI, ByVal un As Long) As Long
Declare Function CloseDesktop Lib "user32" Alias "CloseDesktop" (ByVal hDesktop As Long) As Long
Declare Function CloseWindowStation Lib "user32" Alias "CloseWindowStation" (ByVal hWinSta As Long) As Long

Declare Function CopyImage Lib "user32" Alias "CopyImage" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function CreateDesktop Lib "user32" Alias "CreateDesktopA" (ByVal lpszDesktop As String, ByVal lpszDevice As String, pDevmode As DEVMODE, ByVal dwFlags As Long, ByVal dwDesiredAccess As Long, lpsa As SECURITY_ATTRIBUTES) As Long
Declare Function CreateDialogIndirectParam Lib "user32" Alias "CreateDialogIndirectParamA" (ByVal hInstance As Long, lpTemplate As DLGTEMPLATE, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
Declare Function CreateDialogParam Lib "user32" Alias "CreateDialogParamA" (ByVal hInstance As Long, ByVal lpName As String, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal lParamInit As Long) As Long
Declare Function CreateIconFromResource Lib "user32" Alias "CreateIconFromResource" (presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Boolean, ByVal dwVer As Long) As Long
Declare Function DialogBoxIndirectParam Lib "user32" Alias "DialogBoxIndirectParamA" (ByVal hInstance As Long, hDialogTemplate As DLGTEMPLATE, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long

Declare Function DragDetect Lib "user32" Alias "DragDetect" (ByVal hWnd As Long, ByVal pt As POINTAPI) As Long
Declare Function DragObject Lib "user32" Alias "DragObject" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal un As Long, ByVal dw As Long, ByVal hCursor As Long) As Long
Declare Function DrawAnimatedRects Lib "user32" Alias "DrawAnimatedRects" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As Rect, lprcTo As Rect) As Long
Declare Function DrawCaption Lib "user32" Alias "DrawCaption" (ByVal hWnd As Long, ByVal hDC As Long, pcRect As Rect, ByVal un As Long) As Long
Declare Function DrawEdge Lib "user32" Alias "DrawEdge" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Declare Function DrawFrameControl Lib "user32" Alias "DrawFrameControl" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Declare Function DrawIconEx Lib "user32" Alias "DrawIconEx" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long

Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Declare Function EnumChildWindows Lib "user32" Alias "EnumChildWindows" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hwinsta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumDesktopWindows Lib "user32" Alias "EnumDesktopWindows" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Declare Function EnumPropsEx Lib "user32" Alias "EnumPropsExA" (ByVal hWnd As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumProps Lib "user32" Alias "EnumPropsA" (ByVal hWnd As Long, ByVal lpEnumFunc As Long) As Long
Declare Function EnumThreadWindows Lib "user32" Alias "EnumThreadWindows" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Declare Function GetKeyboardLayoutList Lib "user32" Alias "GetKeyboardLayoutList" (ByVal nBuff As Long, lpList As Long) As Long
Declare Function GetKeyboardLayout Lib "user32" Alias "GetKeyboardLayout" (ByVal dwLayout As Long) As Long
Declare Function GetMenuContextHelpId Lib "user32" Alias "GetMenuContextHelpId" (ByVal hMenu As Long) As Long
Declare Function GetMenuDefaultItem Lib "user32" Alias "GetMenuDefaultItem" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Declare Function GetMenuItemRect Lib "user32" Alias "GetMenuItemRect" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long

Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Declare Function GetScrollInfo Lib "user32" Alias "GetScrollInfo" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Declare Function GetSysColorBrush Lib "user32" Alias "GetSysColorBrush" (ByVal nIndex As Long) As Long
Declare Function GetUserObjectInformation Lib "user32" Alias "GetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Declare Function GetWindowContextHelpId Lib "user32" Alias "GetWindowContextHelpId" (ByVal hWnd As Long) As Long
Declare Function GetWindowRgn Lib "user32" Alias "GetWindowRgn" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Declare Function GrayString Lib "user32" Alias "GrayStringA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpOutputFunc As Long, ByVal lpData As Long, ByVal nCount As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByVal lpcMenuItemInfo As MENUITEMINFO) As Long
Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function LookupIconIdFromDirectoryEx Lib "user32" Alias "LookupIconIdFromDirectoryEx" (presbits As Byte, ByVal fIcon As Boolean, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Declare Function MapVirtualKeyEx Lib "user32" Alias "MapVirtualKeyExA" (ByVal uCode As Long, ByVal uMapType As Long, ByVal dwhkl As Long) As Long
Declare Function MenuItemFromPoint Lib "user32" Alias "MenuItemFromPoint" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal ptScreen As POINTAPI) As Long

Type MSGBOXPARAMS
    cbSize As Long
    hwndOwner As Long
    hInstance As Long
    lpszText As String
    lpszCaption As String
    dwStyle As Long
    lpszIcon As String
    dwContextHelpId As Long
    lpfnMsgBoxCallback As Long
    dwLanguageId As Long
End Type

Declare Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectA" (lpMsgBoxParams As MSGBOXPARAMS) As Long
Declare Function OpenDesktop Lib "user32" Alias "OpenDesktopA" (ByVal lpszDesktop As String, ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenInputDesktop Lib "user32" Alias "OpenInputDesktop" (ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenWindowStation Lib "user32" Alias "OpenWindowStationA" (ByVal lpszWinSta As String, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Declare Function PaintDesktop Lib "user32" Alias "PaintDesktop" (ByVal hdc As Long) As Long

Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long
End Type

Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Declare Function SetMenuContextHelpId Lib "user32" Alias "SetMenuContextHelpId" (ByVal hMenu As Long, ByVal dw As Long) As Long
Declare Function SetMenuDefaultItem Lib "user32" Alias "SetMenuDefaultItem" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Declare Function SetMessageExtraInfo Lib "user32" Alias "SetMessageExtraInfo" (ByVal lParam As Long) As Long
Declare Function SetMessageQueue Lib "user32" Alias "SetMessageQueue" (ByVal cMessagesMax As Long) As Long
Declare Function SetProcessWindowStation Lib "user32" Alias "SetProcessWindowStation" (ByVal hWinSta As Long) As Long
Declare Function SetScrollInfo Lib "user32" Alias "SetScrollInfo" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Declare Function SetSystemCursor Lib "user32" Alias "SetSystemCursor" (ByVal hcur As Long, ByVal id As Long) As Long
Declare Function SetThreadDesktop Lib "user32" Alias "SetThreadDesktop" (ByVal hDesktop As Long) As Long
Declare Function SetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function SetUserObjectInformation Lib "user32" Alias "SetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long) As Long
Declare Function SetWindowContextHelpId Lib "user32" Alias "SetWindowContextHelpId" (ByVal hWnd As Long, ByVal dw As Long) As Long
Declare Function SetWindowRgn Lib "user32" Alias "SetWindowRgn" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As Long) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function ShowWindowAsync Lib "user32" Alias "ShowWindowAsync" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SwitchDesktop Lib "user32" Alias "SwitchDesktop" (ByVal hDesktop As Long) As Long
Declare Function TileWindows Lib "user32" Alias "TileWindows" (ByVal hwndParent As Long, ByVal wHow As Long, lpRect As Rect, ByVal cKids As Long, lpKids As Long) As Integer
Declare Function ToAsciiEx Lib "user32" Alias "ToAsciiEx" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpKeyState As Byte, lpChar As Integer, ByVal uFlags As Long, ByVal dwhkl As Long) As Long

Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
End Type

Declare Function TrackPopupMenuEx Lib "user32" Alias "TrackPopupMenuEx" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hWnd As Long, lpTPMParams As TPMPARAMS) As Long
Declare Function UnhookWindowsHook Lib "user32" Alias "UnhookWindowsHook" (ByVal nCode As Long, ByVal pfnFilterProc As Long) As Long
Declare Function VkKeyScanEx Lib "user32" Alias "VkKeyScanExA" (ByVal ch As Byte, ByVal dwhkl As Long) As Integer
Declare Function WNetGetUniversalName Lib "mpr" Alias "WNetGetUniversalNameA" (ByVal lpLocalPath As String, ByVal dwInfoLevel As Long, lpBuffer As Any, lpBufferSize As Long) As Long
Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Declare Function WNetOpenEnum Lib "mpr" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As NETRESOURCE, lphEnum As Long) As Long

Private Const INVALID_HANDLE_VALUE = -1

'DrawEdge Constants
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC
Private Const BDR_RAISED = &H5
Private Const BDR_SUNKEN = &HA

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const BF_DIAGONAL = &H10

' For diagonal lines, the BF_RECT flags specify the end point of
' the vector bounded by the rectangle parameter.
Private Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

Private Const BF_MIDDLE = &H800    ' Fill in the middle.
Private Const BF_SOFT = &H1000     ' Use for softer buttons.
Private Const BF_ADJUST = &H2000   ' Calculate the space left over.
Private Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Private Const BF_MONO = &H8000     ' For monochrome borders.



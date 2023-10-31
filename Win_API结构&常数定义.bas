Attribute VB_Name = "����ģ��"
Option Explicit
Const BF_MONO = &H8000     ' For monochrome borders.
Const DELETE = &H10000
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_ALL = &H1F0000
Const SPECIFIC_RIGHTS_ALL = &HFFFF
Const SID_REVISION = (1)                         '  Current revision level
Const SID_MAX_SUB_AUTHORITIES = (15)
Const SID_RECOMMENDED_SUB_AUTHORITIES = (1)    ' Will change to around 6 in a future release.
Const SidTypeUser = 1
Const SidTypeGroup = 2
Const SidTypeDomain = 3
Const SidTypeAlias = 4
Const SidTypeWellKnownGroup = 5
Const SidTypeDeletedAccount = 6
Const SidTypeInvalid = 7
Const SidTypeUnknown = 8
Const SECURITY_NULL_RID = &H0
Const SECURITY_WORLD_RID = &H0
Const SECURITY_LOCAL_RID = &H0
Const SECURITY_CREATOR_OWNER_RID = &H0
Const SECURITY_CREATOR_GROUP_RID = &H1
Const SECURITY_DIALUP_RID = &H1
Const SECURITY_NETWORK_RID = &H2
Const SECURITY_BATCH_RID = &H3
Const SECURITY_INTERACTIVE_RID = &H4
Const SECURITY_SERVICE_RID = &H6
Const SECURITY_ANONYMOUS_LOGON_RID = &H7
Const SECURITY_LOGON_IDS_RID = &H5
Const SECURITY_LOCAL_SYSTEM_RID = &H12
Const SECURITY_NT_NON_UNIQUE = &H15
Const SECURITY_BUILTIN_DOMAIN_RID = &H20
Const DOMAIN_USER_RID_ADMIN = &H1F4
Const DOMAIN_USER_RID_GUEST = &H1F5
Const DOMAIN_GROUP_RID_ADMINS = &H200
Const DOMAIN_GROUP_RID_USERS = &H201
Const DOMAIN_GROUP_RID_GUESTS = &H202
Const DOMAIN_ALIAS_RID_ADMINS = &H220
Const DOMAIN_ALIAS_RID_USERS = &H221
Const DOMAIN_ALIAS_RID_GUESTS = &H222
Const DOMAIN_ALIAS_RID_POWER_USERS = &H223
Const DOMAIN_ALIAS_RID_ACCOUNT_OPS = &H224
Const DOMAIN_ALIAS_RID_SYSTEM_OPS = &H225
Const DOMAIN_ALIAS_RID_PRINT_OPS = &H226
Const DOMAIN_ALIAS_RID_BACKUP_OPS = &H227
Const DOMAIN_ALIAS_RID_REPLICATOR = &H228
Const SE_GROUP_MANDATORY = &H1
Const SE_GROUP_ENABLED_BY_DEFAULT = &H2
Const SE_GROUP_ENABLED = &H4
Const SE_GROUP_OWNER = &H8
Const SE_GROUP_LOGON_ID = &HC0000000
Const FILE_BEGIN = 0
Const FILE_CURRENT = 1
Const FILE_END = 2
Const FILE_FLAG_WRITE_THROUGH = &H80000000
Const FILE_FLAG_OVERLAPPED = &H40000000
Const FILE_FLAG_NO_BUFFERING = &H20000000
Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
Const CREATE_NEW = 1
Const CREATE_ALWAYS = 2
Const OPEN_EXISTING = 3
Const OPEN_ALWAYS = 4
Const TRUNCATE_EXISTING = 5
Const PIPE_ACCESS_INBOUND = &H1
Const PIPE_ACCESS_OUTBOUND = &H2
Const PIPE_ACCESS_DUPLEX = &H3
Const PIPE_CLIENT_END = &H0
Const PIPE_SERVER_END = &H1
Const PIPE_WAIT = &H0
Const PIPE_NOWAIT = &H1
Const PIPE_READMODE_BYTE = &H0
Const PIPE_READMODE_MESSAGE = &H2
Const PIPE_TYPE_BYTE = &H0
Const PIPE_TYPE_MESSAGE = &H4
Const PIPE_UNLIMITED_INSTANCES = 255
Const SECURITY_CONTEXT_TRACKING = &H40000
Const SECURITY_EFFECTIVE_ONLY = &H80000
Const SECURITY_SQOS_PRESENT = &H100000
Const SECURITY_VALID_SQOS_FLAGS = &H1F0000
Const SP_SERIALCOMM = &H1&
Const PST_UNSPECIFIED = &H0&
Const PST_RS232 = &H1&
Const PST_PARALLELPORT = &H2&
Const PST_RS422 = &H3&
Const PST_RS423 = &H4&
Const PST_RS449 = &H5&
Const PST_FAX = &H21&
Const PST_SCANNER = &H22&
Const PST_NETWORK_BRIDGE = &H100&
Const PST_LAT = &H101&
Const PST_TCPIP_TELNET = &H102&
Const PST_X25 = &H103&
Const PCF_DTRDSR = &H1&
Const PCF_RTSCTS = &H2&
Const PCF_RLSD = &H4&
Const PCF_PARITY_CHECK = &H8&
Const PCF_XONXOFF = &H10&
Const PCF_SETXCHAR = &H20&
Const PCF_TOTALTIMEOUTS = &H40&
Const PCF_INTTIMEOUTS = &H80&
Const PCF_SPECIALCHARS = &H100&
Const PCF_16BITMODE = &H200&
Const SP_PARITY = &H1&
Const SP_BAUD = &H2&
Const SP_DATABITS = &H4&
Const SP_STOPBITS = &H8&
Const SP_HANDSHAKING = &H10&
Const SP_PARITY_CHECK = &H20&
Const SP_RLSD = &H40&
Const BAUD_075 = &H1&
Const BAUD_110 = &H2&
Const BAUD_134_5 = &H4&
Const BAUD_150 = &H8&
Const BAUD_300 = &H10&
Const BAUD_600 = &H20&
Const BAUD_1200 = &H40&
Const BAUD_1800 = &H80&
Const BAUD_2400 = &H100&
Const BAUD_4800 = &H200&
Const BAUD_7200 = &H400&
Const BAUD_9600 = &H800&
Const BAUD_14400 = &H1000&
Const BAUD_19200 = &H2000&
Const BAUD_38400 = &H4000&
Const BAUD_56K = &H8000&
Const BAUD_128K = &H10000
Const BAUD_115200 = &H20000
Const BAUD_57600 = &H40000
Const BAUD_USER = &H10000000
Const DATABITS_5 = &H1&
Const DATABITS_6 = &H2&
Const DATABITS_7 = &H4&
Const DATABITS_8 = &H8&
Const DATABITS_16 = &H10&
Const DATABITS_16X = &H20&
Const STOPBITS_10 = &H1&
Const STOPBITS_15 = &H2&
Const STOPBITS_20 = &H4&
Const PARITY_NONE = &H100&
Const PARITY_ODD = &H200&
Const PARITY_EVEN = &H400&
Const PARITY_MARK = &H800&
Const PARITY_SPACE = &H1000&
Const DTR_CONTROL_DISABLE = &H0
Const DTR_CONTROL_ENABLE = &H1
Const DTR_CONTROL_HANDSHAKE = &H2
Const RTS_CONTROL_DISABLE = &H0
Const RTS_CONTROL_ENABLE = &H1
Const RTS_CONTROL_HANDSHAKE = &H2
Const RTS_CONTROL_TOGGLE = &H3
Const GMEM_FIXED = &H0
Const GMEM_MOVEABLE = &H2
Const GMEM_NOCOMPACT = &H10
Const GMEM_NODISCARD = &H20
Const GMEM_ZEROINIT = &H40
Const GMEM_MODIFY = &H80
Const GMEM_DISCARDABLE = &H100
Const GMEM_NOT_BANKED = &H1000
Const GMEM_SHARE = &H2000
Const GMEM_DDESHARE = &H2000
Const GMEM_NOTIFY = &H4000
Const GMEM_LOWER = GMEM_NOT_BANKED
Const GMEM_VALID_FLAGS = &H7F72
Const GMEM_INVALID_HANDLE = &H8000
Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Const GMEM_DISCARDED = &H4000
Const GMEM_LOCKCOUNT = &HFF
Const LMEM_FIXED = &H0
Const LMEM_MOVEABLE = &H2
Const LMEM_NOCOMPACT = &H10
Const LMEM_NODISCARD = &H20
Const LMEM_ZEROINIT = &H40
Const LMEM_MODIFY = &H80
Const LMEM_DISCARDABLE = &HF00
Const LMEM_VALID_FLAGS = &HF72
Const LMEM_INVALID_HANDLE = &H8000
Const LHND = (LMEM_MOVEABLE + LMEM_ZEROINIT)
Const LPTR = (LMEM_FIXED + LMEM_ZEROINIT)
Const NONZEROLHND = (LMEM_MOVEABLE)
Const NONZEROLPTR = (LMEM_FIXED)
Const LMEM_DISCARDED = &H4000
Const LMEM_LOCKCOUNT = &HFF
Const DEBUG_PROCESS = &H1
Const DEBUG_ONLY_THIS_PROCESS = &H2
Const CREATE_SUSPENDED = &H4
Const DETACHED_PROCESS = &H8
Const CREATE_NEW_CONSOLE = &H10
Const NORMAL_PRIORITY_CLASS = &H20
Const IDLE_PRIORITY_CLASS = &H40
Const HIGH_PRIORITY_CLASS = &H80
Const REALTIME_PRIORITY_CLASS = &H100
Const CREATE_NEW_PROCESS_GROUP = &H200
Const CREATE_NO_WINDOW = &H8000000
Const PROFILE_USER = &H10000000
Const PROFILE_KERNEL = &H20000000
Const PROFILE_SERVER = &H40000000
Const MAXLONG = &H7FFFFFFF
Const THREAD_BASE_PRIORITY_MIN = -2
Const THREAD_BASE_PRIORITY_MAX = 2
Const THREAD_BASE_PRIORITY_LOWRT = 15
Const THREAD_BASE_PRIORITY_IDLE = -15
Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
Const THREAD_PRIORITY_NORMAL = 0
Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Const THREAD_PRIORITY_ERROR_RETURN = (MAXLONG)
Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
Const APPLICATION_ERROR_MASK = &H20000000
Const ERROR_SEVERITY_SUCCESS = &H0
Const ERROR_SEVERITY_INFORMATIONAL = &H40000000
Const ERROR_SEVERITY_WARNING = &H80000000
Const ERROR_SEVERITY_ERROR = &HC0000000
Const MINCHAR = &H80
Const MAXCHAR = &H7F
Const MINSHORT = &H8000
Const MAXSHORT = &H7FFF
Const MINLONG = &H80000000
Const MAXByte = &HFF
Const MAXWORD = &HFFFF
Const MAXDWORD = &HFFFF
Const LANG_NEUTRAL = &H0
Const LANG_BULGARIAN = &H2
Const LANG_CHINESE = &H4
Const LANG_CROATIAN = &H1A
Const LANG_CZECH = &H5
Const LANG_DANISH = &H6
Const LANG_DUTCH = &H13
Const LANG_ENGLISH = &H9
Const LANG_FINNISH = &HB
Const LANG_FRENCH = &HC
Const LANG_GERMAN = &H7
Const LANG_GREEK = &H8
Const LANG_HUNGARIAN = &HE
Const LANG_ICELANDIC = &HF
Const LANG_ITALIAN = &H10
Const LANG_JAPANESE = &H11
Const LANG_KOREAN = &H12
Const LANG_NORWEGIAN = &H14
Const LANG_POLISH = &H15
Const LANG_PORTUGUESE = &H16
Const LANG_ROMANIAN = &H18
Const LANG_RUSSIAN = &H19
Const LANG_SLOVAK = &H1B
Const LANG_SLOVENIAN = &H24
Const LANG_SPANISH = &HA
Const LANG_SWEDISH = &H1D
Const LANG_TURKISH = &H1F
Const SUBLANG_NEUTRAL = &H0                       '  language neutral
Const SUBLANG_DEFAULT = &H1                       '  user default
Const SUBLANG_SYS_DEFAULT = &H2                   '  system default
Const SUBLANG_CHINESE_TRADITIONAL = &H1           '  Chinese (Taiwan)
Const SUBLANG_CHINESE_SIMPLIFIED = &H2            '  Chinese (PR China)
Const SUBLANG_CHINESE_HONGKONG = &H3              '  Chinese (Hong Kong)
Const SUBLANG_CHINESE_SINGAPORE = &H4             '  Chinese (Singapore)
Const SUBLANG_DUTCH = &H1                         '  Dutch
Const SUBLANG_DUTCH_BELGIAN = &H2                 '  Dutch (Belgian)
Const SUBLANG_ENGLISH_US = &H1                    '  English (USA)
Const SUBLANG_ENGLISH_UK = &H2                    '  English (UK)
Const SUBLANG_ENGLISH_AUS = &H3                   '  English (Australian)
Const SUBLANG_ENGLISH_CAN = &H4                   '  English (Canadian)
Const SUBLANG_ENGLISH_NZ = &H5                    '  English (New Zealand)
Const SUBLANG_ENGLISH_EIRE = &H6                  '  English (Irish)
Const SUBLANG_FRENCH = &H1                        '  French
Const SUBLANG_FRENCH_BELGIAN = &H2                '  French (Belgian)
Const SUBLANG_FRENCH_CANADIAN = &H3               '  French (Canadian)
Const SUBLANG_FRENCH_SWISS = &H4                  '  French (Swiss)
Const SUBLANG_GERMAN = &H1                        '  German
Const SUBLANG_GERMAN_SWISS = &H2                  '  German (Swiss)
Const SUBLANG_GERMAN_AUSTRIAN = &H3               '  German (Austrian)
Const SUBLANG_ITALIAN = &H1                       '  Italian
Const SUBLANG_ITALIAN_SWISS = &H2                 '  Italian (Swiss)
Const SUBLANG_NORWEGIAN_BOKMAL = &H1              '  Norwegian (Bokma
Const SUBLANG_NORWEGIAN_NYNORSK = &H2             '  Norwegian (Nynorsk)
Const SUBLANG_PORTUGUESE = &H2                    '  Portuguese
Const SUBLANG_PORTUGUESE_BRAZILIAN = &H1          '  Portuguese (Brazilian)
Const SUBLANG_SPANISH = &H1                       '  Spanish (Castilian)
Const SUBLANG_SPANISH_MEXICAN = &H2               '  Spanish (Mexican)
Const SUBLANG_SPANISH_MODERN = &H3                '  Spanish (Modern)
Const SORT_DEFAULT = &H0                          '  sorting default
Const SORT_JAPANESE_XJIS = &H0                    '  Japanese0xJIS order
Const SORT_JAPANESE_UNICODE = &H1                 '  Japanese Unicode order
Const SORT_CHINESE_BIG5 = &H0                     '  Chinese BIG5 order
Const SORT_CHINESE_UNICODE = &H1                  '  Chinese Unicode order
Const SORT_KOREAN_KSC = &H0                       '  Korean KSC order
Const SORT_KOREAN_UNICODE = &H1                   '  Korean Unicode order
Const FILE_READ_DATA = (&H1)                     '  file pipe
Const FILE_LIST_DIRECTORY = (&H1)                '  directory
Const FILE_WRITE_DATA = (&H2)                    '  file pipe
Const FILE_ADD_FILE = (&H2)                      '  directory
Const FILE_APPEND_DATA = (&H4)                   '  file
Const FILE_ADD_SUBDIRECTORY = (&H4)              '  directory
Const FILE_CREATE_PIPE_INSTANCE = (&H4)          '  named pipe
Const FILE_READ_EA = (&H8)                       '  file directory
Const FILE_READ_PROPERTIES = FILE_READ_EA
Const FILE_WRITE_EA = (&H10)                     '  file directory
Const FILE_WRITE_PROPERTIES = FILE_WRITE_EA
Const FILE_EXECUTE = (&H20)                      '  file
Const FILE_TRAVERSE = (&H20)                     '  directory
Const FILE_DELETE_CHILD = (&H40)                 '  directory
Const FILE_READ_ATTRIBUTES = (&H80)              '  all
Const FILE_WRITE_ATTRIBUTES = (&H100)            '  all
Const FILE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H1FF)
Const FILE_GENERIC_READ = (STANDARD_RIGHTS_READ Or FILE_READ_DATA Or FILE_READ_ATTRIBUTES Or FILE_READ_EA Or SYNCHRONIZE)
Const FILE_GENERIC_WRITE = (STANDARD_RIGHTS_WRITE Or FILE_WRITE_DATA Or FILE_WRITE_ATTRIBUTES Or FILE_WRITE_EA Or FILE_APPEND_DATA Or SYNCHRONIZE)
Const FILE_GENERIC_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or FILE_READ_ATTRIBUTES Or FILE_EXECUTE Or SYNCHRONIZE)
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const FILE_ATTRIBUTE_COMPRESSED = &H800
Const FILE_NOTIFY_CHANGE_FILE_NAME = &H1
Const FILE_NOTIFY_CHANGE_DIR_NAME = &H2
Const FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4
Const FILE_NOTIFY_CHANGE_SIZE = &H8
Const FILE_NOTIFY_CHANGE_LAST_WRITE = &H10
Const FILE_NOTIFY_CHANGE_SECURITY = &H100
Const MAILSLOT_NO_MESSAGE = (-1)
Const MAILSLOT_WAIT_FOREVER = (-1)
Const FILE_CASE_SENSITIVE_SEARCH = &H1
Const FILE_CASE_PRESERVED_NAMES = &H2
Const FILE_UNICODE_ON_DISK = &H4
Const FILE_PERSISTENT_ACLS = &H8
Const FILE_FILE_COMPRESSION = &H10
Const FILE_VOLUME_IS_COMPRESSED = &H8000
Const IO_COMPLETION_MODIFY_STATE = &H2
Const IO_COMPLETION_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3)
Const DUPLICATE_CLOSE_SOURCE = &H1
Const DUPLICATE_SAME_ACCESS = &H2
Const ACCESS_SYSTEM_SECURITY = &H1000000
Const MAXIMUM_ALLOWED = &H2000000
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Const GENERIC_EXECUTE = &H20000000
Const GENERIC_ALL = &H10000000
Const ACL_REVISION = (2)
Const ACL_REVISION1 = (1)
Const ACL_REVISION2 = (2)
Const ACCESS_ALLOWED_ACE_TYPE = &H0
Const ACCESS_DENIED_ACE_TYPE = &H1
Const SYSTEM_AUDIT_ACE_TYPE = &H2
Const SYSTEM_ALARM_ACE_TYPE = &H3
Const OBJECT_INHERIT_ACE = &H1
Const CONTAINER_INHERIT_ACE = &H2
Const NO_PROPAGATE_INHERIT_ACE = &H4
Const INHERIT_ONLY_ACE = &H8
Const VALID_INHERIT_FLAGS = &HF
Const SUCCESSFUL_ACCESS_ACE_FLAG = &H40
Const FAILED_ACCESS_ACE_FLAG = &H80
Const AclRevisionInformation = 1
Const AclSizeInformation = 2
Const SECURITY_DESCRIPTOR_REVISION = (1)
Const SECURITY_DESCRIPTOR_REVISION1 = (1)
Const SECURITY_DESCRIPTOR_MIN_LENGTH = (20)
Const SE_OWNER_DEFAULTED = &H1
Const SE_GROUP_DEFAULTED = &H2
Const SE_DACL_PRESENT = &H4
Const SE_DACL_DEFAULTED = &H8
Const SE_SACL_PRESENT = &H10
Const SE_SACL_DEFAULTED = &H20
Const SE_SELF_RELATIVE = &H8000
Const SE_PRIVILEGE_ENABLED_BY_DEFAULT = &H1
Const SE_PRIVILEGE_ENABLED = &H2
Const SE_PRIVILEGE_USED_FOR_ACCESS = &H80000000
Const PRIVILEGE_SET_ALL_NECESSARY = (1)
Const SE_CREATE_TOKEN_NAME = "SeCreateTokenPrivilege"
Const SE_ASSIGNPRIMARYTOKEN_NAME = "SeAssignPrimaryTokenPrivilege"
Const SE_LOCK_MEMORY_NAME = "SeLockMemoryPrivilege"
Const SE_INCREASE_QUOTA_NAME = "SeIncreaseQuotaPrivilege"
Const SE_UNSOLICITED_INPUT_NAME = "SeUnsolicitedInputPrivilege"
Const SE_MACHINE_ACCOUNT_NAME = "SeMachineAccountPrivilege"
Const SE_TCB_NAME = "SeTcbPrivilege"
Const SE_SECURITY_NAME = "SeSecurityPrivilege"
Const SE_TAKE_OWNERSHIP_NAME = "SeTakeOwnershipPrivilege"
Const SE_LOAD_DRIVER_NAME = "SeLoadDriverPrivilege"
Const SE_SYSTEM_PROFILE_NAME = "SeSystemProfilePrivilege"
Const SE_SYSTEMTIME_NAME = "SeSystemtimePrivilege"
Const SE_PROF_SINGLE_PROCESS_NAME = "SeProfileSingleProcessPrivilege"
Const SE_INC_BASE_PRIORITY_NAME = "SeIncreaseBasePriorityPrivilege"
Const SE_CREATE_PAGEFILE_NAME = "SeCreatePagefilePrivilege"
Const SE_CREATE_PERMANENT_NAME = "SeCreatePermanentPrivilege"
Const SE_BACKUP_NAME = "SeBackupPrivilege"
Const SE_RESTORE_NAME = "SeRestorePrivilege"
Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Const SE_DEBUG_NAME = "SeDebugPrivilege"
Const SE_AUDIT_NAME = "SeAuditPrivilege"
Const SE_SYSTEM_ENVIRONMENT_NAME = "SeSystemEnvironmentPrivilege"
Const SE_CHANGE_NOTIFY_NAME = "SeChangeNotifyPrivilege"
Const SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege"
Const SecurityAnonymous = 1
Const SecurityIdentification = 2
Const REG_NONE = 0                       ' No value type
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Const REG_BINARY = 3                     ' Free form binary
Const REG_DWORD = 4                      ' 32-bit number
Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Const REG_LINK = 6                       ' Symbolic Link (unicode)
Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Const REG_RESOURCE_LIST = 8              ' Resource list in the resource map
Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Const REG_CREATED_NEW_KEY = &H1                      ' New Registry Key created
Const REG_OPENED_EXISTING_KEY = &H2                      ' Existing Key opened
Const REG_WHOLE_HIVE_VOLATILE = &H1                      ' Restore whole hive volatile
Const REG_REFRESH_HIVE = &H2                      ' Unwind changes to last flush
Const REG_NOTIFY_CHANGE_NAME = &H1                      ' Create or delete (child)
Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Const REG_NOTIFY_CHANGE_LAST_SET = &H4                      ' Time stamp
Const REG_NOTIFY_CHANGE_SECURITY = &H8
Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Const EXCEPTION_DEBUG_EVENT = 1
Const CREATE_THREAD_DEBUG_EVENT = 2
Const CREATE_PROCESS_DEBUG_EVENT = 3
Const EXIT_THREAD_DEBUG_EVENT = 4
Const EXIT_PROCESS_DEBUG_EVENT = 5
Const LOAD_DLL_DEBUG_EVENT = 6
Const UNLOAD_DLL_DEBUG_EVENT = 7
Const OUTPUT_DEBUG_STRING_EVENT = 8
Const RIP_EVENT = 9
Const EXCEPTION_MAXIMUM_PARAMETERS = 15
Const DRIVE_REMOVABLE = 2
Const DRIVE_FIXED = 3
Const DRIVE_REMOTE = 4
Const DRIVE_CDROM = 5
Const DRIVE_RAMDISK = 6
Const FILE_TYPE_UNKNOWN = &H0
Const FILE_TYPE_DISK = &H1
Const FILE_TYPE_CHAR = &H2
Const FILE_TYPE_PIPE = &H3
Const FILE_TYPE_REMOTE = &H8000
Const STD_INPUT_HANDLE = -10&
Const STD_OUTPUT_HANDLE = -11&
Const STD_ERROR_HANDLE = -12&
Const NOPARITY = 0
Const ODDPARITY = 1
Const EVENPARITY = 2
Const MARKPARITY = 3
Const SPACEPARITY = 4
Const ONESTOPBIT = 0
Const ONE5STOPBITS = 1
Const TWOSTOPBITS = 2
Const IGNORE = 0    '  Ignore signal
Const INFINITE = &HFFFF      '  Infinite timeout
Const CBR_110 = 110
Const CBR_300 = 300
Const CBR_600 = 600
Const CBR_1200 = 1200
Const CBR_2400 = 2400
Const CBR_4800 = 4800
Const CBR_9600 = 9600
Const CBR_14400 = 14400
Const CBR_19200 = 19200
Const CBR_38400 = 38400
Const CBR_56000 = 56000
Const CBR_57600 = 57600
Const CBR_115200 = 115200
Const CBR_128000 = 128000
Const CBR_256000 = 256000
Const CE_RXOVER = &H1                '  Receive Queue overflow
Const CE_OVERRUN = &H2               '  Receive Overrun Error
Const CE_RXPARITY = &H4              '  Receive Parity Error
Const CE_FRAME = &H8                 '  Receive Framing error
Const CE_BREAK = &H10                '  Break Detected
Const CE_TXFULL = &H100              '  TX Queue is full
Const CE_PTO = &H200                 '  LPTx Timeout
Const CE_IOE = &H400                 '  LPTx I/O Error
Const CE_DNS = &H800                 '  LPTx Device not selected
Const CE_OOP = &H1000                '  LPTx Out-Of-Paper
Const CE_MODE = &H8000               '  Requested mode unsupported
Const IE_BADID = (-1)                '  Invalid or unsupported id
Const IE_OPEN = (-2)                 '  Device Already Open
Const IE_NOPEN = (-3)                '  Device Not Open
Const IE_MEMORY = (-4)               '  Unable to allocate queues
Const IE_DEFAULT = (-5)              '  Error in default parameters
Const IE_HARDWARE = (-10)            '  Hardware Not Present
Const IE_BYTESIZE = (-11)            '  Illegal Byte Size
Const IE_BAUDRATE = (-12)            '  Unsupported BaudRate
Const EV_RXCHAR = &H1                '  Any Character received
Const EV_RXFLAG = &H2                '  Received certain character
Const EV_TXEMPTY = &H4               '  Transmitt Queue Empty
Const EV_CTS = &H8                   '  CTS changed state
Const EV_DSR = &H10                  '  DSR changed state
Const EV_RLSD = &H20                 '  RLSD changed state
Const EV_BREAK = &H40                '  BREAK received
Const EV_ERR = &H80                  '  Line status error occurred
Const EV_RING = &H100                '  Ring signal detected
Const EV_PERR = &H200                '  Printer error occured
Const EV_RX80FULL = &H400            '  Receive buffer is 80 percent full
Const EV_EVENT1 = &H800              '  Provider specific event 1
Const EV_EVENT2 = &H1000             '  Provider specific event 2
Const SETXOFF = 1  '  Simulate XOFF received
Const SETXON = 2    '  Simulate XON received
Const SETRTS = 3    '  Set RTS high
Const CLRRTS = 4    '  Set RTS low
Const SETDTR = 5    '  Set DTR high
Const CLRDTR = 6    '  Set DTR low
Const RESETDEV = 7       '  Reset device if possible
Const SETBREAK = 8  'Set the device break line
Const CLRBREAK = 9    ' Clear the device break line
Const PURGE_TXABORT = &H1     '  Kill the pending/current writes to the comm port.
Const PURGE_RXABORT = &H2     '  Kill the pending/current reads to the comm port.
Const PURGE_TXCLEAR = &H4     '  Kill the transmit queue if there.
Const PURGE_RXCLEAR = &H8     '  Kill the typeahead buffer if there.
Const LPTx = &H80        '  Set if ID is for LPT device
Const MS_CTS_ON = &H10&
Const MS_DSR_ON = &H20&
Const MS_RING_ON = &H40&
Const MS_RLSD_ON = &H80&
Const S_QUEUEEMPTY = 0
Const S_THRESHOLD = 1
Const S_ALLTHRESHOLD = 2
Const S_NORMAL = 0
Const S_LEGATO = 1
Const S_STACCATO = 2
Const S_PERIOD512 = 0    '  Freq = N/512 high pitch, less coarse hiss
Const S_PERIOD1024 = 1   '  Freq = N/1024
Const S_PERIOD2048 = 2   '  Freq = N/2048 low pitch, more coarse hiss
Const S_PERIODVOICE = 3  '  Source is frequency from voice channel (3)
Const S_WHITE512 = 4     '  Freq = N/512 high pitch, less coarse hiss
Const S_WHITE1024 = 5    '  Freq = N/1024
Const S_WHITE2048 = 6    '  Freq = N/2048 low pitch, more coarse hiss
Const S_WHITEVOICE = 7   '  Source is frequency from voice channel (3)
Const S_SERDVNA = (-1)   '  Device not available
Const S_SEROFM = (-2)    '  Out of memory
Const S_SERMACT = (-3)   '  Music active
Const S_SERQFUL = (-4)   '  Queue full
Const S_SERBDNT = (-5)   '  Invalid note
Const S_SERDLN = (-6)    '  Invalid note length
Const S_SERDCC = (-7)    '  Invalid note count
Const S_SERDTP = (-8)    '  Invalid tempo
Const S_SERDVL = (-9)    '  Invalid volume
Const S_SERDMD = (-10)   '  Invalid mode
Const S_SERDSH = (-11)   '  Invalid shape
Const S_SERDPT = (-12)   '  Invalid pitch
Const S_SERDFQ = (-13)   '  Invalid frequency
Const S_SERDDR = (-14)   '  Invalid duration
Const S_SERDSR = (-15)   '  Invalid source
Const S_SERDST = (-16)   '  Invalid state
Const NMPWAIT_WAIT_FOREVER = &HFFFF
Const NMPWAIT_NOWAIT = &H1
Const NMPWAIT_USE_DEFAULT_WAIT = &H0
Const FS_CASE_IS_PRESERVED = FILE_CASE_PRESERVED_NAMES
Const FS_CASE_SENSITIVE = FILE_CASE_SENSITIVE_SEARCH
Const FS_UNICODE_STORED_ON_DISK = FILE_UNICODE_ON_DISK
Const FS_PERSISTENT_ACLS = FILE_PERSISTENT_ACLS
Const SECTION_QUERY = &H1
Const SECTION_MAP_WRITE = &H2
Const SECTION_MAP_READ = &H4
Const SECTION_MAP_EXECUTE = &H8
Const SECTION_EXTEND_SIZE = &H10
Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
Const FILE_MAP_COPY = SECTION_QUERY
Const FILE_MAP_WRITE = SECTION_MAP_WRITE
Const FILE_MAP_READ = SECTION_MAP_READ
Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS
Const OF_READ = &H0
Const OF_WRITE = &H1
Const OF_READWRITE = &H2
Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40
Const OF_PARSE = &H100
Const OF_DELETE = &H200
Const OF_VERIFY = &H400
Const OF_CANCEL = &H800
Const OF_CREATE = &H1000
Const OF_PROMPT = &H2000
Const OF_EXIST = &H4000
Const OF_REOPEN = &H8000
Const OFS_MAXPATHNAME = 128
Const PROCESSOR_INTEL_386 = 386
Const PROCESSOR_INTEL_486 = 486
Const PROCESSOR_INTEL_PENTIUM = 586
Const PROCESSOR_MIPS_R4000 = 4000
Const PROCESSOR_ALPHA_21064 = 21064
Const PROCESSOR_ARCHITECTURE_INTEL = 0
Const PROCESSOR_ARCHITECTURE_MIPS = 1
Const PROCESSOR_ARCHITECTURE_ALPHA = 2
Const PROCESSOR_ARCHITECTURE_PPC = 3
Const PROCESSOR_ARCHITECTURE_UNKNOWN = &HFFFF
Const DFC_CAPTION = 1
Const DFC_MENU = 2
Const DFC_SCROLL = 3
Const DFC_BUTTON = 4
Const DFCS_CAPTIONCLOSE = &H0
Const DFCS_CAPTIONMIN = &H1
Const DFCS_CAPTIONMAX = &H2
Const DFCS_CAPTIONRESTORE = &H3
Const DFCS_CAPTIONHELP = &H4
Const DFCS_MENUARROW = &H0
Const DFCS_MENUCHECK = &H1
Const DFCS_MENUBULLET = &H2
Const DFCS_MENUARROWRIGHT = &H4
Const DFCS_SCROLLUP = &H0
Const DFCS_SCROLLDOWN = &H1
Const DFCS_SCROLLLEFT = &H2
Const DFCS_SCROLLRIGHT = &H3
Const DFCS_SCROLLCOMBOBOX = &H5
Const DFCS_SCROLLSIZEGRIP = &H8
Const DFCS_SCROLLSIZEGRIPRIGHT = &H10
Const DFCS_BUTTONCHECK = &H0
Const DFCS_BUTTONRADIOIMAGE = &H1
Const DFCS_BUTTONRADIOMASK = &H2
Const DFCS_BUTTONRADIO = &H4
Const DFCS_BUTTON3STATE = &H8
Const DFCS_BUTTONPUSH = &H10
Const DFCS_INACTIVE = &H100
Const DFCS_PUSHED = &H200
Const DFCS_CHECKED = &H400
Const DFCS_ADJUSTRECT = &H2000
Const DFCS_FLAT = &H4000
Const DFCS_MONO = &H8000
Const DONT_RESOLVE_DLL_REFERENCES = &H1
Const TF_FORCEDRIVE = &H80
Const LOCKFILE_FAIL_IMMEDIATELY = &H1
Const LOCKFILE_EXCLUSIVE_LOCK = &H2
Const LNOTIFY_OUTOFMEM = 0
Const LNOTIFY_MOVE = 1
Const LNOTIFY_DISCARD = 2
Const SLE_ERROR = &H1
Const SLE_MINORERROR = &H2
Const SLE_WARNING = &H3
Const SEM_FAILCRITICALERRORS = &H1
Const SEM_NOGPFAULTERRORBOX = &H2
Const SEM_NOOPENFILEERRORBOX = &H8000
Const RT_CURSOR = 1&
Const RT_BITMAP = 2&
Const RT_ICON = 3&
Const RT_MENU = 4&
Const RT_DIALOG = 5&
Const RT_STRING = 6&
Const RT_FONTDIR = 7&
Const RT_FONT = 8&
Const RT_ACCELERATOR = 9&
Const RT_RCDATA = 10&
Const DDD_RAW_TARGET_PATH = &H1
Const DDD_REMOVE_DEFINITION = &H2
Const DDD_EXACT_MATCH_ON_REMOVE = &H4
Const MAX_PATH = 260
Const MOVEFILE_REPLACE_EXISTING = &H1
Const MOVEFILE_COPY_ALLOWED = &H2
Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4
Const TokenUser = 1
Const TokenGroups = 2
Const TokenPrivileges = 3
Const TokenOwner = 4
Const TokenPrimaryGroup = 5
Const TokenDefaultDacl = 6
Const TokenSource = 7
Const TokenType = 8
Const TokenImpersonationLevel = 9
Const TokenStatistics = 10
Const GET_TAPE_MEDIA_INFORMATION = 0
Const GET_TAPE_DRIVE_INFORMATION = 1
Const SET_TAPE_MEDIA_INFORMATION = 0
Const SET_TAPE_DRIVE_INFORMATION = 1
Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Const FORMAT_MESSAGE_FROM_STRING = &H400
Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Const TLS_OUT_OF_INDEXES = &HFFFF
Const BACKUP_DATA = &H1
Const BACKUP_EA_DATA = &H2
Const BACKUP_SECURITY_DATA = &H3
Const BACKUP_ALTERNATE_DATA = &H4
Const BACKUP_LINK = &H5
Const STREAM_MODIFIED_WHEN_READ = &H1
Const STREAM_CONTAINS_SECURITY = &H2
Const STARTF_USESHOWWINDOW = &H1
Const STARTF_USESIZE = &H2
Const STARTF_USEPOSITION = &H4
Const STARTF_USECOUNTCHARS = &H8
Const STARTF_USEFILLATTRIBUTE = &H10
Const STARTF_RUNFULLSCREEN = &H20        '  ignored for non-x86 platforms
Const STARTF_FORCEONFEEDBACK = &H40
Const STARTF_FORCEOFFFEEDBACK = &H80
Const STARTF_USESTDHANDLES = &H100
Const SHUTDOWN_NORETRY = &H1
Const TC_NORMAL = 0
Const TC_HARDERR = 1
Const TC_GP_TRAP = 2
Const TC_SIGNAL = 3
Const MAX_LEADBYTES = 12  '  5 ranges, 2 bytes ea., 0 term.
Const MB_PRECOMPOSED = &H1         '  use precomposed chars
Const MB_COMPOSITE = &H2         '  use composite chars
Const MB_USEGLYPHCHARS = &H4         '  use glyph chars, not ctrl chars
Const WC_DEFAULTCHECK = &H100       '  check for default char
Const WC_COMPOSITECHECK = &H200       '  convert composite to precomposed
Const WC_DISCARDNS = &H10        '  discard non-spacing chars
Const WC_SEPCHARS = &H20        '  generate separate chars
Const WC_DEFAULTCHAR = &H40        '  replace w/ default char
Const CT_CTYPE1 = &H1         '  ctype 1 information
Const CT_CTYPE2 = &H2         '  ctype 2 information
Const CT_CTYPE3 = &H4         '  ctype 3 information
Const C1_UPPER = &H1     '  upper case
Const C1_LOWER = &H2     '  lower case
Const C1_DIGIT = &H4     '  decimal digits
Const C1_SPACE = &H8     '  spacing characters
Const C1_PUNCT = &H10    '  punctuation characters
Const C1_CNTRL = &H20    '  control characters
Const C1_BLANK = &H40    '  blank characters
Const C1_XDIGIT = &H80    '  other digits
Const C1_ALPHA = &H100   '  any letter
Const C2_LEFTTORIGHT = &H1     '  left to right
Const C2_RIGHTTOLEFT = &H2     '  right to left
Const C2_EUROPENUMBER = &H3     '  European number, digit
Const C2_EUROPESEPARATOR = &H4     '  European numeric separator
Const C2_EUROPETERMINATOR = &H5     '  European numeric terminator
Const C2_ARABICNUMBER = &H6     '  Arabic number
Const C2_COMMONSEPARATOR = &H7     '  common numeric separator
Const C2_BLOCKSEPARATOR = &H8     '  block separator
Const C2_SEGMENTSEPARATOR = &H9     '  segment separator
Const C2_WHITESPACE = &HA     '  white space
Const C2_OTHERNEUTRAL = &HB     '  other neutrals
Const C2_NOTAPPLICABLE = &H0     '  no implicit directionality
Const C3_NONSPACING = &H1     '  nonspacing character
Const C3_DIACRITIC = &H2     '  diacritic mark
Const C3_VOWELMARK = &H4     '  vowel mark
Const C3_SYMBOL = &H8     '  symbols
Const C3_NOTAPPLICABLE = &H0     '  ctype 3 is not applicable
Const NORM_IGNORECASE = &H1         '  ignore case
Const NORM_IGNORENONSPACE = &H2         '  ignore nonspacing chars
Const NORM_IGNORESYMBOLS = &H4         '  ignore symbols
Const MAP_FOLDCZONE = &H10        '  fold compatibility zone chars
Const MAP_PRECOMPOSED = &H20        '  convert to precomposed chars
Const MAP_COMPOSITE = &H40        '  convert to composite chars
Const MAP_FOLDDIGITS = &H80        '  all digits to ASCII 0-9
Const LCMAP_LOWERCASE = &H100       '  lower case letters
Const LCMAP_UPPERCASE = &H200       '  upper case letters
Const LCMAP_SORTKEY = &H400       '  WC sort key (normalize)
Const LCMAP_BYTEREV = &H800       '  byte reversal
Const SORT_STRINGSORT = &H1000      '  use string sort method
Const CP_ACP = 0  '  default to ANSI code page
Const CP_OEMCP = 1  '  default to OEM  code page
Const CTRY_DEFAULT = 0
Const CTRY_AUSTRALIA = 61  '  Australia
Const CTRY_AUSTRIA = 43  '  Austria
Const CTRY_BELGIUM = 32  '  Belgium
Const CTRY_BRAZIL = 55  '  Brazil
Const CTRY_CANADA = 2  '  Canada
Const CTRY_DENMARK = 45  '  Denmark
Const CTRY_FINLAND = 358  '  Finland
Const CTRY_FRANCE = 33  '  France
Const CTRY_GERMANY = 49  '  Germany
Const CTRY_ICELAND = 354  '  Iceland
Const CTRY_IRELAND = 353  '  Ireland
Const CTRY_ITALY = 39  '  Italy
Const CTRY_JAPAN = 81  '  Japan
Const CTRY_MEXICO = 52  '  Mexico
Const CTRY_NETHERLANDS = 31  '  Netherlands
Const CTRY_NEW_ZEALAND = 64  '  New Zealand
Const CTRY_NORWAY = 47  '  Norway
Const CTRY_PORTUGAL = 351  '  Portugal
Const CTRY_PRCHINA = 86  '  PR China
Const CTRY_SOUTH_KOREA = 82  '  South Korea
Const CTRY_SPAIN = 34  '  Spain
Const CTRY_SWEDEN = 46  '  Sweden
Const CTRY_SWITZERLAND = 41  '  Switzerland
Const CTRY_TAIWAN = 886  '  Taiwan
Const CTRY_UNITED_KINGDOM = 44  '  United Kingdom
Const CTRY_UNITED_STATES = 1  '  United States
Const LOCALE_NOUSEROVERRIDE = &H80000000  '  do not use user overrides
Const LOCALE_ILANGUAGE = &H1         '  language id
Const LOCALE_SLANGUAGE = &H2         '  localized name of language
Const LOCALE_SENGLANGUAGE = &H1001      '  English name of language
Const LOCALE_SABBREVLANGNAME = &H3         '  abbreviated language name
Const LOCALE_SNATIVELANGNAME = &H4         '  native name of language
Const LOCALE_ICOUNTRY = &H5         '  country/region code
Const LOCALE_SCOUNTRY = &H6         '  localized name of country/region
Const LOCALE_SENGCOUNTRY = &H1002      '  English name of country/region
Const LOCALE_SABBREVCTRYNAME = &H7         '  abbreviated country/region name
Const LOCALE_SNATIVECTRYNAME = &H8         '  native name of country/region
Const LOCALE_IDEFAULTLANGUAGE = &H9         '  default language id
Const LOCALE_IDEFAULTCOUNTRY = &HA         '  default country/region code
Const LOCALE_IDEFAULTCODEPAGE = &HB         '  default code page
Const LOCALE_SLIST = &HC         '  list item separator
Const LOCALE_IMEASURE = &HD         '  0 = metric, 1 = US
Const LOCALE_SDECIMAL = &HE         '  decimal separator
Const LOCALE_STHOUSAND = &HF         '  thousand separator
Const LOCALE_SGROUPING = &H10        '  digit grouping
Const LOCALE_IDIGITS = &H11        '  number of fractional digits
Const LOCALE_ILZERO = &H12        '  leading zeros for decimal
Const LOCALE_SNATIVEDIGITS = &H13        '  native ascii 0-9
Const LOCALE_SCURRENCY = &H14        '  local monetary symbol
Const LOCALE_SINTLSYMBOL = &H15        '  intl monetary symbol
Const LOCALE_SMONDECIMALSEP = &H16        '  monetary decimal separator
Const LOCALE_SMONTHOUSANDSEP = &H17        '  monetary thousand separator
Const LOCALE_SMONGROUPING = &H18        '  monetary grouping
Const LOCALE_ICURRDIGITS = &H19        '  # local monetary digits
Const LOCALE_IINTLCURRDIGITS = &H1A        '  # intl monetary digits
Const LOCALE_ICURRENCY = &H1B        '  positive currency mode
Const LOCALE_INEGCURR = &H1C        '  negative currency mode
Const LOCALE_SDATE = &H1D        '  date separator
Const LOCALE_STIME = &H1E        '  time separator
Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Const LOCALE_SLONGDATE = &H20        '  long date format string
Const LOCALE_STIMEFORMAT = &H1003      '  time format string
Const LOCALE_IDATE = &H21        '  short date format ordering
Const LOCALE_ILDATE = &H22        '  long date format ordering
Const LOCALE_ITIME = &H23        '  time format specifier
Const LOCALE_ICENTURY = &H24        '  century format specifier
Const LOCALE_ITLZERO = &H25        '  leading zeros in time field
Const LOCALE_IDAYLZERO = &H26        '  leading zeros in day field
Const LOCALE_IMONLZERO = &H27        '  leading zeros in month field
Const LOCALE_S1159 = &H28        '  AM designator
Const LOCALE_S2359 = &H29        '  PM designator
Const LOCALE_SDAYNAME1 = &H2A        '  long name for Monday
Const LOCALE_SDAYNAME2 = &H2B        '  long name for Tuesday
Const LOCALE_SDAYNAME3 = &H2C        '  long name for Wednesday
Const LOCALE_SDAYNAME4 = &H2D        '  long name for Thursday
Const LOCALE_SDAYNAME5 = &H2E        '  long name for Friday
Const LOCALE_SDAYNAME6 = &H2F        '  long name for Saturday
Const LOCALE_SDAYNAME7 = &H30        '  long name for Sunday
Const LOCALE_SABBREVDAYNAME1 = &H31        '  abbreviated name for Monday
Const LOCALE_SABBREVDAYNAME2 = &H32        '  abbreviated name for Tuesday
Const LOCALE_SABBREVDAYNAME3 = &H33        '  abbreviated name for Wednesday
Const LOCALE_SABBREVDAYNAME4 = &H34        '  abbreviated name for Thursday
Const LOCALE_SABBREVDAYNAME5 = &H35        '  abbreviated name for Friday
Const LOCALE_SABBREVDAYNAME6 = &H36        '  abbreviated name for Saturday
Const LOCALE_SABBREVDAYNAME7 = &H37        '  abbreviated name for Sunday
Const LOCALE_SMONTHNAME1 = &H38        '  long name for January
Const LOCALE_SMONTHNAME2 = &H39        '  long name for February
Const LOCALE_SMONTHNAME3 = &H3A        '  long name for March
Const LOCALE_SMONTHNAME4 = &H3B        '  long name for April
Const LOCALE_SMONTHNAME5 = &H3C        '  long name for May
Const LOCALE_SMONTHNAME6 = &H3D        '  long name for June
Const LOCALE_SMONTHNAME7 = &H3E        '  long name for July
Const LOCALE_SMONTHNAME8 = &H3F        '  long name for August
Const LOCALE_SMONTHNAME9 = &H40        '  long name for September
Const LOCALE_SMONTHNAME10 = &H41        '  long name for October
Const LOCALE_SMONTHNAME11 = &H42        '  long name for November
Const LOCALE_SMONTHNAME12 = &H43        '  long name for December
Const LOCALE_SABBREVMONTHNAME1 = &H44        '  abbreviated name for January
Const LOCALE_SABBREVMONTHNAME2 = &H45        '  abbreviated name for February
Const LOCALE_SABBREVMONTHNAME3 = &H46        '  abbreviated name for March
Const LOCALE_SABBREVMONTHNAME4 = &H47        '  abbreviated name for April
Const LOCALE_SABBREVMONTHNAME5 = &H48        '  abbreviated name for May
Const LOCALE_SABBREVMONTHNAME6 = &H49        '  abbreviated name for June
Const LOCALE_SABBREVMONTHNAME7 = &H4A        '  abbreviated name for July
Const LOCALE_SABBREVMONTHNAME8 = &H4B        '  abbreviated name for August
Const LOCALE_SABBREVMONTHNAME9 = &H4C        '  abbreviated name for September
Const LOCALE_SABBREVMONTHNAME10 = &H4D        '  abbreviated name for October
Const LOCALE_SABBREVMONTHNAME11 = &H4E        '  abbreviated name for November
Const LOCALE_SABBREVMONTHNAME12 = &H4F        '  abbreviated name for December
Const LOCALE_SABBREVMONTHNAME13 = &H100F
Const LOCALE_SPOSITIVESIGN = &H50        '  positive sign
Const LOCALE_SNEGATIVESIGN = &H51        '  negative sign
Const LOCALE_IPOSSIGNPOSN = &H52        '  positive sign position
Const LOCALE_INEGSIGNPOSN = &H53        '  negative sign position
Const LOCALE_IPOSSYMPRECEDES = &H54        '  mon sym precedes pos amt
Const LOCALE_IPOSSEPBYSPACE = &H55        '  mon sym sep by space from pos amt
Const LOCALE_INEGSYMPRECEDES = &H56        '  mon sym precedes neg amt
Const LOCALE_INEGSEPBYSPACE = &H57        '  mon sym sep by space from neg amt
Const TIME_NOMINUTESORSECONDS = &H1         '  do not use minutes or seconds
Const TIME_NOSECONDS = &H2         '  do not use seconds
Const TIME_NOTIMEMARKER = &H4         '  do not use time marker
Const TIME_FORCE24HOURFORMAT = &H8         '  always use 24 hour format
Const DATE_SHORTDATE = &H1         '  use short date picture
Const DATE_LONGDATE = &H2         '  use long date picture
Const MAX_DEFAULTCHAR = 2
Const CAL_ICALINTVALUE = &H1                     '  calendar type
Const CAL_SCALNAME = &H2                         '  native name of calendar
Const CAL_IYEAROFFSETRANGE = &H3                 '  starting years of eras
Const CAL_SERASTRING = &H4                       '  era name for IYearOffsetRanges
Const CAL_SSHORTDATE = &H5                       '  Integer date format string
Const CAL_SLONGDATE = &H6                        '  long date format string
Const CAL_SDAYNAME1 = &H7                        '  native name for Monday
Const CAL_SDAYNAME2 = &H8                        '  native name for Tuesday
Const CAL_SDAYNAME3 = &H9                        '  native name for Wednesday
Const CAL_SDAYNAME4 = &HA                        '  native name for Thursday
Const CAL_SDAYNAME5 = &HB                        '  native name for Friday
Const CAL_SDAYNAME6 = &HC                        '  native name for Saturday
Const CAL_SDAYNAME7 = &HD                        '  native name for Sunday
Const CAL_SABBREVDAYNAME1 = &HE                  '  abbreviated name for Monday
Const CAL_SABBREVDAYNAME2 = &HF                  '  abbreviated name for Tuesday
Const CAL_SABBREVDAYNAME3 = &H10                 '  abbreviated name for Wednesday
Const CAL_SABBREVDAYNAME4 = &H11                 '  abbreviated name for Thursday
Const CAL_SABBREVDAYNAME5 = &H12                 '  abbreviated name for Friday
Const CAL_SABBREVDAYNAME6 = &H13                 '  abbreviated name for Saturday
Const CAL_SABBREVDAYNAME7 = &H14                 '  abbreviated name for Sunday
Const CAL_SMONTHNAME1 = &H15                     '  native name for January
Const CAL_SMONTHNAME2 = &H16                     '  native name for February
Const CAL_SMONTHNAME3 = &H17                     '  native name for March
Const CAL_SMONTHNAME4 = &H18                     '  native name for April
Const CAL_SMONTHNAME5 = &H19                     '  native name for May
Const CAL_SMONTHNAME6 = &H1A                     '  native name for June
Const CAL_SMONTHNAME7 = &H1B                     '  native name for July
Const CAL_SMONTHNAME8 = &H1C                     '  native name for August
Const CAL_SMONTHNAME9 = &H1D                     '  native name for September
Const CAL_SMONTHNAME10 = &H1E                    '  native name for October
Const CAL_SMONTHNAME11 = &H1F                    '  native name for November
Const CAL_SMONTHNAME12 = &H20                    '  native name for December
Const CAL_SMONTHNAME13 = &H21                    '  native name for 13th month (if any)
Const CAL_SABBREVMONTHNAME1 = &H22               '  abbreviated name for January
Const CAL_SABBREVMONTHNAME2 = &H23               '  abbreviated name for February
Const CAL_SABBREVMONTHNAME3 = &H24               '  abbreviated name for March
Const CAL_SABBREVMONTHNAME4 = &H25               '  abbreviated name for April
Const CAL_SABBREVMONTHNAME5 = &H26               '  abbreviated name for May
Const CAL_SABBREVMONTHNAME6 = &H27               '  abbreviated name for June
Const CAL_SABBREVMONTHNAME7 = &H28               '  abbreviated name for July
Const CAL_SABBREVMONTHNAME8 = &H29               '  abbreviated name for August
Const CAL_SABBREVMONTHNAME9 = &H2A               '  abbreviated name for September
Const CAL_SABBREVMONTHNAME10 = &H2B              '  abbreviated name for October
Const CAL_SABBREVMONTHNAME11 = &H2C              '  abbreviated name for November
Const CAL_SABBREVMONTHNAME12 = &H2D              '  abbreviated name for December
Const CAL_SABBREVMONTHNAME13 = &H2E              '  abbreviated name for 13th month (if any)
Const ENUM_ALL_CALENDARS = &HFFFF                '  enumerate all calendars
Const CAL_GREGORIAN = 1                 '  Gregorian (localized) calendar
Const CAL_GREGORIAN_US = 2              '  Gregorian (U.S.) calendar
Const CAL_JAPAN = 3                     '  Japanese Emperor Era calendar
Const CAL_TAIWAN = 4                    '  Taiwan Region Era calendar
Const CAL_KOREA = 5                     '  Korean Tangun Era calendar
Const RIGHT_ALT_PRESSED = &H1     '  the right alt key is pressed.
Const LEFT_ALT_PRESSED = &H2     '  the left alt key is pressed.
Const RIGHT_CTRL_PRESSED = &H4     '  the right ctrl key is pressed.
Const LEFT_CTRL_PRESSED = &H8     '  the left ctrl key is pressed.
Const SHIFT_PRESSED = &H10    '  the shift key is pressed.
Const NUMLOCK_ON = &H20    '  the numlock light is on.
Const SCROLLLOCK_ON = &H40    '  the scrolllock light is on.
Const CAPSLOCK_ON = &H80    '  the capslock light is on.
Const ENHANCED_KEY = &H100   '  the key is enhanced.
Const FROM_LEFT_1ST_BUTTON_PRESSED = &H1
Const RIGHTMOST_BUTTON_PRESSED = &H2
Const FROM_LEFT_2ND_BUTTON_PRESSED = &H4
Const FROM_LEFT_3RD_BUTTON_PRESSED = &H8
Const FROM_LEFT_4TH_BUTTON_PRESSED = &H10
Const MOUSE_MOVED = &H1
Const DOUBLE_CLICK = &H2
Const KEY_EVENT = &H1     '  Event contains key event record
Const mouse_eventC = &H2     '  Event contains mouse event record
Const WINDOW_BUFFER_SIZE_EVENT = &H4     '  Event contains window change event record
Const MENU_EVENT = &H8     '  Event contains menu event record
Const FOCUS_EVENT = &H10    '  event contains focus change
Const FOREGROUND_BLUE = &H1     '  text color contains blue.
Const FOREGROUND_GREEN = &H2     '  text color contains green.
Const FOREGROUND_RED = &H4     '  text color contains red.
Const FOREGROUND_INTENSITY = &H8     '  text color is intensified.
Const BACKGROUND_BLUE = &H10    '  background color contains blue.
Const BACKGROUND_GREEN = &H20    '  background color contains green.
Const BACKGROUND_RED = &H40    '  background color contains red.
Const BACKGROUND_INTENSITY = &H80    '  background color is intensified.
Const CTRL_C_EVENT = 0
Const CTRL_BREAK_EVENT = 1
Const CTRL_CLOSE_EVENT = 2
Const CTRL_LOGOFF_EVENT = 5
Const CTRL_SHUTDOWN_EVENT = 6
Const ENABLE_PROCESSED_INPUT = &H1
Const ENABLE_LINE_INPUT = &H2
Const ENABLE_ECHO_INPUT = &H4
Const ENABLE_WINDOW_INPUT = &H8
Const ENABLE_MOUSE_INPUT = &H10
Const ENABLE_PROCESSED_OUTPUT = &H1
Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2
Const CONSOLE_TEXTMODE_BUFFER = 1
Const R2_BLACK = 1       '   0
Const R2_NOTMERGEPEN = 2    '  DPon
Const R2_MASKNOTPEN = 3  '  DPna
Const R2_NOTCOPYPEN = 4  '  PN
Const R2_MASKPENNOT = 5  '  PDna
Const R2_NOT = 6    '  Dn
Const R2_XORPEN = 7      '  DPx
Const R2_NOTMASKPEN = 8  '  DPan
Const R2_MASKPEN = 9     '  DPa
Const R2_NOTXORPEN = 10  '  DPxn
Const R2_NOP = 11        '  D
Const R2_MERGENOTPEN = 12        '  DPno
Const R2_COPYPEN = 13    '  P
Const R2_MERGEPENNOT = 14        '  PDno
Const R2_MERGEPEN = 15   '  DPo
Const R2_WHITE = 16      '   1
Const R2_LAST = 16
Const SRCCOPY = &HCC0020    ' (DWORD) dest = source
Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Const PATCOPY = &HF00021    ' (DWORD) dest = pattern
Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Const BLACKNESS = &H42    ' (DWORD) dest = BLACK
Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE
Const GDI_ERROR = &HFFFF
Const HGDI_ERROR = &HFFFF
Const ERRORAPI = 0
Const NULLREGION = 1
Const SIMPLEREGION = 2
Const COMPLEXREGION = 3
Const RGN_AND = 1
Const RGN_OR = 2
Const RGN_XOR = 3
Const RGN_DIFF = 4
Const RGN_COPY = 5
Const RGN_MIN = RGN_AND
Const RGN_MAX = RGN_COPY
Const BLACKONWHITE = 1
Const WHITEONBLACK = 2
Const COLORONCOLOR = 3
Const HALFTONE = 4
Const MAXSTRETCHBLTMODE = 4
Const ALTERNATE = 1
Const WINDING = 2
Const POLYFILL_LAST = 2
Const TA_NOUPDATECP = 0
Const TA_UPDATECP = 1
Const TA_LEFT = 0
Const TA_RIGHT = 2
Const TA_CENTER = 6
Const TA_TOP = 0
Const TA_BOTTOM = 8
Const TA_BASELINE = 24
Const TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
Const VTA_BASELINE = TA_BASELINE
Const VTA_LEFT = TA_BOTTOM
Const VTA_RIGHT = TA_TOP
Const VTA_CENTER = TA_CENTER
Const VTA_BOTTOM = TA_RIGHT
Const VTA_TOP = TA_LEFT
Const ETO_GRAYED = 1
Const ETO_OPAQUE = 2
Const ETO_CLIPPED = 4
Const ASPECT_FILTERING = &H1
Const DCB_RESET = &H1
Const DCB_ACCUMULATE = &H2
Const DCB_DIRTY = DCB_ACCUMULATE
Const DCB_SET = (DCB_RESET Or DCB_ACCUMULATE)
Const DCB_ENABLE = &H4
Const DCB_DISABLE = &H8
Const META_SETBKCOLOR = &H201
Const META_SETBKMODE = &H102
Const META_SETMAPMODE = &H103
Const META_SETROP2 = &H104
Const META_SETRELABS = &H105
Const META_SETPOLYFILLMODE = &H106
Const META_SETSTRETCHBLTMODE = &H107
Const META_SETTEXTCHAREXTRA = &H108
Const META_SETTEXTCOLOR = &H209
Const META_SETTEXTJUSTIFICATION = &H20A
Const META_SETWINDOWORG = &H20B
Const META_SETWINDOWEXT = &H20C
Const META_SETVIEWPORTORG = &H20D
Const META_SETVIEWPORTEXT = &H20E
Const META_OFFSETWINDOWORG = &H20F
Const META_SCALEWINDOWEXT = &H410
Const META_OFFSETVIEWPORTORG = &H211
Const META_SCALEVIEWPORTEXT = &H412
Const META_LINETO = &H213
Const META_MOVETO = &H214
Const META_EXCLUDECLIPRECT = &H415
Const META_INTERSECTCLIPRECT = &H416
Const META_ARC = &H817
Const META_ELLIPSE = &H418
Const META_FLOODFILL = &H419
Const META_PIE = &H81A
Const META_RECTANGLE = &H41B
Const META_ROUNDRECT = &H61C
Const META_PATBLT = &H61D
Const META_SAVEDC = &H1E
Const META_SETPIXEL = &H41F
Const META_OFFSETCLIPRGN = &H220
Const META_TEXTOUT = &H521
Const META_BITBLT = &H922
Const META_STRETCHBLT = &HB23
Const META_POLYGON = &H324
Const META_POLYLINE = &H325
Const META_ESCAPE = &H626
Const META_RESTOREDC = &H127
Const META_FILLREGION = &H228
Const META_FRAMEREGION = &H429
Const META_INVERTREGION = &H12A
Const META_PAINTREGION = &H12B
Const META_SELECTCLIPREGION = &H12C
Const META_SELECTOBJECT = &H12D
Const META_SETTEXTALIGN = &H12E
Const META_CHORD = &H830
Const META_SETMAPPERFLAGS = &H231
Const META_EXTTEXTOUT = &HA32
Const META_SETDIBTODEV = &HD33
Const META_SELECTPALETTE = &H234
Const META_REALIZEPALETTE = &H35
Const META_ANIMATEPALETTE = &H436
Const META_SETPALENTRIES = &H37
Const META_POLYPOLYGON = &H538
Const META_RESIZEPALETTE = &H139
Const META_DIBBITBLT = &H940
Const META_DIBSTRETCHBLT = &HB41
Const META_DIBCREATEPATTERNBRUSH = &H142
Const META_STRETCHDIB = &HF43
Const META_EXTFLOODFILL = &H548
Const META_DELETEOBJECT = &H1F0
Const META_CREATEPALETTE = &HF7
Const META_CREATEPATTERNBRUSH = &H1F9
Const META_CREATEPENINDIRECT = &H2FA
Const META_CREATEFONTINDIRECT = &H2FB
Const META_CREATEBRUSHINDIRECT = &H2FC
Const META_CREATEREGION = &H6FF
Const NEWFRAME = 1
Const AbortDocC = 2
Const NEXTBAND = 3
Const SETCOLORTABLE = 4
Const GETCOLORTABLE = 5
Const FLUSHOUTPUT = 6
Const DRAFTMODE = 7
Const QUERYESCSUPPORT = 8
Const SETABORTPROC = 9
Const StartDocC = 10
Const EndDocC = 11
Const GETPHYSPAGESIZE = 12
Const GETPRINTINGOFFSET = 13
Const GETSCALINGFACTOR = 14
Const MFCOMMENT = 15
Const GETPENWIDTH = 16
Const SETCOPYCOUNT = 17
Const SELECTPAPERSOURCE = 18
Const DEVICEDATA = 19
Const PASSTHROUGH = 19
Const GETTECHNOLGY = 20
Const GETTECHNOLOGY = 20
Const SETLINECAP = 21
Const SETLINEJOIN = 22
Const SetMiterLimitC = 23
Const BANDINFO = 24
Const DRAWPATTERNRECT = 25
Const GETVECTORPENSIZE = 26
Const GETVECTORBRUSHSIZE = 27
Const ENABLEDUPLEX = 28
Const GETSETPAPERBINS = 29
Const GETSETPRINTORIENT = 30
Const ENUMPAPERBINS = 31
Const SETDIBSCALING = 32
Const EPSPRINTING = 33
Const ENUMPAPERMETRICS = 34
Const GETSETPAPERMETRICS = 35
Const POSTSCRIPT_DATA = 37
Const POSTSCRIPT_IGNORE = 38
Const MOUSETRAILS = 39
Const GETDEVICEUNITS = 42
Const GETEXTENDEDTEXTMETRICS = 256
Const GETEXTENTTABLE = 257
Const GETPAIRKERNTABLE = 258
Const GETTRACKKERNTABLE = 259
Const ExtTextOutC = 512
Const GETFACENAME = 513
Const DOWNLOADFACE = 514
Const ENABLERELATIVEWIDTHS = 768
Const ENABLEPAIRKERNING = 769
Const SETKERNTRACK = 770
Const SETALLJUSTVALUES = 771
Const SETCHARSET = 772
Const StretchBltC = 2048
Const GETSETSCREENPARAMS = 3072
Const BEGIN_PATH = 4096
Const CLIP_TO_PATH = 4097
Const END_PATH = 4098
Const EXT_DEVICE_CAPS = 4099
Const RESTORE_CTM = 4100
Const SAVE_CTM = 4101
Const SET_ARC_DIRECTION = 4102
Const SET_BACKGROUND_COLOR = 4103
Const SET_POLY_MODE = 4104
Const SET_SCREEN_ANGLE = 4105
Const SET_SPREAD = 4106
Const TRANSFORM_CTM = 4107
Const SET_CLIP_BOX = 4108
Const SET_BOUNDS = 4109
Const SET_MIRROR_MODE = 4110
Const OPENCHANNEL = 4110
Const DOWNLOADHEADER = 4111
Const CLOSECHANNEL = 4112
Const POSTSCRIPT_PASSTHROUGH = 4115
Const ENCAPSULATED_POSTSCRIPT = 4116
Const SP_NOTREPORTED = &H4000
Const SP_ERROR = (-1)
Const SP_APPABORT = (-2)
Const SP_USERABORT = (-3)
Const SP_OUTOFDISK = (-4)
Const SP_OUTOFMEMORY = (-5)
Const PR_JOBSTATUS = &H0
Const OBJ_PEN = 1
Const OBJ_BRUSH = 2
Const OBJ_DC = 3
Const OBJ_METADC = 4
Const OBJ_PAL = 5
Const OBJ_FONT = 6
Const OBJ_BITMAP = 7
Const OBJ_REGION = 8
Const OBJ_METAFILE = 9
Const OBJ_MEMDC = 10
Const OBJ_EXTPEN = 11
Const OBJ_ENHMETADC = 12
Const OBJ_ENHMETAFILE = 13
Const MWT_IDENTITY = 1
Const MWT_LEFTMULTIPLY = 2
Const MWT_RIGHTMULTIPLY = 3
Const MWT_MIN = MWT_IDENTITY
Const MWT_MAX = MWT_RIGHTMULTIPLY
Const BI_RGB = 0&
Const BI_RLE8 = 1&
Const BI_RLE4 = 2&
Const BI_bitfields = 3&
Const NTM_REGULAR = &H40&
Const NTM_BOLD = &H20&
Const NTM_ITALIC = &H1&
Const TMPF_FIXED_PITCH = &H1
Const TMPF_VECTOR = &H2
Const TMPF_DEVICE = &H8
Const TMPF_TRUETYPE = &H4
Const LF_FACESIZE = 32
Const LF_FULLFACESIZE = 64
Const OUT_DEFAULT_PRECIS = 0
Const OUT_STRING_PRECIS = 1
Const OUT_CHARACTER_PRECIS = 2
Const OUT_STROKE_PRECIS = 3
Const OUT_TT_PRECIS = 4
Const OUT_DEVICE_PRECIS = 5
Const OUT_RASTER_PRECIS = 6
Const OUT_TT_ONLY_PRECIS = 7
Const OUT_OUTLINE_PRECIS = 8
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_CHARACTER_PRECIS = 1
Const CLIP_STROKE_PRECIS = 2
Const CLIP_MASK = &HF
Const CLIP_LH_ANGLES = 16
Const CLIP_TT_ALWAYS = 32
Const CLIP_EMBEDDED = 128
Const DEFAULT_QUALITY = 0
Const DRAFT_QUALITY = 1
Const PROOF_QUALITY = 2
Const DEFAULT_PITCH = 0
Const FIXED_PITCH = 1
Const VARIABLE_PITCH = 2
Const ANSI_CHARSET = 0
Const DEFAULT_CHARSET = 1
Const SYMBOL_CHARSET = 2
Const SHIFTJIS_CHARSET = 128
Const HANGEUL_CHARSET = 129
Const CHINESEBIG5_CHARSET = 136
Const OEM_CHARSET = 255
Const FF_DONTCARE = 0    '  Don't care or don't know.
Const FF_ROMAN = 16      '  Variable stroke width, serifed.
Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.
Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.
Const FF_SCRIPT = 64     '  Cursive, etc.
Const FF_DECORATIVE = 80    '  Old English, etc.
Const FW_DONTCARE = 0
Const FW_THIN = 100
Const FW_EXTRALIGHT = 200
Const FW_LIGHT = 300
Const FW_NORMAL = 400
Const FW_MEDIUM = 500
Const FW_SEMIBOLD = 600
Const FW_BOLD = 700
Const FW_EXTRABOLD = 800
Const FW_HEAVY = 900
Const FW_ULTRALIGHT = FW_EXTRALIGHT
Const FW_REGULAR = FW_NORMAL
Const FW_DEMIBOLD = FW_SEMIBOLD
Const FW_ULTRABOLD = FW_EXTRABOLD
Const FW_BLACK = FW_HEAVY
Const PANOSE_COUNT = 10
Const PAN_FAMILYTYPE_INDEX = 0
Const PAN_SERIFSTYLE_INDEX = 1
Const PAN_WEIGHT_INDEX = 2
Const PAN_PROPORTION_INDEX = 3
Const PAN_CONTRAST_INDEX = 4
Const PAN_STROKEVARIATION_INDEX = 5
Const PAN_ARMSTYLE_INDEX = 6
Const PAN_LETTERFORM_INDEX = 7
Const PAN_MIDLINE_INDEX = 8
Const PAN_XHEIGHT_INDEX = 9
Const PAN_CULTURE_LATIN = 0
Const PAN_ANY = 0  '  Any
Const PAN_NO_FIT = 1  '  No Fit
Const PAN_FAMILY_TEXT_DISPLAY = 2  '  Text and Display
Const PAN_FAMILY_SCRIPT = 3  '  Script
Const PAN_FAMILY_DECORATIVE = 4  '  Decorative
Const PAN_FAMILY_PICTORIAL = 5  '  Pictorial
Const PAN_SERIF_COVE = 2  '  Cove
Const PAN_SERIF_OBTUSE_COVE = 3  '  Obtuse Cove
Const PAN_SERIF_SQUARE_COVE = 4  '  Square Cove
Const PAN_SERIF_OBTUSE_SQUARE_COVE = 5  '  Obtuse Square Cove
Const PAN_SERIF_SQUARE = 6  '  Square
Const PAN_SERIF_THIN = 7  '  Thin
Const PAN_SERIF_BONE = 8  '  Bone
Const PAN_SERIF_EXAGGERATED = 9  '  Exaggerated
Const PAN_SERIF_TRIANGLE = 10  '  Triangle
Const PAN_SERIF_NORMAL_SANS = 11  '  Normal Sans
Const PAN_SERIF_OBTUSE_SANS = 12  '  Obtuse Sans
Const PAN_SERIF_PERP_SANS = 13  '  Prep Sans
Const PAN_SERIF_FLARED = 14  '  Flared
Const PAN_SERIF_ROUNDED = 15  '  Rounded
Const PAN_WEIGHT_VERY_LIGHT = 2  '  Very Light
Const PAN_WEIGHT_LIGHT = 3  '  Light
Const PAN_WEIGHT_THIN = 4  '  Thin
Const PAN_WEIGHT_BOOK = 5  '  Book
Const PAN_WEIGHT_MEDIUM = 6  '  Medium
Const PAN_WEIGHT_DEMI = 7  '  Demi
Const PAN_WEIGHT_BOLD = 8  '  Bold
Const PAN_WEIGHT_HEAVY = 9  '  Heavy
Const PAN_WEIGHT_BLACK = 10  '  Black
Const PAN_WEIGHT_NORD = 11  '  Nord
Const PAN_PROP_OLD_STYLE = 2  '  Old Style
Const PAN_PROP_MODERN = 3  '  Modern
Const PAN_PROP_EVEN_WIDTH = 4  '  Even Width
Const PAN_PROP_EXPANDED = 5  '  Expanded
Const PAN_PROP_CONDENSED = 6  '  Condensed
Const PAN_PROP_VERY_EXPANDED = 7  '  Very Expanded
Const PAN_PROP_VERY_CONDENSED = 8  '  Very Condensed
Const PAN_PROP_MONOSPACED = 9  '  Monospaced
Const PAN_CONTRAST_NONE = 2  '  None
Const PAN_CONTRAST_VERY_LOW = 3  '  Very Low
Const PAN_CONTRAST_LOW = 4  '  Low
Const PAN_CONTRAST_MEDIUM_LOW = 5  '  Medium Low
Const PAN_CONTRAST_MEDIUM = 6  '  Medium
Const PAN_CONTRAST_MEDIUM_HIGH = 7  '  Mediim High
Const PAN_CONTRAST_HIGH = 8  '  High
Const PAN_CONTRAST_VERY_HIGH = 9  '  Very High
Const PAN_STROKE_GRADUAL_DIAG = 2  '  Gradual/Diagonal
Const PAN_STROKE_GRADUAL_TRAN = 3  '  Gradual/Transitional
Const PAN_STROKE_GRADUAL_VERT = 4  '  Gradual/Vertical
Const PAN_STROKE_GRADUAL_HORZ = 5  '  Gradual/Horizontal
Const PAN_STROKE_RAPID_VERT = 6  '  Rapid/Vertical
Const PAN_STROKE_RAPID_HORZ = 7  '  Rapid/Horizontal
Const PAN_STROKE_INSTANT_VERT = 8  '  Instant/Vertical
Const PAN_STRAIGHT_ARMS_HORZ = 2  '  Straight Arms/Horizontal
Const PAN_STRAIGHT_ARMS_WEDGE = 3  '  Straight Arms/Wedge
Const PAN_STRAIGHT_ARMS_VERT = 4  '  Straight Arms/Vertical
Const PAN_STRAIGHT_ARMS_SINGLE_SERIF = 5    '  Straight Arms/Single-Serif
Const PAN_STRAIGHT_ARMS_DOUBLE_SERIF = 6    '  Straight Arms/Double-Serif
Const PAN_BENT_ARMS_HORZ = 7  '  Non-Straight Arms/Horizontal
Const PAN_BENT_ARMS_WEDGE = 8  '  Non-Straight Arms/Wedge
Const PAN_BENT_ARMS_VERT = 9  '  Non-Straight Arms/Vertical
Const PAN_BENT_ARMS_SINGLE_SERIF = 10  '  Non-Straight Arms/Single-Serif
Const PAN_BENT_ARMS_DOUBLE_SERIF = 11  '  Non-Straight Arms/Double-Serif
Const PAN_LETT_NORMAL_CONTACT = 2  '  Normal/Contact
Const PAN_LETT_NORMAL_WEIGHTED = 3  '  Normal/Weighted
Const PAN_LETT_NORMAL_BOXED = 4  '  Normal/Boxed
Const PAN_LETT_NORMAL_FLATTENED = 5  '  Normal/Flattened
Const PAN_LETT_NORMAL_ROUNDED = 6  '  Normal/Rounded
Const PAN_LETT_NORMAL_OFF_CENTER = 7  '  Normal/Off Center
Const PAN_LETT_NORMAL_SQUARE = 8  '  Normal/Square
Const PAN_LETT_OBLIQUE_CONTACT = 9  '  Oblique/Contact
Const PAN_LETT_OBLIQUE_WEIGHTED = 10  '  Oblique/Weighted
Const PAN_LETT_OBLIQUE_BOXED = 11  '  Oblique/Boxed
Const PAN_LETT_OBLIQUE_FLATTENED = 12  '  Oblique/Flattened
Const PAN_LETT_OBLIQUE_ROUNDED = 13  '  Oblique/Rounded
Const PAN_LETT_OBLIQUE_OFF_CENTER = 14  '  Oblique/Off Center
Const PAN_LETT_OBLIQUE_SQUARE = 15  '  Oblique/Square
Const PAN_MIDLINE_STANDARD_TRIMMED = 2  '  Standard/Trimmed
Const PAN_MIDLINE_STANDARD_POINTED = 3  '  Standard/Pointed
Const PAN_MIDLINE_STANDARD_SERIFED = 4  '  Standard/Serifed
Const PAN_MIDLINE_HIGH_TRIMMED = 5  '  High/Trimmed
Const PAN_MIDLINE_HIGH_POINTED = 6  '  High/Pointed
Const PAN_MIDLINE_HIGH_SERIFED = 7  '  High/Serifed
Const PAN_MIDLINE_CONSTANT_TRIMMED = 8  '  Constant/Trimmed
Const PAN_MIDLINE_CONSTANT_POINTED = 9  '  Constant/Pointed
Const PAN_MIDLINE_CONSTANT_SERIFED = 10  '  Constant/Serifed
Const PAN_MIDLINE_LOW_TRIMMED = 11  '  Low/Trimmed
Const PAN_MIDLINE_LOW_POINTED = 12  '  Low/Pointed
Const PAN_MIDLINE_LOW_SERIFED = 13  '  Low/Serifed
Const PAN_XHEIGHT_CONSTANT_SMALL = 2  '  Constant/Small
Const PAN_XHEIGHT_CONSTANT_STD = 3  '  Constant/Standard
Const PAN_XHEIGHT_CONSTANT_LARGE = 4  '  Constant/Large
Const PAN_XHEIGHT_DUCKING_SMALL = 5  '  Ducking/Small
Const PAN_XHEIGHT_DUCKING_STD = 6  '  Ducking/Standard
Const PAN_XHEIGHT_DUCKING_LARGE = 7  '  Ducking/Large
Const ELF_VENDOR_SIZE = 4
Const ELF_VERSION = 0
Const ELF_CULTURE_LATIN = 0
Const RASTER_FONTTYPE = &H1
Const DEVICE_FONTTYPE = &H2
Const TRUETYPE_FONTTYPE = &H4
Const PC_RESERVED = &H1  '  palette index used for animation
Const PC_EXPLICIT = &H2  '  palette index is explicit to device
Const PC_NOCOLLAPSE = &H4        '  do not match color to system palette
Const TRANSPARENT = 1
Const OPAQUE = 2
Const BKMODE_LAST = 2
Const GM_COMPATIBLE = 1
Const GM_ADVANCED = 2
Const GM_LAST = 2
Const PT_CLOSEFIGURE = &H1
Const PT_LINETO = &H2
Const PT_BEZIERTO = &H4
Const PT_MOVETO = &H6
Const MM_TEXT = 1
Const MM_LOMETRIC = 2
Const MM_HIMETRIC = 3
Const MM_LOENGLISH = 4
Const MM_HIENGLISH = 5
Const MM_TWIPS = 6
Const MM_ISOTROPIC = 7
Const MM_ANISOTROPIC = 8
Const MM_MIN = MM_TEXT
Const MM_MAX = MM_ANISOTROPIC
Const MM_MAX_FIXEDSCALE = MM_TWIPS
Const ABSOLUTE = 1
Const RELATIVE = 2
Const WHITE_BRUSH = 0
Const LTGRAY_BRUSH = 1
Const GRAY_BRUSH = 2
Const DKGRAY_BRUSH = 3
Const BLACK_BRUSH = 4
Const NULL_BRUSH = 5
Const HOLLOW_BRUSH = NULL_BRUSH
Const WHITE_PEN = 6
Const BLACK_PEN = 7
Const NULL_PEN = 8
Const OEM_FIXED_FONT = 10
Const ANSI_FIXED_FONT = 11
Const ANSI_VAR_FONT = 12
Const SYSTEM_FONT = 13
Const DEVICE_DEFAULT_FONT = 14
Const DEFAULT_PALETTE = 15
Const SYSTEM_FIXED_FONT = 16
Const STOCK_LAST = 16
Const CLR_INVALID = &HFFFF
Const BS_SOLID = 0
Const BS_NULL = 1
Const BS_HOLLOW = BS_NULL
Const BS_HATCHED = 2
Const BS_PATTERN = 3
Const BS_INDEXED = 4
Const BS_DIBPATTERN = 5
Const BS_DIBPATTERNPT = 6
Const BS_PATTERN8X8 = 7
Const BS_DIBPATTERN8X8 = 8
Const HS_HORIZONTAL = 0              '  -----
Const HS_VERTICAL = 1                '  |||||
Const HS_FDIAGONAL = 2               '  \\\\\
Const HS_BDIAGONAL = 3               '  /////
Const HS_CROSS = 4                   '  +++++
Const HS_DIAGCROSS = 5               '  xxxxx
Const HS_FDIAGONAL1 = 6
Const HS_BDIAGONAL1 = 7
Const HS_SOLID = 8
Const HS_DENSE1 = 9
Const HS_DENSE2 = 10
Const HS_DENSE3 = 11
Const HS_DENSE4 = 12
Const HS_DENSE5 = 13
Const HS_DENSE6 = 14
Const HS_DENSE7 = 15
Const HS_DENSE8 = 16
Const HS_NOSHADE = 17
Const HS_HALFTONE = 18
Const HS_SOLIDCLR = 19
Const HS_DITHEREDCLR = 20
Const HS_SOLIDTEXTCLR = 21
Const HS_DITHEREDTEXTCLR = 22
Const HS_SOLIDBKCLR = 23
Const HS_DITHEREDBKCLR = 24
Const HS_API_MAX = 25
Const PS_SOLID = 0
Const PS_DASH = 1                    '  -------
Const PS_DOT = 2                     '  .......
Const PS_DASHDOT = 3                 '  _._._._
Const PS_DASHDOTDOT = 4              '  _.._.._
Const PS_NULL = 5
Const PS_INSIDEFRAME = 6
Const PS_USERSTYLE = 7
Const PS_ALTERNATE = 8
Const PS_STYLE_MASK = &HF
Const PS_ENDCAP_ROUND = &H0
Const PS_ENDCAP_SQUARE = &H100
Const PS_ENDCAP_FLAT = &H200
Const PS_ENDCAP_MASK = &HF00
Const PS_JOIN_ROUND = &H0
Const PS_JOIN_BEVEL = &H1000
Const PS_JOIN_MITER = &H2000
Const PS_JOIN_MASK = &HF000
Const PS_COSMETIC = &H0
Const PS_GEOMETRIC = &H10000
Const PS_TYPE_MASK = &HF0000
Const AD_COUNTERCLOCKWISE = 1
Const AD_CLOCKWISE = 2
Const DRIVERVERSION = 0      '  Device driver version
Const TECHNOLOGY = 2         '  Device classification
Const HORZSIZE = 4           '  Horizontal size in millimeters
Const VERTSIZE = 6           '  Vertical size in millimeters
Const HORZRES = 8            '  Horizontal width in pixels
Const VERTRES = 10           '  Vertical width in pixels
Const BITSPIXEL = 12         '  Number of bits per pixel
Const PLANES = 14            '  Number of planes
Const NUMBRUSHES = 16        '  Number of brushes the device has
Const NUMPENS = 18           '  Number of pens the device has
Const NUMMARKERS = 20        '  Number of markers the device has
Const NUMFONTS = 22          '  Number of fonts the device has
Const NUMCOLORS = 24         '  Number of colors the device supports
Const PDEVICESIZE = 26       '  Size required for device descriptor
Const CURVECAPS = 28         '  Curve capabilities
Const LINECAPS = 30          '  Line capabilities
Const POLYGONALCAPS = 32     '  Polygonal capabilities
Const TEXTCAPS = 34          '  Text capabilities
Const CLIPCAPS = 36          '  Clipping capabilities
Const RASTERCAPS = 38        '  Bitblt capabilities
Const ASPECTX = 40           '  Length of the X leg
Const ASPECTY = 42           '  Length of the Y leg
Const ASPECTXY = 44          '  Length of the hypotenuse
Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Const SIZEPALETTE = 104      '  Number of entries in physical palette
Const NUMRESERVED = 106      '  Number of reserved entries in palette
Const COLORRES = 108         '  Actual color resolution
Const PHYSICALWIDTH = 110    '  Physical Width in device units
Const PHYSICALHEIGHT = 111    '  Physical Height in device units
Const PHYSICALOFFSETX = 112    '  Physical Printable Area x margin
Const PHYSICALOFFSETY = 113    '  Physical Printable Area y margin
Const SCALINGFACTORX = 114    '  Scaling factor x
Const SCALINGFACTORY = 115    '  Scaling factor y
Const DT_PLOTTER = 0             '  Vector plotter
Const DT_RASDISPLAY = 1          '  Raster display
Const DT_RASPRINTER = 2          '  Raster printer
Const DT_RASCAMERA = 3           '  Raster camera
Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Const DT_METAFILE = 5            '  Metafile, VDM
Const DT_DISPFILE = 6            '  Display-file
Const CC_NONE = 0                '  Curves not supported
Const CC_CIRCLES = 1             '  Can do circles
Const CC_PIE = 2                 '  Can do pie wedges
Const CC_CHORD = 4               '  Can do chord arcs
Const CC_ELLIPSES = 8            '  Can do ellipese
Const CC_WIDE = 16               '  Can do wide lines
Const CC_STYLED = 32             '  Can do styled lines
Const CC_WIDESTYLED = 64         '  Can do wide styled lines
Const CC_INTERIORS = 128    '  Can do interiors
Const CC_ROUNDRECT = 256    '
Const LC_NONE = 0                '  Lines not supported
Const LC_POLYLINE = 2            '  Can do polylines
Const LC_MARKER = 4              '  Can do markers
Const LC_POLYMARKER = 8          '  Can do polymarkers
Const LC_WIDE = 16               '  Can do wide lines
Const LC_STYLED = 32             '  Can do styled lines
Const LC_WIDESTYLED = 64         '  Can do wide styled lines
Const LC_INTERIORS = 128    '  Can do interiors
Const PC_NONE = 0                '  Polygonals not supported
Const PC_POLYGON = 1             '  Can do polygons
Const PC_RECTANGLE = 2           '  Can do rectangles
Const PC_WINDPOLYGON = 4         '  Can do winding polygons
Const PC_TRAPEZOID = 4           '  Can do trapezoids
Const PC_SCANLINE = 8            '  Can do scanlines
Const PC_WIDE = 16               '  Can do wide borders
Const PC_STYLED = 32             '  Can do styled borders
Const PC_WIDESTYLED = 64         '  Can do wide styled borders
Const PC_INTERIORS = 128    '  Can do interiors
Const CP_NONE = 0                '  No clipping of output
Const CP_RECTANGLE = 1           '  Output clipped to rects
Const CP_REGION = 2              '
Const TC_OP_CHARACTER = &H1              '  Can do OutputPrecision   CHARACTER
Const TC_OP_STROKE = &H2                 '  Can do OutputPrecision   STROKE
Const TC_CP_STROKE = &H4                 '  Can do ClipPrecision     STROKE
Const TC_CR_90 = &H8                     '  Can do CharRotAbility    90
Const TC_CR_ANY = &H10                   '  Can do CharRotAbility    ANY
Const TC_SF_X_YINDEP = &H20              '  Can do ScaleFreedom      X_YINDEPENDENT
Const TC_SA_DOUBLE = &H40                '  Can do ScaleAbility      DOUBLE
Const TC_SA_INTEGER = &H80               '  Can do ScaleAbility      INTEGER
Const TC_SA_CONTIN = &H100               '  Can do ScaleAbility      CONTINUOUS
Const TC_EA_DOUBLE = &H200               '  Can do EmboldenAbility   DOUBLE
Const TC_IA_ABLE = &H400                 '  Can do ItalisizeAbility  ABLE
Const TC_UA_ABLE = &H800                 '  Can do UnderlineAbility  ABLE
Const TC_SO_ABLE = &H1000                '  Can do StrikeOutAbility  ABLE
Const TC_RA_ABLE = &H2000                '  Can do RasterFontAble    ABLE
Const TC_VA_ABLE = &H4000                '  Can do VectorFontAble    ABLE
Const TC_RESERVED = &H8000
Const TC_SCROLLBLT = &H10000             '  do text scroll with blt
Const RC_NONE = 0
Const RC_BITBLT = 1                  '  Can do standard BLT.
Const RC_BANDING = 2                 '  Device requires banding support
Const RC_SCALING = 4                 '  Device requires scaling support
Const RC_BITMAP64 = 8                '  Device can support >64K bitmap
Const RC_GDI20_OUTPUT = &H10             '  has 2.0 output calls
Const RC_GDI20_STATE = &H20
Const RC_SAVEBITMAP = &H40
Const RC_DI_BITMAP = &H80                '  supports DIB to memory
Const RC_PALETTE = &H100                 '  supports a palette
Const RC_DIBTODEV = &H200                '  supports DIBitsToDevice
Const RC_BIGFONT = &H400                 '  supports >64K fonts
Const RC_STRETCHBLT = &H800              '  supports StretchBlt
Const RC_FLOODFILL = &H1000              '  supports FloodFill
Const RC_STRETCHDIB = &H2000             '  supports StretchDIBits
Const RC_OP_DX_OUTPUT = &H4000
Const RC_DEVBITS = &H8000
Const DIB_RGB_COLORS = 0    '  color table in RGBs
Const DIB_PAL_COLORS = 1    '  color table in palette indices
Const DIB_PAL_INDICES = 2    '  No color table indices into surf palette
Const DIB_PAL_PHYSINDICES = 2    '  No color table indices into surf palette
Const DIB_PAL_LOGINDICES = 4    '  No color table indices into DC palette
Const SYSPAL_ERROR = 0
Const SYSPAL_STATIC = 1
Const SYSPAL_NOSTATIC = 2
Const CBM_CREATEDIB = &H2      '  create DIB bitmap
Const CBM_INIT = &H4           '  initialize bitmap
Const FLOODFILLBORDER = 0
Const FLOODFILLSURFACE = 1
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const DM_SPECVERSION = &H320
Const DM_ORIENTATION = &H1&
Const DM_PAPERSIZE = &H2&
Const DM_PAPERLENGTH = &H4&
Const DM_PAPERWIDTH = &H8&
Const DM_SCALE = &H10&
Const DM_COPIES = &H100&
Const DM_DEFAULTSOURCE = &H200&
Const DM_PRINTQUALITY = &H400&
Const DM_COLOR = &H800&
Const DM_DUPLEX = &H1000&
Const DM_YRESOLUTION = &H2000&
Const DM_TTOPTION = &H4000&
Const DM_COLLATE As Long = &H8000
Const DM_FORMNAME As Long = &H10000
Const DMORIENT_PORTRAIT = 1
Const DMORIENT_LANDSCAPE = 2
Const DMPAPER_LETTER = 1
Const DMPAPER_FIRST = DMPAPER_LETTER
Const DMPAPER_LETTERSMALL = 2            '  Letter Small 8 1/2 x 11 in
Const DMPAPER_TABLOID = 3                '  Tabloid 11 x 17 in
Const DMPAPER_LEDGER = 4                 '  Ledger 17 x 11 in
Const DMPAPER_LEGAL = 5                  '  Legal 8 1/2 x 14 in
Const DMPAPER_STATEMENT = 6              '  Statement 5 1/2 x 8 1/2 in
Const DMPAPER_EXECUTIVE = 7              '  Executive 7 1/4 x 10 1/2 in
Const DMPAPER_A3 = 8                     '  A3 297 x 420 mm
Const DMPAPER_A4 = 9                     '  A4 210 x 297 mm
Const DMPAPER_A4SMALL = 10               '  A4 Small 210 x 297 mm
Const DMPAPER_A5 = 11                    '  A5 148 x 210 mm
Const DMPAPER_B4 = 12                    '  B4 250 x 354
Const DMPAPER_B5 = 13                    '  B5 182 x 257 mm
Const DMPAPER_FOLIO = 14                 '  Folio 8 1/2 x 13 in
Const DMPAPER_QUARTO = 15                '  Quarto 215 x 275 mm
Const DMPAPER_10X14 = 16                 '  10x14 in
Const DMPAPER_11X17 = 17                 '  11x17 in
Const DMPAPER_NOTE = 18                  '  Note 8 1/2 x 11 in
Const DMPAPER_ENV_9 = 19                 '  Envelope #9 3 7/8 x 8 7/8
Const DMPAPER_ENV_10 = 20                '  Envelope #10 4 1/8 x 9 1/2
Const DMPAPER_ENV_11 = 21                '  Envelope #11 4 1/2 x 10 3/8
Const DMPAPER_ENV_12 = 22                '  Envelope #12 4 \276 x 11
Const DMPAPER_ENV_14 = 23                '  Envelope #14 5 x 11 1/2
Const DMPAPER_CSHEET = 24                '  C size sheet
Const DMPAPER_DSHEET = 25                '  D size sheet
Const DMPAPER_ESHEET = 26                '  E size sheet
Const DMPAPER_ENV_DL = 27                '  Envelope DL 110 x 220mm
Const DMPAPER_ENV_C5 = 28                '  Envelope C5 162 x 229 mm
Const DMPAPER_ENV_C3 = 29                '  Envelope C3  324 x 458 mm
Const DMPAPER_ENV_C4 = 30                '  Envelope C4  229 x 324 mm
Const DMPAPER_ENV_C6 = 31                '  Envelope C6  114 x 162 mm
Const DMPAPER_ENV_C65 = 32               '  Envelope C65 114 x 229 mm
Const DMPAPER_ENV_B4 = 33                '  Envelope B4  250 x 353 mm
Const DMPAPER_ENV_B5 = 34                '  Envelope B5  176 x 250 mm
Const DMPAPER_ENV_B6 = 35                '  Envelope B6  176 x 125 mm
Const DMPAPER_ENV_ITALY = 36             '  Envelope 110 x 230 mm
Const DMPAPER_ENV_MONARCH = 37           '  Envelope Monarch 3.875 x 7.5 in
Const DMPAPER_ENV_PERSONAL = 38          '  6 3/4 Envelope 3 5/8 x 6 1/2 in
Const DMPAPER_FANFOLD_US = 39            '  US Std Fanfold 14 7/8 x 11 in
Const DMPAPER_FANFOLD_STD_GERMAN = 40    '  German Std Fanfold 8 1/2 x 12 in
Const DMPAPER_FANFOLD_LGL_GERMAN = 41    '  German Legal Fanfold 8 1/2 x 13 in
Const DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN
Const DMPAPER_USER = 256
Const DMBIN_UPPER = 1
Const DMBIN_FIRST = DMBIN_UPPER
Const DMBIN_ONLYONE = 1
Const DMBIN_LOWER = 2
Const DMBIN_MIDDLE = 3
Const DMBIN_MANUAL = 4
Const DMBIN_ENVELOPE = 5
Const DMBIN_ENVMANUAL = 6
Const DMBIN_AUTO = 7
Const DMBIN_TRACTOR = 8
Const DMBIN_SMALLFMT = 9
Const DMBIN_LARGEFMT = 10
Const DMBIN_LARGECAPACITY = 11
Const DMBIN_CASSETTE = 14
Const DMBIN_LAST = DMBIN_CASSETTE
Const DMBIN_USER = 256               '  device specific bins start here
Const DMRES_DRAFT = (-1)
Const DMRES_LOW = (-2)
Const DMRES_MEDIUM = (-3)
Const DMRES_HIGH = (-4)
Const DMCOLOR_MONOCHROME = 1
Const DMCOLOR_COLOR = 2
Const DMDUP_SIMPLEX = 1
Const DMDUP_VERTICAL = 2
Const DMDUP_HORIZONTAL = 3
Const DMTT_BITMAP = 1            '  print TT fonts as graphics
Const DMTT_DOWNLOAD = 2          '  download TT fonts as soft fonts
Const DMTT_SUBDEV = 3            '  substitute device fonts for TT fonts
Const DMCOLLATE_FALSE = 0
Const DMCOLLATE_TRUE = 1
Const DM_GRAYSCALE = &H1
Const DM_INTERLACED = &H2
Const RDH_RECTANGLES = 1
Const GGO_METRICS = 0
Const GGO_BITMAP = 1
Const GGO_NATIVE = 2
Const TT_POLYGON_TYPE = 24
Const TT_PRIM_LINE = 1
Const TT_PRIM_QSPLINE = 2
Const TT_AVAILABLE = &H1
Const TT_ENABLED = &H2
Const DM_UPDATE = 1
Const DM_COPY = 2
Const DM_PROMPT = 4
Const DM_MODIFY = 8
Const DM_IN_BUFFER = DM_MODIFY
Const DM_IN_PROMPT = DM_PROMPT
Const DM_OUT_BUFFER = DM_COPY
Const DM_OUT_DEFAULT = DM_UPDATE
Const DC_FIELDS = 1
Const DC_PAPERS = 2
Const DC_PAPERSIZE = 3
Const DC_MINEXTENT = 4
Const DC_MAXEXTENT = 5
Const DC_BINS = 6
Const DC_DUPLEX = 7
Const DC_SIZE = 8
Const DC_EXTRA = 9
Const DC_VERSION = 10
Const DC_DRIVER = 11
Const DC_BINNAMES = 12
Const DC_ENUMRESOLUTIONS = 13
Const DC_FILEDEPENDENCIES = 14
Const DC_TRUETYPE = 15
Const DC_PAPERNAMES = 16
Const DC_ORIENTATION = 17
Const DC_COPIES = 18
Const DCTT_BITMAP = &H1&
Const DCTT_DOWNLOAD = &H2&
Const DCTT_SUBDEV = &H4&
Const CA_NEGATIVE = &H1
Const CA_LOG_FILTER = &H2
Const ILLUMINANT_DEVICE_DEFAULT = 0
Const ILLUMINANT_A = 1
Const ILLUMINANT_B = 2
Const ILLUMINANT_C = 3
Const ILLUMINANT_D50 = 4
Const ILLUMINANT_D55 = 5
Const ILLUMINANT_D65 = 6
Const ILLUMINANT_D75 = 7
Const ILLUMINANT_F2 = 8
Const ILLUMINANT_MAX_INDEX = ILLUMINANT_F2
Const ILLUMINANT_TUNGSTEN = ILLUMINANT_A
Const ILLUMINANT_DAYLIGHT = ILLUMINANT_C
Const ILLUMINANT_FLUORESCENT = ILLUMINANT_F2
Const ILLUMINANT_NTSC = ILLUMINANT_C
Const RGB_GAMMA_MIN = 2500    'words
Const RGB_GAMMA_MAX = 65000
Const REFERENCE_WHITE_MIN = 6000    'words
Const REFERENCE_WHITE_MAX = 10000
Const REFERENCE_BLACK_MIN = 0
Const REFERENCE_BLACK_MAX = 4000
Const COLOR_ADJ_MIN = -100    'shorts
Const COLOR_ADJ_MAX = 100
Const FONTMAPPER_MAX = 10
Const ENHMETA_SIGNATURE = &H464D4520
Const ENHMETA_STOCK_OBJECT = &H80000000
Const EMR_HEADER = 1
Const EMR_POLYBEZIER = 2
Const EMR_POLYGON = 3
Const EMR_POLYLINE = 4
Const EMR_POLYBEZIERTO = 5
Const EMR_POLYLINETO = 6
Const EMR_POLYPOLYLINE = 7
Const EMR_POLYPOLYGON = 8
Const EMR_SETWINDOWEXTEX = 9
Const EMR_SETWINDOWORGEX = 10
Const EMR_SETVIEWPORTEXTEX = 11
Const EMR_SETVIEWPORTORGEX = 12
Const EMR_SETBRUSHORGEX = 13
Const EMR_EOF = 14
Const EMR_SETPIXELV = 15
Const EMR_SETMAPPERFLAGS = 16
Const EMR_SETMAPMODE = 17
Const EMR_SETBKMODE = 18
Const EMR_SETPOLYFILLMODE = 19
Const EMR_SETROP2 = 20
Const EMR_SETSTRETCHBLTMODE = 21
Const EMR_SETTEXTALIGN = 22
Const EMR_SETCOLORADJUSTMENT = 23
Const EMR_SETTEXTCOLOR = 24
Const EMR_SETBKCOLOR = 25
Const EMR_OFFSETCLIPRGN = 26
Const EMR_MOVETOEX = 27
Const EMR_SETMETARGN = 28
Const EMR_EXCLUDECLIPRECT = 29
Const EMR_INTERSECTCLIPRECT = 30
Const EMR_SCALEVIEWPORTEXTEX = 31
Const EMR_SCALEWINDOWEXTEX = 32
Const EMR_SAVEDC = 33
Const EMR_RESTOREDC = 34
Const EMR_SETWORLDTRANSFORM = 35
Const EMR_MODIFYWORLDTRANSFORM = 36
Const EMR_SELECTOBJECT = 37
Const EMR_CREATEPEN = 38
Const EMR_CREATEBRUSHINDIRECT = 39
Const EMR_DELETEOBJECT = 40
Const EMR_ANGLEARC = 41
Const EMR_ELLIPSE = 42
Const EMR_RECTANGLE = 43
Const EMR_ROUNDRECT = 44
Const EMR_ARC = 45
Const EMR_CHORD = 46
Const EMR_PIE = 47
Const EMR_SELECTPALETTE = 48
Const EMR_CREATEPALETTE = 49
Const EMR_SETPALETTEENTRIES = 50
Const EMR_RESIZEPALETTE = 51
Const EMR_REALIZEPALETTE = 52
Const EMR_EXTFLOODFILL = 53
Const EMR_LINETO = 54
Const EMR_ARCTO = 55
Const EMR_POLYDRAW = 56
Const EMR_SETARCDIRECTION = 57
Const EMR_SETMITERLIMIT = 58
Const EMR_BEGINPATH = 59
Const EMR_ENDPATH = 60
Const EMR_CLOSEFIGURE = 61
Const EMR_FILLPATH = 62
Const EMR_STROKEANDFILLPATH = 63
Const EMR_STROKEPATH = 64
Const EMR_FLATTENPATH = 65
Const EMR_WIDENPATH = 66
Const EMR_SELECTCLIPPATH = 67
Const EMR_ABORTPATH = 68
Const EMR_GDICOMMENT = 70
Const EMR_FILLRGN = 71
Const EMR_FRAMERGN = 72
Const EMR_INVERTRGN = 73
Const EMR_PAINTRGN = 74
Const EMR_EXTSELECTCLIPRGN = 75
Const EMR_BITBLT = 76
Const EMR_STRETCHBLT = 77
Const EMR_MASKBLT = 78
Const EMR_PLGBLT = 79
Const EMR_SETDIBITSTODEVICE = 80
Const EMR_STRETCHDIBITS = 81
Const EMR_EXTCREATEFONTINDIRECTW = 82
Const EMR_EXTTEXTOUTA = 83
Const EMR_EXTTEXTOUTW = 84
Const EMR_POLYBEZIER16 = 85
Const EMR_POLYGON16 = 86
Const EMR_POLYLINE16 = 87
Const EMR_POLYBEZIERTO16 = 88
Const EMR_POLYLINETO16 = 89
Const EMR_POLYPOLYLINE16 = 90
Const EMR_POLYPOLYGON16 = 91
Const EMR_POLYDRAW16 = 92
Const EMR_CREATEMONOBRUSH = 93
Const EMR_CREATEDIBPATTERNBRUSHPT = 94
Const EMR_EXTCREATEPEN = 95
Const EMR_POLYTEXTOUTA = 96
Const EMR_POLYTEXTOUTW = 97
Const EMR_MIN = 1
Const EMR_MAX = 97
Const STRETCH_ANDSCANS = 1
Const STRETCH_ORSCANS = 2
Const STRETCH_DELETESCANS = 3
Const STRETCH_HALFTONE = 4
Const TCI_SRCCHARSET = 1
Const TCI_SRCCODEPAGE = 2
Const TCI_SRCFONTSIG = 3
Const MONO_FONT = 8
Const JOHAB_CHARSET = 130
Const HEBREW_CHARSET = 177
Const ARABIC_CHARSET = 178
Const GREEK_CHARSET = 161
Const TURKISH_CHARSET = 162
Const THAI_CHARSET = 222
Const EASTEUROPE_CHARSET = 238
Const RUSSIAN_CHARSET = 204
Const MAC_CHARSET = 77
Const BALTIC_CHARSET = 186
Const FS_LATIN1 = &H1&
Const FS_LATIN2 = &H2&
Const FS_CYRILLIC = &H4&
Const FS_GREEK = &H8&
Const FS_TURKISH = &H10&
Const FS_HEBREW = &H20&
Const FS_ARABIC = &H40&
Const FS_BALTIC = &H80&
Const FS_THAI = &H10000
Const FS_JISJAPAN = &H20000
Const FS_CHINESESIMP = &H40000
Const FS_WANSUNG = &H80000
Const FS_CHINESETRAD = &H100000
Const FS_JOHAB = &H200000
Const FS_SYMBOL = &H80000000
Const DEFAULT_GUI_FONT = 17
Const DM_RESERVED1 = &H800000
Const DM_RESERVED2 = &H1000000
Const DM_ICMMETHOD = &H2000000
Const DM_ICMINTENT = &H4000000
Const DM_MEDIATYPE = &H8000000
Const DM_DITHERTYPE = &H10000000
Const DMPAPER_ISO_B4 = 42                '  B4 (ISO) 250 x 353 mm
Const DMPAPER_JAPANESE_POSTCARD = 43     '  Japanese Postcard 100 x 148 mm
Const DMPAPER_9X11 = 44                  '  9 x 11 in
Const DMPAPER_10X11 = 45                 '  10 x 11 in
Const DMPAPER_15X11 = 46                 '  15 x 11 in
Const DMPAPER_ENV_INVITE = 47            '  Envelope Invite 220 x 220 mm
Const DMPAPER_RESERVED_48 = 48           '  RESERVED--DO NOT USE
Const DMPAPER_RESERVED_49 = 49           '  RESERVED--DO NOT USE
Const DMPAPER_LETTER_EXTRA = 50              '  Letter Extra 9 \275 x 12 in
Const DMPAPER_LEGAL_EXTRA = 51               '  Legal Extra 9 \275 x 15 in
Const DMPAPER_TABLOID_EXTRA = 52              '  Tabloid Extra 11.69 x 18 in
Const DMPAPER_A4_EXTRA = 53                   '  A4 Extra 9.27 x 12.69 in
Const DMPAPER_LETTER_TRANSVERSE = 54     '  Letter Transverse 8 \275 x 11 in
Const DMPAPER_A4_TRANSVERSE = 55         '  A4 Transverse 210 x 297 mm
Const DMPAPER_LETTER_EXTRA_TRANSVERSE = 56    '  Letter Extra Transverse 9\275 x 12 in
Const DMPAPER_A_PLUS = 57                '  SuperA/SuperA/A4 227 x 356 mm
Const DMPAPER_B_PLUS = 58                '  SuperB/SuperB/A3 305 x 487 mm
Const DMPAPER_LETTER_PLUS = 59           '  Letter Plus 8.5 x 12.69 in
Const DMPAPER_A4_PLUS = 60               '  A4 Plus 210 x 330 mm
Const DMPAPER_A5_TRANSVERSE = 61         '  A5 Transverse 148 x 210 mm
Const DMPAPER_B5_TRANSVERSE = 62         '  B5 (JIS) Transverse 182 x 257 mm
Const DMPAPER_A3_EXTRA = 63              '  A3 Extra 322 x 445 mm
Const DMPAPER_A5_EXTRA = 64              '  A5 Extra 174 x 235 mm
Const DMPAPER_B5_EXTRA = 65              '  B5 (ISO) Extra 201 x 276 mm
Const DMPAPER_A2 = 66                    '  A2 420 x 594 mm
Const DMPAPER_A3_TRANSVERSE = 67         '  A3 Transverse 297 x 420 mm
Const DMPAPER_A3_EXTRA_TRANSVERSE = 68   '  A3 Extra Transverse 322 x 445 mm
Const DMTT_DOWNLOAD_OUTLINE = 4    '  download TT fonts as outline soft fonts
Const DMICMMETHOD_NONE = 1       '  ICM disabled
Const DMICMMETHOD_SYSTEM = 2     '  ICM handled by system
Const DMICMMETHOD_DRIVER = 3     '  ICM handled by driver
Const DMICMMETHOD_DEVICE = 4     '  ICM handled by device
Const DMICMMETHOD_USER = 256     '  Device-specific methods start here
Const DMICM_SATURATE = 1         '  Maximize color saturation
Const DMICM_CONTRAST = 2         '  Maximize color contrast
Const DMICM_COLORMETRIC = 3      '  Use specific color metric
Const DMICM_USER = 256           '  Device-specific intents start here
Const DMMEDIA_STANDARD = 1         '  Standard paper
Const DMMEDIA_GLOSSY = 2           '  Glossy paper
Const DMMEDIA_TRANSPARENCY = 3     '  Transparency
Const DMMEDIA_USER = 256           '  Device-specific media start here
Const DMDITHER_NONE = 1          '  No dithering
Const DMDITHER_COARSE = 2        '  Dither with a coarse brush
Const DMDITHER_FINE = 3          '  Dither with a fine brush
Const DMDITHER_LINEART = 4       '  LineArt dithering
Const DMDITHER_GRAYSCALE = 5     '  Device does grayscaling
Const DMDITHER_USER = 256        '  Device-specific dithers start here
Const GGO_GRAY2_BITMAP = 4
Const GGO_GRAY4_BITMAP = 5
Const GGO_GRAY8_BITMAP = 6
Const GGO_GLYPH_INDEX = &H80
Const GCP_DBCS = &H1
Const GCP_REORDER = &H2
Const GCP_USEKERNING = &H8
Const GCP_GLYPHSHAPE = &H10
Const GCP_LIGATE = &H20
Const GCP_DIACRITIC = &H100
Const GCP_KASHIDA = &H400
Const GCP_ERROR = &H8000
Const FLI_MASK = &H103B
Const GCP_JUSTIFY = &H10000
Const GCP_NODIACRITICS = &H20000
Const FLI_GLYPHS = &H40000
Const GCP_CLASSIN = &H80000
Const GCP_MAXEXTENT = &H100000
Const GCP_JUSTIFYIN = &H200000
Const GCP_DISPLAYZWG = &H400000
Const GCP_SYMSWAPOFF = &H800000
Const GCP_NUMERICOVERRIDE = &H1000000
Const GCP_NEUTRALOVERRIDE = &H2000000
Const GCP_NUMERICSLATIN = &H4000000
Const GCP_NUMERICSLOCAL = &H8000000
Const GCPCLASS_LATIN = 1
Const GCPCLASS_HEBREW = 2
Const GCPCLASS_ARABIC = 2
Const GCPCLASS_NEUTRAL = 3
Const GCPCLASS_LOCALNUMBER = 4
Const GCPCLASS_LATINNUMBER = 5
Const GCPCLASS_LATINNUMERICTERMINATOR = 6
Const GCPCLASS_LATINNUMERICSEPARATOR = 7
Const GCPCLASS_NUMERICSEPARATOR = 8
Const GCPCLASS_PREBOUNDRTL = &H80
Const GCPCLASS_PREBOUNDLTR = &H40
Const DC_BINADJUST = 19
Const DC_EMF_COMPLIANT = 20
Const DC_DATATYPE_PRODUCED = 21
Const DC_COLLATE = 22
Const DCTT_DOWNLOAD_OUTLINE = &H8&
Const DCBA_FACEUPNONE = &H0
Const DCBA_FACEUPCENTER = &H1
Const DCBA_FACEUPLEFT = &H2
Const DCBA_FACEUPRIGHT = &H3
Const DCBA_FACEDOWNNONE = &H100
Const DCBA_FACEDOWNCENTER = &H101
Const DCBA_FACEDOWNLEFT = &H102
Const DCBA_FACEDOWNRIGHT = &H103
Const ICM_OFF = 1
Const ICM_ON = 2
Const ICM_QUERY = 3
Const EMR_SETICMMODE = 98
Const EMR_CREATECOLORSPACE = 99
Const EMR_SETCOLORSPACE = 100
Const EMR_DELETECOLORSPACE = 101
Const SB_HORZ = 0
Const SB_VERT = 1
Const SB_CTL = 2
Const SB_BOTH = 3
Const SB_LINEUP = 0
Const SB_LINELEFT = 0
Const SB_LINEDOWN = 1
Const SB_LINERIGHT = 1
Const SB_PAGEUP = 2
Const SB_PAGELEFT = 2
Const SB_PAGEDOWN = 3
Const SB_PAGERIGHT = 3
Const SB_THUMBPOSITION = 4
Const SB_THUMBTRACK = 5
Const SB_TOP = 6
Const SB_LEFT = 6
Const SB_BOTTOM = 7
Const SB_RIGHT = 7
Const SB_ENDSCROLL = 8
Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1
Const SW_NORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
Const SW_MAXIMIZE = 3
Const SW_SHOWNOACTIVATE = 4
Const SW_SHOW = 5
Const SW_MINIMIZE = 6
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNA = 8
Const SW_RESTORE = 9
Const SW_SHOWDEFAULT = 10
Const SW_MAX = 10
Const HIDE_WINDOW = 0
Const SHOW_OPENWINDOW = 1
Const SHOW_ICONWINDOW = 2
Const SHOW_FULLSCREEN = 3
Const SHOW_OPENNOACTIVATE = 4
Const SW_PARENTCLOSING = 1
Const SW_OTHERZOOM = 2
Const SW_PARENTOPENING = 3
Const SW_OTHERUNZOOM = 4
Const KF_EXTENDED = &H100
Const KF_DLGMODE = &H800
Const KF_MENUMODE = &H1000
Const KF_ALTDOWN = &H2000
Const KF_REPEAT = &H4000
Const KF_UP = &H8000
Const VK_LBUTTON = &H1
Const VK_RBUTTON = &H2
Const VK_CANCEL = &H3
Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON
Const VK_BACK = &H8
Const VK_TAB = &H9
Const VK_CLEAR = &HC
Const VK_RETURN = &HD
Const VK_SHIFT = &H10
Const VK_CONTROL = &H11
Const VK_MENU = &H12
Const VK_PAUSE = &H13
Const VK_CAPITAL = &H14
Const VK_ESCAPE = &H1B
Const VK_SPACE = &H20
Const VK_PRIOR = &H21
Const VK_NEXT = &H22
Const VK_END = &H23
Const VK_HOME = &H24
Const VK_LEFT = &H25
Const VK_UP = &H26
Const VK_RIGHT = &H27
Const VK_DOWN = &H28
Const VK_SELECT = &H29
Const VK_PRINT = &H2A
Const VK_EXECUTE = &H2B
Const VK_SNAPSHOT = &H2C
Const VK_INSERT = &H2D
Const VK_DELETE = &H2E
Const VK_HELP = &H2F
Const VK_NUMPAD0 = &H60
Const VK_NUMPAD1 = &H61
Const VK_NUMPAD2 = &H62
Const VK_NUMPAD3 = &H63
Const VK_NUMPAD4 = &H64
Const VK_NUMPAD5 = &H65
Const VK_NUMPAD6 = &H66
Const VK_NUMPAD7 = &H67
Const VK_NUMPAD8 = &H68
Const VK_NUMPAD9 = &H69
Const VK_MULTIPLY = &H6A
Const VK_ADD = &H6B
Const VK_SEPARATOR = &H6C
Const VK_SUBTRACT = &H6D
Const VK_DECIMAL = &H6E
Const VK_DIVIDE = &H6F
Const VK_F1 = &H70
Const VK_F2 = &H71
Const VK_F3 = &H72
Const VK_F4 = &H73
Const VK_F5 = &H74
Const VK_F6 = &H75
Const VK_F7 = &H76
Const VK_F8 = &H77
Const VK_F9 = &H78
Const VK_F10 = &H79
Const VK_F11 = &H7A
Const VK_F12 = &H7B
Const VK_F13 = &H7C
Const VK_F14 = &H7D
Const VK_F15 = &H7E
Const VK_F16 = &H7F
Const VK_F17 = &H80
Const VK_F18 = &H81
Const VK_F19 = &H82
Const VK_F20 = &H83
Const VK_F21 = &H84
Const VK_F22 = &H85
Const VK_F23 = &H86
Const VK_F24 = &H87
Const VK_NUMLOCK = &H90
Const VK_SCROLL = &H91
Const VK_LSHIFT = &HA0
Const VK_RSHIFT = &HA1
Const VK_LCONTROL = &HA2
Const VK_RCONTROL = &HA3
Const VK_LMENU = &HA4
Const VK_RMENU = &HA5
Const VK_ATTN = &HF6
Const VK_CRSEL = &HF7
Const VK_EXSEL = &HF8
Const VK_EREOF = &HF9
Const VK_PLAY = &HFA
Const VK_ZOOM = &HFB
Const VK_NONAME = &HFC
Const VK_PA1 = &HFD
Const VK_OEM_CLEAR = &HFE
Const WH_MIN = (-1)
Const WH_MSGFILTER = (-1)
Const WH_JOURNALRECORD = 0
Const WH_JOURNALPLAYBACK = 1
Const WH_KEYBOARD = 2
Const WH_GETMESSAGE = 3
Const WH_CALLWNDPROC = 4
Const WH_CBT = 5
Const WH_SYSMSGFILTER = 6
Const WH_MOUSE = 7
Const WH_HARDWARE = 8
Const WH_DEBUG = 9
Const WH_SHELL = 10
Const WH_FOREGROUNDIDLE = 11
Const WH_MAX = 11
Const HC_ACTION = 0
Const HC_GETNEXT = 1
Const HC_SKIP = 2
Const HC_NOREMOVE = 3
Const HC_NOREM = HC_NOREMOVE
Const HC_SYSMODALON = 4
Const HC_SYSMODALOFF = 5
Const HCBT_MOVESIZE = 0
Const HCBT_MINMAX = 1
Const HCBT_QS = 2
Const HCBT_CREATEWND = 3
Const HCBT_DESTROYWND = 4
Const HCBT_ACTIVATE = 5
Const HCBT_CLICKSKIPPED = 6
Const HCBT_KEYSKIPPED = 7
Const HCBT_SYSCOMMAND = 8
Const HCBT_SETFOCUS = 9
Const MSGF_DIALOGBOX = 0
Const MSGF_MESSAGEBOX = 1
Const MSGF_MENU = 2
Const MSGF_MOVE = 3
Const MSGF_SIZE = 4
Const MSGF_SCROLLBAR = 5
Const MSGF_NEXTWINDOW = 6
Const MSGF_MAINLOOP = 8
Const MSGF_MAX = 8
Const MSGF_USER = 4096
Const HSHELL_WINDOWCREATED = 1
Const HSHELL_WINDOWDESTROYED = 2
Const HSHELL_ACTIVATESHELLWINDOW = 3
Const HKL_PREV = 0
Const HKL_NEXT = 1
Const KLF_ACTIVATE = &H1
Const KLF_SUBSTITUTE_OK = &H2
Const KLF_UNLOADPREVIOUS = &H4
Const KLF_REORDER = &H8
Const KL_NAMELENGTH = 9
Const DESKTOP_READOBJECTS = &H1&
Const DESKTOP_CREATEWINDOW = &H2&
Const DESKTOP_CREATEMENU = &H4&
Const DESKTOP_HOOKCONTROL = &H8&
Const DESKTOP_JOURNALRECORD = &H10&
Const DESKTOP_JOURNALPLAYBACK = &H20&
Const DESKTOP_ENUMERATE = &H40&
Const DESKTOP_WRITEOBJECTS = &H80&
Const WINSTA_ENUMDESKTOPS = &H1&
Const WINSTA_READATTRIBUTES = &H2&
Const WINSTA_ACCESSCLIPBOARD = &H4&
Const WINSTA_CREATEDESKTOP = &H8&
Const WINSTA_WRITEATTRIBUTES = &H10&
Const WINSTA_ACCESSPUBLICATOMS = &H20&
Const WINSTA_EXITWINDOWS = &H40&
Const WINSTA_ENUMERATE = &H100&
Const WINSTA_READSCREEN = &H200&
Const GWL_WNDPROC = (-4)
Const GWL_HINSTANCE = (-6)
Const GWL_HWNDPARENT = (-8)
Const GWL_STYLE = (-16)
Const GWL_EXSTYLE = (-20)
Const GWL_USERDATA = (-21)
Const GWL_ID = (-12)
Const GCL_MENUNAME = (-8)
Const GCL_HBRBACKGROUND = (-10)
Const GCL_HCURSOR = (-12)
Const GCL_HICON = (-14)
Const GCL_HMODULE = (-16)
Const GCL_CBWNDEXTRA = (-18)
Const GCL_CBCLSEXTRA = (-20)
Const GCL_WNDPROC = (-24)
Const GCL_STYLE = (-26)
Const GCW_ATOM = (-32)
Const WM_NULL = &H0
Const WM_CREATE = &H1
Const WM_DESTROY = &H2
Const WM_MOVE = &H3
Const WM_SIZE = &H5
Const WM_ACTIVATE = &H6
Const WA_INACTIVE = 0
Const WA_ACTIVE = 1
Const WA_CLICKACTIVE = 2
Const WM_SETFOCUS = &H7
Const WM_KILLFOCUS = &H8
Const WM_ENABLE = &HA
Const WM_SETREDRAW = &HB
Const WM_SETTEXT = &HC
Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE
Const WM_PAINT = &HF
Const WM_CLOSE = &H10
Const WM_QUERYENDSESSION = &H11
Const WM_QUIT = &H12
Const WM_QUERYOPEN = &H13
Const WM_ERASEBKGND = &H14
Const WM_SYSCOLORCHANGE = &H15
Const WM_ENDSESSION = &H16
Const WM_SHOWWINDOW = &H18
Const WM_WININICHANGE = &H1A
Const WM_DEVMODECHANGE = &H1B
Const WM_ACTIVATEAPP = &H1C
Const WM_FONTCHANGE = &H1D
Const WM_TIMECHANGE = &H1E
Const WM_CANCELMODE = &H1F
Const WM_SETCURSOR = &H20
Const WM_MOUSEACTIVATE = &H21
Const WM_CHILDACTIVATE = &H22
Const WM_QUEUESYNC = &H23
Const WM_GETMINMAXINFO = &H24
Const WM_PAINTICON = &H26
Const WM_ICONERASEBKGND = &H27
Const WM_NEXTDLGCTL = &H28
Const WM_SPOOLERSTATUS = &H2A
Const WM_DRAWITEM = &H2B
Const WM_MEASUREITEM = &H2C
Const WM_DELETEITEM = &H2D
Const WM_VKEYTOITEM = &H2E
Const WM_CHARTOITEM = &H2F
Const WM_SETFONT = &H30
Const WM_GETFONT = &H31
Const WM_SETHOTKEY = &H32
Const WM_GETHOTKEY = &H33
Const WM_QUERYDRAGICON = &H37
Const WM_COMPAREITEM = &H39
Const WM_COMPACTING = &H41
Const WM_OTHERWINDOWCREATED = &H42               '  no longer suported
Const WM_OTHERWINDOWDESTROYED = &H43             '  no longer suported
Const WM_COMMNOTIFY = &H44                       '  no longer suported
Const CN_RECEIVE = &H1
Const CN_TRANSMIT = &H2
Const CN_EVENT = &H4
Const WM_WINDOWPOSCHANGING = &H46
Const WM_WINDOWPOSCHANGED = &H47
Const WM_POWER = &H48
Const PWR_OK = 1
Const PWR_FAIL = (-1)
Const PWR_SUSPENDREQUEST = 1
Const PWR_SUSPENDRESUME = 2
Const PWR_CRITICALRESUME = 3
Const WM_COPYDATA = &H4A
Const WM_CANCELJOURNAL = &H4B
Const WM_NCCREATE = &H81
Const WM_NCDESTROY = &H82
Const WM_NCCALCSIZE = &H83
Const WM_NCHITTEST = &H84
Const WM_NCPAINT = &H85
Const WM_NCACTIVATE = &H86
Const WM_GETDLGCODE = &H87
Const WM_NCMOUSEMOVE = &HA0
Const WM_NCLBUTTONDOWN = &HA1
Const WM_NCLBUTTONUP = &HA2
Const WM_NCLBUTTONDBLCLK = &HA3
Const WM_NCRBUTTONDOWN = &HA4
Const WM_NCRBUTTONUP = &HA5
Const WM_NCRBUTTONDBLCLK = &HA6
Const WM_NCMBUTTONDOWN = &HA7
Const WM_NCMBUTTONUP = &HA8
Const WM_NCMBUTTONDBLCLK = &HA9
Const WM_KEYFIRST = &H100
Const WM_KEYDOWN = &H100
Const WM_KEYUP = &H101
Const WM_CHAR = &H102
Const WM_DEADCHAR = &H103
Const WM_SYSKEYDOWN = &H104
Const WM_SYSKEYUP = &H105
Const WM_SYSCHAR = &H106
Const WM_SYSDEADCHAR = &H107
Const WM_KEYLAST = &H108
Const WM_INITDIALOG = &H110
Const WM_COMMAND = &H111
Const WM_SYSCOMMAND = &H112
Const WM_TIMER = &H113
Const WM_HSCROLL = &H114
Const WM_VSCROLL = &H115
Const WM_INITMENU = &H116
Const WM_INITMENUPOPUP = &H117
Const WM_MENUSELECT = &H11F
Const WM_MENUCHAR = &H120
Const WM_ENTERIDLE = &H121
Const WM_CTLCOLORMSGBOX = &H132
Const WM_CTLCOLOREDIT = &H133
Const WM_CTLCOLORLISTBOX = &H134
Const WM_CTLCOLORBTN = &H135
Const WM_CTLCOLORDLG = &H136
Const WM_CTLCOLORSCROLLBAR = &H137
Const WM_CTLCOLORSTATIC = &H138
Const WM_MOUSEFIRST = &H200
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209
Const WM_MOUSELAST = &H209
Const WM_PARENTNOTIFY = &H210
Const WM_ENTERMENULOOP = &H211
Const WM_EXITMENULOOP = &H212
Const WM_MDICREATE = &H220
Const WM_MDIDESTROY = &H221
Const WM_MDIACTIVATE = &H222
Const WM_MDIRESTORE = &H223
Const WM_MDINEXT = &H224
Const WM_MDIMAXIMIZE = &H225
Const WM_MDITILE = &H226
Const WM_MDICASCADE = &H227
Const WM_MDIICONARRANGE = &H228
Const WM_MDIGETACTIVE = &H229
Const WM_MDISETMENU = &H230
Const WM_DROPFILES = &H233
Const WM_MDIREFRESHMENU = &H234
Const WM_CUT = &H300
Const WM_COPY = &H301
Const WM_PASTE = &H302
Const WM_CLEAR = &H303
Const WM_UNDO = &H304
Const WM_RENDERFORMAT = &H305
Const WM_RENDERALLFORMATS = &H306
Const WM_DESTROYCLIPBOARD = &H307
Const WM_DRAWCLIPBOARD = &H308
Const WM_PAINTCLIPBOARD = &H309
Const WM_VSCROLLCLIPBOARD = &H30A
Const WM_SIZECLIPBOARD = &H30B
Const WM_ASKCBFORMATNAME = &H30C
Const WM_CHANGECBCHAIN = &H30D
Const WM_HSCROLLCLIPBOARD = &H30E
Const WM_QUERYNEWPALETTE = &H30F
Const WM_PALETTEISCHANGING = &H310
Const WM_PALETTECHANGED = &H311
Const WM_HOTKEY = &H312
Const WM_PENWINFIRST = &H380
Const WM_PENWINLAST = &H38F
Const WM_USER = &H400
Const ST_BEGINSWP = 0
Const ST_ENDSWP = 1
Const HTERROR = (-2)
Const HTTRANSPARENT = (-1)
Const HTNOWHERE = 0
Const HTCLIENT = 1
Const HTCAPTION = 2
Const HTSYSMENU = 3
Const HTGROWBOX = 4
Const HTSIZE = HTGROWBOX
Const HTMENU = 5
Const HTHSCROLL = 6
Const HTVSCROLL = 7
Const HTMINBUTTON = 8
Const HTMAXBUTTON = 9
Const HTLEFT = 10
Const HTRIGHT = 11
Const HTTOP = 12
Const HTTOPLEFT = 13
Const HTTOPRIGHT = 14
Const HTBOTTOM = 15
Const HTBOTTOMLEFT = 16
Const HTBOTTOMRIGHT = 17
Const HTBORDER = 18
Const HTREDUCE = HTMINBUTTON
Const HTZOOM = HTMAXBUTTON
Const HTSIZEFIRST = HTLEFT
Const HTSIZELAST = HTBOTTOMRIGHT
Const SMTO_NORMAL = &H0
Const SMTO_BLOCK = &H1
Const SMTO_ABORTIFHUNG = &H2
Const MA_ACTIVATE = 1
Const MA_ACTIVATEANDEAT = 2
Const MA_NOACTIVATE = 3
Const MA_NOACTIVATEANDEAT = 4
Const SIZE_RESTORED = 0
Const SIZE_MINIMIZED = 1
Const SIZE_MAXIMIZED = 2
Const SIZE_MAXSHOW = 3
Const SIZE_MAXHIDE = 4
Const SIZENORMAL = SIZE_RESTORED
Const SIZEICONIC = SIZE_MINIMIZED
Const SIZEFULLSCREEN = SIZE_MAXIMIZED
Const SIZEZOOMSHOW = SIZE_MAXSHOW
Const SIZEZOOMHIDE = SIZE_MAXHIDE
Const WVR_ALIGNTOP = &H10
Const WVR_ALIGNLEFT = &H20
Const WVR_ALIGNBOTTOM = &H40
Const WVR_ALIGNRIGHT = &H80
Const WVR_HREDRAW = &H100
Const WVR_VREDRAW = &H200
Const WVR_REDRAW = (WVR_HREDRAW Or WVR_VREDRAW)
Const WVR_VALIDRECTS = &H400
Const MK_LBUTTON = &H1
Const MK_RBUTTON = &H2
Const MK_SHIFT = &H4
Const MK_CONTROL = &H8
Const MK_MBUTTON = &H10
Const WS_OVERLAPPED = &H0&
Const WS_POPUP = &H80000000
Const WS_CHILD = &H40000000
Const WS_MINIMIZE = &H20000000
Const WS_VISIBLE = &H10000000
Const WS_DISABLED = &H8000000
Const WS_CLIPSIBLINGS = &H4000000
Const WS_CLIPCHILDREN = &H2000000
Const WS_MAXIMIZE = &H1000000
Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Const WS_BORDER = &H800000
Const WS_DLGFRAME = &H400000
Const WS_VSCROLL = &H200000
Const WS_HSCROLL = &H100000
Const WS_SYSMENU = &H80000
Const WS_THICKFRAME = &H40000
Const WS_GROUP = &H20000
Const WS_TABSTOP = &H10000
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const WS_TILED = WS_OVERLAPPED
Const WS_ICONIC = WS_MINIMIZE
Const WS_SIZEBOX = WS_THICKFRAME
Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Const WS_CHILDWINDOW = (WS_CHILD)
Const WS_EX_DLGMODALFRAME = &H1&
Const WS_EX_NOPARENTNOTIFY = &H4&
Const WS_EX_TOPMOST = &H8&
Const WS_EX_ACCEPTFILES = &H10&
Const WS_EX_TRANSPARENT = &H20&
Const CS_VREDRAW = &H1
Const CS_HREDRAW = &H2
Const CS_KEYCVTWINDOW = &H4
Const CS_DBLCLKS = &H8
Const CS_OWNDC = &H20
Const CS_CLASSDC = &H40
Const CS_PARENTDC = &H80
Const CS_NOKEYCVT = &H100
Const CS_NOCLOSE = &H200
Const CS_SAVEBITS = &H800
Const CS_BYTEALIGNCLIENT = &H1000
Const CS_BYTEALIGNWINDOW = &H2000
Const CS_PUBLICCLASS = &H4000
Const CF_TEXT = 1
Const CF_BITMAP = 2
Const CF_METAFILEPICT = 3
Const CF_SYLK = 4
Const CF_DIF = 5
Const CF_TIFF = 6
Const CF_OEMTEXT = 7
Const CF_DIB = 8
Const CF_PALETTE = 9
Const CF_PENDATA = 10
Const CF_RIFF = 11
Const CF_WAVE = 12
Const CF_UNICODETEXT = 13
Const CF_ENHMETAFILE = 14
Const CF_OWNERDISPLAY = &H80
Const CF_DSPTEXT = &H81
Const CF_DSPBITMAP = &H82
Const CF_DSPMETAFILEPICT = &H83
Const CF_DSPENHMETAFILE = &H8E
Const CF_PRIVATEFIRST = &H200
Const CF_PRIVATELAST = &H2FF
Const CF_GDIOBJFIRST = &H300
Const CF_GDIOBJLAST = &H3FF
Const FVIRTKEY = True          '  Assumed to be == TRUE
Const FNOINVERT = &H2
Const FSHIFT = &H4
Const FCONTROL = &H8
Const FALT = &H10
Const WPF_SETMINPOSITION = &H1
Const WPF_RESTORETOMAXIMIZED = &H2
Const ODT_MENU = 1
Const ODT_LISTBOX = 2
Const ODT_COMBOBOX = 3
Const ODT_BUTTON = 4
Const ODA_DRAWENTIRE = &H1
Const ODA_SELECT = &H2
Const ODA_FOCUS = &H4
Const ODS_SELECTED = &H1
Const ODS_GRAYED = &H2
Const ODS_DISABLED = &H4
Const ODS_CHECKED = &H8
Const ODS_FOCUS = &H10
Const PM_NOREMOVE = &H0
Const PM_REMOVE = &H1
Const PM_NOYIELD = &H2
Const MOD_ALT = &H1
Const MOD_CONTROL = &H2
Const MOD_SHIFT = &H4
Const IDHOT_SNAPWINDOW = (-1)    '  SHIFT-PRINTSCRN
Const IDHOT_SNAPDESKTOP = (-2)    '  PRINTSCRN
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Const READAPI = 0        '  Flags for _lopen
Const WRITEAPI = 1
Const READ_WRITE = 2
Const HWND_BROADCAST = &HFFFF&
Const CW_USEDEFAULT = &H80000000
Const HWND_DESKTOP = 0
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Const SWP_NOREDRAW = &H8
Const SWP_NOACTIVATE = &H10
Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOCOPYBITS = &H100
Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const DLGWINDOWEXTRA = 30        '  Window extra bytes needed for private dialog classes
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
Const MOUSEEVENTF_MOVE = &H1    '  mouse move
Const MOUSEEVENTF_LEFTDOWN = &H2    '  left button down
Const MOUSEEVENTF_LEFTUP = &H4    '  left button up
Const MOUSEEVENTF_RIGHTDOWN = &H8    '  right button down
Const MOUSEEVENTF_RIGHTUP = &H10    '  right button up
Const MOUSEEVENTF_MIDDLEDOWN = &H20    '  middle button down
Const MOUSEEVENTF_MIDDLEUP = &H40    '  middle button up
Const MOUSEEVENTF_ABSOLUTE = &H8000    '  absolute move
Const QS_KEY = &H1
Const QS_MOUSEMOVE = &H2
Const QS_MOUSEBUTTON = &H4
Const QS_POSTMESSAGE = &H8
Const QS_TIMER = &H10
Const QS_PAINT = &H20
Const QS_SENDMESSAGE = &H40
Const QS_HOTKEY = &H80
Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Const SM_CXSCREEN = 0
Const SM_CYSCREEN = 1
Const SM_CXVSCROLL = 2
Const SM_CYHSCROLL = 3
Const SM_CYCAPTION = 4
Const SM_CXBORDER = 5
Const SM_CYBORDER = 6
Const SM_CXDLGFRAME = 7
Const SM_CYDLGFRAME = 8
Const SM_CYVTHUMB = 9
Const SM_CXHTHUMB = 10
Const SM_CXICON = 11
Const SM_CYICON = 12
Const SM_CXCURSOR = 13
Const SM_CYCURSOR = 14
Const SM_CYMENU = 15
Const SM_CXFULLSCREEN = 16
Const SM_CYFULLSCREEN = 17
Const SM_CYKANJIWINDOW = 18
Const SM_MOUSEPRESENT = 19
Const SM_CYVSCROLL = 20
Const SM_CXHSCROLL = 21
Const SM_DEBUG = 22
Const SM_SWAPBUTTON = 23
Const SM_RESERVED1 = 24
Const SM_RESERVED2 = 25
Const SM_RESERVED3 = 26
Const SM_RESERVED4 = 27
Const SM_CXMIN = 28
Const SM_CYMIN = 29
Const SM_CXSIZE = 30
Const SM_CYSIZE = 31
Const SM_CXFRAME = 32
Const SM_CYFRAME = 33
Const SM_CXMINTRACK = 34
Const SM_CYMINTRACK = 35
Const SM_CXDOUBLECLK = 36
Const SM_CYDOUBLECLK = 37
Const SM_CXICONSPACING = 38
Const SM_CYICONSPACING = 39
Const SM_MENUDROPALIGNMENT = 40
Const SM_PENWINDOWS = 41
Const SM_DBCSENABLED = 42
Const SM_CMOUSEBUTTONS = 43
Const SM_CMETRICS = 44
Const SM_CXSIZEFRAME = SM_CXFRAME
Const SM_CYSIZEFRAME = SM_CYFRAME
Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Const TPM_LEFTBUTTON = &H0&
Const TPM_RIGHTBUTTON = &H2&
Const TPM_LEFTALIGN = &H0&
Const TPM_CENTERALIGN = &H4&
Const TPM_RIGHTALIGN = &H8&
Const DT_TOP = &H0
Const DT_LEFT = &H0
Const DT_CENTER = &H1
Const DT_RIGHT = &H2
Const DT_VCENTER = &H4
Const DT_BOTTOM = &H8
Const DT_WORDBREAK = &H10
Const DT_SINGLELINE = &H20
Const DT_EXPANDTABS = &H40
Const DT_TABSTOP = &H80
Const DT_NOCLIP = &H100
Const DT_EXTERNALLEADING = &H200
Const DT_CALCRECT = &H400
Const DT_NOPREFIX = &H800
Const DT_INTERNAL = &H1000
Const DCX_WINDOW = &H1&
Const DCX_CACHE = &H2&
Const DCX_NORESETATTRS = &H4&
Const DCX_CLIPCHILDREN = &H8&
Const DCX_CLIPSIBLINGS = &H10&
Const DCX_PARENTCLIP = &H20&
Const DCX_EXCLUDERGN = &H40&
Const DCX_INTERSECTRGN = &H80&
Const DCX_EXCLUDEUPDATE = &H100&
Const DCX_INTERSECTUPDATE = &H200&
Const DCX_LOCKWINDOWUPDATE = &H400&
Const DCX_NORECOMPUTE = &H100000
Const DCX_VALIDATE = &H200000
Const RDW_INVALIDATE = &H1
Const RDW_INTERNALPAINT = &H2
Const RDW_ERASE = &H4
Const RDW_VALIDATE = &H8
Const RDW_NOINTERNALPAINT = &H10
Const RDW_NOERASE = &H20
Const RDW_NOCHILDREN = &H40
Const RDW_ALLCHILDREN = &H80
Const RDW_UPDATENOW = &H100
Const RDW_ERASENOW = &H200
Const RDW_FRAME = &H400
Const RDW_NOFRAME = &H800
Const SW_SCROLLCHILDREN = &H1
Const SW_INVALIDATE = &H2
Const SW_ERASE = &H4
Const ESB_ENABLE_BOTH = &H0
Const ESB_DISABLE_BOTH = &H3
Const ESB_DISABLE_LEFT = &H1
Const ESB_DISABLE_RIGHT = &H2
Const ESB_DISABLE_UP = &H1
Const ESB_DISABLE_DOWN = &H2
Const ESB_DISABLE_LTUP = ESB_DISABLE_LEFT
Const ESB_DISABLE_RTDN = ESB_DISABLE_RIGHT
Const MB_OK = &H0&
Const MB_OKCANCEL = &H1&
Const MB_ABORTRETRYIGNORE = &H2&
Const MB_YESNOCANCEL = &H3&
Const MB_YESNO = &H4&
Const MB_RETRYCANCEL = &H5&
Const MB_ICONHAND = &H10&
Const MB_ICONQUESTION = &H20&
Const MB_ICONEXCLAMATION = &H30&
Const MB_ICONASTERISK = &H40&
Const MB_ICONINFORMATION = MB_ICONASTERISK
Const MB_ICONSTOP = MB_ICONHAND
Const MB_DEFBUTTON1 = &H0&
Const MB_DEFBUTTON2 = &H100&
Const MB_DEFBUTTON3 = &H200&
Const MB_APPLMODAL = &H0&
Const MB_SYSTEMMODAL = &H1000&
Const MB_TASKMODAL = &H2000&
Const MB_NOFOCUS = &H8000&
Const MB_SETFOREGROUND = &H10000
Const MB_DEFAULT_DESKTOP_ONLY = &H20000
Const MB_TYPEMASK = &HF&
Const MB_ICONMASK = &HF0&
Const MB_DEFMASK = &HF00&
Const MB_MODEMASK = &H3000&
Const MB_MISCMASK = &HC000&
Const CTLCOLOR_MSGBOX = 0
Const CTLCOLOR_EDIT = 1
Const CTLCOLOR_LISTBOX = 2
Const CTLCOLOR_BTN = 3
Const CTLCOLOR_DLG = 4
Const CTLCOLOR_SCROLLBAR = 5
Const CTLCOLOR_STATIC = 6
Const CTLCOLOR_MAX = 8   '  three bits max
Const COLOR_SCROLLBAR = 0
Const COLOR_BACKGROUND = 1
Const COLOR_ACTIVECAPTION = 2
Const COLOR_INACTIVECAPTION = 3
Const COLOR_MENU = 4
Const COLOR_WINDOW = 5
Const COLOR_WINDOWFRAME = 6
Const COLOR_MENUTEXT = 7
Const COLOR_WINDOWTEXT = 8
Const COLOR_CAPTIONTEXT = 9
Const COLOR_ACTIVEBORDER = 10
Const COLOR_INACTIVEBORDER = 11
Const COLOR_APPWORKSPACE = 12
Const COLOR_HIGHLIGHT = 13
Const COLOR_HIGHLIGHTTEXT = 14
Const COLOR_BTNFACE = 15
Const COLOR_BTNSHADOW = 16
Const COLOR_GRAYTEXT = 17
Const COLOR_BTNTEXT = 18
Const COLOR_INACTIVECAPTIONTEXT = 19
Const COLOR_BTNHIGHLIGHT = 20
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Const GW_HWNDPREV = 3
Const GW_OWNER = 4
Const GW_CHILD = 5
Const GW_MAX = 5
Const MF_INSERT = &H0&
Const MF_CHANGE = &H80&
Const MF_APPEND = &H100&
Const MF_DELETE = &H200&
Const MF_REMOVE = &H1000&
Const MF_BYCOMMAND = &H0&
Const MF_BYPOSITION = &H400&
Const MF_SEPARATOR = &H800&
Const MF_ENABLED = &H0&
Const MF_GRAYED = &H1&
Const MF_DISABLED = &H2&
Const MF_UNCHECKED = &H0&
Const MF_CHECKED = &H8&
Const MF_USECHECKBITMAPS = &H200&
Const MF_STRING = &H0&
Const MF_BITMAP = &H4&
Const MF_OWNERDRAW = &H100&
Const MF_POPUP = &H10&
Const MF_MENUBARBREAK = &H20&
Const MF_MENUBREAK = &H40&
Const MF_UNHILITE = &H0&
Const MF_HILITE = &H80&
Const MF_SYSMENU = &H2000&
Const MF_HELP = &H4000&
Const MF_MOUSESELECT = &H8000&
Const MF_END = &H80
Const SC_SIZE = &HF000&
Const SC_MOVE = &HF010&
Const SC_MINIMIZE = &HF020&
Const SC_MAXIMIZE = &HF030&
Const SC_NEXTWINDOW = &HF040&
Const SC_PREVWINDOW = &HF050&
Const SC_CLOSE = &HF060&
Const SC_VSCROLL = &HF070&
Const SC_HSCROLL = &HF080&
Const SC_MOUSEMENU = &HF090&
Const SC_KEYMENU = &HF100&
Const SC_ARRANGE = &HF110&
Const SC_RESTORE = &HF120&
Const SC_TASKLIST = &HF130&
Const SC_SCREENSAVE = &HF140&
Const SC_HOTKEY = &HF150&
Const SC_ICON = SC_MINIMIZE
Const SC_ZOOM = SC_MAXIMIZE
Const IDC_ARROW = 32512&
Const IDC_IBEAM = 32513&
Const IDC_WAIT = 32514&
Const IDC_CROSS = 32515&
Const IDC_UPARROW = 32516&
Const IDC_SIZE = 32640&
Const IDC_ICON = 32641&
Const IDC_SIZENWSE = 32642&
Const IDC_SIZENESW = 32643&
Const IDC_SIZEWE = 32644&
Const IDC_SIZENS = 32645&
Const IDC_SIZEALL = 32646&
Const IDC_NO = 32648&
Const IDC_APPSTARTING = 32650&
Const OBM_CLOSE = 32754
Const OBM_UPARROW = 32753
Const OBM_DNARROW = 32752
Const OBM_RGARROW = 32751
Const OBM_LFARROW = 32750
Const OBM_REDUCE = 32749
Const OBM_ZOOM = 32748
Const OBM_RESTORE = 32747
Const OBM_REDUCED = 32746
Const OBM_ZOOMD = 32745
Const OBM_RESTORED = 32744
Const OBM_UPARROWD = 32743
Const OBM_DNARROWD = 32742
Const OBM_RGARROWD = 32741
Const OBM_LFARROWD = 32740
Const OBM_MNARROW = 32739
Const OBM_COMBO = 32738
Const OBM_UPARROWI = 32737
Const OBM_DNARROWI = 32736
Const OBM_RGARROWI = 32735
Const OBM_LFARROWI = 32734
Const OBM_OLD_CLOSE = 32767
Const OBM_SIZE = 32766
Const OBM_OLD_UPARROW = 32765
Const OBM_OLD_DNARROW = 32764
Const OBM_OLD_RGARROW = 32763
Const OBM_OLD_LFARROW = 32762
Const OBM_BTSIZE = 32761
Const OBM_CHECK = 32760
Const OBM_CHECKBOXES = 32759
Const OBM_BTNCORNERS = 32758
Const OBM_OLD_REDUCE = 32757
Const OBM_OLD_ZOOM = 32756
Const OBM_OLD_RESTORE = 32755
Const OCR_NORMAL = 32512
Const OCR_IBEAM = 32513
Const OCR_WAIT = 32514
Const OCR_CROSS = 32515
Const OCR_UP = 32516
Const OCR_SIZE = 32640
Const OCR_ICON = 32641
Const OCR_SIZENWSE = 32642
Const OCR_SIZENESW = 32643
Const OCR_SIZEWE = 32644
Const OCR_SIZENS = 32645
Const OCR_SIZEALL = 32646
Const OCR_ICOCUR = 32647
Const OCR_NO = 32648    ' not in win3.1
Const OIC_SAMPLE = 32512
Const OIC_HAND = 32513
Const OIC_QUES = 32514
Const OIC_BANG = 32515
Const OIC_NOTE = 32516
Const ORD_LANGDRIVER = 1    '  The ordinal number for the entry point of
Const IDI_APPLICATION = 32512&
Const IDI_HAND = 32513&
Const IDI_QUESTION = 32514&
Const IDI_EXCLAMATION = 32515&
Const IDI_ASTERISK = 32516&
Const IDOK = 1
Const IDCANCEL = 2
Const IDABORT = 3
Const IDRETRY = 4
Const IDIGNORE = 5
Const IDYES = 6
Const IDNO = 7
Const ES_LEFT = &H0&
Const ES_CENTER = &H1&
Const ES_RIGHT = &H2&
Const ES_MULTILINE = &H4&
Const ES_UPPERCASE = &H8&
Const ES_LOWERCASE = &H10&
Const ES_PASSWORD = &H20&
Const ES_AUTOVSCROLL = &H40&
Const ES_AUTOHSCROLL = &H80&
Const ES_NOHIDESEL = &H100&
Const ES_OEMCONVERT = &H400&
Const ES_READONLY = &H800&
Const ES_WANTRETURN = &H1000&
Const EN_SETFOCUS = &H100
Const EN_KILLFOCUS = &H200
Const EN_CHANGE = &H300
Const EN_UPDATE = &H400
Const EN_ERRSPACE = &H500
Const EN_MAXTEXT = &H501
Const EN_HSCROLL = &H601
Const EN_VSCROLL = &H602
Const EM_GETSEL = &HB0
Const EM_SETSEL = &HB1
Const EM_GETRECT = &HB2
Const EM_SETRECT = &HB3
Const EM_SETRECTNP = &HB4
Const EM_SCROLL = &HB5
Const EM_LINESCROLL = &HB6
Const EM_SCROLLCARET = &HB7
Const EM_GETMODIFY = &HB8
Const EM_SETMODIFY = &HB9
Const EM_GETLINECOUNT = &HBA
Const EM_LINEINDEX = &HBB
Const EM_SETHANDLE = &HBC
Const EM_GETHANDLE = &HBD
Const EM_GETTHUMB = &HBE
Const EM_LINELENGTH = &HC1
Const EM_REPLACESEL = &HC2
Const EM_GETLINE = &HC4
Const EM_LIMITTEXT = &HC5
Const EM_CANUNDO = &HC6
Const EM_UNDO = &HC7
Const EM_FMTLINES = &HC8
Const EM_LINEFROMCHAR = &HC9
Const EM_SETTABSTOPS = &HCB
Const EM_SETPASSWORDCHAR = &HCC
Const EM_EMPTYUNDOBUFFER = &HCD
Const EM_GETFIRSTVISIBLELINE = &HCE
Const EM_SETREADONLY = &HCF
Const EM_SETWORDBREAKPROC = &HD0
Const EM_GETWORDBREAKPROC = &HD1
Const EM_GETPASSWORDCHAR = &HD2
Const WB_LEFT = 0
Const WB_RIGHT = 1
Const WB_ISDELIMITER = 2
Const BS_PUSHBUTTON = &H0&
Const BS_DEFPUSHBUTTON = &H1&
Const BS_CHECKBOX = &H2&
Const BS_AUTOCHECKBOX = &H3&
Const BS_RADIOBUTTON = &H4&
Const BS_3STATE = &H5&
Const BS_AUTO3STATE = &H6&
Const BS_GROUPBOX = &H7&
Const BS_USERBUTTON = &H8&
Const BS_AUTORADIOBUTTON = &H9&
Const BS_OWNERDRAW = &HB&
Const BS_LEFTTEXT = &H20&
Const BN_CLICKED = 0
Const BN_PAINT = 1
Const BN_HILITE = 2
Const BN_UNHILITE = 3
Const BN_DISABLE = 4
Const BN_DOUBLECLICKED = 5
Const BM_GETCHECK = &HF0
Const BM_SETCHECK = &HF1
Const BM_GETSTATE = &HF2
Const BM_SETSTATE = &HF3
Const BM_SETSTYLE = &HF4
Const SS_LEFT = &H0&
Const SS_CENTER = &H1&
Const SS_RIGHT = &H2&
Const SS_ICON = &H3&
Const SS_BLACKRECT = &H4&
Const SS_GRAYRECT = &H5&
Const SS_WHITERECT = &H6&
Const SS_BLACKFRAME = &H7&
Const SS_GRAYFRAME = &H8&
Const SS_WHITEFRAME = &H9&
Const SS_USERITEM = &HA&
Const SS_SIMPLE = &HB&
Const SS_LEFTNOWORDWRAP = &HC&
Const SS_NOPREFIX = &H80           '  Don't do "&" character translation
Const STM_SETICON = &H170
Const STM_GETICON = &H171
Const STM_MSGMAX = &H172
Const WC_DIALOG = 8002&
Const DWL_MSGRESULT = 0
Const DWL_DLGPROC = 4
Const DWL_USER = 8
Const DDL_READWRITE = &H0
Const DDL_READONLY = &H1
Const DDL_HIDDEN = &H2
Const DDL_SYSTEM = &H4
Const DDL_DIRECTORY = &H10
Const DDL_ARCHIVE = &H20
Const DDL_POSTMSGS = &H2000
Const DDL_DRIVES = &H4000
Const DDL_EXCLUSIVE = &H8000
Const DS_ABSALIGN = &H1&
Const DS_SYSMODAL = &H2&
Const DS_LOCALEDIT = &H20          '  Edit items get Local storage.
Const DS_SETFONT = &H40            '  User specified font for Dlg controls
Const DS_MODALFRAME = &H80         '  Can be combined with WS_CAPTION
Const DS_NOIDLEMSG = &H100         '  WM_ENTERIDLE message will not be sent
Const DS_SETFOREGROUND = &H200     '  not in win3.1
Const DM_GETDEFID = WM_USER + 0
Const DM_SETDEFID = WM_USER + 1
Const DC_HASDEFID = &H534      '0x534B
Const DLGC_WANTARROWS = &H1              '  Control wants arrow keys
Const DLGC_WANTTAB = &H2                 '  Control wants tab keys
Const DLGC_WANTALLKEYS = &H4             '  Control wants all keys
Const DLGC_WANTMESSAGE = &H4             '  Pass message to control
Const DLGC_HASSETSEL = &H8               '  Understands EM_SETSEL message
Const DLGC_DEFPUSHBUTTON = &H10          '  Default pushbutton
Const DLGC_UNDEFPUSHBUTTON = &H20        '  Non-default pushbutton
Const DLGC_RADIOBUTTON = &H40            '  Radio button
Const DLGC_WANTCHARS = &H80              '  Want WM_CHAR messages
Const DLGC_STATIC = &H100                '  Static item: don't include
Const DLGC_BUTTON = &H2000               '  Button item: can be checked
Const LB_CTLCODE = 0&
Const LB_OKAY = 0
Const LB_ERR = (-1)
Const LB_ERRSPACE = (-2)
Const LBN_ERRSPACE = (-2)
Const LBN_SELCHANGE = 1
Const LBN_DBLCLK = 2
Const LBN_SELCANCEL = 3
Const LBN_SETFOCUS = 4
Const LBN_KILLFOCUS = 5
Const LB_ADDSTRING = &H180
Const LB_INSERTSTRING = &H181
Const LB_DELETESTRING = &H182
Const LB_SELITEMRANGEEX = &H183
Const LB_RESETCONTENT = &H184
Const LB_SETSEL = &H185
Const LB_SETCURSEL = &H186
Const LB_GETSEL = &H187
Const LB_GETCURSEL = &H188
Const LB_GETTEXT = &H189
Const LB_GETTEXTLEN = &H18A
Const LB_GETCOUNT = &H18B
Const LB_SELECTSTRING = &H18C
Const LB_DIR = &H18D
Const LB_GETTOPINDEX = &H18E
Const LB_FINDSTRING = &H18F
Const LB_GETSELCOUNT = &H190
Const LB_GETSELITEMS = &H191
Const LB_SETTABSTOPS = &H192
Const LB_GETHORIZONTALEXTENT = &H193
Const LB_SETHORIZONTALEXTENT = &H194
Const LB_SETCOLUMNWIDTH = &H195
Const LB_ADDFILE = &H196
Const LB_SETTOPINDEX = &H197
Const LB_GETITEMRECT = &H198
Const LB_GETITEMDATA = &H199
Const LB_SETITEMDATA = &H19A
Const LB_SELITEMRANGE = &H19B
Const LB_SETANCHORINDEX = &H19C
Const LB_GETANCHORINDEX = &H19D
Const LB_SETCARETINDEX = &H19E
Const LB_GETCARETINDEX = &H19F
Const LB_SETITEMHEIGHT = &H1A0
Const LB_GETITEMHEIGHT = &H1A1
Const LB_FINDSTRINGEXACT = &H1A2
Const LB_SETLOCALE = &H1A5
Const LB_GETLOCALE = &H1A6
Const LB_SETCOUNT = &H1A7
Const LB_MSGMAX = &H1A8
Const LBS_NOTIFY = &H1&
Const LBS_SORT = &H2&
Const LBS_NOREDRAW = &H4&
Const LBS_MULTIPLESEL = &H8&
Const LBS_OWNERDRAWFIXED = &H10&
Const LBS_OWNERDRAWVARIABLE = &H20&
Const LBS_HASSTRINGS = &H40&
Const LBS_USETABSTOPS = &H80&
Const LBS_NOINTEGRALHEIGHT = &H100&
Const LBS_MULTICOLUMN = &H200&
Const LBS_WANTKEYBOARDINPUT = &H400&
Const LBS_EXTENDEDSEL = &H800&
Const LBS_DISABLENOSCROLL = &H1000&
Const LBS_NODATA = &H2000&
Const LBS_STANDARD = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)
Const CB_OKAY = 0
Const CB_ERR = (-1)
Const CB_ERRSPACE = (-2)
Const CBN_ERRSPACE = (-1)
Const CBN_SELCHANGE = 1
Const CBN_DBLCLK = 2
Const CBN_SETFOCUS = 3
Const CBN_KILLFOCUS = 4
Const CBN_EDITCHANGE = 5
Const CBN_EDITUPDATE = 6
Const CBN_DROPDOWN = 7
Const CBN_CLOSEUP = 8
Const CBN_SELENDOK = 9
Const CBN_SELENDCANCEL = 10
Const CBS_SIMPLE = &H1&
Const CBS_DROPDOWN = &H2&
Const CBS_DROPDOWNLIST = &H3&
Const CBS_OWNERDRAWFIXED = &H10&
Const CBS_OWNERDRAWVARIABLE = &H20&
Const CBS_AUTOHSCROLL = &H40&
Const CBS_OEMCONVERT = &H80&
Const CBS_SORT = &H100&
Const CBS_HASSTRINGS = &H200&
Const CBS_NOINTEGRALHEIGHT = &H400&
Const CBS_DISABLENOSCROLL = &H800&
Const CB_GETEDITSEL = &H140
Const CB_LIMITTEXT = &H141
Const CB_SETEDITSEL = &H142
Const CB_ADDSTRING = &H143
Const CB_DELETESTRING = &H144
Const CB_DIR = &H145
Const CB_GETCOUNT = &H146
Const CB_GETCURSEL = &H147
Const CB_GETLBTEXT = &H148
Const CB_GETLBTEXTLEN = &H149
Const CB_INSERTSTRING = &H14A
Const CB_RESETCONTENT = &H14B
Const CB_FINDSTRING = &H14C
Const CB_SELECTSTRING = &H14D
Const CB_SETCURSEL = &H14E
Const CB_SHOWDROPDOWN = &H14F
Const CB_GETITEMDATA = &H150
Const CB_SETITEMDATA = &H151
Const CB_GETDROPPEDCONTROLRECT = &H152
Const CB_SETITEMHEIGHT = &H153
Const CB_GETITEMHEIGHT = &H154
Const CB_SETEXTENDEDUI = &H155
Const CB_GETEXTENDEDUI = &H156
Const CB_GETDROPPEDSTATE = &H157
Const CB_FINDSTRINGEXACT = &H158
Const CB_SETLOCALE = &H159
Const CB_GETLOCALE = &H15A
Const CB_MSGMAX = &H15B
Const SBS_HORZ = &H0&
Const SBS_VERT = &H1&
Const SBS_TOPALIGN = &H2&
Const SBS_LEFTALIGN = &H2&
Const SBS_BOTTOMALIGN = &H4&
Const SBS_RIGHTALIGN = &H4&
Const SBS_SIZEBOXTOPLEFTALIGN = &H2&
Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
Const SBS_SIZEBOX = &H8&
Const SBM_SETPOS = &HE0    ' not in win3.1
Const SBM_GETPOS = &HE1    ' not in win3.1
Const SBM_SETRANGE = &HE2    ' not in win3.1
Const SBM_SETRANGEREDRAW = &HE6    ' not in win3.1
Const SBM_GETRANGE = &HE3    ' not in win3.1
Const SBM_ENABLE_ARROWS = &HE4    ' not in win3.1
Const MDIS_ALLCHILDSTYLES = &H1
Const MDITILE_VERTICAL = &H0
Const MDITILE_HORIZONTAL = &H1
Const MDITILE_SKIPDISABLED = &H2
Const HELP_CONTEXT = &H1          '  Display topic in ulTopic
Const HELP_QUIT = &H2             '  Terminate help
Const HELP_INDEX = &H3            '  Display index
Const HELP_CONTENTS = &H3&
Const HELP_HELPONHELP = &H4       '  Display help on using help
Const HELP_SETINDEX = &H5         '  Set current Index for multi index help
Const HELP_SETCONTENTS = &H5&
Const HELP_CONTEXTPOPUP = &H8&
Const HELP_FORCEFILE = &H9&
Const HELP_KEY = &H101            '  Display topic for keyword in offabData
Const HELP_COMMAND = &H102&
Const HELP_PARTIALKEY = &H105&
Const HELP_MULTIKEY = &H201&
Const HELP_SETWINPOS = &H203&
Const SPI_GETBEEP = 1
Const SPI_SETBEEP = 2
Const SPI_GETMOUSE = 3
Const SPI_SETMOUSE = 4
Const SPI_GETBORDER = 5
Const SPI_SETBORDER = 6
Const SPI_GETKEYBOARDSPEED = 10
Const SPI_SETKEYBOARDSPEED = 11
Const SPI_LANGDRIVER = 12
Const SPI_ICONHORIZONTALSPACING = 13
Const SPI_GETSCREENSAVETIMEOUT = 14
Const SPI_SETSCREENSAVETIMEOUT = 15
Const SPI_GETSCREENSAVEACTIVE = 16
Const SPI_SETSCREENSAVEACTIVE = 17
Const SPI_GETGRIDGRANULARITY = 18
Const SPI_SETGRIDGRANULARITY = 19
Const SPI_SETDESKWALLPAPER = 20
Const SPI_SETDESKPATTERN = 21
Const SPI_GETKEYBOARDDELAY = 22
Const SPI_SETKEYBOARDDELAY = 23
Const SPI_ICONVERTICALSPACING = 24
Const SPI_GETICONTITLEWRAP = 25
Const SPI_SETICONTITLEWRAP = 26
Const SPI_GETMENUDROPALIGNMENT = 27
Const SPI_SETMENUDROPALIGNMENT = 28
Const SPI_SETDOUBLECLKWIDTH = 29
Const SPI_SETDOUBLECLKHEIGHT = 30
Const SPI_GETICONTITLELOGFONT = 31
Const SPI_SETDOUBLECLICKTIME = 32
Const SPI_SETMOUSEBUTTONSWAP = 33
Const SPI_SETICONTITLELOGFONT = 34
Const SPI_GETFASTTASKSWITCH = 35
Const SPI_SETFASTTASKSWITCH = 36
Const SPI_SETDRAGFULLWINDOWS = 37
Const SPI_GETDRAGFULLWINDOWS = 38
Const SPI_GETNONCLIENTMETRICS = 41
Const SPI_SETNONCLIENTMETRICS = 42
Const SPI_GETMINIMIZEDMETRICS = 43
Const SPI_SETMINIMIZEDMETRICS = 44
Const SPI_GETICONMETRICS = 45
Const SPI_SETICONMETRICS = 46
Const SPI_SETWORKAREA = 47
Const SPI_GETWORKAREA = 48
Const SPI_SETPENWINDOWS = 49
Const SPI_GETFILTERKEYS = 50
Const SPI_SETFILTERKEYS = 51
Const SPI_GETTOGGLEKEYS = 52
Const SPI_SETTOGGLEKEYS = 53
Const SPI_GETMOUSEKEYS = 54
Const SPI_SETMOUSEKEYS = 55
Const SPI_GETSHOWSOUNDS = 56
Const SPI_SETSHOWSOUNDS = 57
Const SPI_GETSTICKYKEYS = 58
Const SPI_SETSTICKYKEYS = 59
Const SPI_GETACCESSTIMEOUT = 60
Const SPI_SETACCESSTIMEOUT = 61
Const SPI_GETSERIALKEYS = 62
Const SPI_SETSERIALKEYS = 63
Const SPI_GETSOUNDSENTRY = 64
Const SPI_SETSOUNDSENTRY = 65
Const SPI_GETHIGHCONTRAST = 66
Const SPI_SETHIGHCONTRAST = 67
Const SPI_GETKEYBOARDPREF = 68
Const SPI_SETKEYBOARDPREF = 69
Const SPI_GETSCREENREADER = 70
Const SPI_SETSCREENREADER = 71
Const SPI_GETANIMATION = 72
Const SPI_SETANIMATION = 73
Const SPI_GETFONTSMOOTHING = 74
Const SPI_SETFONTSMOOTHING = 75
Const SPI_SETDRAGWIDTH = 76
Const SPI_SETDRAGHEIGHT = 77
Const SPI_SETHANDHELD = 78
Const SPI_GETLOWPOWERTIMEOUT = 79
Const SPI_GETPOWEROFFTIMEOUT = 80
Const SPI_SETLOWPOWERTIMEOUT = 81
Const SPI_SETPOWEROFFTIMEOUT = 82
Const SPI_GETLOWPOWERACTIVE = 83
Const SPI_GETPOWEROFFACTIVE = 84
Const SPI_SETLOWPOWERACTIVE = 85
Const SPI_SETPOWEROFFACTIVE = 86
Const SPI_SETCURSORS = 87
Const SPI_SETICONS = 88
Const SPI_GETDEFAULTINPUTLANG = 89
Const SPI_SETDEFAULTINPUTLANG = 90
Const SPI_SETLANGTOGGLE = 91
Const SPI_GETWINDOWSEXTENSION = 92
Const SPI_SETMOUSETRAILS = 93
Const SPI_GETMOUSETRAILS = 94
Const SPI_SCREENSAVERRUNNING = 97
Const SPIF_UPDATEINIFILE = &H1
Const SPIF_SENDWININICHANGE = &H2
Const WM_DDE_FIRST = &H3E0
Const WM_DDE_INITIATE = (WM_DDE_FIRST)
Const WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
Const WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
Const WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
Const WM_DDE_ACK = (WM_DDE_FIRST + 4)
Const WM_DDE_DATA = (WM_DDE_FIRST + 5)
Const WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
Const WM_DDE_POKE = (WM_DDE_FIRST + 7)
Const WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
Const WM_DDE_LAST = (WM_DDE_FIRST + 8)
Const XST_NULL = 0  '  quiescent states
Const XST_INCOMPLETE = 1
Const XST_CONNECTED = 2
Const XST_INIT1 = 3  '  mid-initiation states
Const XST_INIT2 = 4
Const XST_REQSENT = 5  '  active conversation states
Const XST_DATARCVD = 6
Const XST_POKESENT = 7
Const XST_POKEACKRCVD = 8
Const XST_EXECSENT = 9
Const XST_EXECACKRCVD = 10
Const XST_ADVSENT = 11
Const XST_UNADVSENT = 12
Const XST_ADVACKRCVD = 13
Const XST_UNADVACKRCVD = 14
Const XST_ADVDATASENT = 15
Const XST_ADVDATAACKRCVD = 16
Const CADV_LATEACK = &HFFFF
Const ST_CONNECTED = &H1
Const ST_ADVISE = &H2
Const ST_ISLOCAL = &H4
Const ST_BLOCKED = &H8
Const ST_CLIENT = &H10
Const ST_TERMINATED = &H20
Const ST_INLIST = &H40
Const ST_BLOCKNEXT = &H80
Const ST_ISSELF = &H100
Const DDE_FACK = &H8000
Const DDE_FBUSY = &H4000
Const DDE_FDEFERUPD = &H4000
Const DDE_FACKREQ = &H8000
Const DDE_FRELEASE = &H2000
Const DDE_FREQUESTED = &H1000
Const DDE_FAPPSTATUS = &HFF
Const DDE_FNOTPROCESSED = &H0
Const DDE_FACKRESERVED = (Not (DDE_FACK Or DDE_FBUSY Or DDE_FAPPSTATUS))
Const DDE_FADVRESERVED = (Not (DDE_FACKREQ Or DDE_FDEFERUPD))
Const DDE_FDATRESERVED = (Not (DDE_FACKREQ Or DDE_FRELEASE Or DDE_FREQUESTED))
Const DDE_FPOKRESERVED = (Not (DDE_FRELEASE))
Const MSGF_DDEMGR = &H8001
Const CP_WINANSI = 1004  '  default codepage for windows old DDE convs.
Const CP_WINUNICODE = 1200
Const XTYPF_NOBLOCK = &H2     '  CBR_BLOCK will not work
Const XTYPF_NODATA = &H4     '  DDE_FDEFERUPD
Const XTYPF_ACKREQ = &H8     '  DDE_FACKREQ
Const XCLASS_MASK = &HFC00
Const XCLASS_BOOL = &H1000
Const XCLASS_DATA = &H2000
Const XCLASS_FLAGS = &H4000
Const XCLASS_NOTIFICATION = &H8000
Const XTYP_ERROR = (&H0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Const XTYP_ADVDATA = (&H10 Or XCLASS_FLAGS)
Const XTYP_ADVREQ = (&H20 Or XCLASS_DATA Or XTYPF_NOBLOCK)
Const XTYP_ADVSTART = (&H30 Or XCLASS_BOOL)
Const XTYP_ADVSTOP = (&H40 Or XCLASS_NOTIFICATION)
Const XTYP_EXECUTE = (&H50 Or XCLASS_FLAGS)
Const XTYP_CONNECT = (&H60 Or XCLASS_BOOL Or XTYPF_NOBLOCK)
Const XTYP_CONNECT_CONFIRM = (&H70 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Const XTYP_XACT_COMPLETE = (&H80 Or XCLASS_NOTIFICATION)
Const XTYP_POKE = (&H90 Or XCLASS_FLAGS)
Const XTYP_REGISTER = (&HA0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Const XTYP_REQUEST = (&HB0 Or XCLASS_DATA)
Const XTYP_DISCONNECT = (&HC0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Const XTYP_UNREGISTER = (&HD0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Const XTYP_WILDCONNECT = (&HE0 Or XCLASS_DATA Or XTYPF_NOBLOCK)
Const XTYP_MASK = &HF0
Const XTYP_SHIFT = 4  '  shift to turn XTYP_ into an index
Const TIMEOUT_ASYNC = &HFFFF
Const QID_SYNC = &HFFFF
Const SZDDESYS_TOPIC = "System"
Const SZDDESYS_ITEM_TOPICS = "Topics"
Const SZDDESYS_ITEM_SYSITEMS = "SysItems"
Const SZDDESYS_ITEM_RTNMSG = "ReturnMessage"
Const SZDDESYS_ITEM_STATUS = "Status"
Const SZDDESYS_ITEM_FORMATS = "Formats"
Const SZDDESYS_ITEM_HELP = "Help"
Const SZDDE_ITEM_ITEMLIST = "TopicItemList"
Const CBR_BLOCK = &HFFFF
Const CBF_FAIL_SELFCONNECTIONS = &H1000
Const CBF_FAIL_CONNECTIONS = &H2000
Const CBF_FAIL_ADVISES = &H4000
Const CBF_FAIL_EXECUTES = &H8000
Const CBF_FAIL_POKES = &H10000
Const CBF_FAIL_REQUESTS = &H20000
Const CBF_FAIL_ALLSVRXACTIONS = &H3F000
Const CBF_SKIP_CONNECT_CONFIRMS = &H40000
Const CBF_SKIP_REGISTRATIONS = &H80000
Const CBF_SKIP_UNREGISTRATIONS = &H100000
Const CBF_SKIP_DISCONNECTS = &H200000
Const CBF_SKIP_ALLNOTIFICATIONS = &H3C0000
Const APPCMD_CLIENTONLY = &H10&
Const APPCMD_FILTERINITS = &H20&
Const APPCMD_MASK = &HFF0&
Const APPCLASS_STANDARD = &H0&
Const APPCLASS_MASK = &HF&
Const EC_ENABLEALL = 0
Const EC_ENABLEONE = ST_BLOCKNEXT
Const EC_DISABLE = ST_BLOCKED
Const EC_QUERYWAITING = 2
Const DNS_REGISTER = &H1
Const DNS_UNREGISTER = &H2
Const DNS_FILTERON = &H4
Const DNS_FILTEROFF = &H8
Const HDATA_APPOWNED = &H1
Const DMLERR_NO_ERROR = 0                           '  must be 0
Const DMLERR_FIRST = &H4000
Const DMLERR_ADVACKTIMEOUT = &H4000
Const DMLERR_BUSY = &H4001
Const DMLERR_DATAACKTIMEOUT = &H4002
Const DMLERR_DLL_NOT_INITIALIZED = &H4003
Const DMLERR_DLL_USAGE = &H4004
Const DMLERR_EXECACKTIMEOUT = &H4005
Const DMLERR_INVALIDPARAMETER = &H4006
Const DMLERR_LOW_MEMORY = &H4007
Const DMLERR_MEMORY_ERROR = &H4008
Const DMLERR_NOTPROCESSED = &H4009
Const DMLERR_NO_CONV_ESTABLISHED = &H400A
Const DMLERR_POKEACKTIMEOUT = &H400B
Const DMLERR_POSTMSG_FAILED = &H400C
Const DMLERR_REENTRANCY = &H400D
Const DMLERR_SERVER_DIED = &H400E
Const DMLERR_SYS_ERROR = &H400F
Const DMLERR_UNADVACKTIMEOUT = &H4010
Const DMLERR_UNFOUND_QUEUE_ID = &H4011
Const DMLERR_LAST = &H4011
Const MH_CREATE = 1
Const MH_KEEP = 2
Const MH_DELETE = 3
Const MH_CLEANUP = 4
Const MAX_MONITORS = 4
Const APPCLASS_MONITOR = &H1&
Const XTYP_MONITOR = (&HF0 Or XCLASS_NOTIFICATION Or XTYPF_NOBLOCK)
Const MF_HSZ_INFO = &H1000000
Const MF_SENDMSGS = &H2000000
Const MF_POSTMSGS = &H4000000
Const MF_CALLBACKS = &H8000000
Const MF_ERRORS = &H10000000
Const MF_LINKS = &H20000000
Const MF_CONV = &H40000000
Const MF_MASK = &HFF000000
Const NO_ERROR = 0    '  dderror
Const ERROR_SUCCESS = 0&
Const ERROR_INVALID_FUNCTION = 1    '  dderror
Const ERROR_FILE_NOT_FOUND = 2&
Const ERROR_PATH_NOT_FOUND = 3&
Const ERROR_TOO_MANY_OPEN_FILES = 4&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_INVALID_HANDLE = 6&
Const ERROR_ARENA_TRASHED = 7&
Const ERROR_NOT_ENOUGH_MEMORY = 8    '  dderror
Const ERROR_INVALID_BLOCK = 9&
Const ERROR_BAD_ENVIRONMENT = 10&
Const ERROR_BAD_FORMAT = 11&
Const ERROR_INVALID_ACCESS = 12&
Const ERROR_INVALID_DATA = 13&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_DRIVE = 15&
Const ERROR_CURRENT_DIRECTORY = 16&
Const ERROR_NOT_SAME_DEVICE = 17&
Const ERROR_NO_MORE_FILES = 18&
Const ERROR_WRITE_PROTECT = 19&
Const ERROR_BAD_UNIT = 20&
Const ERROR_NOT_READY = 21&
Const ERROR_BAD_COMMAND = 22&
Const ERROR_CRC = 23&
Const ERROR_BAD_LENGTH = 24&
Const ERROR_SEEK = 25&
Const ERROR_NOT_DOS_DISK = 26&
Const ERROR_SECTOR_NOT_FOUND = 27&
Const ERROR_OUT_OF_PAPER = 28&
Const ERROR_WRITE_FAULT = 29&
Const ERROR_READ_FAULT = 30&
Const ERROR_GEN_FAILURE = 31&
Const ERROR_SHARING_VIOLATION = 32&
Const ERROR_LOCK_VIOLATION = 33&
Const ERROR_WRONG_DISK = 34&
Const ERROR_SHARING_BUFFER_EXCEEDED = 36&
Const ERROR_HANDLE_EOF = 38&
Const ERROR_HANDLE_DISK_FULL = 39&
Const ERROR_NOT_SUPPORTED = 50&
Const ERROR_REM_NOT_LIST = 51&
Const ERROR_DUP_NAME = 52&
Const ERROR_BAD_NETPATH = 53&
Const ERROR_NETWORK_BUSY = 54&
Const ERROR_DEV_NOT_EXIST = 55    '  dderror
Const ERROR_TOO_MANY_CMDS = 56&
Const ERROR_ADAP_HDW_ERR = 57&
Const ERROR_BAD_NET_RESP = 58&
Const ERROR_UNEXP_NET_ERR = 59&
Const ERROR_BAD_REM_ADAP = 60&
Const ERROR_PRINTQ_FULL = 61&
Const ERROR_NO_SPOOL_SPACE = 62&
Const ERROR_PRINT_CANCELLED = 63&
Const ERROR_NETNAME_DELETED = 64&
Const ERROR_NETWORK_ACCESS_DENIED = 65&
Const ERROR_BAD_DEV_TYPE = 66&
Const ERROR_BAD_NET_NAME = 67&
Const ERROR_TOO_MANY_NAMES = 68&
Const ERROR_TOO_MANY_SESS = 69&
Const ERROR_SHARING_PAUSED = 70&
Const ERROR_REQ_NOT_ACCEP = 71&
Const ERROR_REDIR_PAUSED = 72&
Const ERROR_FILE_EXISTS = 80&
Const ERROR_CANNOT_MAKE = 82&
Const ERROR_FAIL_I24 = 83&
Const ERROR_OUT_OF_STRUCTURES = 84&
Const ERROR_ALREADY_ASSIGNED = 85&
Const ERROR_INVALID_PASSWORD = 86&
Const ERROR_INVALID_PARAMETER = 87    '  dderror
Const ERROR_NET_WRITE_FAULT = 88&
Const ERROR_NO_PROC_SLOTS = 89&
Const ERROR_TOO_MANY_SEMAPHORES = 100&
Const ERROR_EXCL_SEM_ALREADY_OWNED = 101&
Const ERROR_SEM_IS_SET = 102&
Const ERROR_TOO_MANY_SEM_REQUESTS = 103&
Const ERROR_INVALID_AT_INTERRUPT_TIME = 104&
Const ERROR_SEM_OWNER_DIED = 105&
Const ERROR_SEM_USER_LIMIT = 106&
Const ERROR_DISK_CHANGE = 107&
Const ERROR_DRIVE_LOCKED = 108&
Const ERROR_BROKEN_PIPE = 109&
Const ERROR_OPEN_FAILED = 110&
Const ERROR_BUFFER_OVERFLOW = 111&
Const ERROR_DISK_FULL = 112&
Const ERROR_NO_MORE_SEARCH_HANDLES = 113&
Const ERROR_INVALID_TARGET_HANDLE = 114&
Const ERROR_INVALID_CATEGORY = 117&
Const ERROR_INVALID_VERIFY_SWITCH = 118&
Const ERROR_BAD_DRIVER_LEVEL = 119&
Const ERROR_CALL_NOT_IMPLEMENTED = 120&
Const ERROR_SEM_TIMEOUT = 121&
Const ERROR_INSUFFICIENT_BUFFER = 122    '  dderror
Const ERROR_INVALID_NAME = 123&
Const ERROR_INVALID_LEVEL = 124&
Const ERROR_NO_VOLUME_LABEL = 125&
Const ERROR_MOD_NOT_FOUND = 126&
Const ERROR_PROC_NOT_FOUND = 127&
Const ERROR_WAIT_NO_CHILDREN = 128&
Const ERROR_CHILD_NOT_COMPLETE = 129&
Const ERROR_DIRECT_ACCESS_HANDLE = 130&
Const ERROR_NEGATIVE_SEEK = 131&
Const ERROR_SEEK_ON_DEVICE = 132&
Const ERROR_IS_JOIN_TARGET = 133&
Const ERROR_IS_JOINED = 134&
Const ERROR_IS_SUBSTED = 135&
Const ERROR_NOT_JOINED = 136&
Const ERROR_NOT_SUBSTED = 137&
Const ERROR_JOIN_TO_JOIN = 138&
Const ERROR_SUBST_TO_SUBST = 139&
Const ERROR_JOIN_TO_SUBST = 140&
Const ERROR_SUBST_TO_JOIN = 141&
Const ERROR_BUSY_DRIVE = 142&
Const ERROR_SAME_DRIVE = 143&
Const ERROR_DIR_NOT_ROOT = 144&
Const ERROR_DIR_NOT_EMPTY = 145&
Const ERROR_IS_SUBST_PATH = 146&
Const ERROR_IS_JOIN_PATH = 147&
Const ERROR_PATH_BUSY = 148&
Const ERROR_IS_SUBST_TARGET = 149&
Const ERROR_SYSTEM_TRACE = 150&
Const ERROR_INVALID_EVENT_COUNT = 151&
Const ERROR_TOO_MANY_MUXWAITERS = 152&
Const ERROR_INVALID_LIST_FORMAT = 153&
Const ERROR_LABEL_TOO_LONG = 154&
Const ERROR_TOO_MANY_TCBS = 155&
Const ERROR_SIGNAL_REFUSED = 156&
Const ERROR_DISCARDED = 157&
Const ERROR_NOT_LOCKED = 158&
Const ERROR_BAD_THREADID_ADDR = 159&
Const ERROR_BAD_ARGUMENTS = 160&
Const ERROR_BAD_PATHNAME = 161&
Const ERROR_SIGNAL_PENDING = 162&
Const ERROR_MAX_THRDS_REACHED = 164&
Const ERROR_LOCK_FAILED = 167&
Const ERROR_BUSY = 170&
Const ERROR_CANCEL_VIOLATION = 173&
Const ERROR_ATOMIC_LOCKS_NOT_SUPPORTED = 174&
Const ERROR_INVALID_SEGMENT_NUMBER = 180&
Const ERROR_INVALID_ORDINAL = 182&
Const ERROR_ALREADY_EXISTS = 183&
Const ERROR_INVALID_FLAG_NUMBER = 186&
Const ERROR_SEM_NOT_FOUND = 187&
Const ERROR_INVALID_STARTING_CODESEG = 188&
Const ERROR_INVALID_STACKSEG = 189&
Const ERROR_INVALID_MODULETYPE = 190&
Const ERROR_INVALID_EXE_SIGNATURE = 191&
Const ERROR_EXE_MARKED_INVALID = 192&
Const ERROR_BAD_EXE_FORMAT = 193&
Const ERROR_ITERATED_DATA_EXCEEDS_64k = 194&
Const ERROR_INVALID_MINALLOCSIZE = 195&
Const ERROR_DYNLINK_FROM_INVALID_RING = 196&
Const ERROR_IOPL_NOT_ENABLED = 197&
Const ERROR_INVALID_SEGDPL = 198&
Const ERROR_AUTODATASEG_EXCEEDS_64k = 199&
Const ERROR_RING2SEG_MUST_BE_MOVABLE = 200&
Const ERROR_RELOC_CHAIN_XEEDS_SEGLIM = 201&
Const ERROR_INFLOOP_IN_RELOC_CHAIN = 202&
Const ERROR_ENVVAR_NOT_FOUND = 203&
Const ERROR_NO_SIGNAL_SENT = 205&
Const ERROR_FILENAME_EXCED_RANGE = 206&
Const ERROR_RING2_STACK_IN_USE = 207&
Const ERROR_META_EXPANSION_TOO_LONG = 208&
Const ERROR_INVALID_SIGNAL_NUMBER = 209&
Const ERROR_THREAD_1_INACTIVE = 210&
Const ERROR_LOCKED = 212&
Const ERROR_TOO_MANY_MODULES = 214&
Const ERROR_NESTING_NOT_ALLOWED = 215&
Const ERROR_BAD_PIPE = 230&
Const ERROR_PIPE_BUSY = 231&
Const ERROR_NO_DATA = 232&
Const ERROR_PIPE_NOT_CONNECTED = 233&
Const ERROR_MORE_DATA = 234    '  dderror
Const ERROR_VC_DISCONNECTED = 240&
Const ERROR_INVALID_EA_NAME = 254&
Const ERROR_EA_LIST_INCONSISTENT = 255&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_CANNOT_COPY = 266&
Const ERROR_DIRECTORY = 267&
Const ERROR_EAS_DIDNT_FIT = 275&
Const ERROR_EA_FILE_CORRUPT = 276&
Const ERROR_EA_TABLE_FULL = 277&
Const ERROR_INVALID_EA_HANDLE = 278&
Const ERROR_EAS_NOT_SUPPORTED = 282&
Const ERROR_NOT_OWNER = 288&
Const ERROR_TOO_MANY_POSTS = 298&
Const ERROR_MR_MID_NOT_FOUND = 317&
Const ERROR_INVALID_ADDRESS = 487&
Const ERROR_ARITHMETIC_OVERFLOW = 534&
Const ERROR_PIPE_CONNECTED = 535&
Const ERROR_PIPE_LISTENING = 536&
Const ERROR_EA_ACCESS_DENIED = 994&
Const ERROR_OPERATION_ABORTED = 995&
Const ERROR_IO_INCOMPLETE = 996&
Const ERROR_IO_PENDING = 997    '  dderror
Const ERROR_NOACCESS = 998&
Const ERROR_SWAPERROR = 999&
Const ERROR_STACK_OVERFLOW = 1001&
Const ERROR_INVALID_MESSAGE = 1002&
Const ERROR_CAN_NOT_COMPLETE = 1003&
Const ERROR_INVALID_FLAGS = 1004&
Const ERROR_UNRECOGNIZED_VOLUME = 1005&
Const ERROR_FILE_INVALID = 1006&
Const ERROR_FULLSCREEN_MODE = 1007&
Const ERROR_NO_TOKEN = 1008&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_REGISTRY_RECOVERED = 1014&
Const ERROR_REGISTRY_CORRUPT = 1015&
Const ERROR_REGISTRY_IO_FAILED = 1016&
Const ERROR_NOT_REGISTRY_FILE = 1017&
Const ERROR_KEY_DELETED = 1018&
Const ERROR_NO_LOG_SPACE = 1019&
Const ERROR_KEY_HAS_CHILDREN = 1020&
Const ERROR_CHILD_MUST_BE_VOLATILE = 1021&
Const ERROR_NOTIFY_ENUM_DIR = 1022&
Const ERROR_DEPENDENT_SERVICES_RUNNING = 1051&
Const ERROR_INVALID_SERVICE_CONTROL = 1052&
Const ERROR_SERVICE_REQUEST_TIMEOUT = 1053&
Const ERROR_SERVICE_NO_THREAD = 1054&
Const ERROR_SERVICE_DATABASE_LOCKED = 1055&
Const ERROR_SERVICE_ALREADY_RUNNING = 1056&
Const ERROR_INVALID_SERVICE_ACCOUNT = 1057&
Const ERROR_SERVICE_DISABLED = 1058&
Const ERROR_CIRCULAR_DEPENDENCY = 1059&
Const ERROR_SERVICE_DOES_NOT_EXIST = 1060&
Const ERROR_SERVICE_CANNOT_ACCEPT_CTRL = 1061&
Const ERROR_SERVICE_NOT_ACTIVE = 1062&
Const ERROR_FAILED_SERVICE_CONTROLLER_CONNECT = 1063&
Const ERROR_EXCEPTION_IN_SERVICE = 1064&
Const ERROR_DATABASE_DOES_NOT_EXIST = 1065&
Const ERROR_SERVICE_SPECIFIC_ERROR = 1066&
Const ERROR_PROCESS_ABORTED = 1067&
Const ERROR_SERVICE_DEPENDENCY_FAIL = 1068&
Const ERROR_SERVICE_LOGON_FAILED = 1069&
Const ERROR_SERVICE_START_HANG = 1070&
Const ERROR_INVALID_SERVICE_LOCK = 1071&
Const ERROR_SERVICE_MARKED_FOR_DELETE = 1072&
Const ERROR_SERVICE_EXISTS = 1073&
Const ERROR_ALREADY_RUNNING_LKG = 1074&
Const ERROR_SERVICE_DEPENDENCY_DELETED = 1075&
Const ERROR_BOOT_ALREADY_ACCEPTED = 1076&
Const ERROR_SERVICE_NEVER_STARTED = 1077&
Const ERROR_DUPLICATE_SERVICE_NAME = 1078&
Const ERROR_END_OF_MEDIA = 1100&
Const ERROR_FILEMARK_DETECTED = 1101&
Const ERROR_BEGINNING_OF_MEDIA = 1102&
Const ERROR_SETMARK_DETECTED = 1103&
Const ERROR_NO_DATA_DETECTED = 1104&
Const ERROR_PARTITION_FAILURE = 1105&
Const ERROR_INVALID_BLOCK_LENGTH = 1106&
Const ERROR_DEVICE_NOT_PARTITIONED = 1107&
Const ERROR_UNABLE_TO_LOCK_MEDIA = 1108&
Const ERROR_UNABLE_TO_UNLOAD_MEDIA = 1109&
Const ERROR_MEDIA_CHANGED = 1110&
Const ERROR_BUS_RESET = 1111&
Const ERROR_NO_MEDIA_IN_DRIVE = 1112&
Const ERROR_NO_UNICODE_TRANSLATION = 1113&
Const ERROR_DLL_INIT_FAILED = 1114&
Const ERROR_SHUTDOWN_IN_PROGRESS = 1115&
Const ERROR_NO_SHUTDOWN_IN_PROGRESS = 1116&
Const ERROR_IO_DEVICE = 1117&
Const ERROR_SERIAL_NO_DEVICE = 1118&
Const ERROR_IRQ_BUSY = 1119&
Const ERROR_MORE_WRITES = 1120&
Const ERROR_COUNTER_TIMEOUT = 1121&
Const ERROR_FLOPPY_ID_MARK_NOT_FOUND = 1122&
Const ERROR_FLOPPY_WRONG_CYLINDER = 1123&
Const ERROR_FLOPPY_UNKNOWN_ERROR = 1124&
Const ERROR_FLOPPY_BAD_REGISTERS = 1125&
Const ERROR_DISK_RECALIBRATE_FAILED = 1126&
Const ERROR_DISK_OPERATION_FAILED = 1127&
Const ERROR_DISK_RESET_FAILED = 1128&
Const ERROR_EOM_OVERFLOW = 1129&
Const ERROR_NOT_ENOUGH_SERVER_MEMORY = 1130&
Const ERROR_POSSIBLE_DEADLOCK = 1131&
Const ERROR_MAPPED_ALIGNMENT = 1132&
Const ERROR_INVALID_PIXEL_FORMAT = 2000
Const ERROR_BAD_DRIVER = 2001
Const ERROR_INVALID_WINDOW_STYLE = 2002
Const ERROR_METAFILE_NOT_SUPPORTED = 2003
Const ERROR_TRANSFORM_NOT_SUPPORTED = 2004
Const ERROR_CLIPPING_NOT_SUPPORTED = 2005
Const ERROR_UNKNOWN_PRINT_MONITOR = 3000
Const ERROR_PRINTER_DRIVER_IN_USE = 3001
Const ERROR_SPOOL_FILE_NOT_FOUND = 3002
Const ERROR_SPL_NO_STARTDOC = 3003
Const ERROR_SPL_NO_ADDJOB = 3004
Const ERROR_PRINT_PROCESSOR_ALREADY_INSTALLED = 3005
Const ERROR_PRINT_MONITOR_ALREADY_INSTALLED = 3006
Const ERROR_WINS_INTERNAL = 4000
Const ERROR_CAN_NOT_DEL_LOCAL_WINS = 4001
Const ERROR_STATIC_INIT = 4002
Const ERROR_INC_BACKUP = 4003
Const ERROR_FULL_BACKUP = 4004
Const ERROR_REC_NON_EXISTENT = 4005
Const ERROR_RPL_NOT_ALLOWED = 4006
Const SEVERITY_SUCCESS = 0
Const SEVERITY_ERROR = 1
Const FACILITY_NT_BIT = &H10000000
Const NOERROR = 0
Const E_UNEXPECTED = &H8000FFFF
Const E_NOTIMPL = &H80004001
Const E_OUTOFMEMORY = &H8007000E
Const E_INVALIDARG = &H80070057
Const E_NOINTERFACE = &H80004002
Const E_POINTER = &H80004003
Const E_HANDLE = &H80070006
Const E_ABORT = &H80004004
Const E_FAIL = &H80004005
Const E_ACCESSDENIED = &H80070005
Const CO_E_INIT_TLS = &H80004006
Const CO_E_INIT_SHARED_ALLOCATOR = &H80004007
Const CO_E_INIT_MEMORY_ALLOCATOR = &H80004008
Const CO_E_INIT_CLASS_CACHE = &H80004009
Const CO_E_INIT_RPC_CHANNEL = &H8000400A
Const CO_E_INIT_TLS_SET_CHANNEL_CONTROL = &H8000400B
Const CO_E_INIT_TLS_CHANNEL_CONTROL = &H8000400C
Const CO_E_INIT_UNACCEPTED_USER_ALLOCATOR = &H8000400D
Const CO_E_INIT_SCM_MUTEX_EXISTS = &H8000400E
Const CO_E_INIT_SCM_FILE_MAPPING_EXISTS = &H8000400F
Const CO_E_INIT_SCM_MAP_VIEW_OF_FILE = &H80004010
Const CO_E_INIT_SCM_EXEC_FAILURE = &H80004011
Const CO_E_INIT_ONLY_SINGLE_THREADED = &H80004012
Const S_OK = &H0
Const S_FALSE = &H1
Const OLE_E_FIRST = &H80040000
Const OLE_E_LAST = &H800400FF
Const OLE_S_FIRST = &H40000
Const OLE_S_LAST = &H400FF
Const OLE_E_OLEVERB = &H80040000
Const OLE_E_ADVF = &H80040001
Const OLE_E_ENUM_NOMORE = &H80040002
Const OLE_E_ADVISENOTSUPPORTED = &H80040003
Const OLE_E_NOCONNECTION = &H80040004
Const OLE_E_NOTRUNNING = &H80040005
Const OLE_E_NOCACHE = &H80040006
Const OLE_E_BLANK = &H80040007
Const OLE_E_CLASSDIFF = &H80040008
Const OLE_E_CANT_GETMONIKER = &H80040009
Const OLE_E_CANT_BINDTOSOURCE = &H8004000A
Const OLE_E_STATIC = &H8004000B
Const OLE_E_PROMPTSAVECANCELLED = &H8004000C
Const OLE_E_INVALIDRECT = &H8004000D
Const OLE_E_WRONGCOMPOBJ = &H8004000E
Const OLE_E_INVALIDHWND = &H8004000F
Const OLE_E_NOT_INPLACEACTIVE = &H80040010
Const OLE_E_CANTCONVERT = &H80040011
Const OLE_E_NOSTORAGE = &H80040012
Const DV_E_FORMATETC = &H80040064
Const DV_E_DVTARGETDEVICE = &H80040065
Const DV_E_STGMEDIUM = &H80040066
Const DV_E_STATDATA = &H80040067
Const DV_E_LINDEX = &H80040068
Const DV_E_TYMED = &H80040069
Const DV_E_CLIPFORMAT = &H8004006A
Const DV_E_DVASPECT = &H8004006B
Const DV_E_DVTARGETDEVICE_SIZE = &H8004006C
Const DV_E_NOIVIEWOBJECT = &H8004006D
Const DRAGDROP_E_FIRST = &H80040100
Const DRAGDROP_E_LAST = &H8004010F
Const DRAGDROP_S_FIRST = &H40100
Const DRAGDROP_S_LAST = &H4010F
Const DRAGDROP_E_NOTREGISTERED = &H80040100
Const DRAGDROP_E_ALREADYREGISTERED = &H80040101
Const DRAGDROP_E_INVALIDHWND = &H80040102
Const CLASSFACTORY_E_FIRST = &H80040110
Const CLASSFACTORY_E_LAST = &H8004011F
Const CLASSFACTORY_S_FIRST = &H40110
Const CLASSFACTORY_S_LAST = &H4011F
Const CLASS_E_NOAGGREGATION = &H80040110
Const CLASS_E_CLASSNOTAVAILABLE = &H80040111
Const MARSHAL_E_FIRST = &H80040120
Const MARSHAL_E_LAST = &H8004012F
Const MARSHAL_S_FIRST = &H40120
Const MARSHAL_S_LAST = &H4012F
Const DATA_E_FIRST = &H80040130
Const DATA_E_LAST = &H8004013F
Const DATA_S_FIRST = &H40130
Const DATA_S_LAST = &H4013F
Const VIEW_E_FIRST = &H80040140
Const VIEW_E_LAST = &H8004014F
Const VIEW_S_FIRST = &H40140
Const VIEW_S_LAST = &H4014F
Const VIEW_E_DRAW = &H80040140
Const REGDB_E_FIRST = &H80040150
Const REGDB_E_LAST = &H8004015F
Const REGDB_S_FIRST = &H40150
Const REGDB_S_LAST = &H4015F
Const REGDB_E_READREGDB = &H80040150
Const REGDB_E_WRITEREGDB = &H80040151
Const REGDB_E_KEYMISSING = &H80040152
Const REGDB_E_INVALIDVALUE = &H80040153
Const REGDB_E_CLASSNOTREG = &H80040154
Const REGDB_E_IIDNOTREG = &H80040155
Const CACHE_E_FIRST = &H80040170
Const CACHE_E_LAST = &H8004017F
Const CACHE_S_FIRST = &H40170
Const CACHE_S_LAST = &H4017F
Const CACHE_E_NOCACHE_UPDATED = &H80040170
Const OLEOBJ_E_FIRST = &H80040180
Const OLEOBJ_E_LAST = &H8004018F
Const OLEOBJ_S_FIRST = &H40180
Const OLEOBJ_S_LAST = &H4018F
Const OLEOBJ_E_NOVERBS = &H80040180
Const OLEOBJ_E_INVALIDVERB = &H80040181
Const CLIENTSITE_E_FIRST = &H80040190
Const CLIENTSITE_E_LAST = &H8004019F
Const CLIENTSITE_S_FIRST = &H40190
Const CLIENTSITE_S_LAST = &H4019F
Const INPLACE_E_NOTUNDOABLE = &H800401A0
Const INPLACE_E_NOTOOLSPACE = &H800401A1
Const INPLACE_E_FIRST = &H800401A0
Const INPLACE_E_LAST = &H800401AF
Const INPLACE_S_FIRST = &H401A0
Const INPLACE_S_LAST = &H401AF
Const ENUM_E_FIRST = &H800401B0
Const ENUM_E_LAST = &H800401BF
Const ENUM_S_FIRST = &H401B0
Const ENUM_S_LAST = &H401BF
Const CONVERT10_E_FIRST = &H800401C0
Const CONVERT10_E_LAST = &H800401CF
Const CONVERT10_S_FIRST = &H401C0
Const CONVERT10_S_LAST = &H401CF
Const CONVERT10_E_OLESTREAM_GET = &H800401C0
Const CONVERT10_E_OLESTREAM_PUT = &H800401C1
Const CONVERT10_E_OLESTREAM_FMT = &H800401C2
Const CONVERT10_E_OLESTREAM_BITMAP_TO_DIB = &H800401C3
Const CONVERT10_E_STG_FMT = &H800401C4
Const CONVERT10_E_STG_NO_STD_STREAM = &H800401C5
Const CONVERT10_E_STG_DIB_TO_BITMAP = &H800401C6
Const CLIPBRD_E_FIRST = &H800401D0
Const CLIPBRD_E_LAST = &H800401DF
Const CLIPBRD_S_FIRST = &H401D0
Const CLIPBRD_S_LAST = &H401DF
Const CLIPBRD_E_CANT_OPEN = &H800401D0
Const CLIPBRD_E_CANT_EMPTY = &H800401D1
Const CLIPBRD_E_CANT_SET = &H800401D2
Const CLIPBRD_E_BAD_DATA = &H800401D3
Const CLIPBRD_E_CANT_CLOSE = &H800401D4
Const MK_E_FIRST = &H800401E0
Const MK_E_LAST = &H800401EF
Const MK_S_FIRST = &H401E0
Const MK_S_LAST = &H401EF
Const MK_E_CONNECTMANUALLY = &H800401E0
Const MK_E_EXCEEDEDDEADLINE = &H800401E1
Const MK_E_NEEDGENERIC = &H800401E2
Const MK_E_UNAVAILABLE = &H800401E3
Const MK_E_SYNTAX = &H800401E4
Const MK_E_NOOBJECT = &H800401E5
Const MK_E_INVALIDEXTENSION = &H800401E6
Const MK_E_INTERMEDIATEINTERFACENOTSUPPORTED = &H800401E7
Const MK_E_NOTBINDABLE = &H800401E8
Const MK_E_NOTBOUND = &H800401E9
Const MK_E_CANTOPENFILE = &H800401EA
Const MK_E_MUSTBOTHERUSER = &H800401EB
Const MK_E_NOINVERSE = &H800401EC
Const MK_E_NOSTORAGE = &H800401ED
Const MK_E_NOPREFIX = &H800401EE
Const MK_E_ENUMERATION_FAILED = &H800401EF
Const CO_E_FIRST = &H800401F0
Const CO_E_LAST = &H800401FF
Const CO_S_FIRST = &H401F0
Const CO_S_LAST = &H401FF
Const CO_E_NOTINITIALIZED = &H800401F0
Const CO_E_ALREADYINITIALIZED = &H800401F1
Const CO_E_CANTDETERMINECLASS = &H800401F2
Const CO_E_CLASSSTRING = &H800401F3
Const CO_E_IIDSTRING = &H800401F4
Const CO_E_APPNOTFOUND = &H800401F5
Const CO_E_APPSINGLEUSE = &H800401F6
Const CO_E_ERRORINAPP = &H800401F7
Const CO_E_DLLNOTFOUND = &H800401F8
Const CO_E_ERRORINDLL = &H800401F9
Const CO_E_WRONGOSFORAPP = &H800401FA
Const CO_E_OBJNOTREG = &H800401FB
Const CO_E_OBJISREG = &H800401FC
Const CO_E_OBJNOTCONNECTED = &H800401FD
Const CO_E_APPDIDNTREG = &H800401FE
Const CO_E_RELEASED = &H800401FF
Const OLE_S_USEREG = &H40000
Const OLE_S_STATIC = &H40001
Const OLE_S_MAC_CLIPFORMAT = &H40002
Const DRAGDROP_S_DROP = &H40100
Const DRAGDROP_S_CANCEL = &H40101
Const DRAGDROP_S_USEDEFAULTCURSORS = &H40102
Const DATA_S_SAMEFORMATETC = &H40130
Const VIEW_S_ALREADY_FROZEN = &H40140
Const CACHE_S_FORMATETC_NOTSUPPORTED = &H40170
Const CACHE_S_SAMECACHE = &H40171
Const CACHE_S_SOMECACHES_NOTUPDATED = &H40172
Const OLEOBJ_S_INVALIDVERB = &H40180
Const OLEOBJ_S_CANNOT_DOVERB_NOW = &H40181
Const OLEOBJ_S_INVALIDHWND = &H40182
Const INPLACE_S_TRUNCATED = &H401A0
Const CONVERT10_S_NO_PRESENTATION = &H401C0
Const MK_S_REDUCED_TO_SELF = &H401E2
Const MK_S_ME = &H401E4
Const MK_S_HIM = &H401E5
Const MK_S_US = &H401E6
Const MK_S_MONIKERALREADYREGISTERED = &H401E7
Const CO_E_CLASS_CREATE_FAILED = &H80080001
Const CO_E_SCM_ERROR = &H80080002
Const CO_E_SCM_RPC_FAILURE = &H80080003
Const CO_E_BAD_PATH = &H80080004
Const CO_E_SERVER_EXEC_FAILURE = &H80080005
Const CO_E_OBJSRV_RPC_FAILURE = &H80080006
Const MK_E_NO_NORMALIZED = &H80080007
Const CO_E_SERVER_STOPPING = &H80080008
Const MEM_E_INVALID_ROOT = &H80080009
Const MEM_E_INVALID_LINK = &H80080010
Const MEM_E_INVALID_SIZE = &H80080011
Const DISP_E_UNKNOWNINTERFACE = &H80020001
Const DISP_E_MEMBERNOTFOUND = &H80020003
Const DISP_E_PARAMNOTFOUND = &H80020004
Const DISP_E_TYPEMISMATCH = &H80020005
Const DISP_E_UNKNOWNNAME = &H80020006
Const DISP_E_NONAMEDARGS = &H80020007
Const DISP_E_BADVARTYPE = &H80020008
Const DISP_E_EXCEPTION = &H80020009
Const DISP_E_OVERFLOW = &H8002000A
Const DISP_E_BADINDEX = &H8002000B
Const DISP_E_UNKNOWNLCID = &H8002000C
Const DISP_E_ARRAYISLOCKED = &H8002000D
Const DISP_E_BADPARAMCOUNT = &H8002000E
Const DISP_E_PARAMNOTOPTIONAL = &H8002000F
Const DISP_E_BADCALLEE = &H80020010
Const DISP_E_NOTACOLLECTION = &H80020011
Const TYPE_E_BUFFERTOOSMALL = &H80028016
Const TYPE_E_INVDATAREAD = &H80028018
Const TYPE_E_UNSUPFORMAT = &H80028019
Const TYPE_E_REGISTRYACCESS = &H8002801C
Const TYPE_E_LIBNOTREGISTERED = &H8002801D
Const TYPE_E_UNDEFINEDTYPE = &H80028027
Const TYPE_E_QUALIFIEDNAMEDISALLOWED = &H80028028
Const TYPE_E_INVALIDSTATE = &H80028029
Const TYPE_E_WRONGTYPEKIND = &H8002802A
Const TYPE_E_ELEMENTNOTFOUND = &H8002802B
Const TYPE_E_AMBIGUOUSNAME = &H8002802C
Const TYPE_E_NAMECONFLICT = &H8002802D
Const TYPE_E_UNKNOWNLCID = &H8002802E
Const TYPE_E_DLLFUNCTIONNOTFOUND = &H8002802F
Const TYPE_E_BADMODULEKIND = &H800288BD
Const TYPE_E_SIZETOOBIG = &H800288C5
Const TYPE_E_DUPLICATEID = &H800288C6
Const TYPE_E_INVALIDID = &H800288CF
Const TYPE_E_TYPEMISMATCH = &H80028CA0
Const TYPE_E_OUTOFBOUNDS = &H80028CA1
Const TYPE_E_IOERROR = &H80028CA2
Const TYPE_E_CANTCREATETMPFILE = &H80028CA3
Const TYPE_E_CANTLOADLIBRARY = &H80029C4A
Const TYPE_E_INCONSISTENTPROPFUNCS = &H80029C83
Const TYPE_E_CIRCULARTYPE = &H80029C84
Const STG_E_INVALIDFUNCTION = &H80030001
Const STG_E_FILENOTFOUND = &H80030002
Const STG_E_PATHNOTFOUND = &H80030003
Const STG_E_TOOMANYOPENFILES = &H80030004
Const STG_E_ACCESSDENIED = &H80030005
Const STG_E_INVALIDHANDLE = &H80030006
Const STG_E_INSUFFICIENTMEMORY = &H80030008
Const STG_E_INVALIDPOINTER = &H80030009
Const STG_E_NOMOREFILES = &H80030012
Const STG_E_DISKISWRITEPROTECTED = &H80030013
Const STG_E_SEEKERROR = &H80030019
Const STG_E_WRITEFAULT = &H8003001D
Const STG_E_READFAULT = &H8003001E
Const STG_E_SHAREVIOLATION = &H80030020
Const STG_E_LOCKVIOLATION = &H80030021
Const STG_E_FILEALREADYEXISTS = &H80030050
Const STG_E_INVALIDPARAMETER = &H80030057
Const STG_E_MEDIUMFULL = &H80030070
Const STG_E_ABNORMALAPIEXIT = &H800300FA
Const STG_E_INVALIDHEADER = &H800300FB
Const STG_E_INVALIDNAME = &H800300FC
Const STG_E_UNKNOWN = &H800300FD
Const STG_E_UNIMPLEMENTEDFUNCTION = &H800300FE
Const STG_E_INVALIDFLAG = &H800300FF
Const STG_E_INUSE = &H80030100
Const STG_E_NOTCURRENT = &H80030101
Const STG_E_REVERTED = &H80030102
Const STG_E_CANTSAVE = &H80030103
Const STG_E_OLDFORMAT = &H80030104
Const STG_E_OLDDLL = &H80030105
Const STG_E_SHAREREQUIRED = &H80030106
Const STG_E_NOTFILEBASEDSTORAGE = &H80030107
Const STG_E_EXTANTMARSHALLINGS = &H80030108
Const STG_S_CONVERTED = &H30200
Const RPC_E_CALL_REJECTED = &H80010001
Const RPC_E_CALL_CANCELED = &H80010002
Const RPC_E_CANTPOST_INSENDCALL = &H80010003
Const RPC_E_CANTCALLOUT_INASYNCCALL = &H80010004
Const RPC_E_CANTCALLOUT_INEXTERNALCALL = &H80010005
Const RPC_E_CONNECTION_TERMINATED = &H80010006
Const RPC_E_SERVER_DIED = &H80010007
Const RPC_E_CLIENT_DIED = &H80010008
Const RPC_E_INVALID_DATAPACKET = &H80010009
Const RPC_E_CANTTRANSMIT_CALL = &H8001000A
Const RPC_E_CLIENT_CANTMARSHAL_DATA = &H8001000B
Const RPC_E_CLIENT_CANTUNMARSHAL_DATA = &H8001000C
Const RPC_E_SERVER_CANTMARSHAL_DATA = &H8001000D
Const RPC_E_SERVER_CANTUNMARSHAL_DATA = &H8001000E
Const RPC_E_INVALID_DATA = &H8001000F
Const RPC_E_INVALID_PARAMETER = &H80010010
Const RPC_E_CANTCALLOUT_AGAIN = &H80010011
Const RPC_E_SERVER_DIED_DNE = &H80010012
Const RPC_E_SYS_CALL_FAILED = &H80010100
Const RPC_E_OUT_OF_RESOURCES = &H80010101
Const RPC_E_ATTEMPTED_MULTITHREAD = &H80010102
Const RPC_E_NOT_REGISTERED = &H80010103
Const RPC_E_FAULT = &H80010104
Const RPC_E_SERVERFAULT = &H80010105
Const RPC_E_CHANGED_MODE = &H80010106
Const RPC_E_INVALIDMETHOD = &H80010107
Const RPC_E_DISCONNECTED = &H80010108
Const RPC_E_RETRY = &H80010109
Const RPC_E_SERVERCALL_RETRYLATER = &H8001010A
Const RPC_E_SERVERCALL_REJECTED = &H8001010B
Const RPC_E_INVALID_CALLDATA = &H8001010C
Const RPC_E_CANTCALLOUT_ININPUTSYNCCALL = &H8001010D
Const RPC_E_WRONG_THREAD = &H8001010E
Const RPC_E_THREAD_NOT_INIT = &H8001010F
Const RPC_E_UNEXPECTED = &H8001FFFF
Const ERROR_BAD_USERNAME = 2202&
Const ERROR_NOT_CONNECTED = 2250&
Const ERROR_OPEN_FILES = 2401&
Const ERROR_DEVICE_IN_USE = 2404&
Const ERROR_BAD_DEVICE = 1200&
Const ERROR_CONNECTION_UNAVAIL = 1201&
Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
Const ERROR_NO_NET_OR_BAD_PATH = 1203&
Const ERROR_BAD_PROVIDER = 1204&
Const ERROR_CANNOT_OPEN_PROFILE = 1205&
Const ERROR_BAD_PROFILE = 1206&
Const ERROR_NOT_CONTAINER = 1207&
Const ERROR_EXTENDED_ERROR = 1208&
Const ERROR_INVALID_GROUPNAME = 1209&
Const ERROR_INVALID_COMPUTERNAME = 1210&
Const ERROR_INVALID_EVENTNAME = 1211&
Const ERROR_INVALID_DOMAINNAME = 1212&
Const ERROR_INVALID_SERVICENAME = 1213&
Const ERROR_INVALID_NETNAME = 1214&
Const ERROR_INVALID_SHARENAME = 1215&
Const ERROR_INVALID_PASSWORDNAME = 1216&
Const ERROR_INVALID_MESSAGENAME = 1217&
Const ERROR_INVALID_MESSAGEDEST = 1218&
Const ERROR_SESSION_CREDENTIAL_CONFLICT = 1219&
Const ERROR_REMOTE_SESSION_LIMIT_EXCEEDED = 1220&
Const ERROR_DUP_DOMAINNAME = 1221&
Const ERROR_NO_NETWORK = 1222&
Const ERROR_NOT_ALL_ASSIGNED = 1300&
Const ERROR_SOME_NOT_MAPPED = 1301&
Const ERROR_NO_QUOTAS_FOR_ACCOUNT = 1302&
Const ERROR_LOCAL_USER_SESSION_KEY = 1303&
Const ERROR_NULL_LM_PASSWORD = 1304&
Const ERROR_UNKNOWN_REVISION = 1305&
Const ERROR_REVISION_MISMATCH = 1306&
Const ERROR_INVALID_OWNER = 1307&
Const ERROR_INVALID_PRIMARY_GROUP = 1308&
Const ERROR_NO_IMPERSONATION_TOKEN = 1309&
Const ERROR_CANT_DISABLE_MANDATORY = 1310&
Const ERROR_NO_LOGON_SERVERS = 1311&
Const ERROR_NO_SUCH_LOGON_SESSION = 1312&
Const ERROR_NO_SUCH_PRIVILEGE = 1313&
Const ERROR_PRIVILEGE_NOT_HELD = 1314&
Const ERROR_INVALID_ACCOUNT_NAME = 1315&
Const ERROR_USER_EXISTS = 1316&
Const ERROR_NO_SUCH_USER = 1317&
Const ERROR_GROUP_EXISTS = 1318&
Const ERROR_NO_SUCH_GROUP = 1319&
Const ERROR_MEMBER_IN_GROUP = 1320&
Const ERROR_MEMBER_NOT_IN_GROUP = 1321&
Const ERROR_LAST_ADMIN = 1322&
Const ERROR_WRONG_PASSWORD = 1323&
Const ERROR_ILL_FORMED_PASSWORD = 1324&
Const ERROR_PASSWORD_RESTRICTION = 1325&
Const ERROR_LOGON_FAILURE = 1326&
Const ERROR_ACCOUNT_RESTRICTION = 1327&
Const ERROR_INVALID_LOGON_HOURS = 1328&
Const ERROR_INVALID_WORKSTATION = 1329&
Const ERROR_PASSWORD_EXPIRED = 1330&
Const ERROR_ACCOUNT_DISABLED = 1331&
Const ERROR_NONE_MAPPED = 1332&
Const ERROR_TOO_MANY_LUIDS_REQUESTED = 1333&
Const ERROR_LUIDS_EXHAUSTED = 1334&
Const ERROR_INVALID_SUB_AUTHORITY = 1335&
Const ERROR_INVALID_ACL = 1336&
Const ERROR_INVALID_SID = 1337&
Const ERROR_INVALID_SECURITY_DESCR = 1338&
Const ERROR_BAD_INHERITANCE_ACL = 1340&
Const ERROR_SERVER_DISABLED = 1341&
Const ERROR_SERVER_NOT_DISABLED = 1342&
Const ERROR_INVALID_ID_AUTHORITY = 1343&
Const ERROR_ALLOTTED_SPACE_EXCEEDED = 1344&
Const ERROR_INVALID_GROUP_ATTRIBUTES = 1345&
Const ERROR_BAD_IMPERSONATION_LEVEL = 1346&
Const ERROR_CANT_OPEN_ANONYMOUS = 1347&
Const ERROR_BAD_VALIDATION_CLASS = 1348&
Const ERROR_BAD_TOKEN_TYPE = 1349&
Const ERROR_NO_SECURITY_ON_OBJECT = 1350&
Const ERROR_CANT_ACCESS_DOMAIN_INFO = 1351&
Const ERROR_INVALID_SERVER_STATE = 1352&
Const ERROR_INVALID_DOMAIN_STATE = 1353&
Const ERROR_INVALID_DOMAIN_ROLE = 1354&
Const ERROR_NO_SUCH_DOMAIN = 1355&
Const ERROR_DOMAIN_EXISTS = 1356&
Const ERROR_DOMAIN_LIMIT_EXCEEDED = 1357&
Const ERROR_INTERNAL_DB_CORRUPTION = 1358&
Const ERROR_INTERNAL_ERROR = 1359&
Const ERROR_GENERIC_NOT_MAPPED = 1360&
Const ERROR_BAD_DESCRIPTOR_FORMAT = 1361&
Const ERROR_NOT_LOGON_PROCESS = 1362&
Const ERROR_LOGON_SESSION_EXISTS = 1363&
Const ERROR_NO_SUCH_PACKAGE = 1364&
Const ERROR_BAD_LOGON_SESSION_STATE = 1365&
Const ERROR_LOGON_SESSION_COLLISION = 1366&
Const ERROR_INVALID_LOGON_TYPE = 1367&
Const ERROR_CANNOT_IMPERSONATE = 1368&
Const ERROR_RXACT_INVALID_STATE = 1369&
Const ERROR_RXACT_COMMIT_FAILURE = 1370&
Const ERROR_SPECIAL_ACCOUNT = 1371&
Const ERROR_SPECIAL_GROUP = 1372&
Const ERROR_SPECIAL_USER = 1373&
Const ERROR_MEMBERS_PRIMARY_GROUP = 1374&
Const ERROR_TOKEN_ALREADY_IN_USE = 1375&
Const ERROR_NO_SUCH_ALIAS = 1376&
Const ERROR_MEMBER_NOT_IN_ALIAS = 1377&
Const ERROR_MEMBER_IN_ALIAS = 1378&
Const ERROR_ALIAS_EXISTS = 1379&
Const ERROR_LOGON_NOT_GRANTED = 1380&
Const ERROR_TOO_MANY_SECRETS = 1381&
Const ERROR_SECRET_TOO_LONG = 1382&
Const ERROR_INTERNAL_DB_ERROR = 1383&
Const ERROR_TOO_MANY_CONTEXT_IDS = 1384&
Const ERROR_LOGON_TYPE_NOT_GRANTED = 1385&
Const ERROR_NT_CROSS_ENCRYPTION_REQUIRED = 1386&
Const ERROR_NO_SUCH_MEMBER = 1387&
Const ERROR_INVALID_MEMBER = 1388&
Const ERROR_TOO_MANY_SIDS = 1389&
Const ERROR_LM_CROSS_ENCRYPTION_REQUIRED = 1390&
Const ERROR_NO_INHERITANCE = 1391&
Const ERROR_FILE_CORRUPT = 1392&
Const ERROR_DISK_CORRUPT = 1393&
Const ERROR_NO_USER_SESSION_KEY = 1394&
Const ERROR_INVALID_WINDOW_HANDLE = 1400&
Const ERROR_INVALID_MENU_HANDLE = 1401&
Const ERROR_INVALID_CURSOR_HANDLE = 1402&
Const ERROR_INVALID_ACCEL_HANDLE = 1403&
Const ERROR_INVALID_HOOK_HANDLE = 1404&
Const ERROR_INVALID_DWP_HANDLE = 1405&
Const ERROR_TLW_WITH_WSCHILD = 1406&
Const ERROR_CANNOT_FIND_WND_CLASS = 1407&
Const ERROR_WINDOW_OF_OTHER_THREAD = 1408&
Const ERROR_HOTKEY_ALREADY_REGISTERED = 1409&
Const ERROR_CLASS_ALREADY_EXISTS = 1410&
Const ERROR_CLASS_DOES_NOT_EXIST = 1411&
Const ERROR_CLASS_HAS_WINDOWS = 1412&
Const ERROR_INVALID_INDEX = 1413&
Const ERROR_INVALID_ICON_HANDLE = 1414&
Const ERROR_PRIVATE_DIALOG_INDEX = 1415&
Const ERROR_LISTBOX_ID_NOT_FOUND = 1416&
Const ERROR_NO_WILDCARD_CHARACTERS = 1417&
Const ERROR_CLIPBOARD_NOT_OPEN = 1418&
Const ERROR_HOTKEY_NOT_REGISTERED = 1419&
Const ERROR_WINDOW_NOT_DIALOG = 1420&
Const ERROR_CONTROL_ID_NOT_FOUND = 1421&
Const ERROR_INVALID_COMBOBOX_MESSAGE = 1422&
Const ERROR_WINDOW_NOT_COMBOBOX = 1423&
Const ERROR_INVALID_EDIT_HEIGHT = 1424&
Const ERROR_DC_NOT_FOUND = 1425&
Const ERROR_INVALID_HOOK_FILTER = 1426&
Const ERROR_INVALID_FILTER_PROC = 1427&
Const ERROR_HOOK_NEEDS_HMOD = 1428&
Const ERROR_PUBLIC_ONLY_HOOK = 1429&
Const ERROR_JOURNAL_HOOK_SET = 1430&
Const ERROR_HOOK_NOT_INSTALLED = 1431&
Const ERROR_INVALID_LB_MESSAGE = 1432&
Const ERROR_SETCOUNT_ON_BAD_LB = 1433&
Const ERROR_LB_WITHOUT_TABSTOPS = 1434&
Const ERROR_DESTROY_OBJECT_OF_OTHER_THREAD = 1435&
Const ERROR_CHILD_WINDOW_MENU = 1436&
Const ERROR_NO_SYSTEM_MENU = 1437&
Const ERROR_INVALID_MSGBOX_STYLE = 1438&
Const ERROR_INVALID_SPI_VALUE = 1439&
Const ERROR_SCREEN_ALREADY_LOCKED = 1440&
Const ERROR_HWNDS_HAVE_DIFF_PARENT = 1441&
Const ERROR_NOT_CHILD_WINDOW = 1442&
Const ERROR_INVALID_GW_COMMAND = 1443&
Const ERROR_INVALID_THREAD_ID = 1444&
Const ERROR_NON_MDICHILD_WINDOW = 1445&
Const ERROR_POPUP_ALREADY_ACTIVE = 1446&
Const ERROR_NO_SCROLLBARS = 1447&
Const ERROR_INVALID_SCROLLBAR_RANGE = 1448&
Const ERROR_INVALID_SHOWWIN_COMMAND = 1449&
Const ERROR_EVENTLOG_FILE_CORRUPT = 1500&
Const ERROR_EVENTLOG_CANT_START = 1501&
Const ERROR_LOG_FILE_FULL = 1502&
Const ERROR_EVENTLOG_FILE_CHANGED = 1503&
Const RPC_S_INVALID_STRING_BINDING = 1700&
Const RPC_S_WRONG_KIND_OF_BINDING = 1701&
Const RPC_S_INVALID_BINDING = 1702&
Const RPC_S_PROTSEQ_NOT_SUPPORTED = 1703&
Const RPC_S_INVALID_RPC_PROTSEQ = 1704&
Const RPC_S_INVALID_STRING_UUID = 1705&
Const RPC_S_INVALID_ENDPOINT_FORMAT = 1706&
Const RPC_S_INVALID_NET_ADDR = 1707&
Const RPC_S_NO_ENDPOINT_FOUND = 1708&
Const RPC_S_INVALID_TIMEOUT = 1709&
Const RPC_S_OBJECT_NOT_FOUND = 1710&
Const RPC_S_ALREADY_REGISTERED = 1711&
Const RPC_S_TYPE_ALREADY_REGISTERED = 1712&
Const RPC_S_ALREADY_LISTENING = 1713&
Const RPC_S_NO_PROTSEQS_REGISTERED = 1714&
Const RPC_S_NOT_LISTENING = 1715&
Const RPC_S_UNKNOWN_MGR_TYPE = 1716&
Const RPC_S_UNKNOWN_IF = 1717&
Const RPC_S_NO_BINDINGS = 1718&
Const RPC_S_NO_PROTSEQS = 1719&
Const RPC_S_CANT_CREATE_ENDPOINT = 1720&
Const RPC_S_OUT_OF_RESOURCES = 1721&
Const RPC_S_SERVER_UNAVAILABLE = 1722&
Const RPC_S_SERVER_TOO_BUSY = 1723&
Const RPC_S_INVALID_NETWORK_OPTIONS = 1724&
Const RPC_S_NO_CALL_ACTIVE = 1725&
Const RPC_S_CALL_FAILED = 1726&
Const RPC_S_CALL_FAILED_DNE = 1727&
Const RPC_S_PROTOCOL_ERROR = 1728&
Const RPC_S_UNSUPPORTED_TRANS_SYN = 1730&
Const RPC_S_UNSUPPORTED_TYPE = 1732&
Const RPC_S_INVALID_TAG = 1733&
Const RPC_S_INVALID_BOUND = 1734&
Const RPC_S_NO_ENTRY_NAME = 1735&
Const RPC_S_INVALID_NAME_SYNTAX = 1736&
Const RPC_S_UNSUPPORTED_NAME_SYNTAX = 1737&
Const RPC_S_UUID_NO_ADDRESS = 1739&
Const RPC_S_DUPLICATE_ENDPOINT = 1740&
Const RPC_S_UNKNOWN_AUTHN_TYPE = 1741&
Const RPC_S_MAX_CALLS_TOO_SMALL = 1742&
Const RPC_S_STRING_TOO_LONG = 1743&
Const RPC_S_PROTSEQ_NOT_FOUND = 1744&
Const RPC_S_PROCNUM_OUT_OF_RANGE = 1745&
Const RPC_S_BINDING_HAS_NO_AUTH = 1746&
Const RPC_S_UNKNOWN_AUTHN_SERVICE = 1747&
Const RPC_S_UNKNOWN_AUTHN_LEVEL = 1748&
Const RPC_S_INVALID_AUTH_IDENTITY = 1749&
Const RPC_S_UNKNOWN_AUTHZ_SERVICE = 1750&
Const EPT_S_INVALID_ENTRY = 1751&
Const EPT_S_CANT_PERFORM_OP = 1752&
Const EPT_S_NOT_REGISTERED = 1753&
Const RPC_S_NOTHING_TO_EXPORT = 1754&
Const RPC_S_INCOMPLETE_NAME = 1755&
Const RPC_S_INVALID_VERS_OPTION = 1756&
Const RPC_S_NO_MORE_MEMBERS = 1757&
Const RPC_S_NOT_ALL_OBJS_UNEXPORTED = 1758&
Const RPC_S_INTERFACE_NOT_FOUND = 1759&
Const RPC_S_ENTRY_ALREADY_EXISTS = 1760&
Const RPC_S_ENTRY_NOT_FOUND = 1761&
Const RPC_S_NAME_SERVICE_UNAVAILABLE = 1762&
Const RPC_S_INVALID_NAF_ID = 1763&
Const RPC_S_CANNOT_SUPPORT = 1764&
Const RPC_S_NO_CONTEXT_AVAILABLE = 1765&
Const RPC_S_INTERNAL_ERROR = 1766&
Const RPC_S_ZERO_DIVIDE = 1767&
Const RPC_S_ADDRESS_ERROR = 1768&
Const RPC_S_FP_DIV_ZERO = 1769&
Const RPC_S_FP_UNDERFLOW = 1770&
Const RPC_S_FP_OVERFLOW = 1771&
Const RPC_X_NO_MORE_ENTRIES = 1772&
Const RPC_X_SS_CHAR_TRANS_OPEN_FAIL = 1773&
Const RPC_X_SS_CHAR_TRANS_SHORT_FILE = 1774&
Const RPC_X_SS_IN_NULL_CONTEXT = 1775&
Const RPC_X_SS_CONTEXT_DAMAGED = 1777&
Const RPC_X_SS_HANDLES_MISMATCH = 1778&
Const RPC_X_SS_CANNOT_GET_CALL_HANDLE = 1779&
Const RPC_X_NULL_REF_POINTER = 1780&
Const RPC_X_ENUM_VALUE_OUT_OF_RANGE = 1781&
Const RPC_X_BYTE_COUNT_TOO_SMALL = 1782&
Const RPC_X_BAD_STUB_DATA = 1783&
Const ERROR_INVALID_USER_BUFFER = 1784&
Const ERROR_UNRECOGNIZED_MEDIA = 1785&
Const ERROR_NO_TRUST_LSA_SECRET = 1786&
Const ERROR_NO_TRUST_SAM_ACCOUNT = 1787&
Const ERROR_TRUSTED_DOMAIN_FAILURE = 1788&
Const ERROR_TRUSTED_RELATIONSHIP_FAILURE = 1789&
Const ERROR_TRUST_FAILURE = 1790&
Const RPC_S_CALL_IN_PROGRESS = 1791&
Const ERROR_NETLOGON_NOT_STARTED = 1792&
Const ERROR_ACCOUNT_EXPIRED = 1793&
Const ERROR_REDIRECTOR_HAS_OPEN_HANDLES = 1794&
Const ERROR_PRINTER_DRIVER_ALREADY_INSTALLED = 1795&
Const ERROR_UNKNOWN_PORT = 1796&
Const ERROR_UNKNOWN_PRINTER_DRIVER = 1797&
Const ERROR_UNKNOWN_PRINTPROCESSOR = 1798&
Const ERROR_INVALID_SEPARATOR_FILE = 1799&
Const ERROR_INVALID_PRIORITY = 1800&
Const ERROR_INVALID_PRINTER_NAME = 1801&
Const ERROR_PRINTER_ALREADY_EXISTS = 1802&
Const ERROR_INVALID_PRINTER_COMMAND = 1803&
Const ERROR_INVALID_DATATYPE = 1804&
Const ERROR_INVALID_ENVIRONMENT = 1805&
Const RPC_S_NO_MORE_BINDINGS = 1806&
Const ERROR_NOLOGON_INTERDOMAIN_TRUST_ACCOUNT = 1807&
Const ERROR_NOLOGON_WORKSTATION_TRUST_ACCOUNT = 1808&
Const ERROR_NOLOGON_SERVER_TRUST_ACCOUNT = 1809&
Const ERROR_DOMAIN_TRUST_INCONSISTENT = 1810&
Const ERROR_SERVER_HAS_OPEN_HANDLES = 1811&
Const ERROR_RESOURCE_DATA_NOT_FOUND = 1812&
Const ERROR_RESOURCE_TYPE_NOT_FOUND = 1813&
Const ERROR_RESOURCE_NAME_NOT_FOUND = 1814&
Const ERROR_RESOURCE_LANG_NOT_FOUND = 1815&
Const ERROR_NOT_ENOUGH_QUOTA = 1816&
Const RPC_S_GROUP_MEMBER_NOT_FOUND = 1898&
Const EPT_S_CANT_CREATE = 1899&
Const RPC_S_INVALID_OBJECT = 1900&
Const ERROR_INVALID_TIME = 1901&
Const ERROR_INVALID_FORM_NAME = 1902&
Const ERROR_INVALID_FORM_SIZE = 1903&
Const ERROR_ALREADY_WAITING = 1904&
Const ERROR_PRINTER_DELETED = 1905&
Const ERROR_INVALID_PRINTER_STATE = 1906&
Const ERROR_NO_BROWSER_SERVERS_FOUND = 6118&
Const MAXPNAMELEN = 32  '  max product name length (including NULL)
Const MAXERRORLENGTH = 128  '  max error text length (including final NULL)
Const TIME_MS = &H1     '  time in Milliseconds
Const TIME_SAMPLES = &H2     '  number of wave samples
Const TIME_BYTES = &H4     '  current byte offset
Const TIME_SMPTE = &H8     '  SMPTE time
Const TIME_MIDI = &H10    '  MIDI time
Const MM_JOY1MOVE = &H3A0  '  joystick
Const MM_JOY2MOVE = &H3A1
Const MM_JOY1ZMOVE = &H3A2
Const MM_JOY2ZMOVE = &H3A3
Const MM_JOY1BUTTONDOWN = &H3B5
Const MM_JOY2BUTTONDOWN = &H3B6
Const MM_JOY1BUTTONUP = &H3B7
Const MM_JOY2BUTTONUP = &H3B8
Const MM_MCINOTIFY = &H3B9  '  MCI
Const MM_MCISYSTEM_STRING = &H3CA
Const MM_WOM_OPEN = &H3BB  '  waveform output
Const MM_WOM_CLOSE = &H3BC
Const MM_WOM_DONE = &H3BD
Const MM_WIM_OPEN = &H3BE  '  waveform input
Const MM_WIM_CLOSE = &H3BF
Const MM_WIM_DATA = &H3C0
Const MM_MIM_OPEN = &H3C1  '  MIDI input
Const MM_MIM_CLOSE = &H3C2
Const MM_MIM_DATA = &H3C3
Const MM_MIM_LONGDATA = &H3C4
Const MM_MIM_ERROR = &H3C5
Const MM_MIM_LONGERROR = &H3C6
Const MM_MOM_OPEN = &H3C7  '  MIDI output
Const MM_MOM_CLOSE = &H3C8
Const MM_MOM_DONE = &H3C9
Const MMSYSERR_BASE = 0
Const WAVERR_BASE = 32
Const MIDIERR_BASE = 64
Const TIMERR_BASE = 96   '  was 128, changed to match Win 31 Sonic
Const JOYERR_BASE = 160
Const MCIERR_BASE = 256
Const MCI_STRING_OFFSET = 512  '  if this number is changed you MUST
Const MCI_VD_OFFSET = 1024
Const MCI_CD_OFFSET = 1088
Const MCI_WAVE_OFFSET = 1152
Const MCI_SEQ_OFFSET = 1216
Const MMSYSERR_NOERROR = 0  '  no error
Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)  '  unspecified error
Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)  '  device ID out of range
Const MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)  '  driver failed enable
Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)  '  device already allocated
Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)  '  device handle is invalid
Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)  '  no device driver present
Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)  '  memory allocation error
Const MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)  '  function isn't supported
Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)  '  error value out of range
Const MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10)    '  invalid flag passed
Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)    '  invalid parameter passed
Const MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12)    '  handle being used
Const MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13)    '  "Specified alias not found in WIN.INI
Const MMSYSERR_LASTERROR = (MMSYSERR_BASE + 13)    '  last error in range
Const MM_MOM_POSITIONCB = &H3CA              '  Callback for MEVT_POSITIONCB
Const MM_MCISIGNAL = &H3CB
Const MM_MIM_MOREDATA = &H3CC                '  MIM_DONE w/ pending events
Const MIDICAPS_STREAM = &H8               '  driver supports midiStreamOut directly
Const MEVT_F_SHORT = &H0&
Const MEVT_F_LONG = &H80000000
Const MEVT_F_CALLBACK = &H40000000
Const MIDISTRM_ERROR = -2
Const MIDIPROP_SET = &H80000000
Const MIDIPROP_GET = &H40000000
Const MIDIPROP_TIMEDIV = &H1&
Const MIDIPROP_TEMPO = &H2&
Const MIXER_SHORT_NAME_CHARS = 16
Const MIXER_LONG_NAME_CHARS = 64
Const MIXERR_BASE = 1024
Const MIXERR_INVALLINE = (MIXERR_BASE + 0)
Const MIXERR_INVALCONTROL = (MIXERR_BASE + 1)
Const MIXERR_INVALVALUE = (MIXERR_BASE + 2)
Const MIXERR_LASTERROR = (MIXERR_BASE + 2)
Const MIXER_OBJECTF_HANDLE = &H80000000
Const MIXER_OBJECTF_MIXER = &H0&
Const MIXER_OBJECTF_HMIXER = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)
Const MIXER_OBJECTF_WAVEOUT = &H10000000
Const MIXER_OBJECTF_HWAVEOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)
Const MIXER_OBJECTF_WAVEIN = &H20000000
Const MIXER_OBJECTF_HWAVEIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)
Const MIXER_OBJECTF_MIDIOUT = &H30000000
Const MIXER_OBJECTF_HMIDIOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)
Const MIXER_OBJECTF_MIDIIN = &H40000000
Const MIXER_OBJECTF_HMIDIIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIIN)
Const MIXER_OBJECTF_AUX = &H50000000
Const MIXERLINE_LINEF_ACTIVE = &H1&
Const MIXERLINE_LINEF_DISCONNECTED = &H8000&
Const MIXERLINE_LINEF_SOURCE = &H80000000
Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 0)
Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 1)
Const MIXERLINE_COMPONENTTYPE_DST_LINE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 2)
Const MIXERLINE_COMPONENTTYPE_DST_MONITOR = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 3)
Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 5)
Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 6)
Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Const MIXERLINE_COMPONENTTYPE_DST_LAST = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 0)
Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)
Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)
Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
Const MIXERLINE_COMPONENTTYPE_SRC_LAST = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
Const MIXERLINE_TARGETTYPE_UNDEFINED = 0
Const MIXERLINE_TARGETTYPE_WAVEOUT = 1
Const MIXERLINE_TARGETTYPE_WAVEIN = 2
Const MIXERLINE_TARGETTYPE_MIDIOUT = 3
Const MIXERLINE_TARGETTYPE_MIDIIN = 4
Const MIXERLINE_TARGETTYPE_AUX = 5
Const MIXER_GETLINEINFOF_DESTINATION = &H0&
Const MIXER_GETLINEINFOF_SOURCE = &H1&
Const MIXER_GETLINEINFOF_LINEID = &H2&
Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Const MIXER_GETLINEINFOF_TARGETTYPE = &H4&
Const MIXER_GETLINEINFOF_QUERYMASK = &HF&
Const MIXERCONTROL_CONTROLF_UNIFORM = &H1&
Const MIXERCONTROL_CONTROLF_MULTIPLE = &H2&
Const MIXERCONTROL_CONTROLF_DISABLED = &H80000000
Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Const MIXERCONTROL_CT_CLASS_TIME = &H60000000
Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000
Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Const MIXERCONTROL_CT_UNITS_MASK = &HFF0000
Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000    '  in 10ths
Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000    '  in 10ths
Const MIXERCONTROL_CONTROLTYPE_CUSTOM = (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Const MIXERCONTROL_CONTROLTYPE_ONOFF = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Const MIXERCONTROL_CONTROLTYPE_MONO = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
Const MIXERCONTROL_CONTROLTYPE_STEREOENH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
Const MIXERCONTROL_CONTROLTYPE_BUTTON = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Const MIXERCONTROL_CONTROLTYPE_DECIBELS = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
Const MIXERCONTROL_CONTROLTYPE_SIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)
Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Const MIXERCONTROL_CONTROLTYPE_PERCENT = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
Const MIXERCONTROL_CONTROLTYPE_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
Const MIXERCONTROL_CONTROLTYPE_PAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Const MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Const MIXERCONTROL_CONTROLTYPE_MIXER = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Const MIXERCONTROL_CONTROLTYPE_MICROTIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Const MIXERCONTROL_CONTROLTYPE_MILLITIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Const MIXER_GETLINECONTROLSF_ALL = &H0&
Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&
Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Const MIXER_GETCONTROLDETAILSF_LISTTEXT = &H1&
Const MIXER_GETCONTROLDETAILSF_QUERYMASK = &HF&
Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Const MIXER_SETCONTROLDETAILSF_CUSTOM = &H1&
Const MIXER_SETCONTROLDETAILSF_QUERYMASK = &HF&
Const JOY_BUTTON5 = &H10&
Const JOY_BUTTON6 = &H20&
Const JOY_BUTTON7 = &H40&
Const JOY_BUTTON8 = &H80&
Const JOY_BUTTON9 = &H100&
Const JOY_BUTTON10 = &H200&
Const JOY_BUTTON11 = &H400&
Const JOY_BUTTON12 = &H800&
Const JOY_BUTTON13 = &H1000&
Const JOY_BUTTON14 = &H2000&
Const JOY_BUTTON15 = &H4000&
Const JOY_BUTTON16 = &H8000&
Const JOY_BUTTON17 = &H10000
Const JOY_BUTTON18 = &H20000
Const JOY_BUTTON19 = &H40000
Const JOY_BUTTON20 = &H80000
Const JOY_BUTTON21 = &H100000
Const JOY_BUTTON22 = &H200000
Const JOY_BUTTON23 = &H400000
Const JOY_BUTTON24 = &H800000
Const JOY_BUTTON25 = &H1000000
Const JOY_BUTTON26 = &H2000000
Const JOY_BUTTON27 = &H4000000
Const JOY_BUTTON28 = &H8000000
Const JOY_BUTTON29 = &H10000000
Const JOY_BUTTON30 = &H20000000
Const JOY_BUTTON31 = &H40000000
Const JOY_BUTTON32 = &H80000000
Const JOY_POVCENTERED = -1
Const JOY_POVFORWARD = 0
Const JOY_POVRIGHT = 9000
Const JOY_POVBACKWARD = 18000
Const JOY_POVLEFT = 27000
Const JOY_RETURNX = &H1&
Const JOY_RETURNY = &H2&
Const JOY_RETURNZ = &H4&
Const JOY_RETURNR = &H8&
Const JOY_RETURNU = &H10                             '  axis 5
Const JOY_RETURNV = &H20                             '  axis 6
Const JOY_RETURNPOV = &H40&
Const JOY_RETURNBUTTONS = &H80&
Const JOY_RETURNRAWDATA = &H100&
Const JOY_RETURNPOVCTS = &H200&
Const JOY_RETURNCENTERED = &H400&
Const JOY_USEDEADZONE = &H800&
Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)
Const JOY_CAL_READALWAYS = &H10000
Const JOY_CAL_READXYONLY = &H20000
Const JOY_CAL_READ3 = &H40000
Const JOY_CAL_READ4 = &H80000
Const JOY_CAL_READXONLY = &H100000
Const JOY_CAL_READYONLY = &H200000
Const JOY_CAL_READ5 = &H400000
Const JOY_CAL_READ6 = &H800000
Const JOY_CAL_READZONLY = &H1000000
Const JOY_CAL_READRONLY = &H2000000
Const JOY_CAL_READUONLY = &H4000000
Const JOY_CAL_READVONLY = &H8000000
Const WAVE_FORMAT_QUERY = &H1
Const SND_PURGE = &H40               '  purge non-static events for task
Const SND_APPLICATION = &H80         '  look for application specific association
Const WAVE_MAPPED = &H4
Const WAVE_FORMAT_DIRECT = &H8
Const WAVE_FORMAT_DIRECT_QUERY = (WAVE_FORMAT_QUERY Or WAVE_FORMAT_DIRECT)
Const MIM_MOREDATA = MM_MIM_MOREDATA
Const MOM_POSITIONCB = MM_MOM_POSITIONCB
Const MIDI_IO_STATUS = &H20&
Const DRV_LOAD = &H1
Const DRV_ENABLE = &H2
Const DRV_OPEN = &H3
Const DRV_CLOSE = &H4
Const DRV_DISABLE = &H5
Const DRV_FREE = &H6
Const DRV_CONFIGURE = &H7
Const DRV_QUERYCONFIGURE = &H8
Const DRV_INSTALL = &H9
Const DRV_REMOVE = &HA
Const DRV_EXITSESSION = &HB
Const DRV_POWER = &HF
Const DRV_RESERVED = &H800
Const DRV_USER = &H4000
Const DRVCNF_CANCEL = &H0
Const DRVCNF_OK = &H1
Const DRVCNF_RESTART = &H2
Const DRV_CANCEL = DRVCNF_CANCEL
Const DRV_OK = DRVCNF_OK
Const DRV_RESTART = DRVCNF_RESTART
Const DRV_MCI_FIRST = DRV_RESERVED
Const DRV_MCI_LAST = DRV_RESERVED + &HFFF
Const CALLBACK_TYPEMASK = &H70000      '  callback type mask
Const CALLBACK_NULL = &H0        '  no callback
Const CALLBACK_WINDOW = &H10000      '  dwCallback is a HWND
Const CALLBACK_TASK = &H20000      '  dwCallback is a HTASK
Const CALLBACK_FUNCTION = &H30000      '  dwCallback is a FARPROC
Const MM_MICROSOFT = 1  '  Microsoft Corp.
Const MM_MIDI_MAPPER = 1  '  MIDI Mapper
Const MM_WAVE_MAPPER = 2  '  Wave Mapper
Const MM_SNDBLST_MIDIOUT = 3  '  Sound Blaster MIDI output port
Const MM_SNDBLST_MIDIIN = 4  '  Sound Blaster MIDI input port
Const MM_SNDBLST_SYNTH = 5  '  Sound Blaster internal synthesizer
Const MM_SNDBLST_WAVEOUT = 6  '  Sound Blaster waveform output
Const MM_SNDBLST_WAVEIN = 7  '  Sound Blaster waveform input
Const MM_ADLIB = 9  '  Ad Lib-compatible synthesizer
Const MM_MPU401_MIDIOUT = 10  '  MPU401-compatible MIDI output port
Const MM_MPU401_MIDIIN = 11  '  MPU401-compatible MIDI input port
Const MM_PC_JOYSTICK = 12  '  Joystick adapter
Const SND_SYNC = &H0         '  play synchronously (default)
Const SND_ASYNC = &H1         '  play asynchronously
Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Const SND_FILENAME = &H20000     '  name is a file name
Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Const SND_ALIAS_START = 0  '  must be > 4096 to keep strings in same section of resource file
Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Const SND_VALID = &H1F        '  valid flags          / ;Internal /
Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Const SND_VALIDFLAGS = &H17201F    '  Set of valid flag bits.  Anything outside
Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
Const SND_TYPE_MASK = &H170007
Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)    '  unsupported wave format
Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)    '  still something playing
Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)    '  header not prepared
Const WAVERR_SYNC = (WAVERR_BASE + 3)    '  device is synchronous
Const WAVERR_LASTERROR = (WAVERR_BASE + 3)    '  last error in range
Const WOM_OPEN = MM_WOM_OPEN
Const WOM_CLOSE = MM_WOM_CLOSE
Const WOM_DONE = MM_WOM_DONE
Const WIM_OPEN = MM_WIM_OPEN
Const WIM_CLOSE = MM_WIM_CLOSE
Const WIM_DATA = MM_WIM_DATA
Const WAVE_MAPPER = -1&
Const WAVE_ALLOWSYNC = &H2
Const WAVE_VALID = &H3         '  ;Internal
Const WHDR_DONE = &H1         '  done bit
Const WHDR_PREPARED = &H2         '  set if this header has been prepared
Const WHDR_BEGINLOOP = &H4         '  loop start block
Const WHDR_ENDLOOP = &H8         '  loop end block
Const WHDR_INQUEUE = &H10        '  reserved for driver
Const WHDR_VALID = &H1F        '  valid flags      / ;Internal /
Const WAVECAPS_PITCH = &H1         '  supports pitch control
Const WAVECAPS_PLAYBACKRATE = &H2         '  supports playback rate control
Const WAVECAPS_VOLUME = &H4         '  supports volume control
Const WAVECAPS_LRVOLUME = &H8         '  separate left-right volume control
Const WAVECAPS_SYNC = &H10
Const WAVE_INVALIDFORMAT = &H0              '  invalid format
Const WAVE_FORMAT_1M08 = &H1              '  11.025 kHz, Mono,   8-bit
Const WAVE_FORMAT_1S08 = &H2              '  11.025 kHz, Stereo, 8-bit
Const WAVE_FORMAT_1M16 = &H4              '  11.025 kHz, Mono,   16-bit
Const WAVE_FORMAT_1S16 = &H8              '  11.025 kHz, Stereo, 16-bit
Const WAVE_FORMAT_2M08 = &H10             '  22.05  kHz, Mono,   8-bit
Const WAVE_FORMAT_2S08 = &H20             '  22.05  kHz, Stereo, 8-bit
Const WAVE_FORMAT_2M16 = &H40             '  22.05  kHz, Mono,   16-bit
Const WAVE_FORMAT_2S16 = &H80             '  22.05  kHz, Stereo, 16-bit
Const WAVE_FORMAT_4M08 = &H100            '  44.1   kHz, Mono,   8-bit
Const WAVE_FORMAT_4S08 = &H200            '  44.1   kHz, Stereo, 8-bit
Const WAVE_FORMAT_4M16 = &H400            '  44.1   kHz, Mono,   16-bit
Const WAVE_FORMAT_4S16 = &H800            '  44.1   kHz, Stereo, 16-bit
Const WAVE_FORMAT_PCM = 1  '  Needed in resource files so outside #ifndef RC_INVOKED
Const MIDIERR_UNPREPARED = (MIDIERR_BASE + 0)   '  header not prepared
Const MIDIERR_STILLPLAYING = (MIDIERR_BASE + 1)   '  still something playing
Const MIDIERR_NOMAP = (MIDIERR_BASE + 2)   '  no current map
Const MIDIERR_NOTREADY = (MIDIERR_BASE + 3)   '  hardware is still busy
Const MIDIERR_NODEVICE = (MIDIERR_BASE + 4)   '  port no longer connected
Const MIDIERR_INVALIDSETUP = (MIDIERR_BASE + 5)   '  invalid setup
Const MIDIERR_LASTERROR = (MIDIERR_BASE + 5)   '  last error in range
Const MIM_OPEN = MM_MIM_OPEN
Const MIM_CLOSE = MM_MIM_CLOSE
Const MIM_DATA = MM_MIM_DATA
Const MIM_LONGDATA = MM_MIM_LONGDATA
Const MIM_ERROR = MM_MIM_ERROR
Const MIM_LONGERROR = MM_MIM_LONGERROR
Const MOM_OPEN = MM_MOM_OPEN
Const MOM_CLOSE = MM_MOM_CLOSE
Const MOM_DONE = MM_MOM_DONE
Const MIDIMAPPER = (-1)  '  Cannot be cast to DWORD as RC complains
Const MIDI_MAPPER = -1&
Const MIDI_CACHE_ALL = 1
Const MIDI_CACHE_BESTFIT = 2
Const MIDI_CACHE_QUERY = 3
Const MIDI_UNCACHE = 4
Const MIDI_CACHE_VALID = (MIDI_CACHE_ALL Or MIDI_CACHE_BESTFIT Or MIDI_CACHE_QUERY Or MIDI_UNCACHE)  '  ;Internal
Const MOD_MIDIPORT = 1  '  output port
Const MOD_SYNTH = 2  '  generic internal synth
Const MOD_SQSYNTH = 3  '  square wave internal synth
Const MOD_FMSYNTH = 4  '  FM internal synth
Const MOD_MAPPER = 5  '  MIDI mapper
Const MIDICAPS_VOLUME = &H1         '  supports volume control
Const MIDICAPS_LRVOLUME = &H2         '  separate left-right volume control
Const MIDICAPS_CACHE = &H4
Const MHDR_DONE = &H1         '  done bit
Const MHDR_PREPARED = &H2         '  set if header prepared
Const MHDR_INQUEUE = &H4         '  reserved for driver
Const MHDR_VALID = &H7         '  valid flags / ;Internal /
Const AUX_MAPPER = -1&
Const AUXCAPS_CDAUDIO = 1  '  audio from internal CD-ROM drive
Const AUXCAPS_AUXIN = 2  '  audio from auxiliary input jacks
Const AUXCAPS_VOLUME = &H1         '  supports volume control
Const AUXCAPS_LRVOLUME = &H2         '  separate left-right volume control
Const TIMERR_NOERROR = (0)  '  no error
Const TIMERR_NOCANDO = (TIMERR_BASE + 1)    '  request not completed
Const TIMERR_STRUCT = (TIMERR_BASE + 33)    '  time struct size
Const TIME_ONESHOT = 0  '  program timer for single event
Const TIME_PERIODIC = 1  '  program for continuous periodic event
Const JOYERR_NOERROR = (0)  '  no error
Const JOYERR_PARMS = (JOYERR_BASE + 5)    '  bad parameters
Const JOYERR_NOCANDO = (JOYERR_BASE + 6)    '  request not completed
Const JOYERR_UNPLUGGED = (JOYERR_BASE + 7)    '  joystick is unplugged
Const JOY_BUTTON1 = &H1
Const JOY_BUTTON2 = &H2
Const JOY_BUTTON3 = &H4
Const JOY_BUTTON4 = &H8
Const JOY_BUTTON1CHG = &H100
Const JOY_BUTTON2CHG = &H200
Const JOY_BUTTON3CHG = &H400
Const JOY_BUTTON4CHG = &H800
Const JOYSTICKID1 = 0
Const JOYSTICKID2 = 1
Const MMIOERR_BASE = 256
Const MMIOERR_FILENOTFOUND = (MMIOERR_BASE + 1)  '  file not found
Const MMIOERR_OUTOFMEMORY = (MMIOERR_BASE + 2)  '  out of memory
Const MMIOERR_CANNOTOPEN = (MMIOERR_BASE + 3)  '  cannot open
Const MMIOERR_CANNOTCLOSE = (MMIOERR_BASE + 4)  '  cannot close
Const MMIOERR_CANNOTREAD = (MMIOERR_BASE + 5)  '  cannot read
Const MMIOERR_CANNOTWRITE = (MMIOERR_BASE + 6)    '  cannot write
Const MMIOERR_CANNOTSEEK = (MMIOERR_BASE + 7)  '  cannot seek
Const MMIOERR_CANNOTEXPAND = (MMIOERR_BASE + 8)  '  cannot expand file
Const MMIOERR_CHUNKNOTFOUND = (MMIOERR_BASE + 9)  '  chunk not found
Const MMIOERR_UNBUFFERED = (MMIOERR_BASE + 10)    '  file is unbuffered
Const CFSEPCHAR = "+"  '  compound file name separator char.
Const MMIO_RWMODE = &H3         '  mask to get bits used for opening
Const MMIO_SHAREMODE = &H70        '  file sharing mode number
Const MMIO_CREATE = &H1000      '  create new file (or truncate file)
Const MMIO_PARSE = &H100       '  parse new file returning path
Const MMIO_DELETE = &H200       '  create new file (or truncate file)
Const MMIO_EXIST = &H4000      '  checks for existence of file
Const MMIO_ALLOCBUF = &H10000     '  mmioOpen() should allocate a buffer
Const MMIO_GETTEMP = &H20000     '  mmioOpen() should retrieve temp name
Const MMIO_DIRTY = &H10000000  '  I/O buffer is dirty
Const MMIO_OPEN_VALID = &H3FFFF     '  valid flags for mmioOpen / ;Internal /
Const MMIO_READ = &H0         '  open file for reading only
Const MMIO_WRITE = &H1         '  open file for writing only
Const MMIO_READWRITE = &H2         '  open file for reading and writing
Const MMIO_COMPAT = &H0         '  compatibility mode
Const MMIO_EXCLUSIVE = &H10        '  exclusive-access mode
Const MMIO_DENYWRITE = &H20        '  deny writing to other processes
Const MMIO_DENYREAD = &H30        '  deny reading to other processes
Const MMIO_DENYNONE = &H40        '  deny nothing to other processes
Const MMIO_FHOPEN = &H10    '  mmioClose(): keep file handle open
Const MMIO_EMPTYBUF = &H10    '  mmioFlush(): empty the I/O buffer
Const MMIO_TOUPPER = &H10    '  mmioStringToFOURCC(): cvt. to u-case
Const MMIO_INSTALLPROC = &H10000     '  mmioInstallIOProc(): install MMIOProc
Const MMIO_PUBLICPROC = &H10000000  '  mmioInstallIOProc: install Globally
Const MMIO_UNICODEPROC = &H1000000   '  mmioInstallIOProc(): Unicode MMIOProc
Const MMIO_REMOVEPROC = &H20000     '  mmioInstallIOProc(): remove MMIOProc
Const MMIO_FINDPROC = &H40000     '  mmioInstallIOProc(): find an MMIOProc
Const MMIO_FINDCHUNK = &H10    '  mmioDescend(): find a chunk by ID
Const MMIO_FINDRIFF = &H20    '  mmioDescend(): find a LIST chunk
Const MMIO_FINDLIST = &H40    '  mmioDescend(): find a RIFF chunk
Const MMIO_CREATERIFF = &H20    '  mmioCreateChunk(): make a LIST chunk
Const MMIO_CREATELIST = &H40    '  mmioCreateChunk(): make a RIFF chunk
Const MMIO_VALIDPROC = &H11070000  '  valid for mmioInstallIOProc / ;Internal /
Const MMIOM_READ = MMIO_READ  '  read (must equal MMIO_READ!)
Const MMIOM_WRITE = MMIO_WRITE  '  write (must equal MMIO_WRITE!)
Const MMIOM_SEEK = 2  '  seek to a new position in file
Const MMIOM_OPEN = 3  '  open file
Const MMIOM_CLOSE = 4  '  close file
Const MMIOM_WRITEFLUSH = 5  '  write and flush
Const MMIOM_RENAME = 6  '  rename specified file
Const MMIOM_USER = &H8000  '  beginning of user-defined messages
Const SEEK_SET = 0  '  seek to an absolute position
Const SEEK_CUR = 1  '  seek relative to current position
Const SEEK_END = 2  '  seek relative to end of file
Const MMIO_DEFAULTBUFFER = 8192  '  default buffer size
Const MCIERR_INVALID_DEVICE_ID = (MCIERR_BASE + 1)
Const MCIERR_UNRECOGNIZED_KEYWORD = (MCIERR_BASE + 3)
Const MCIERR_UNRECOGNIZED_COMMAND = (MCIERR_BASE + 5)
Const MCIERR_HARDWARE = (MCIERR_BASE + 6)
Const MCIERR_INVALID_DEVICE_NAME = (MCIERR_BASE + 7)
Const MCIERR_OUT_OF_MEMORY = (MCIERR_BASE + 8)
Const MCIERR_DEVICE_OPEN = (MCIERR_BASE + 9)
Const MCIERR_CANNOT_LOAD_DRIVER = (MCIERR_BASE + 10)
Const MCIERR_MISSING_COMMAND_STRING = (MCIERR_BASE + 11)
Const MCIERR_PARAM_OVERFLOW = (MCIERR_BASE + 12)
Const MCIERR_MISSING_STRING_ARGUMENT = (MCIERR_BASE + 13)
Const MCIERR_BAD_INTEGER = (MCIERR_BASE + 14)
Const MCIERR_PARSER_INTERNAL = (MCIERR_BASE + 15)
Const MCIERR_DRIVER_INTERNAL = (MCIERR_BASE + 16)
Const MCIERR_MISSING_PARAMETER = (MCIERR_BASE + 17)
Const MCIERR_UNSUPPORTED_FUNCTION = (MCIERR_BASE + 18)
Const MCIERR_FILE_NOT_FOUND = (MCIERR_BASE + 19)
Const MCIERR_DEVICE_NOT_READY = (MCIERR_BASE + 20)
Const MCIERR_INTERNAL = (MCIERR_BASE + 21)
Const MCIERR_DRIVER = (MCIERR_BASE + 22)
Const MCIERR_CANNOT_USE_ALL = (MCIERR_BASE + 23)
Const MCIERR_MULTIPLE = (MCIERR_BASE + 24)
Const MCIERR_EXTENSION_NOT_FOUND = (MCIERR_BASE + 25)
Const MCIERR_OUTOFRANGE = (MCIERR_BASE + 26)
Const MCIERR_FLAGS_NOT_COMPATIBLE = (MCIERR_BASE + 28)
Const MCIERR_FILE_NOT_SAVED = (MCIERR_BASE + 30)
Const MCIERR_DEVICE_TYPE_REQUIRED = (MCIERR_BASE + 31)
Const MCIERR_DEVICE_LOCKED = (MCIERR_BASE + 32)
Const MCIERR_DUPLICATE_ALIAS = (MCIERR_BASE + 33)
Const MCIERR_BAD_CONSTANT = (MCIERR_BASE + 34)
Const MCIERR_MUST_USE_SHAREABLE = (MCIERR_BASE + 35)
Const MCIERR_MISSING_DEVICE_NAME = (MCIERR_BASE + 36)
Const MCIERR_BAD_TIME_FORMAT = (MCIERR_BASE + 37)
Const MCIERR_NO_CLOSING_QUOTE = (MCIERR_BASE + 38)
Const MCIERR_DUPLICATE_FLAGS = (MCIERR_BASE + 39)
Const MCIERR_INVALID_FILE = (MCIERR_BASE + 40)
Const MCIERR_NULL_PARAMETER_BLOCK = (MCIERR_BASE + 41)
Const MCIERR_UNNAMED_RESOURCE = (MCIERR_BASE + 42)
Const MCIERR_NEW_REQUIRES_ALIAS = (MCIERR_BASE + 43)
Const MCIERR_NOTIFY_ON_AUTO_OPEN = (MCIERR_BASE + 44)
Const MCIERR_NO_ELEMENT_ALLOWED = (MCIERR_BASE + 45)
Const MCIERR_NONAPPLICABLE_FUNCTION = (MCIERR_BASE + 46)
Const MCIERR_ILLEGAL_FOR_AUTO_OPEN = (MCIERR_BASE + 47)
Const MCIERR_FILENAME_REQUIRED = (MCIERR_BASE + 48)
Const MCIERR_EXTRA_CHARACTERS = (MCIERR_BASE + 49)
Const MCIERR_DEVICE_NOT_INSTALLED = (MCIERR_BASE + 50)
Const MCIERR_GET_CD = (MCIERR_BASE + 51)
Const MCIERR_SET_CD = (MCIERR_BASE + 52)
Const MCIERR_SET_DRIVE = (MCIERR_BASE + 53)
Const MCIERR_DEVICE_LENGTH = (MCIERR_BASE + 54)
Const MCIERR_DEVICE_ORD_LENGTH = (MCIERR_BASE + 55)
Const MCIERR_NO_INTEGER = (MCIERR_BASE + 56)
Const MCIERR_WAVE_OUTPUTSINUSE = (MCIERR_BASE + 64)
Const MCIERR_WAVE_SETOUTPUTINUSE = (MCIERR_BASE + 65)
Const MCIERR_WAVE_INPUTSINUSE = (MCIERR_BASE + 66)
Const MCIERR_WAVE_SETINPUTINUSE = (MCIERR_BASE + 67)
Const MCIERR_WAVE_OUTPUTUNSPECIFIED = (MCIERR_BASE + 68)
Const MCIERR_WAVE_INPUTUNSPECIFIED = (MCIERR_BASE + 69)
Const MCIERR_WAVE_OUTPUTSUNSUITABLE = (MCIERR_BASE + 70)
Const MCIERR_WAVE_SETOUTPUTUNSUITABLE = (MCIERR_BASE + 71)
Const MCIERR_WAVE_INPUTSUNSUITABLE = (MCIERR_BASE + 72)
Const MCIERR_WAVE_SETINPUTUNSUITABLE = (MCIERR_BASE + 73)
Const MCIERR_SEQ_DIV_INCOMPATIBLE = (MCIERR_BASE + 80)
Const MCIERR_SEQ_PORT_INUSE = (MCIERR_BASE + 81)
Const MCIERR_SEQ_PORT_NONEXISTENT = (MCIERR_BASE + 82)
Const MCIERR_SEQ_PORT_MAPNODEVICE = (MCIERR_BASE + 83)
Const MCIERR_SEQ_PORT_MISCERROR = (MCIERR_BASE + 84)
Const MCIERR_SEQ_TIMER = (MCIERR_BASE + 85)
Const MCIERR_SEQ_PORTUNSPECIFIED = (MCIERR_BASE + 86)
Const MCIERR_SEQ_NOMIDIPRESENT = (MCIERR_BASE + 87)
Const MCIERR_NO_WINDOW = (MCIERR_BASE + 90)
Const MCIERR_CREATEWINDOW = (MCIERR_BASE + 91)
Const MCIERR_FILE_READ = (MCIERR_BASE + 92)
Const MCIERR_FILE_WRITE = (MCIERR_BASE + 93)
Const MCIERR_CUSTOM_DRIVER_BASE = (MCIERR_BASE + 256)
Const MCI_FIRST = &H800
Const MCI_OPEN = &H803
Const MCI_CLOSE = &H804
Const MCI_ESCAPE = &H805
Const MCI_PLAY = &H806
Const MCI_SEEK = &H807
Const MCI_STOP = &H808
Const MCI_PAUSE = &H809
Const MCI_INFO = &H80A
Const MCI_GETDEVCAPS = &H80B
Const MCI_SPIN = &H80C
Const MCI_SET = &H80D
Const MCI_STEP = &H80E
Const MCI_RECORD = &H80F
Const MCI_SYSINFO = &H810
Const MCI_BREAK = &H811
Const MCI_SOUND = &H812
Const MCI_SAVE = &H813
Const MCI_STATUS = &H814
Const MCI_CUE = &H830
Const MCI_REALIZE = &H840
Const MCI_WINDOW = &H841
Const MCI_PUT = &H842
Const MCI_WHERE = &H843
Const MCI_FREEZE = &H844
Const MCI_UNFREEZE = &H845
Const MCI_LOAD = &H850
Const MCI_CUT = &H851
Const MCI_COPY = &H852
Const MCI_PASTE = &H853
Const MCI_UPDATE = &H854
Const MCI_RESUME = &H855
Const MCI_DELETE = &H856
Const MCI_LAST = &HFFF
Const MCI_USER_MESSAGES = (&H400 + MCI_FIRST)
Const MCI_ALL_DEVICE_ID = -1   '  Matches all MCI devices
Const MCI_DEVTYPE_VCR = 513
Const MCI_DEVTYPE_VIDEODISC = 514
Const MCI_DEVTYPE_OVERLAY = 515
Const MCI_DEVTYPE_CD_AUDIO = 516
Const MCI_DEVTYPE_DAT = 517
Const MCI_DEVTYPE_SCANNER = 518
Const MCI_DEVTYPE_ANIMATION = 519
Const MCI_DEVTYPE_DIGITAL_VIDEO = 520
Const MCI_DEVTYPE_OTHER = 521
Const MCI_DEVTYPE_WAVEFORM_AUDIO = 522
Const MCI_DEVTYPE_SEQUENCER = 523
Const MCI_DEVTYPE_FIRST = MCI_DEVTYPE_VCR
Const MCI_DEVTYPE_LAST = MCI_DEVTYPE_SEQUENCER
Const MCI_DEVTYPE_FIRST_USER = &H1000
Const MCI_MODE_NOT_READY = (MCI_STRING_OFFSET + 12)
Const MCI_MODE_STOP = (MCI_STRING_OFFSET + 13)
Const MCI_MODE_PLAY = (MCI_STRING_OFFSET + 14)
Const MCI_MODE_RECORD = (MCI_STRING_OFFSET + 15)
Const MCI_MODE_SEEK = (MCI_STRING_OFFSET + 16)
Const MCI_MODE_PAUSE = (MCI_STRING_OFFSET + 17)
Const MCI_MODE_OPEN = (MCI_STRING_OFFSET + 18)
Const MCI_FORMAT_MILLISECONDS = 0
Const MCI_FORMAT_HMS = 1
Const MCI_FORMAT_MSF = 2
Const MCI_FORMAT_FRAMES = 3
Const MCI_FORMAT_SMPTE_24 = 4
Const MCI_FORMAT_SMPTE_25 = 5
Const MCI_FORMAT_SMPTE_30 = 6
Const MCI_FORMAT_SMPTE_30DROP = 7
Const MCI_FORMAT_BYTES = 8
Const MCI_FORMAT_SAMPLES = 9
Const MCI_FORMAT_TMSF = 10
Const MCI_NOTIFY_SUCCESSFUL = &H1
Const MCI_NOTIFY_SUPERSEDED = &H2
Const MCI_NOTIFY_ABORTED = &H4
Const MCI_NOTIFY_FAILURE = &H8
Const MCI_NOTIFY = &H1&
Const MCI_WAIT = &H2&
Const MCI_FROM = &H4&
Const MCI_TO = &H8&
Const MCI_TRACK = &H10&
Const MCI_OPEN_SHAREABLE = &H100&
Const MCI_OPEN_ELEMENT = &H200&
Const MCI_OPEN_ALIAS = &H400&
Const MCI_OPEN_ELEMENT_ID = &H800&
Const MCI_OPEN_TYPE_ID = &H1000&
Const MCI_OPEN_TYPE = &H2000&
Const MCI_SEEK_TO_START = &H100&
Const MCI_SEEK_TO_END = &H200&
Const MCI_STATUS_ITEM = &H100&
Const MCI_STATUS_START = &H200&
Const MCI_STATUS_LENGTH = &H1&
Const MCI_STATUS_POSITION = &H2&
Const MCI_STATUS_NUMBER_OF_TRACKS = &H3&
Const MCI_STATUS_MODE = &H4&
Const MCI_STATUS_MEDIA_PRESENT = &H5&
Const MCI_STATUS_TIME_FORMAT = &H6&
Const MCI_STATUS_READY = &H7&
Const MCI_STATUS_CURRENT_TRACK = &H8&
Const MCI_INFO_PRODUCT = &H100&
Const MCI_INFO_FILE = &H200&
Const MCI_GETDEVCAPS_ITEM = &H100&
Const MCI_GETDEVCAPS_CAN_RECORD = &H1&
Const MCI_GETDEVCAPS_HAS_AUDIO = &H2&
Const MCI_GETDEVCAPS_HAS_VIDEO = &H3&
Const MCI_GETDEVCAPS_DEVICE_TYPE = &H4&
Const MCI_GETDEVCAPS_USES_FILES = &H5&
Const MCI_GETDEVCAPS_COMPOUND_DEVICE = &H6&
Const MCI_GETDEVCAPS_CAN_EJECT = &H7&
Const MCI_GETDEVCAPS_CAN_PLAY = &H8&
Const MCI_GETDEVCAPS_CAN_SAVE = &H9&
Const MCI_SYSINFO_QUANTITY = &H100&
Const MCI_SYSINFO_OPEN = &H200&
Const MCI_SYSINFO_NAME = &H400&
Const MCI_SYSINFO_INSTALLNAME = &H800&
Const MCI_SET_DOOR_OPEN = &H100&
Const MCI_SET_DOOR_CLOSED = &H200&
Const MCI_SET_TIME_FORMAT = &H400&
Const MCI_SET_AUDIO = &H800&
Const MCI_SET_VIDEO = &H1000&
Const MCI_SET_ON = &H2000&
Const MCI_SET_OFF = &H4000&
Const MCI_SET_AUDIO_ALL = &H4001&
Const MCI_SET_AUDIO_LEFT = &H4002&
Const MCI_SET_AUDIO_RIGHT = &H4003&
Const MCI_BREAK_KEY = &H100&
Const MCI_BREAK_HWND = &H200&
Const MCI_BREAK_OFF = &H400&
Const MCI_RECORD_INSERT = &H100&
Const MCI_RECORD_OVERWRITE = &H200&
Const MCI_SOUND_NAME = &H100&
Const MCI_SAVE_FILE = &H100&
Const MCI_LOAD_FILE = &H100&
Const MCI_VD_MODE_PARK = (MCI_VD_OFFSET + 1)
Const MCI_VD_MEDIA_CLV = (MCI_VD_OFFSET + 2)
Const MCI_VD_MEDIA_CAV = (MCI_VD_OFFSET + 3)
Const MCI_VD_MEDIA_OTHER = (MCI_VD_OFFSET + 4)
Const MCI_VD_FORMAT_TRACK = &H4001
Const MCI_VD_PLAY_REVERSE = &H10000
Const MCI_VD_PLAY_FAST = &H20000
Const MCI_VD_PLAY_SPEED = &H40000
Const MCI_VD_PLAY_SCAN = &H80000
Const MCI_VD_PLAY_SLOW = &H100000
Const MCI_VD_SEEK_REVERSE = &H10000
Const MCI_VD_STATUS_SPEED = &H4002&
Const MCI_VD_STATUS_FORWARD = &H4003&
Const MCI_VD_STATUS_MEDIA_TYPE = &H4004&
Const MCI_VD_STATUS_SIDE = &H4005&
Const MCI_VD_STATUS_DISC_SIZE = &H4006&
Const MCI_VD_GETDEVCAPS_CLV = &H10000
Const MCI_VD_GETDEVCAPS_CAV = &H20000
Const MCI_VD_SPIN_UP = &H10000
Const MCI_VD_SPIN_DOWN = &H20000
Const MCI_VD_GETDEVCAPS_CAN_REVERSE = &H4002&
Const MCI_VD_GETDEVCAPS_FAST_RATE = &H4003&
Const MCI_VD_GETDEVCAPS_SLOW_RATE = &H4004&
Const MCI_VD_GETDEVCAPS_NORMAL_RATE = &H4005&
Const MCI_VD_STEP_FRAMES = &H10000
Const MCI_VD_STEP_REVERSE = &H20000
Const MCI_VD_ESCAPE_STRING = &H100&
Const MCI_WAVE_PCM = (MCI_WAVE_OFFSET + 0)
Const MCI_WAVE_MAPPER = (MCI_WAVE_OFFSET + 1)
Const MCI_WAVE_OPEN_BUFFER = &H10000
Const MCI_WAVE_SET_FORMATTAG = &H10000
Const MCI_WAVE_SET_CHANNELS = &H20000
Const MCI_WAVE_SET_SAMPLESPERSEC = &H40000
Const MCI_WAVE_SET_AVGBYTESPERSEC = &H80000
Const MCI_WAVE_SET_BLOCKALIGN = &H100000
Const MCI_WAVE_SET_BITSPERSAMPLE = &H200000
Const MCI_WAVE_INPUT = &H400000
Const MCI_WAVE_OUTPUT = &H800000
Const MCI_WAVE_STATUS_FORMATTAG = &H4001&
Const MCI_WAVE_STATUS_CHANNELS = &H4002&
Const MCI_WAVE_STATUS_SAMPLESPERSEC = &H4003&
Const MCI_WAVE_STATUS_AVGBYTESPERSEC = &H4004&
Const MCI_WAVE_STATUS_BLOCKALIGN = &H4005&
Const MCI_WAVE_STATUS_BITSPERSAMPLE = &H4006&
Const MCI_WAVE_STATUS_LEVEL = &H4007&
Const MCI_WAVE_SET_ANYINPUT = &H4000000
Const MCI_WAVE_SET_ANYOUTPUT = &H8000000
Const MCI_WAVE_GETDEVCAPS_INPUTS = &H4001&
Const MCI_WAVE_GETDEVCAPS_OUTPUTS = &H4002&
Const MCI_SEQ_DIV_PPQN = (0 + MCI_SEQ_OFFSET)
Const MCI_SEQ_DIV_SMPTE_24 = (1 + MCI_SEQ_OFFSET)
Const MCI_SEQ_DIV_SMPTE_25 = (2 + MCI_SEQ_OFFSET)
Const MCI_SEQ_DIV_SMPTE_30DROP = (3 + MCI_SEQ_OFFSET)
Const MCI_SEQ_DIV_SMPTE_30 = (4 + MCI_SEQ_OFFSET)
Const MCI_SEQ_FORMAT_SONGPTR = &H4001
Const MCI_SEQ_FILE = &H4002
Const MCI_SEQ_MIDI = &H4003
Const MCI_SEQ_SMPTE = &H4004
Const MCI_SEQ_NONE = 65533
Const MCI_SEQ_MAPPER = 65535
Const MCI_SEQ_STATUS_TEMPO = &H4002&
Const MCI_SEQ_STATUS_PORT = &H4003&
Const MCI_SEQ_STATUS_SLAVE = &H4007&
Const MCI_SEQ_STATUS_MASTER = &H4008&
Const MCI_SEQ_STATUS_OFFSET = &H4009&
Const MCI_SEQ_STATUS_DIVTYPE = &H400A&
Const MCI_SEQ_SET_TEMPO = &H10000
Const MCI_SEQ_SET_PORT = &H20000
Const MCI_SEQ_SET_SLAVE = &H40000
Const MCI_SEQ_SET_MASTER = &H80000
Const MCI_SEQ_SET_OFFSET = &H1000000
Const MCI_ANIM_OPEN_WS = &H10000
Const MCI_ANIM_OPEN_PARENT = &H20000
Const MCI_ANIM_OPEN_NOSTATIC = &H40000
Const MCI_ANIM_PLAY_SPEED = &H10000
Const MCI_ANIM_PLAY_REVERSE = &H20000
Const MCI_ANIM_PLAY_FAST = &H40000
Const MCI_ANIM_PLAY_SLOW = &H80000
Const MCI_ANIM_PLAY_SCAN = &H100000
Const MCI_ANIM_STEP_REVERSE = &H10000
Const MCI_ANIM_STEP_FRAMES = &H20000
Const MCI_ANIM_STATUS_SPEED = &H4001&
Const MCI_ANIM_STATUS_FORWARD = &H4002&
Const MCI_ANIM_STATUS_HWND = &H4003&
Const MCI_ANIM_STATUS_HPAL = &H4004&
Const MCI_ANIM_STATUS_STRETCH = &H4005&
Const MCI_ANIM_INFO_TEXT = &H10000
Const MCI_ANIM_GETDEVCAPS_CAN_REVERSE = &H4001&
Const MCI_ANIM_GETDEVCAPS_FAST_RATE = &H4002&
Const MCI_ANIM_GETDEVCAPS_SLOW_RATE = &H4003&
Const MCI_ANIM_GETDEVCAPS_NORMAL_RATE = &H4004&
Const MCI_ANIM_GETDEVCAPS_PALETTES = &H4006&
Const MCI_ANIM_GETDEVCAPS_CAN_STRETCH = &H4007&
Const MCI_ANIM_GETDEVCAPS_MAX_WINDOWS = &H4008&
Const MCI_ANIM_REALIZE_NORM = &H10000
Const MCI_ANIM_REALIZE_BKGD = &H20000
Const MCI_ANIM_WINDOW_HWND = &H10000
Const MCI_ANIM_WINDOW_STATE = &H40000
Const MCI_ANIM_WINDOW_TEXT = &H80000
Const MCI_ANIM_WINDOW_ENABLE_STRETCH = &H100000
Const MCI_ANIM_WINDOW_DISABLE_STRETCH = &H200000
Const MCI_ANIM_WINDOW_DEFAULT = &H0&
Const MCI_ANIM_RECT = &H10000
Const MCI_ANIM_PUT_SOURCE = &H20000      '  also  MCI_WHERE
Const MCI_ANIM_PUT_DESTINATION = &H40000      '  also  MCI_WHERE
Const MCI_ANIM_WHERE_SOURCE = &H20000
Const MCI_ANIM_WHERE_DESTINATION = &H40000
Const MCI_ANIM_UPDATE_HDC = &H20000
Const MCI_OVLY_OPEN_WS = &H10000
Const MCI_OVLY_OPEN_PARENT = &H20000
Const MCI_OVLY_STATUS_HWND = &H4001&
Const MCI_OVLY_STATUS_STRETCH = &H4002&
Const MCI_OVLY_INFO_TEXT = &H10000
Const MCI_OVLY_GETDEVCAPS_CAN_STRETCH = &H4001&
Const MCI_OVLY_GETDEVCAPS_CAN_FREEZE = &H4002&
Const MCI_OVLY_GETDEVCAPS_MAX_WINDOWS = &H4003&
Const MCI_OVLY_WINDOW_HWND = &H10000
Const MCI_OVLY_WINDOW_STATE = &H40000
Const MCI_OVLY_WINDOW_TEXT = &H80000
Const MCI_OVLY_WINDOW_ENABLE_STRETCH = &H100000
Const MCI_OVLY_WINDOW_DISABLE_STRETCH = &H200000
Const MCI_OVLY_WINDOW_DEFAULT = &H0&
Const MCI_OVLY_RECT = &H10000
Const MCI_OVLY_PUT_SOURCE = &H20000
Const MCI_OVLY_PUT_DESTINATION = &H40000
Const MCI_OVLY_PUT_FRAME = &H80000
Const MCI_OVLY_PUT_VIDEO = &H100000
Const MCI_OVLY_WHERE_SOURCE = &H20000
Const MCI_OVLY_WHERE_DESTINATION = &H40000
Const MCI_OVLY_WHERE_FRAME = &H80000
Const MCI_OVLY_WHERE_VIDEO = &H100000
Const CAPS1 = 94              '  other caps
Const C1_TRANSPARENT = &H1     '  new raster cap
Const NEWTRANSPARENT = 3  '  use with SetBkMode()
Const QUERYROPSUPPORT = 40  '  use to determine ROP support
Const SELECTDIB = 41  '  DIB.DRV select dib escape
Const SE_ERR_SHARE = 26
Const SE_ERR_ASSOCINCOMPLETE = 27
Const SE_ERR_DDETIMEOUT = 28
Const SE_ERR_DDEFAIL = 29
Const SE_ERR_DDEBUSY = 30
Const SE_ERR_NOASSOC = 31
Const PRINTER_CONTROL_PAUSE = 1
Const PRINTER_CONTROL_RESUME = 2
Const PRINTER_CONTROL_PURGE = 3
Const PRINTER_STATUS_PAUSED = &H1
Const PRINTER_STATUS_ERROR = &H2
Const PRINTER_STATUS_PENDING_DELETION = &H4
Const PRINTER_STATUS_PAPER_JAM = &H8
Const PRINTER_STATUS_PAPER_OUT = &H10
Const PRINTER_STATUS_MANUAL_FEED = &H20
Const PRINTER_STATUS_PAPER_PROBLEM = &H40
Const PRINTER_STATUS_OFFLINE = &H80
Const PRINTER_STATUS_IO_ACTIVE = &H100
Const PRINTER_STATUS_BUSY = &H200
Const PRINTER_STATUS_PRINTING = &H400
Const PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
Const PRINTER_STATUS_NOT_AVAILABLE = &H1000
Const PRINTER_STATUS_WAITING = &H2000
Const PRINTER_STATUS_PROCESSING = &H4000
Const PRINTER_STATUS_INITIALIZING = &H8000
Const PRINTER_STATUS_WARMING_UP = &H10000
Const PRINTER_STATUS_TONER_LOW = &H20000
Const PRINTER_STATUS_NO_TONER = &H40000
Const PRINTER_STATUS_PAGE_PUNT = &H80000
Const PRINTER_STATUS_USER_INTERVENTION = &H100000
Const PRINTER_STATUS_OUT_OF_MEMORY = &H200000
Const PRINTER_STATUS_DOOR_OPEN = &H400000
Const PRINTER_ATTRIBUTE_QUEUED = &H1
Const PRINTER_ATTRIBUTE_DIRECT = &H2
Const PRINTER_ATTRIBUTE_DEFAULT = &H4
Const PRINTER_ATTRIBUTE_SHARED = &H8
Const PRINTER_ATTRIBUTE_NETWORK = &H10
Const PRINTER_ATTRIBUTE_HIDDEN = &H20
Const PRINTER_ATTRIBUTE_LOCAL = &H40
Const NO_PRIORITY = 0
Const MAX_PRIORITY = 99
Const MIN_PRIORITY = 1
Const DEF_PRIORITY = 1
Const JOB_CONTROL_PAUSE = 1
Const JOB_CONTROL_RESUME = 2
Const JOB_CONTROL_CANCEL = 3
Const JOB_CONTROL_RESTART = 4
Const JOB_STATUS_PAUSED = &H1
Const JOB_STATUS_ERROR = &H2
Const JOB_STATUS_DELETING = &H4
Const JOB_STATUS_SPOOLING = &H8
Const JOB_STATUS_PRINTING = &H10
Const JOB_STATUS_OFFLINE = &H20
Const JOB_STATUS_PAPEROUT = &H40
Const JOB_STATUS_PRINTED = &H80
Const JOB_POSITION_UNSPECIFIED = 0
Const FORM_BUILTIN = &H1
Const PRINTER_CONTROL_SET_STATUS = 4
Const PRINTER_ATTRIBUTE_WORK_OFFLINE = &H400
Const PRINTER_ATTRIBUTE_ENABLE_BIDI = &H800
Const JOB_CONTROL_DELETE = 5
Const JOB_STATUS_USER_INTERVENTION = &H10000
Const DI_CHANNEL = 1                  '  start direct read/write channel,
Const DI_READ_SPOOL_JOB = 3
Const PORT_TYPE_WRITE = &H1
Const PORT_TYPE_READ = &H2
Const PORT_TYPE_REDIRECTED = &H4
Const PORT_TYPE_NET_ATTACHED = &H8
Const PRINTER_ENUM_DEFAULT = &H1
Const PRINTER_ENUM_LOCAL = &H2
Const PRINTER_ENUM_CONNECTIONS = &H4
Const PRINTER_ENUM_FAVORITE = &H4
Const PRINTER_ENUM_NAME = &H8
Const PRINTER_ENUM_REMOTE = &H10
Const PRINTER_ENUM_SHARED = &H20
Const PRINTER_ENUM_NETWORK = &H40
Const PRINTER_ENUM_EXPAND = &H4000
Const PRINTER_ENUM_CONTAINER = &H8000
Const PRINTER_ENUM_ICONMASK = &HFF0000
Const PRINTER_ENUM_ICON1 = &H10000
Const PRINTER_ENUM_ICON2 = &H20000
Const PRINTER_ENUM_ICON3 = &H40000
Const PRINTER_ENUM_ICON4 = &H80000
Const PRINTER_ENUM_ICON5 = &H100000
Const PRINTER_ENUM_ICON6 = &H200000
Const PRINTER_ENUM_ICON7 = &H400000
Const PRINTER_ENUM_ICON8 = &H800000
Const PRINTER_CHANGE_ADD_PRINTER = &H1
Const PRINTER_CHANGE_SET_PRINTER = &H2
Const PRINTER_CHANGE_DELETE_PRINTER = &H4
Const PRINTER_CHANGE_PRINTER = &HFF
Const PRINTER_CHANGE_ADD_JOB = &H100
Const PRINTER_CHANGE_SET_JOB = &H200
Const PRINTER_CHANGE_DELETE_JOB = &H400
Const PRINTER_CHANGE_WRITE_JOB = &H800
Const PRINTER_CHANGE_JOB = &HFF00
Const PRINTER_CHANGE_ADD_FORM = &H10000
Const PRINTER_CHANGE_SET_FORM = &H20000
Const PRINTER_CHANGE_DELETE_FORM = &H40000
Const PRINTER_CHANGE_FORM = &H70000
Const PRINTER_CHANGE_ADD_PORT = &H100000
Const PRINTER_CHANGE_CONFIGURE_PORT = &H200000
Const PRINTER_CHANGE_DELETE_PORT = &H400000
Const PRINTER_CHANGE_PORT = &H700000
Const PRINTER_CHANGE_ADD_PRINT_PROCESSOR = &H1000000
Const PRINTER_CHANGE_DELETE_PRINT_PROCESSOR = &H4000000
Const PRINTER_CHANGE_PRINT_PROCESSOR = &H7000000
Const PRINTER_CHANGE_ADD_PRINTER_DRIVER = &H10000000
Const PRINTER_CHANGE_DELETE_PRINTER_DRIVER = &H40000000
Const PRINTER_CHANGE_PRINTER_DRIVER = &H70000000
Const PRINTER_CHANGE_TIMEOUT = &H80000000
Const PRINTER_CHANGE_ALL = &H7777FFFF
Const PRINTER_ERROR_INFORMATION = &H80000000
Const PRINTER_ERROR_WARNING = &H40000000
Const PRINTER_ERROR_SEVERE = &H20000000
Const PRINTER_ERROR_OUTOFPAPER = &H1
Const PRINTER_ERROR_JAM = &H2
Const PRINTER_ERROR_OUTOFTONER = &H4
Const SERVER_ACCESS_ADMINISTER = &H1
Const SERVER_ACCESS_ENUMERATE = &H2
Const PRINTER_ACCESS_ADMINISTER = &H4
Const PRINTER_ACCESS_USE = &H8
Const JOB_ACCESS_ADMINISTER = &H10
Const SERVER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVER_ACCESS_ADMINISTER Or SERVER_ACCESS_ENUMERATE)
Const SERVER_READ = (STANDARD_RIGHTS_READ Or SERVER_ACCESS_ENUMERATE)
Const SERVER_WRITE = (STANDARD_RIGHTS_WRITE Or SERVER_ACCESS_ADMINISTER Or SERVER_ACCESS_ENUMERATE)
Const SERVER_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or SERVER_ACCESS_ENUMERATE)
Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
Const PRINTER_READ = (STANDARD_RIGHTS_READ Or PRINTER_ACCESS_USE)
Const PRINTER_WRITE = (STANDARD_RIGHTS_WRITE Or PRINTER_ACCESS_USE)
Const PRINTER_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or PRINTER_ACCESS_USE)
Const JOB_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or JOB_ACCESS_ADMINISTER)
Const JOB_READ = (STANDARD_RIGHTS_READ Or JOB_ACCESS_ADMINISTER)
Const JOB_WRITE = (STANDARD_RIGHTS_WRITE Or JOB_ACCESS_ADMINISTER)
Const JOB_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or JOB_ACCESS_ADMINISTER)
Const RESOURCE_CONNECTED = &H1
Const RESOURCE_PUBLICNET = &H2
Const RESOURCE_REMEMBERED = &H3
Const RESOURCETYPE_ANY = &H0
Const RESOURCETYPE_DISK = &H1
Const RESOURCETYPE_PRINT = &H2
Const RESOURCETYPE_UNKNOWN = &HFFFF
Const RESOURCEUSAGE_CONNECTABLE = &H1
Const RESOURCEUSAGE_CONTAINER = &H2
Const RESOURCEUSAGE_RESERVED = &H80000000
Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Const RESOURCEDISPLAYTYPE_SERVER = &H2
Const RESOURCEDISPLAYTYPE_SHARE = &H3
Const RESOURCEDISPLAYTYPE_FILE = &H4
Const RESOURCEDISPLAYTYPE_GROUP = &H5
Const CONNECT_UPDATE_PROFILE = &H1
Const WN_SUCCESS = NO_ERROR
Const WN_NOT_SUPPORTED = ERROR_NOT_SUPPORTED
Const WN_NET_ERROR = ERROR_UNEXP_NET_ERR
Const WN_MORE_DATA = ERROR_MORE_DATA
Const WN_BAD_POINTER = ERROR_INVALID_ADDRESS
Const WN_BAD_VALUE = ERROR_INVALID_PARAMETER
Const WN_BAD_PASSWORD = ERROR_INVALID_PASSWORD
Const WN_ACCESS_DENIED = ERROR_ACCESS_DENIED
Const WN_FUNCTION_BUSY = ERROR_BUSY
Const WN_WINDOWS_ERROR = ERROR_UNEXP_NET_ERR
Const WN_BAD_USER = ERROR_BAD_USERNAME
Const WN_OUT_OF_MEMORY = ERROR_NOT_ENOUGH_MEMORY
Const WN_NO_NETWORK = ERROR_NO_NETWORK
Const WN_EXTENDED_ERROR = ERROR_EXTENDED_ERROR
Const WN_NOT_CONNECTED = ERROR_NOT_CONNECTED
Const WN_OPEN_FILES = ERROR_OPEN_FILES
Const WN_DEVICE_IN_USE = ERROR_DEVICE_IN_USE
Const WN_BAD_NETNAME = ERROR_BAD_NET_NAME
Const WN_BAD_LOCALNAME = ERROR_BAD_DEVICE
Const WN_ALREADY_CONNECTED = ERROR_ALREADY_ASSIGNED
Const WN_DEVICE_ERROR = ERROR_GEN_FAILURE
Const WN_CONNECTION_CLOSED = ERROR_CONNECTION_UNAVAIL
Const WN_NO_NET_OR_BAD_PATH = ERROR_NO_NET_OR_BAD_PATH
Const WN_BAD_PROVIDER = ERROR_BAD_PROVIDER
Const WN_CANNOT_OPEN_PROFILE = ERROR_CANNOT_OPEN_PROFILE
Const WN_BAD_PROFILE = ERROR_BAD_PROFILE
Const WN_BAD_HANDLE = ERROR_INVALID_HANDLE
Const WN_NO_MORE_ENTRIES = ERROR_NO_MORE_ITEMS
Const WN_NOT_CONTAINER = ERROR_NOT_CONTAINER
Const WN_NO_ERROR = NO_ERROR
Const NCBNAMSZ = 16  '  absolute length of a net name
Const MAX_LANA = 254  '  lana's in range 0 to MAX_LANA
Const NAME_FLAGS_MASK = &H87
Const GROUP_NAME = &H80
Const UNIQUE_NAME = &H0
Const REGISTERING = &H0
Const REGISTERED = &H4
Const DEREGISTERED = &H5
Const DUPLICATE = &H6
Const DUPLICATE_DEREG = &H7
Const LISTEN_OUTSTANDING = &H1
Const CALL_PENDING = &H2
Const SESSION_ESTABLISHED = &H3
Const HANGUP_PENDING = &H4
Const HANGUP_COMPLETE = &H5
Const SESSION_ABORTED = &H6
Const ALL_TRANSPORTS = "M\0\0\0"
Const MS_NBF = "MNBF"
Const NCBCALL = &H10  '  NCB CALL
Const NCBLISTEN = &H11  '  NCB LISTEN
Const NCBHANGUP = &H12  '  NCB HANG UP
Const NCBSEND = &H14  '  NCB SEND
Const NCBRECV = &H15  '  NCB RECEIVE
Const NCBRECVANY = &H16  '  NCB RECEIVE ANY
Const NCBCHAINSEND = &H17  '  NCB CHAIN SEND
Const NCBDGSEND = &H20  '  NCB SEND DATAGRAM
Const NCBDGRECV = &H21  '  NCB RECEIVE DATAGRAM
Const NCBDGSENDBC = &H22  '  NCB SEND BROADCAST DATAGRAM
Const NCBDGRECVBC = &H23  '  NCB RECEIVE BROADCAST DATAGRAM
Const NCBADDNAME = &H30  '  NCB ADD NAME
Const NCBDELNAME = &H31  '  NCB DELETE NAME
Const NCBRESET = &H32  '  NCB RESET
Const NCBASTAT = &H33  '  NCB ADAPTER STATUS
Const NCBSSTAT = &H34  '  NCB SESSION STATUS
Const NCBCANCEL = &H35  '  NCB CANCEL
Const NCBADDGRNAME = &H36  '  NCB ADD GROUP NAME
Const NCBENUM = &H37  '  NCB ENUMERATE LANA NUMBERS
Const NCBUNLINK = &H70  '  NCB UNLINK
Const NCBSENDNA = &H71  '  NCB SEND NO ACK
Const NCBCHAINSENDNA = &H72  '  NCB CHAIN SEND NO ACK
Const NCBLANSTALERT = &H73  '  NCB LAN STATUS ALERT
Const NCBACTION = &H77  '  NCB ACTION
Const NCBFINDNAME = &H78  '  NCB FIND NAME
Const NCBTRACE = &H79  '  NCB TRACE
Const ASYNCH = &H80  '  high bit set == asynchronous
Const NRC_GOODRET = &H0   '  good return
Const NRC_BUFLEN = &H1   '  illegal buffer length
Const NRC_ILLCMD = &H3   '  illegal command
Const NRC_CMDTMO = &H5   '  command timed out
Const NRC_INCOMP = &H6   '  message incomplete, issue another command
Const NRC_BADDR = &H7   '  illegal buffer address
Const NRC_SNUMOUT = &H8   '  session number out of range
Const NRC_NORES = &H9   '  no resource available
Const NRC_SCLOSED = &HA   '  session closed
Const NRC_CMDCAN = &HB   '  command cancelled
Const NRC_DUPNAME = &HD   '  duplicate name
Const NRC_NAMTFUL = &HE   '  name table full
Const NRC_ACTSES = &HF   '  no deletions, name has active sessions
Const NRC_LOCTFUL = &H11  '  local session table full
Const NRC_REMTFUL = &H12  '  remote session table full
Const NRC_ILLNN = &H13  '  illegal name number
Const NRC_NOCALL = &H14  '  no callname
Const NRC_NOWILD = &H15  '  cannot put  in NCB_NAME
Const NRC_INUSE = &H16  '  name in use on remote adapter
Const NRC_NAMERR = &H17  '  name deleted
Const NRC_SABORT = &H18  '  session ended abnormally
Const NRC_NAMCONF = &H19  '  name conflict detected
Const NRC_IFBUSY = &H21  '  interface busy, IRET before retrying
Const NRC_TOOMANY = &H22  '  too many commands outstanding, retry later
Const NRC_BRIDGE = &H23  '  ncb_lana_num field invalid
Const NRC_CANOCCR = &H24  '  command completed while cancel occurring
Const NRC_CANCEL = &H26  '  command not valid to cancel
Const NRC_DUPENV = &H30  '  name defined by anther local process
Const NRC_ENVNOTDEF = &H34  '  environment undefined. RESET required
Const NRC_OSRESNOTAV = &H35  '  required OS resources exhausted
Const NRC_MAXAPPS = &H36  '  max number of applications exceeded
Const NRC_NOSAPS = &H37  '  no saps available for netbios
Const NRC_NORESOURCES = &H38  '  requested resources are not available
Const NRC_INVADDRESS = &H39  '  invalid ncb address or length > segment
Const NRC_INVDDID = &H3B  '  invalid NCB DDID
Const NRC_LOCKFAIL = &H3C  '  lock of user area failed
Const NRC_OPENERR = &H3F  '  NETBIOS not loaded
Const NRC_SYSTEM = &H40  '  system error
Const NRC_PENDING = &HFF  '  asynchronous command is not yet finished
Const FILTER_TEMP_DUPLICATE_ACCOUNT As Long = &H1&
Const FILTER_NORMAL_ACCOUNT As Long = &H2&
Const FILTER_PROXY_ACCOUNT As Long = &H4&
Const FILTER_INTERDOMAIN_TRUST_ACCOUNT As Long = &H8&
Const FILTER_WORKSTATION_TRUST_ACCOUNT As Long = &H10&
Const FILTER_SERVER_TRUST_ACCOUNT As Long = &H20&
Const TIMEQ_FOREVER = -1&             '((unsigned long) -1L)
Const USER_MAXSTORAGE_UNLIMITED = -1&    '((unsigned long) -1L)
Const USER_NO_LOGOFF = -1&            '((unsigned long) -1L)
Const UNITS_PER_DAY = 24
Const UNITS_PER_WEEK = UNITS_PER_DAY * 7
Const USER_PRIV_MASK = 3
Const USER_PRIV_GUEST = 0
Const USER_PRIV_USER = 1
Const USER_PRIV_ADMIN = 2
Const UNLEN = 256         ' Maximum username length
Const GNLEN = UNLEN       ' Maximum groupname length
Const CNLEN = 15          ' Maximum computer name length
Const PWLEN = 256         ' Maximum password length
Const LM20_PWLEN = 14     ' LM 2.0 Maximum password length
Const MAXCOMMENTSZ = 256  ' Multipurpose comment length
Const LG_INCLUDE_INDIRECT As Long = &H1&
Const UF_SCRIPT = &H1
Const UF_ACCOUNTDISABLE = &H2
Const UF_HOMEDIR_REQUIRED = &H8
Const UF_LOCKOUT = &H10
Const UF_PASSWD_NOTREQD = &H20
Const UF_PASSWD_CANT_CHANGE = &H40
Const NERR_Success As Long = 0&
Const NERR_BASE = 2100
Const NERR_InvalidComputer = (NERR_BASE + 251)
Const NERR_NotPrimary = (NERR_BASE + 126)
Const NERR_GroupExists = (NERR_BASE + 123)
Const NERR_UserExists = (NERR_BASE + 124)
Const NERR_PasswordTooShort = (NERR_BASE + 145)
Const RESOURCE_GLOBALNET As Long = &H2&
Const RESOURCE_ENUM_ALL As Long = &HFFFF
Const RESOURCEUSAGE_ALL As Long = &H0&
Const EXCEPTION_EXECUTE_HANDLER = 1
Const EXCEPTION_CONTINUE_SEARCH = 0
Const EXCEPTION_CONTINUE_EXECUTION = -1
Const ctlFirst = &H400
Const ctlLast = &H4FF
Const psh1 = &H400
Const psh2 = &H401
Const psh3 = &H402
Const psh4 = &H403
Const psh5 = &H404
Const psh6 = &H405
Const psh7 = &H406
Const psh8 = &H407
Const psh9 = &H408
Const psh10 = &H409
Const psh11 = &H40A
Const psh12 = &H40B
Const psh13 = &H40C
Const psh14 = &H40D
Const psh15 = &H40E
Const pshHelp = psh15
Const psh16 = &H40F
Const chx1 = &H410
Const chx2 = &H411
Const chx3 = &H412
Const chx4 = &H413
Const chx5 = &H414
Const chx6 = &H415
Const chx7 = &H416
Const chx8 = &H417
Const chx9 = &H418
Const chx10 = &H419
Const chx11 = &H41A
Const chx12 = &H41B
Const chx13 = &H41C
Const chx14 = &H41D
Const chx15 = &H41E
Const chx16 = &H41D
Const rad1 = &H420
Const rad2 = &H421
Const rad3 = &H422
Const rad4 = &H423
Const rad5 = &H424
Const rad6 = &H425
Const rad7 = &H426
Const rad8 = &H427
Const rad9 = &H428
Const rad10 = &H429
Const rad11 = &H42A
Const rad12 = &H42B
Const rad13 = &H42C
Const rad14 = &H42D
Const rad15 = &H42E
Const rad16 = &H42F
Const grp1 = &H430
Const grp2 = &H431
Const grp3 = &H432
Const grp4 = &H433
Const frm1 = &H434
Const frm2 = &H435
Const frm3 = &H436
Const frm4 = &H437
Const rct1 = &H438
Const rct2 = &H439
Const rct3 = &H43A
Const rct4 = &H43B
Const ico1 = &H43C
Const ico2 = &H43D
Const ico3 = &H43E
Const ico4 = &H43F
Const stc1 = &H440
Const stc2 = &H441
Const stc3 = &H442
Const stc4 = &H443
Const stc5 = &H444
Const stc6 = &H445
Const stc7 = &H446
Const stc8 = &H447
Const stc9 = &H448
Const stc10 = &H449
Const stc11 = &H44A
Const stc12 = &H44B
Const stc13 = &H44C
Const stc14 = &H44D
Const stc15 = &H44E
Const stc16 = &H44F
Const stc17 = &H450
Const stc18 = &H451
Const stc19 = &H452
Const stc20 = &H453
Const stc21 = &H454
Const stc22 = &H455
Const stc23 = &H456
Const stc24 = &H457
Const stc25 = &H458
Const stc26 = &H459
Const stc27 = &H45A
Const stc28 = &H45B
Const stc29 = &H45C
Const stc30 = &H45D
Const stc31 = &H45E
Const stc32 = &H45F
Const lst1 = &H460
Const lst2 = &H461
Const lst3 = &H462
Const lst4 = &H463
Const lst5 = &H464
Const lst6 = &H465
Const lst7 = &H466
Const lst8 = &H467
Const lst9 = &H468
Const lst10 = &H469
Const lst11 = &H46A
Const lst12 = &H46B
Const lst13 = &H46C
Const lst14 = &H46D
Const lst15 = &H46E
Const lst16 = &H46F
Const cmb1 = &H470
Const cmb2 = &H471
Const cmb3 = &H472
Const cmb4 = &H473
Const cmb5 = &H474
Const cmb6 = &H475
Const cmb7 = &H476
Const cmb8 = &H477
Const cmb9 = &H478
Const cmb10 = &H479
Const cmb11 = &H47A
Const cmb12 = &H47B
Const cmb13 = &H47C
Const cmb14 = &H47D
Const cmb15 = &H47E
Const cmb16 = &H47F
Const edt1 = &H480
Const edt2 = &H481
Const edt3 = &H482
Const edt4 = &H483
Const edt5 = &H484
Const edt6 = &H485
Const edt7 = &H486
Const edt8 = &H487
Const edt9 = &H488
Const edt10 = &H489
Const edt11 = &H48A
Const edt12 = &H48B
Const edt13 = &H48C
Const edt14 = &H48D
Const edt15 = &H48E
Const edt16 = &H48F
Const scr1 = &H490
Const scr2 = &H491
Const scr3 = &H492
Const scr4 = &H493
Const scr5 = &H494
Const scr6 = &H495
Const scr7 = &H496
Const scr8 = &H497
Const FILEOPENORD = 1536
Const MULTIFILEOPENORD = 1537
Const PRINTDLGORD = 1538
Const PRNSETUPDLGORD = 1539
Const FINDDLGORD = 1540
Const REPLACEDLGORD = 1541
Const FONTDLGORD = 1542
Const FORMATDLGORD31 = 1543
Const FORMATDLGORD30 = 1544
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
Const SERVICES_ACTIVE_DATABASE = "ServicesActive"
Const SERVICES_FAILED_DATABASE = "ServicesFailed"
Const SERVICE_NO_CHANGE = &HFFFF
Const SERVICE_ACTIVE = &H1
Const SERVICE_INACTIVE = &H2
Const SERVICE_STATE_ALL = (SERVICE_ACTIVE Or SERVICE_INACTIVE)
Const SERVICE_CONTROL_STOP = &H1
Const SERVICE_CONTROL_PAUSE = &H2
Const SERVICE_CONTROL_CONTINUE = &H3
Const SERVICE_CONTROL_INTERROGATE = &H4
Const SERVICE_CONTROL_SHUTDOWN = &H5
Const SERVICE_STOPPED = &H1
Const SERVICE_START_PENDING = &H2
Const SERVICE_STOP_PENDING = &H3
Const SERVICE_RUNNING = &H4
Const SERVICE_CONTINUE_PENDING = &H5
Const SERVICE_PAUSE_PENDING = &H6
Const SERVICE_PAUSED = &H7
Const SERVICE_ACCEPT_STOP = &H1
Const SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
Const SERVICE_ACCEPT_SHUTDOWN = &H4
Const SC_MANAGER_CONNECT = &H1
Const SC_MANAGER_CREATE_SERVICE = &H2
Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Const SC_MANAGER_LOCK = &H8
Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)
Const SERVICE_QUERY_CONFIG = &H1
Const SERVICE_CHANGE_CONFIG = &H2
Const SERVICE_QUERY_STATUS = &H4
Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Const SERVICE_START = &H10
Const SERVICE_STOP = &H20
Const SERVICE_PAUSE_CONTINUE = &H40
Const SERVICE_INTERROGATE = &H80
Const SERVICE_USER_DEFINED_CONTROL = &H100
Const SC_GROUP_IDENTIFIER = "+"
Const PERF_DATA_VERSION = 1
Const PERF_DATA_REVISION = 1
Const PERF_NO_INSTANCES = -1  '  no instances
Const PERF_SIZE_DWORD = &H0
Const PERF_SIZE_LARGE = &H100
Const PERF_SIZE_ZERO = &H200       '  for Zero Length fields
Const PERF_SIZE_VARIABLE_LEN = &H300       '  length is in CounterLength field of Counter Definition struct
Const PERF_TYPE_NUMBER = &H0         '  a number (not a counter)
Const PERF_TYPE_COUNTER = &H400       '  an increasing numeric value
Const PERF_TYPE_TEXT = &H800       '  a text field
Const PERF_TYPE_ZERO = &HC00       '  displays a zero
Const PERF_NUMBER_HEX = &H0         '  display as HEX value
Const PERF_NUMBER_DECIMAL = &H10000     '  display as a decimal integer
Const PERF_NUMBER_DEC_1000 = &H20000     '  display as a decimal/1000
Const PERF_COUNTER_VALUE = &H0         '  display counter value
Const PERF_COUNTER_RATE = &H10000     '  divide ctr / delta time
Const PERF_COUNTER_FRACTION = &H20000     '  divide ctr / base
Const PERF_COUNTER_BASE = &H30000     '  base value used in fractions
Const PERF_COUNTER_ELAPSED = &H40000     '  subtract counter from current time
Const PERF_COUNTER_QUEUELEN = &H50000     '  Use Queuelen processing func.
Const PERF_COUNTER_HISTOGRAM = &H60000     '  Counter begins or ends a histogram
Const PERF_TEXT_UNICODE = &H0         '  type of text in text field
Const PERF_TEXT_ASCII = &H10000     '  ASCII using the CodePage field
Const PERF_TIMER_TICK = &H0         '  use system perf. freq for base
Const PERF_TIMER_100NS = &H100000    '  use 100 NS timer time base units
Const PERF_OBJECT_TIMER = &H200000    '  use the object timer freq
Const PERF_DELTA_COUNTER = &H400000    '  compute difference first
Const PERF_DELTA_BASE = &H800000    '  compute base diff as well
Const PERF_INVERSE_COUNTER = &H1000000   '  show as 1.00-value (assumes:
Const PERF_MULTI_COUNTER = &H2000000   '  sum of multiple instances
Const PERF_DISPLAY_NO_SUFFIX = &H0         '  no suffix
Const PERF_DISPLAY_PER_SEC = &H10000000  '  "/sec"
Const PERF_DISPLAY_PERCENT = &H20000000  '  "%"
Const PERF_DISPLAY_SECONDS = &H30000000  '  "secs"
Const PERF_DISPLAY_NOSHOW = &H40000000  '  value is not displayed
Const PERF_COUNTER_COUNTER = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_PER_SEC)
Const PERF_COUNTER_TIMER = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_PERCENT)
Const PERF_COUNTER_QUEUELEN_TYPE = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_QUEUELEN Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_NO_SUFFIX)
Const PERF_COUNTER_BULK_COUNT = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_PER_SEC)
Const PERF_COUNTER_TEXT = (PERF_SIZE_VARIABLE_LEN Or PERF_TYPE_TEXT Or PERF_TEXT_UNICODE Or PERF_DISPLAY_NO_SUFFIX)
Const PERF_COUNTER_RAWCOUNT = (PERF_SIZE_DWORD Or PERF_TYPE_NUMBER Or PERF_NUMBER_DECIMAL Or PERF_DISPLAY_NO_SUFFIX)
Const PERF_SAMPLE_FRACTION = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_FRACTION Or PERF_DELTA_COUNTER Or PERF_DELTA_BASE Or PERF_DISPLAY_PERCENT)
Const PERF_SAMPLE_COUNTER = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_DISPLAY_NO_SUFFIX)
Const PERF_COUNTER_NODATA = (PERF_SIZE_ZERO Or PERF_DISPLAY_NOSHOW)
Const PERF_COUNTER_TIMER_INV = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_TICK Or PERF_DELTA_COUNTER Or PERF_INVERSE_COUNTER Or PERF_DISPLAY_PERCENT)
Const PERF_SAMPLE_BASE = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_BASE Or PERF_DISPLAY_NOSHOW Or &H1)         '  for compatibility with pre-beta versions
Const PERF_AVERAGE_TIMER = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_FRACTION Or PERF_DISPLAY_SECONDS)
Const PERF_AVERAGE_BASE = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_BASE Or PERF_DISPLAY_NOSHOW Or &H2)         '  for compatibility with pre-beta versions
Const PERF_AVERAGE_BULK = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_FRACTION Or PERF_DISPLAY_NOSHOW)
Const PERF_100NSEC_TIMER = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_100NS Or PERF_DELTA_COUNTER Or PERF_DISPLAY_PERCENT)
Const PERF_100NSEC_TIMER_INV = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_100NS Or PERF_DELTA_COUNTER Or PERF_INVERSE_COUNTER Or PERF_DISPLAY_PERCENT)
Const PERF_COUNTER_MULTI_TIMER = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_DELTA_COUNTER Or PERF_TIMER_TICK Or PERF_MULTI_COUNTER Or PERF_DISPLAY_PERCENT)
Const PERF_COUNTER_MULTI_TIMER_INV = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_RATE Or PERF_DELTA_COUNTER Or PERF_MULTI_COUNTER Or PERF_TIMER_TICK Or PERF_INVERSE_COUNTER Or PERF_DISPLAY_PERCENT)
Const PERF_COUNTER_MULTI_BASE = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_BASE Or PERF_MULTI_COUNTER Or PERF_DISPLAY_NOSHOW)
Const PERF_100NSEC_MULTI_TIMER = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_DELTA_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_100NS Or PERF_MULTI_COUNTER Or PERF_DISPLAY_PERCENT)
Const PERF_100NSEC_MULTI_TIMER_INV = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_DELTA_COUNTER Or PERF_COUNTER_RATE Or PERF_TIMER_100NS Or PERF_MULTI_COUNTER Or PERF_INVERSE_COUNTER Or PERF_DISPLAY_PERCENT)
Const PERF_RAW_FRACTION = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_FRACTION Or PERF_DISPLAY_PERCENT)
Const PERF_RAW_BASE = (PERF_SIZE_DWORD Or PERF_TYPE_COUNTER Or PERF_COUNTER_BASE Or PERF_DISPLAY_NOSHOW Or &H3)         '  for compatibility with pre-beta versions
Const PERF_ELAPSED_TIME = (PERF_SIZE_LARGE Or PERF_TYPE_COUNTER Or PERF_COUNTER_ELAPSED Or PERF_OBJECT_TIMER Or PERF_DISPLAY_SECONDS)
Const PERF_COUNTER_HISTOGRAM_TYPE = &H80000000  ' Counter begins or ends a histogram
Const PERF_DETAIL_NOVICE = 100    '  The uninformed can understand it
Const PERF_DETAIL_ADVANCED = 200    '  For the advanced user
Const PERF_DETAIL_EXPERT = 300    '  For the expert user
Const PERF_DETAIL_WIZARD = 400    '  For the system designer
Const PERF_NO_UNIQUE_ID = -1
Const CDERR_DIALOGFAILURE = &HFFFF
Const CDERR_GENERALCODES = &H0
Const CDERR_STRUCTSIZE = &H1
Const CDERR_INITIALIZATION = &H2
Const CDERR_NOTEMPLATE = &H3
Const CDERR_NOHINSTANCE = &H4
Const CDERR_LOADSTRFAILURE = &H5
Const CDERR_FINDRESFAILURE = &H6
Const CDERR_LOADRESFAILURE = &H7
Const CDERR_LOCKRESFAILURE = &H8
Const CDERR_MEMALLOCFAILURE = &H9
Const CDERR_MEMLOCKFAILURE = &HA
Const CDERR_NOHOOK = &HB
Const CDERR_REGISTERMSGFAIL = &HC
Const PDERR_PRINTERCODES = &H1000
Const PDERR_SETUPFAILURE = &H1001
Const PDERR_PARSEFAILURE = &H1002
Const PDERR_RETDEFFAILURE = &H1003
Const PDERR_LOADDRVFAILURE = &H1004
Const PDERR_GETDEVMODEFAIL = &H1005
Const PDERR_INITFAILURE = &H1006
Const PDERR_NODEVICES = &H1007
Const PDERR_NODEFAULTPRN = &H1008
Const PDERR_DNDMMISMATCH = &H1009
Const PDERR_CREATEICFAILURE = &H100A
Const PDERR_PRINTERNOTFOUND = &H100B
Const PDERR_DEFAULTDIFFERENT = &H100C
Const CFERR_CHOOSEFONTCODES = &H2000
Const CFERR_NOFONTS = &H2001
Const CFERR_MAXLESSTHANMIN = &H2002
Const FNERR_FILENAMECODES = &H3000
Const FNERR_SUBCLASSFAILURE = &H3001
Const FNERR_INVALIDFILENAME = &H3002
Const FNERR_BUFFERTOOSMALL = &H3003
Const FRERR_FINDREPLACECODES = &H4000
Const FRERR_BUFFERLENGTHZERO = &H4001
Const CCERR_CHOOSECOLORCODES = &H5000
Const LZERROR_BADINHANDLE = (-1)  '  invalid input handle
Const LZERROR_BADOUTHANDLE = (-2)    '  invalid output handle
Const LZERROR_READ = (-3)         '  corrupt compressed file format
Const LZERROR_WRITE = (-4)        '  out of space for output file
Const LZERROR_PUBLICLOC = (-5)    '  insufficient memory for LZFile struct
Const LZERROR_GLOBLOCK = (-6)     '  bad Global handle
Const LZERROR_BADVALUE = (-7)     '  input parameter out of range
Const LZERROR_UNKNOWNALG = (-8)   '  compression algorithm not recognized
Const VK_PROCESSKEY = &HE5
Const STYLE_DESCRIPTION_SIZE = 32
Const WM_CONVERTREQUESTEX = &H108
Const WM_IME_STARTCOMPOSITION = &H10D
Const WM_IME_ENDCOMPOSITION = &H10E
Const WM_IME_COMPOSITION = &H10F
Const WM_IME_KEYLAST = &H10F
Const WM_IME_SETCONTEXT = &H281
Const WM_IME_NOTIFY = &H282
Const WM_IME_CONTROL = &H283
Const WM_IME_COMPOSITIONFULL = &H284
Const WM_IME_SELECT = &H285
Const WM_IME_CHAR = &H286
Const WM_IME_KEYDOWN = &H290
Const WM_IME_KEYUP = &H291
Const IMC_GETCANDIDATEPOS = &H7
Const IMC_SETCANDIDATEPOS = &H8
Const IMC_GETCOMPOSITIONFONT = &H9
Const IMC_SETCOMPOSITIONFONT = &HA
Const IMC_GETCOMPOSITIONWINDOW = &HB
Const IMC_SETCOMPOSITIONWINDOW = &HC
Const IMC_GETSTATUSWINDOWPOS = &HF
Const IMC_SETSTATUSWINDOWPOS = &H10
Const IMC_CLOSESTATUSWINDOW = &H21
Const IMC_OPENSTATUSWINDOW = &H22
Const NI_OPENCANDIDATE = &H10
Const NI_CLOSECANDIDATE = &H11
Const NI_SELECTCANDIDATESTR = &H12
Const NI_CHANGECANDIDATELIST = &H13
Const NI_FINALIZECONVERSIONRESULT = &H14
Const NI_COMPOSITIONSTR = &H15
Const NI_SETCANDIDATE_PAGESTART = &H16
Const NI_SETCANDIDATE_PAGESIZE = &H17
Const ISC_SHOWUICANDIDATEWINDOW = &H1
Const ISC_SHOWUICOMPOSITIONWINDOW = &H80000000
Const ISC_SHOWUIGUIDELINE = &H40000000
Const ISC_SHOWUIALLCANDIDATEWINDOW = &HF
Const ISC_SHOWUIALL = &HC000000F
Const CPS_COMPLETE = &H1
Const CPS_CONVERT = &H2
Const CPS_REVERT = &H3
Const CPS_CANCEL = &H4
Const IME_CHOTKEY_IME_NONIME_TOGGLE = &H10
Const IME_CHOTKEY_SHAPE_TOGGLE = &H11
Const IME_CHOTKEY_SYMBOL_TOGGLE = &H12
Const IME_JHOTKEY_CLOSE_OPEN = &H30
Const IME_KHOTKEY_SHAPE_TOGGLE = &H50
Const IME_KHOTKEY_HANJACONVERT = &H51
Const IME_KHOTKEY_ENGLISH = &H52
Const IME_THOTKEY_IME_NONIME_TOGGLE = &H70
Const IME_THOTKEY_SHAPE_TOGGLE = &H71
Const IME_THOTKEY_SYMBOL_TOGGLE = &H72
Const IME_HOTKEY_DSWITCH_FIRST = &H100
Const IME_HOTKEY_DSWITCH_LAST = &H11F
Const IME_ITHOTKEY_RESEND_RESULTSTR = &H200
Const IME_ITHOTKEY_PREVIOUS_COMPOSITION = &H201
Const IME_ITHOTKEY_UISTYLE_TOGGLE = &H202
Const GCS_COMPREADSTR = &H1
Const GCS_COMPREADATTR = &H2
Const GCS_COMPREADCLAUSE = &H4
Const GCS_COMPSTR = &H8
Const GCS_COMPATTR = &H10
Const GCS_COMPCLAUSE = &H20
Const GCS_CURSORPOS = &H80
Const GCS_DELTASTART = &H100
Const GCS_RESULTREADSTR = &H200
Const GCS_RESULTREADCLAUSE = &H400
Const GCS_RESULTSTR = &H800
Const GCS_RESULTCLAUSE = &H1000
Const CS_INSERTCHAR = &H2000
Const CS_NOMOVECARET = &H4000
Const IME_PROP_AT_CARET = &H10000
Const IME_PROP_SPECIAL_UI = &H20000
Const IME_PROP_CANDLIST_START_FROM_1 = &H40000
Const IME_PROP_UNICODE = &H80000
Const UI_CAP_2700 = &H1
Const UI_CAP_ROT90 = &H2
Const UI_CAP_ROTANY = &H4
Const SCS_CAP_COMPSTR = &H1
Const SCS_CAP_MAKEREAD = &H2
Const SELECT_CAP_CONVERSION = &H1
Const SELECT_CAP_SENTENCE = &H2
Const GGL_LEVEL = &H1
Const GGL_INDEX = &H2
Const GGL_STRING = &H3
Const GGL_PRIVATE = &H4
Const GL_LEVEL_NOGUIDELINE = &H0
Const GL_LEVEL_FATAL = &H1
Const GL_LEVEL_ERROR = &H2
Const GL_LEVEL_WARNING = &H3
Const GL_LEVEL_INFORMATION = &H4
Const GL_ID_UNKNOWN = &H0
Const GL_ID_NOMODULE = &H1
Const GL_ID_NODICTIONARY = &H10
Const GL_ID_CANNOTSAVE = &H11
Const GL_ID_NOCONVERT = &H20
Const GL_ID_TYPINGERROR = &H21
Const GL_ID_TOOMANYSTROKE = &H22
Const GL_ID_READINGCONFLICT = &H23
Const GL_ID_INPUTREADING = &H24
Const GL_ID_INPUTRADICAL = &H25
Const GL_ID_INPUTCODE = &H26
Const GL_ID_INPUTSYMBOL = &H27
Const GL_ID_CHOOSECANDIDATE = &H28
Const GL_ID_REVERSECONVERSION = &H29
Const GL_ID_PRIVATE_FIRST = &H8000
Const GL_ID_PRIVATE_LAST = &HFFFF
Const IGP_PROPERTY = &H4
Const IGP_CONVERSION = &H8
Const IGP_SENTENCE = &HC
Const IGP_UI = &H10
Const IGP_SETCOMPSTR = &H14
Const IGP_SELECT = &H18
Const SCS_SETSTR = (GCS_COMPREADSTR Or GCS_COMPSTR)
Const SCS_CHANGEATTR = (GCS_COMPREADATTR Or GCS_COMPATTR)
Const SCS_CHANGECLAUSE = (GCS_COMPREADCLAUSE Or GCS_COMPCLAUSE)
Const ATTR_INPUT = &H0
Const ATTR_TARGET_CONVERTED = &H1
Const ATTR_CONVERTED = &H2
Const ATTR_TARGET_NOTCONVERTED = &H3
Const ATTR_INPUT_ERROR = &H4
Const CFS_DEFAULT = &H0
Const CFS_RECT = &H1
Const CFS_POINT = &H2
Const CFS_SCREEN = &H4
Const CFS_FORCE_POSITION = &H20
Const CFS_CANDIDATEPOS = &H40
Const CFS_EXCLUDE = &H80
Const GCL_CONVERSION = &H1
Const GCL_REVERSECONVERSION = &H2
Const GCL_REVERSE_LENGTH = &H3
Const IME_CMODE_ALPHANUMERIC = &H0
Const IME_CMODE_NATIVE = &H1
Const IME_CMODE_CHINESE = IME_CMODE_NATIVE
Const IME_CMODE_HANGEUL = IME_CMODE_NATIVE
Const IME_CMODE_JAPANESE = IME_CMODE_NATIVE
Const IME_CMODE_KATAKANA = &H2                   '  only effect under IME_CMODE_NATIVE
Const IME_CMODE_LANGUAGE = &H3
Const IME_CMODE_FULLSHAPE = &H8
Const IME_CMODE_ROMAN = &H10
Const IME_CMODE_CHARCODE = &H20
Const IME_CMODE_HANJACONVERT = &H40
Const IME_CMODE_SOFTKBD = &H80
Const IME_CMODE_NOCONVERSION = &H100
Const IME_CMODE_EUDC = &H200
Const IME_CMODE_SYMBOL = &H400
Const IME_SMODE_NONE = &H0
Const IME_SMODE_PLAURALCLAUSE = &H1
Const IME_SMODE_SINGLECONVERT = &H2
Const IME_SMODE_AUTOMATIC = &H4
Const IME_SMODE_PHRASEPREDICT = &H8
Const IME_CAND_UNKNOWN = &H0
Const IME_CAND_READ = &H1
Const IME_CAND_CODE = &H2
Const IME_CAND_MEANING = &H3
Const IME_CAND_RADICAL = &H4
Const IME_CAND_STROKE = &H5
Const IMN_CLOSESTATUSWINDOW = &H1
Const IMN_OPENSTATUSWINDOW = &H2
Const IMN_CHANGECANDIDATE = &H3
Const IMN_CLOSECANDIDATE = &H4
Const IMN_OPENCANDIDATE = &H5
Const IMN_SETCONVERSIONMODE = &H6
Const IMN_SETSENTENCEMODE = &H7
Const IMN_SETOPENSTATUS = &H8
Const IMN_SETCANDIDATEPOS = &H9
Const IMN_SETCOMPOSITIONFONT = &HA
Const IMN_SETCOMPOSITIONWINDOW = &HB
Const IMN_SETSTATUSWINDOWPOS = &HC
Const IMN_GUIDELINE = &HD
Const IMN_PRIVATE = &HE
Const IMM_ERROR_NODATA = (-1)
Const IMM_ERROR_GENERAL = (-2)
Const IME_CONFIG_GENERAL = 1
Const IME_CONFIG_REGISTERWORD = 2
Const IME_CONFIG_SELECTDICTIONARY = 3
Const IME_ESC_QUERY_SUPPORT = &H3
Const IME_ESC_RESERVED_FIRST = &H4
Const IME_ESC_RESERVED_LAST = &H7FF
Const IME_ESC_PRIVATE_FIRST = &H800
Const IME_ESC_PRIVATE_LAST = &HFFF
Const IME_ESC_SEQUENCE_TO_INTERNAL = &H1001
Const IME_ESC_GET_EUDC_DICTIONARY = &H1003
Const IME_ESC_SET_EUDC_DICTIONARY = &H1004
Const IME_ESC_MAX_KEY = &H1005
Const IME_ESC_IME_NAME = &H1006
Const IME_ESC_SYNC_HOTKEY = &H1007
Const IME_ESC_HANJA_MODE = &H1008
Const IME_REGWORD_STYLE_EUDC = &H1
Const IME_REGWORD_STYLE_USER_FIRST = &H80000000
Const IME_REGWORD_STYLE_USER_LAST = &HFFFF
Const SOFTKEYBOARD_TYPE_T1 = &H1
Const SOFTKEYBOARD_TYPE_C1 = &H2
Const DIALOPTION_BILLING = &H40          '  Supports wait for bong "$"
Const DIALOPTION_QUIET = &H80            '  Supports wait for quiet "@"
Const DIALOPTION_DIALTONE = &H100        '  Supports wait for dial tone "W"
Const MDMVOLFLAG_LOW = &H1
Const MDMVOLFLAG_MEDIUM = &H2
Const MDMVOLFLAG_HIGH = &H4
Const MDMVOL_LOW = &H0
Const MDMVOL_MEDIUM = &H1
Const MDMVOL_HIGH = &H2
Const MDMSPKRFLAG_OFF = &H1
Const MDMSPKRFLAG_DIAL = &H2
Const MDMSPKRFLAG_ON = &H4
Const MDMSPKRFLAG_CALLSETUP = &H8
Const MDMSPKR_OFF = &H0
Const MDMSPKR_DIAL = &H1
Const MDMSPKR_ON = &H2
Const MDMSPKR_CALLSETUP = &H3
Const MDM_COMPRESSION = &H1
Const MDM_ERROR_CONTROL = &H2
Const MDM_FORCED_EC = &H4
Const MDM_CELLULAR = &H8
Const MDM_FLOWCONTROL_HARD = &H10
Const MDM_FLOWCONTROL_SOFT = &H20
Const MDM_CCITT_OVERRIDE = &H40
Const MDM_SPEED_ADJUST = &H80
Const MDM_TONE_DIAL = &H100
Const MDM_BLIND_DIAL = &H200
Const MDM_V23_OVERRIDE = &H400
Const ABM_NEW = &H0
Const ABM_REMOVE = &H1
Const ABM_QUERYPOS = &H2
Const ABM_SETPOS = &H3
Const ABM_GETSTATE = &H4
Const ABM_GETTASKBARPOS = &H5
Const ABM_ACTIVATE = &H6               '  lParam == TRUE/FALSE means activate/deactivate
Const ABM_GETAUTOHIDEBAR = &H7
Const ABM_SETAUTOHIDEBAR = &H8          '  this can fail at any time.  MUST check the result
Const ABM_WINDOWPOSCHANGED = &H9
Const ABN_STATECHANGE = &H0
Const ABN_POSCHANGED = &H1
Const ABN_FULLSCREENAPP = &H2
Const ABN_WINDOWARRANGE = &H3    '  lParam == TRUE means hide
Const ABS_AUTOHIDE = &H1
Const ABS_ALWAYSONTOP = &H2
Const ABE_LEFT = 0
Const ABE_TOP = 1
Const ABE_RIGHT = 2
Const ABE_BOTTOM = 3
Const EIRESID = -1
Const FO_MOVE = &H1
Const FO_COPY = &H2
Const FO_DELETE = &H3
Const FO_RENAME = &H4
Const FOF_MULTIDESTFILES = &H1
Const FOF_CONFIRMMOUSE = &H2
Const FOF_SILENT = &H4                      '  don't create progress/report
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Const FOF_WANTMAPPINGHANDLE = &H20          '  Fill in SHFILEOPSTRUCT.hNameMappings
Const FOF_ALLOWUNDO = &H40
Const FOF_FILESONLY = &H80                  '  on *.*, do only files
Const FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files
Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs
Const PO_DELETE = &H13           '  printer is being deleted
Const PO_RENAME = &H14           '  printer is being renamed
Const PO_PORTCHANGE = &H20       '  port this printer connected to is being changed
Const PO_REN_PORT = &H34         '  PO_RENAME and PO_PORTCHANGE at same time.
Const SE_ERR_FNF = 2                     '  file not found
Const SE_ERR_PNF = 3                     '  path not found
Const SE_ERR_ACCESSDENIED = 5            '  access denied
Const SE_ERR_OOM = 8                     '  out of memory
Const SE_ERR_DLLNOTFOUND = 32
Const SEE_MASK_CLASSNAME = &H1
Const SEE_MASK_CLASSKEY = &H3
Const SEE_MASK_IDLIST = &H4
Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_ICON = &H10
Const SEE_MASK_HOTKEY = &H20
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_CONNECTNETDRV = &H80
Const SEE_MASK_FLAG_DDEWAIT = &H100
Const SEE_MASK_DOENVSUBST = &H200
Const SEE_MASK_FLAG_NO_UI = &H400
Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4
Const SHGFI_ICON = &H100                         '  get icon
Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Const SHGFI_TYPENAME = &H400                     '  get type name
Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Const SHGFI_EXETYPE = &H2000                     '  return exe type
Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Const SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Const SHGFI_LARGEICON = &H0                      '  get large icon
Const SHGFI_SMALLICON = &H1                      '  get small icon
Const SHGFI_OPENICON = &H2                       '  get open icon
Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute
Const SHGNLI_PIDL = &H1                          '  pszLinkTo is a pidl
Const SHGNLI_PREFIXNAME = &H2                    '  Make name "Shortcut to xxx"
Const VS_VERSION_INFO = 1
Const VS_USER_DEFINED = 100
Const VS_FFI_SIGNATURE = &HFEEF04BD
Const VS_FFI_STRUCVERSION = &H10000
Const VS_FFI_FILEFLAGSMASK = &H3F&
Const VS_FF_DEBUG = &H1&
Const VS_FF_PRERELEASE = &H2&
Const VS_FF_PATCHED = &H4&
Const VS_FF_PRIVATEBUILD = &H8&
Const VS_FF_INFOINFERRED = &H10&
Const VS_FF_SPECIALBUILD = &H20&
Const VOS_UNKNOWN = &H0&
Const VOS_DOS = &H10000
Const VOS_OS216 = &H20000
Const VOS_OS232 = &H30000
Const VOS_NT = &H40000
Const VOS__BASE = &H0&
Const VOS__WINDOWS16 = &H1&
Const VOS__PM16 = &H2&
Const VOS__PM32 = &H3&
Const VOS__WINDOWS32 = &H4&
Const VOS_DOS_WINDOWS16 = &H10001
Const VOS_DOS_WINDOWS32 = &H10004
Const VOS_OS216_PM16 = &H20002
Const VOS_OS232_PM32 = &H30003
Const VOS_NT_WINDOWS32 = &H40004
Const VFT_UNKNOWN = &H0&
Const VFT_APP = &H1&
Const VFT_DLL = &H2&
Const VFT_DRV = &H3&
Const VFT_FONT = &H4&
Const VFT_VXD = &H5&
Const VFT_STATIC_LIB = &H7&
Const VFT2_UNKNOWN = &H0&
Const VFT2_DRV_PRINTER = &H1&
Const VFT2_DRV_KEYBOARD = &H2&
Const VFT2_DRV_LANGUAGE = &H3&
Const VFT2_DRV_DISPLAY = &H4&
Const VFT2_DRV_MOUSE = &H5&
Const VFT2_DRV_NETWORK = &H6&
Const VFT2_DRV_SYSTEM = &H7&
Const VFT2_DRV_INSTALLABLE = &H8&
Const VFT2_DRV_SOUND = &H9&
Const VFT2_DRV_COMM = &HA&
Const VFT2_DRV_INPUTMETHOD = &HB&
Const VFT2_FONT_RASTER = &H1&
Const VFT2_FONT_VECTOR = &H2&
Const VFT2_FONT_TRUETYPE = &H3&
Const VFFF_ISSHAREDFILE = &H1
Const VFF_CURNEDEST = &H1
Const VFF_FILEINUSE = &H2
Const VFF_BUFFTOOSMALL = &H4
Const VIFF_FORCEINSTALL = &H1
Const VIFF_DONTDELETEOLD = &H2
Const VIF_TEMPFILE = &H1&
Const VIF_MISMATCH = &H2&
Const VIF_SRCOLD = &H4&
Const VIF_DIFFLANG = &H8&
Const VIF_DIFFCODEPG = &H10&
Const VIF_DIFFTYPE = &H20&
Const VIF_WRITEPROT = &H40&
Const VIF_FILEINUSE = &H80&
Const VIF_OUTOFSPACE = &H100&
Const VIF_ACCESSVIOLATION = &H200&
Const VIF_SHARINGVIOLATION = &H400&
Const VIF_CANNOTCREATE = &H800&
Const VIF_CANNOTDELETE = &H1000&
Const VIF_CANNOTRENAME = &H2000&
Const VIF_CANNOTDELETECUR = &H4000&
Const VIF_OUTOFMEMORY = &H8000&
Const VIF_CANNOTREADSRC = &H10000
Const VIF_CANNOTREADDST = &H20000
Const VIF_BUFFTOOSMALL = &H40000
Const PROCESS_HEAP_REGION = &H1
Const PROCESS_HEAP_UNCOMMITTED_RANGE = &H2
Const PROCESS_HEAP_ENTRY_BUSY = &H4
Const PROCESS_HEAP_ENTRY_MOVEABLE = &H10
Const PROCESS_HEAP_ENTRY_DDESHARE = &H20
Const SCS_32BIT_BINARY = 0
Const SCS_DOS_BINARY = 1
Const SCS_WOW_BINARY = 2
Const SCS_PIF_BINARY = 3
Const SCS_POSIX_BINARY = 4
Const SCS_OS216_BINARY = 5
Const LOGON32_LOGON_INTERACTIVE = 2
Const LOGON32_LOGON_BATCH = 4
Const LOGON32_LOGON_SERVICE = 5
Const LOGON32_PROVIDER_DEFAULT = 0
Const LOGON32_PROVIDER_WINNT35 = 1
Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2
Const AC_LINE_OFFLINE = &H0
Const AC_LINE_ONLINE = &H1
Const AC_LINE_BACKUP_POWER = &H2
Const AC_LINE_UNKNOWN = &HFF
Const BATTERY_FLAG_HIGH = &H1
Const BATTERY_FLAG_LOW = &H2
Const BATTERY_FLAG_CRITICAL = &H4
Const BATTERY_FLAG_CHARGING = &H8
Const BATTERY_FLAG_NO_BATTERY = &H80
Const BATTERY_FLAG_UNKNOWN = &HFF
Const BATTERY_PERCENTAGE_UNKNOWN = &HFF
Const BATTERY_LIFE_UNKNOWN = &HFFFF
Const OFN_READONLY = &H1
Const OFN_OVERWRITEPROMPT = &H2
Const OFN_HIDEREADONLY = &H4
Const OFN_NOCHANGEDIR = &H8
Const OFN_SHOWHELP = &H10
Const OFN_ENABLEHOOK = &H20
Const OFN_ENABLETEMPLATE = &H40
Const OFN_ENABLETEMPLATEHANDLE = &H80
Const OFN_NOVALIDATE = &H100
Const OFN_ALLOWMULTISELECT = &H200
Const OFN_EXTENSIONDIFFERENT = &H400
Const OFN_PATHMUSTEXIST = &H800
Const OFN_FILEMUSTEXIST = &H1000
Const OFN_CREATEPROMPT = &H2000
Const OFN_SHAREAWARE = &H4000
Const OFN_NOREADONLYRETURN = &H8000
Const OFN_NOTESTFILECREATE = &H10000
Const OFN_NONETWORKBUTTON = &H20000
Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Const OFN_EXPLORER = &H80000                         '  new look commdlg
Const OFN_NODEREFERENCELINKS = &H100000
Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Const OFN_SHAREFALLTHROUGH = 2
Const OFN_SHARENOWARN = 1
Const OFN_SHAREWARN = 0
Const CDM_FIRST = (WM_USER + 100)
Const CDM_LAST = (WM_USER + 200)
Const CDM_GETSPEC = (CDM_FIRST + &H0)
Const CDM_GETFILEPATH = (CDM_FIRST + &H1)
Const CDM_GETFOLDERPATH = (CDM_FIRST + &H2)
Const CDM_GETFOLDERIDLIST = (CDM_FIRST + &H3)
Const CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)
Const CDM_HIDECONTROL = (CDM_FIRST + &H5)
Const CDM_SETDEFEXT = (CDM_FIRST + &H6)
Const CC_RGBINIT = &H1
Const CC_FULLOPEN = &H2
Const CC_PREVENTFULLOPEN = &H4
Const CC_SHOWHELP = &H8
Const CC_ENABLEHOOK = &H10
Const CC_ENABLETEMPLATE = &H20
Const CC_ENABLETEMPLATEHANDLE = &H40
Const CC_SOLIDCOLOR = &H80
Const CC_ANYCOLOR = &H100
Const FR_DOWN = &H1
Const FR_WHOLEWORD = &H2
Const FR_MATCHCASE = &H4
Const FR_FINDNEXT = &H8
Const FR_REPLACE = &H10
Const FR_REPLACEALL = &H20
Const FR_DIALOGTERM = &H40
Const FR_SHOWHELP = &H80
Const FR_ENABLEHOOK = &H100
Const FR_ENABLETEMPLATE = &H200
Const FR_NOUPDOWN = &H400
Const FR_NOMATCHCASE = &H800
Const FR_NOWHOLEWORD = &H1000
Const FR_ENABLETEMPLATEHANDLE = &H2000
Const FR_HIDEUPDOWN = &H4000
Const FR_HIDEMATCHCASE = &H8000
Const FR_HIDEWHOLEWORD = &H10000
Const CF_SCREENFONTS = &H1
Const CF_PRINTERFONTS = &H2
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_SHOWHELP = &H4&
Const CF_ENABLEHOOK = &H8&
Const CF_ENABLETEMPLATE = &H10&
Const CF_ENABLETEMPLATEHANDLE = &H20&
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_USESTYLE = &H80&
Const CF_EFFECTS = &H100&
Const CF_APPLY = &H200&
Const CF_ANSIONLY = &H400&
Const CF_SCRIPTSONLY = CF_ANSIONLY
Const CF_NOVECTORFONTS = &H800&
Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Const CF_NOSIMULATIONS = &H1000&
Const CF_LIMITSIZE = &H2000&
Const CF_FIXEDPITCHONLY = &H4000&
Const CF_WYSIWYG = &H8000    '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Const CF_FORCEFONTEXIST = &H10000
Const CF_SCALABLEONLY = &H20000
Const CF_TTONLY = &H40000
Const CF_NOFACESEL = &H80000
Const CF_NOSTYLESEL = &H100000
Const CF_NOSIZESEL = &H200000
Const CF_SELECTSCRIPT = &H400000
Const CF_NOSCRIPTSEL = &H800000
Const CF_NOVERTFONTS = &H1000000
Const SIMULATED_FONTTYPE = &H8000
Const PRINTER_FONTTYPE = &H4000
Const SCREEN_FONTTYPE = &H2000
Const BOLD_FONTTYPE = &H100
Const ITALIC_FONTTYPE = &H200
Const REGULAR_FONTTYPE = &H400
Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)
Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Const SHAREVISTRING = "commdlg_ShareViolation"
Const FILEOKSTRING = "commdlg_FileNameOK"
Const COLOROKSTRING = "commdlg_ColorOK"
Const SETRGBSTRING = "commdlg_SetRGBColor"
Const HELPMSGSTRING = "commdlg_help"
Const FINDMSGSTRING = "commdlg_FindReplace"
Const CD_LBSELNOITEMS = -1
Const CD_LBSELCHANGE = 0
Const CD_LBSELSUB = 1
Const CD_LBSELADD = 2
Const PD_ALLPAGES = &H0
Const PD_SELECTION = &H1
Const PD_PAGENUMS = &H2
Const PD_NOSELECTION = &H4
Const PD_NOPAGENUMS = &H8
Const PD_COLLATE = &H10
Const PD_PRINTTOFILE = &H20
Const PD_PRINTSETUP = &H40
Const PD_NOWARNING = &H80
Const PD_RETURNDC = &H100
Const PD_RETURNIC = &H200
Const PD_RETURNDEFAULT = &H400
Const PD_SHOWHELP = &H800
Const PD_ENABLEPRINTHOOK = &H1000
Const PD_ENABLESETUPHOOK = &H2000
Const PD_ENABLEPRINTTEMPLATE = &H4000
Const PD_ENABLESETUPTEMPLATE = &H8000
Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Const PD_USEDEVMODECOPIES = &H40000
Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Const PD_DISABLEPRINTTOFILE = &H80000
Const PD_HIDEPRINTTOFILE = &H100000
Const PD_NONETWORKBUTTON = &H200000
Const DN_DEFAULTPRN = &H1
Const WM_PSD_PAGESETUPDLG = (WM_USER)
Const WM_PSD_FULLPAGERECT = (WM_USER + 1)
Const WM_PSD_MINMARGINRECT = (WM_USER + 2)
Const WM_PSD_MARGINRECT = (WM_USER + 3)
Const WM_PSD_GREEKTEXTRECT = (WM_USER + 4)
Const WM_PSD_ENVSTAMPRECT = (WM_USER + 5)
Const WM_PSD_YAFULLPAGERECT = (WM_USER + 6)
Const PSD_DEFAULTMINMARGINS = &H0    '  default (printer's)
Const PSD_INWININIINTLMEASURE = &H0    '  1st of 4 possible
Const PSD_MINMARGINS = &H1    '  use caller's
Const PSD_MARGINS = &H2    '  use caller's
Const PSD_INTHOUSANDTHSOFINCHES = &H4    '  2nd of 4 possible
Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8    '  3rd of 4 possible
Const PSD_DISABLEMARGINS = &H10
Const PSD_DISABLEPRINTER = &H20
Const PSD_NOWARNING = &H80    '  must be same as PD_*
Const PSD_DISABLEORIENTATION = &H100
Const PSD_RETURNDEFAULT = &H400    '  must be same as PD_*
Const PSD_DISABLEPAPER = &H200
Const PSD_SHOWHELP = &H800    '  must be same as PD_*
Const PSD_ENABLEPAGESETUPHOOK = &H2000    '  must be same as PD_*
Const PSD_ENABLEPAGESETUPTEMPLATE = &H8000    '  must be same as PD_*
Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE = &H20000    '  must be same as PD_*
Const PSD_ENABLEPAGEPAINTHOOK = &H40000
Const PSD_DISABLEPAGEPAINTING = &H80000
Const INVALID_HANDLE_VALUE = -1
Const BDR_RAISEDOUTER = &H1
Const BDR_SUNKENOUTER = &H2
Const BDR_RAISEDINNER = &H4
Const BDR_SUNKENINNER = &H8
Const BDR_OUTER = &H3
Const BDR_INNER = &HC
Const BDR_RAISED = &H5
Const BDR_SUNKEN = &HA
Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Const BF_LEFT = &H1
Const BF_TOP = &H2
Const BF_RIGHT = &H4
Const BF_BOTTOM = &H8
Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Const BF_DIAGONAL = &H10
Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Const BF_MIDDLE = &H800    ' Fill in the middle.
Const BF_SOFT = &H1000     ' Use for softer buttons.
Const BF_ADJUST = &H2000   ' Calculate the space left over.
Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Const ANYSIZE_ARRAY = 1
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
	X As Integer
	Y As Integer
End Type
Type msg
	hwnd As Long
	message As Long
	wParam As Long
	lParam As Long
	time As Long
	pt As POINTAPI
End Type
Type SID_IDENTIFIER_AUTHORITY
	Value(6) As Byte
End Type
Type SID_AND_ATTRIBUTES
	Sid As Long
	Attributes As Long
End Type
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
Type COMSTAT
	fBitFields As Long
	cbInQue As Long
	cbOutQue As Long
End Type
Type DCB
	DCBlength As Long
	BaudRate As Long
	fBitFields As Long
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
	wReserved1 As Integer
End Type
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
Type GENERIC_MAPPING
	GenericRead As Long
	GenericWrite As Long
	GenericExecute As Long
	GenericAll As Long
End Type
Type Luid
	lowpart As Long
	highpart As Long
End Type

Type LUID_AND_ATTRIBUTES
	pLuid As Luid
	Attributes As Long
End Type
Type ACL
	AclRevision As Byte
	Sbz1 As Byte
	AclSize As Integer
	AceCount As Integer
	Sbz2 As Integer
End Type
Type ACE_HEADER
	AceType As Byte
	AceFlags As Byte
	AceSize As Long
End Type
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
Type ACL_REVISION_INFORMATION
	AclRevision As Long
End Type
Type ACL_SIZE_INFORMATION
	AceCount As Long
	AclBytesInUse As Long
	AclBytesFree As Long
End Type
Type SECURITY_DESCRIPTOR
	Revision As Byte
	Sbz1 As Byte
	Control As Long
	Owner As Long
	Group As Long
	Sacl As ACL
	Dacl As ACL
End Type
Type PRIVILEGE_SET
	PrivilegeCount As Long
	Control As Long
	Privilege(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Type EXCEPTION_RECORD
	ExceptionCode As Long
	ExceptionFlags As Long
	pExceptionRecord As Long
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
Type OFSTRUCT
	cBytes As Byte
	fFixedDisk As Byte
	nErrCode As Integer
	Reserved1 As Integer
	Reserved2 As Integer
	szPathName(OFS_MAXPATHNAME) As Byte
End Type
Type CRITICAL_SECTION
	dummy As Long
End Type
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
Type MEMORY_BASIC_INFORMATION
	BaseAddress As Long
	AllocationBase As Long
	AllocationProtect As Long
	RegionSize As Long
	State As Long
	Protect As Long
	lType As Long
End Type
Type EVENTLOGRECORD
	Length As Long
	Reserved As Long
	RecordNumber As Long
	TimeGenerated As Long
	TimeWritten As Long
	EventID As Long
	EventType As Integer
	NumStrings As Integer
	EventCategory As Integer
	ReservedFlags As Integer
	ClosingRecordNumber As Long
	StringOffset As Long
	UserSidLength As Long
	UserSidOffset As Long
	DataLength As Long
	DataOffset As Long
End Type
Type TOKEN_GROUPS
	GroupCount As Long
	Groups(ANYSIZE_ARRAY) As SID_AND_ATTRIBUTES
End Type
Type TOKEN_PRIVILEGES
	PrivilegeCount As Long
	Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
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
Type LDT_BYTES
	BaseMid As Byte
	Flags1 As Byte
	Flags2 As Byte
	BaseHi As Byte
End Type
Type LDT_ENTRY
	LimitLow As Integer
	BaseLow As Integer
	HighWord As Long
End Type
Type TIME_ZONE_INFORMATION
	Bias As Long
	StandardName(32) As Integer
	StandardDate As SystemTime
	StandardBias As Long
	DaylightName(32) As Integer
	DaylightDate As SystemTime
	DaylightBias As Long
End Type
Type WIN32_STREAM_ID
	dwStreamID As Long
	dwStreamAttributes As Long
	dwStreamSizeLow As Long
	dwStreamSizeHigh As Long
	dwStreamNameSize As Long
	cStreamName As Byte
End Type
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
Type CPINFO
	MaxCharSize As Long
	DefaultChar(MAX_DEFAULTCHAR) As Byte
	LeadByte(MAX_LEADBYTES) As Byte
End Type
Type NUMBERFMT
	NumDigits As Long
	LeadingZero As Long
	Grouping As Long
	lpDecimalSep As String
	lpThousandSep As String
	NegativeOrder As Long
End Type
Type CURRENCYFMT
	NumDigits As Long
	LeadingZero As Long
	Grouping As Long
	lpDecimalSep As String
	lpThousandSep As String
	NegativeOrder As Long
	PositiveOrder As Long
	lpCurrencySymbol As String
End Type
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
Type MOUSE_EVENT_RECORD
	dwMousePosition As COORD
	dwButtonState As Long
	dwControlKeyState As Long
	dwEventFlags As Long
End Type
Type WINDOW_BUFFER_SIZE_RECORD
	dwSize As COORD
End Type
Type MENU_EVENT_RECORD
	dwCommandId As Long
End Type
Type FOCUS_EVENT_RECORD
	bSetFocus As Long
End Type
Type CHAR_INFO
	Char As Integer
	Attributes As Integer
End Type
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
Type xform
	eM11 As Double
	eM12 As Double
	eM21 As Double
	eM22 As Double
	eDx As Double
	eDy As Double
End Type
Type BITMAP
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
Type BITMAPCOREHEADER
	bcSize As Long
	bcWidth As Integer
	bcHeight As Integer
	bcPlanes As Integer
	bcBitCount As Integer
End Type
Type BITMAPINFOHEADER
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
Type PELARRAY
	paXCount As Long
	paYCount As Long
	paXExt As Long
	paYExt As Long
	paRGBs As Integer
End Type
Type LOGBRUSH
	lbStyle As Long
	lbColor As Long
	lbHatch As Long
End Type
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
Type LOGPALETTE
	palVersion As Integer
	palNumEntries As Integer
	palPalEntry(1) As PALETTEENTRY
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
Type EXTLOGFONT
	elfLogFont As LOGFONT
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
Type DOCINFO
	cbSize As Long
	lpszDocName As String
	lpszOutput As String
End Type
Type KERNINGPAIR
	wFirst As Integer
	wSecond As Integer
	iKernAmount As Long
End Type
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
Type EMRSELECTCOLORSPACE
	pEmr As emr
	ihCS As Long
End Type
Type EMRCREATECOLORSPACE
	pEmr As emr
	ihCS As Long
	lcs As LOGCOLORSPACE
End Type
Type CBTACTIVATESTRUCT
	fMouse As Long
	hWndActive As Long
End Type
Type EVENTMSG
	message As Long
	paramL As Long
	paramH As Long
	time As Long
	hwnd As Long
End Type
Type CWPSTRUCT
	lParam As Long
	wParam As Long
	message As Long
	hwnd As Long
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
	hwnd As Long
	wHitTestCode As Long
	dwExtraInfo As Long
End Type
Type MINMAXINFO
	ptReserved As POINTAPI
	ptMaxSize As POINTAPI
	ptMaxPosition As POINTAPI
	ptMinTrackSize As POINTAPI
	ptMaxTrackSize As POINTAPI
End Type
Type COPYDATASTRUCT
	dwData As Long
	cbData As Long
	lpData As Long
End Type
Type WINDOWPOS
	hwnd As Long
	hWndInsertAfter As Long
	X As Long
	Y As Long
	cx As Long
	cy As Long
	Flags As Long
End Type
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
Type MEASUREITEMSTRUCT
	CtlType As Long
	CtlID As Long
	itemID As Long
	itemWidth As Long
	itemHeight As Long
	itemData As Long
End Type
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
Type DELETEITEMSTRUCT
	CtlType As Long
	CtlID As Long
	itemID As Long
	hwndItem As Long
	itemData As Long
End Type
Type COMPAREITEMSTRUCT
	CtlType As Long
	CtlID As Long
	hwndItem As Long
	itemID1 As Long
	itemData1 As Long
	itemID2 As Long
	itemData2 As Long
End Type
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
Type MENUITEMTEMPLATEHEADER
	versionNumber As Integer
	offset As Integer
End Type
Type MENUITEMTEMPLATE
	mtOption As Integer
	mtID As Integer
	mtString As Byte
End Type
Type ICONINFO
	fIcon As Long
	xHotspot As Long
	yHotspot As Long
	hbmMask As Long
	hbmColor As Long
End Type
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
Type MULTIKEYHELP
	mkSize As Long
	mkKeylist As Byte
	szKeyphrase As String * 253
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
Type DDEACK
	bAppReturnCode As Integer
	Reserved As Integer
	fbusy As Integer
	fAck As Integer
End Type
Type DDEADVISE
	Reserved As Integer
	fDeferUpd As Integer
	fAckReq As Integer
	cfFormat As Integer
End Type
Type DDEDATA
	unused As Integer
	fresponse As Integer
	fRelease As Integer
	Reserved As Integer
	fAckReq As Integer
	cfFormat As Integer
	Value(1) As Byte
End Type
Type DDEPOKE
	unused As Integer
	fRelease As Integer
	fReserved As Integer
	cfFormat As Integer
	Value(1) As Byte
End Type
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
Type HSZPAIR
	hszSvc As Long
	hszTopic As Long
End Type
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
	hwnd As Long
	hwndPartner As Long
End Type
Type DDEML_MSG_HOOK_DATA
	uiLo As Long
	uiHi As Long
	cbData As Long
	Data(8) As Long
End Type
Type MONMSGSTRUCT
	cb As Long
	hwndTo As Long
	dwTime As Long
	htask As Long
	wMsg As Long
	wParam As Long
	lParam As Long
	dmhd As DDEML_MSG_HOOK_DATA
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
	cc As CONVCONTEXT
	cbData As Long
	Data(8) As Long
End Type
Type MONHSZSTRUCT
	cb As Long
	fsAction As Long
	dwTime As Long
	hsz As Long
	htask As Long
	str As Byte
End Type
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
	hConvClient As Long
	hConvServer As Long
End Type
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
Type MIDIEVENT
	dwDeltaTime As Long
	dwStreamID As Long
	dwEvent As Long
	dwParms(1) As Long
End Type
Type MIDISTRMBUFFVER
	dwVersion As Long
	dwMid As Long
	dwOEMVersion As Long
End Type
Type MIDIPROPTIMEDIV
	cbStruct As Long
	dwTimeDiv As Long
End Type
Type MIDIPROPTEMPO
	cbStruct As Long
	dwTempo As Long
End Type
Type MIXERCAPS
	wMid As Integer
	wPid As Integer
	vDriverVersion As Long
	szPname As String * MAXPNAMELEN
	fdwSupport As Long
	cDestinations As Long
End Type
Type Target

	dwType As Long
	dwDeviceID As Long
	wMid As Integer
	wPid As Integer
	vDriverVersion As Long
	szPname As String * MAXPNAMELEN
End Type
Type MIXERLINE
	cbStruct As Long
	dwDestination As Long
	dwSource As Long
	dwLineID As Long
	fdwLine As Long
	dwUser As Long
	dwComponentType As Long
	cChannels As Long
	cConnections As Long
	cControls As Long
	szShortName As String * MIXER_SHORT_NAME_CHARS
	szName As String * MIXER_LONG_NAME_CHARS
	lpTarget As Target
End Type
Type MIXERCONTROL
	cbStruct As Long
	dwControlID As Long
	dwControlType As Long
	fdwControl As Long
	cMultipleItems As Long
	szShortName As String * MIXER_SHORT_NAME_CHARS
	szName As String * MIXER_LONG_NAME_CHARS
	Bounds As Double
	Metrics As Long
End Type
Type MIXERLINECONTROLS
	cbStruct As Long
	dwLineID As Long
	dwControl As Long
	cControls As Long
	cbmxctrl As Long
	pamxctrl As MIXERCONTROL
End Type
Type MIXERCONTROLDETAILS
	cbStruct As Long
	dwControlID As Long
	cChannels As Long
	item As Long
	cbDetails As Long
	paDetails As Long
End Type
Type MIXERCONTROLDETAILS_LISTTEXT
	dwParam1 As Long
	dwParam2 As Long
	szName As String * MIXER_LONG_NAME_CHARS
End Type
Type MIXERCONTROLDETAILS_BOOLEAN
	fValue As Long
End Type
Type MIXERCONTROLDETAILS_SIGNED
	lValue As Long
End Type
Type MIXERCONTROLDETAILS_UNSIGNED
	dwValue As Long
End Type
Type JOYINFOEX
	dwSize As Long
	dwFlags As Long
	dwXpos As Long
	dwYpos As Long
	dwZpos As Long
	dwRpos As Long
	dwUpos As Long
	dwVpos As Long
	dwButtons As Long
	dwButtonNumber As Long
	dwPOV As Long
	dwReserved1 As Long
	dwReserved2 As Long
End Type
Type DRVCONFIGINFO
	dwDCISize As Long
	lpszDCISectionName As String
	lpszDCIAliasName As String
	dnDevNode As Long
End Type
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
Type WAVEOUTCAPS
	wMid As Integer
	wPid As Integer
	vDriverVersion As Long
	szPname As String * MAXPNAMELEN
	dwFormats As Long
	wChannels As Integer
	dwSupport As Long
End Type
Type WAVEINCAPS
	wMid As Integer
	wPid As Integer
	vDriverVersion As Long
	szPname As String * MAXPNAMELEN
	dwFormats As Long
	wChannels As Integer
End Type
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
Type AUXCAPS
	wMid As Integer
	wPid As Integer
	vDriverVersion As Long
	szPname As String * MAXPNAMELEN
	wTechnology As Integer
	dwSupport As Long
End Type
Type TIMECAPS
	wPeriodMin As Long
	wPeriodMax As Long
End Type
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
Type MMCKINFO
	ckid As Long
	ckSize As Long
	fccType As Long
	dwDataOffset As Long
	dwFlags As Long
End Type
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
	hwnd As Long
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
	hwnd As Long
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
Type ADDJOB_INFO_1
	path As String
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
Type DRIVER_INFO_3
	cVersion As Long
	pName As String
	pEnvironment As String
	pDriverPath As String
	pDataFile As String
	pConfigFile As String
	pHelpFile As String
	pDependentFiles As String
	pMonitorName As String
	pDefaultDataType As String
End Type
Type DOC_INFO_2
	pDocName As String
	pOutputFile As String
	pDatatype As String
	dwMode As Long
	JobId As Long
End Type
Type PORT_INFO_2
	pPortName As String
	pMonitorName As String
	pDescription As String
	fPortType As Long
	Reserved As Long
End Type
Type PROVIDOR_INFO_1
	pName As String
	pEnvironment As String
	pDLLName As String
End Type
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
	ncb_reserve(10) As Byte
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
	name As String * NCBNAMSZ
	name_num As Integer
	name_flags As Integer
End Type
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
Type CRGB
	bRed As Byte
	bGreen As Byte
	bBlue As Byte
	bExtra As Byte
End Type
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
Type LARGE_INTEGER
	lowpart As Long
	highpart As Long
End Type
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
Type STYLEBUF
	dwStyle As Long
	szDescription As String * STYLE_DESCRIPTION_SIZE
End Type
Type MODEMDEVCAPS
	dwActualSize As Long
	dwRequiredSize As Long
	dwDevSpecificOffset As Long
	dwDevSpecificSize As Long
	dwModemProviderVersion As Long
	dwModemManufacturerOffset As Long
	dwModemManufacturerSize As Long
	dwModemModelOffset As Long
	dwModemModelSize As Long
	dwModemVersionOffset As Long
	dwModemVersionSize As Long
	dwDialOptions As Long
	dwCallSetupFailTimer As Long
	dwInactivityTimeout As Long
	dwSpeakerVolume As Long
	dwSpeakerMode As Long
	dwModemOptions As Long
	dwMaxDTERate As Long
	dwMaxDCERate As Long
	abVariablePortion(1) As Byte
End Type
Type MODEMSETTINGS
	dwActualSize As Long
	dwRequiredSize As Long
	dwDevSpecificOffset As Long
	dwDevSpecificSize As Long
	dwCallSetupFailTimer As Long
	dwInactivityTimeout As Long
	dwSpeakerVolume As Long
	dwSpeakerMode As Long
	dwPreferredModemOptions As Long
	dwNegotiatedModemOptions As Long
	dwNegotiatedDCERate As Long
	abVariablePortion(1) As Byte
End Type
Type DRAGINFO
	uSize As Long
	pt As POINTAPI
	fNC As Long
	lpFileList As String
	grfKeyState As Long
End Type
Type APPBARDATA
	cbSize As Long
	hwnd As Long
	uCallbackMessage As Long
	uEdge As Long
	rc As RECT
	lParam As Long
End Type
Type SHFILEOPSTRUCT
	hwnd As Long
	wFunc As Long     '????????
	pFrom As String    '??????
	pTo As String        '???????
	fFlags As Integer   '????
	fAnyOperationsAborted As Long
	hNameMappings As Long
	lpszProgressTitle As String
End Type
Type SHNAMEMAPPING
	pszOldPath As String
	pszNewPath As String
	cchOldPath As Long
	cchNewPath As Long
End Type
Type SHELLEXECUTEINFO
	cbSize As Long
	fMask As Long
	hwnd As Long
	lpVerb As String
	lpFile As String
	lpParameters As String
	lpDirectory As String
	nShow As Long
	hInstApp As Long
	lpIDList As Long
	lpClass As String
	hkeyClass As Long
	dwHotKey As Long
	hIcon As Long
	hProcess As Long
End Type
Type NOTIFYICONDATA
	cbSize As Long
	hwnd As Long
	uID As Long
	uFlags As Long
	uCallbackMessage As Long
	hIcon As Long
	szTip As String * 64
End Type
Type SHFILEINFO
	hIcon As Long
	iIcon As Long
	dwAttributes As Long
	szDisplayName As String * MAX_PATH
	szTypeName As String * 80
End Type
Type VS_FIXEDFILEINFO
	dwSignature As Long
	dwStrucVersion As Long
	dwFileVersionMS As Long
	dwFileVersionLS As Long
	dwProductVersionMS As Long
	dwProductVersionLS As Long
	dwFileFlagsMask As Long
	dwFileFlags As Long
	dwFileOS As Long
	dwFileType As Long
	dwFileSubtype As Long
	dwFileDateMS As Long
	dwFileDateLS As Long
End Type
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
Type OSVERSIONINFO
	dwOSVersionInfoSize As Long
	dwMajorVersion As Long
	dwMinorVersion As Long
	dwBuildNumber As Long
	dwPlatformId As Long
	szCSDVersion As String * 128
End Type
Type SYSTEM_POWER_STATUS
	ACLineStatus As Byte
	BatteryFlag As Byte
	BatteryLifePercent As Byte
	Reserved1 As Byte
	BatteryLifeTime As Long
	BatteryFullLifeTime As Long
End Type
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
Type NMHDR
	hwndFrom As Long
	idfrom As Long
	code As Long
End Type
Type OFNOTIFY
	hdr As NMHDR
	lpOFN As OPENFILENAME
	pszFile As String
End Type
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
Type FINDREPLACE
	lStructSize As Long
	hwndOwner As Long
	hInstance As Long
	Flags As Long
	lpstrFindWhat As String
	lpstrReplaceWith As String
	wFindWhatLen As Integer
	wReplaceWithLen As Integer
	lCustData As Long
	lpfnHook As Long
	lpTemplateName As String
End Type
Type ChooseFont
	lStructSize As Long
	hwndOwner As Long
	hdc As Long
	lpLogFont As LOGFONT
	iPointSize As Long
	Flags As Long
	rgbColors As Long
	lCustData As Long
	lpfnHook As Long
	lpTemplateName As String
	hInstance As Long
	lpszStyle As String
	nFontType As Integer
	MISSING_ALIGNMENT As Integer
	nSizeMin As Long
	nSizeMax As Long
End Type
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
Type DEVNAMES
	wDriverOffset As Integer
	wDeviceOffset As Integer
	wOutputOffset As Integer
	wDefault As Integer
End Type
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
Type DRAWTEXTPARAMS
	cbSize As Long
	iTabLength As Long
	iLeftMargin As Long
	iRightMargin As Long
	uiLengthDrawn As Long
End Type
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
Type SCROLLINFO
	cbSize As Long
	fMask As Long
	nMin As Long
	nMax As Long
	nPage As Long
	nPos As Long
	nTrackPos As Long
End Type
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
Type TPMPARAMS
	cbSize As Long
	rcExclude As RECT
End Type
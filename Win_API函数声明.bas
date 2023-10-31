Attribute VB_Name = "´úÂëÄ£¿é"
Declare Function AbortDoc Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function AbortPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function AbortPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function AbortSystemShutdown Lib "ADVAPI32.DLL" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long
Declare Function AccessCheck Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As Long, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, PrivilegeSet As PRIVILEGE_SET, PrivilegeSetLength As Long, GrantedAccess As Long, ByVal Status As Long) As Long
Declare Function AccessCheckAndAuditAlarm Lib "ADVAPI32.DLL" Alias "AccessCheckAndAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, SecurityDescriptor As SECURITY_DESCRIPTOR, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, ByVal ObjectCreation As Long, GrantedAccess As Long, ByVal AccessStatus As Long, ByVal pfGenerateOnClose As Long) As Long
Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal Flags As Long) As Long
Declare Function AddAccessAllowedAce Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
Declare Function AddAccessDeniedAce Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
Declare Function AddAce Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwStartingAceIndex As Long, pAceList As Any, ByVal nAceListLength As Long) As Long
Declare Function AddAtom Lib "kernel32" Alias "AddAtomA" (ByVal lpString As String) As Integer
Declare Function AddAuditAccessAce Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwAccessMask As Long, pSid As Any, ByVal bAuditSuccess As Long, ByVal bAuditFailure As Long) As Long
Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte) As Long
Declare Function AddJob Lib "winspool.drv" Alias "AddJobA" (ByVal hPrinter As Long, ByVal Level As Long, pData As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Declare Function AddMonitor Lib "winspool.drv" Alias "AddMonitorA" (ByVal pName As String, ByVal Level As Long, pMonitors As Byte) As Long
Declare Function AddPort Lib "winspool.drv" Alias "AddPortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pMonitorName As String) As Long
Declare Function AddPrinter Lib "winspool.drv" Alias "AddPrinterA" (ByVal pName As String, ByVal Level As Long, pPrinter As Any) As Long
Declare Function AddPrinterConnection Lib "winspool.drv" Alias "AddPrinterConnectionA" (ByVal pName As String) As Long
Declare Function AddPrinterDriver Lib "winspool.drv" Alias "AddPrinterDriverA" (ByVal pName As String, ByVal Level As Long, pDriverInfo As Any) As Long
Declare Function AddPrintProcessor Lib "winspool.drv" Alias "AddPrintProcessorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pPathName As String, ByVal pPrintProcessorName As String) As Long
Declare Function AddPrintProvidor Lib "winspool.drv" Alias "AddPrintProvidorA" (ByVal pName As String, ByVal Level As Long, pProvidorInfo As Byte) As Long
Declare Function AdjustTokenGroups Lib "ADVAPI32.DLL" (ByVal TokenHandle As Long, ByVal ResetToDefault As Long, NewState As TOKEN_GROUPS, ByVal BufferLength As Long, PreviousState As TOKEN_GROUPS, ReturnLength As Long) As Long
Declare Function AdjustTokenPrivileges Lib "ADVAPI32.DLL" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Declare Function AdjustWindowRect Lib "user32" (lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
Declare Function AdjustWindowRectEx Lib "user32" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long
Declare Function AdvancedDocumentProperties Lib "winspool.drv" Alias "AdvancedDocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, pDevModeInput As DEVMODE) As Long
Declare Function AllocateAndInitializeSid Lib "ADVAPI32.DLL" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Declare Function AllocateLocallyUniqueId Lib "ADVAPI32.DLL" (Luid As LARGE_INTEGER) As Long
Declare Function AllocConsole Lib "kernel32" () As Long
Declare Function AngleArc Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dwRadius As Long, ByVal eStartAngle As Double, ByVal eSweepAngle As Double) As Long
Declare Function AnimatePalette Lib "gdi32" Alias "AnimatePaletteA" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteColors As PALETTEENTRY) As Long
Declare Function AnyPopup Lib "user32" () As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function ArcTo Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function AreAllAccessesGranted Lib "ADVAPI32.DLL" (ByVal GrantedAccess As Long, ByVal DesiredAccess As Long) As Long
Declare Function AreAnyAccessesGranted Lib "ADVAPI32.DLL" (ByVal GrantedAccess As Long, ByVal DesiredAccess As Long) As Long
Declare Function ArrangeIconicWindows Lib "user32" (ByVal hwnd As Long) As Long
Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Declare Function auxGetDevCaps Lib "winmm.dll" Alias "auxGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As AUXCAPS, ByVal uSize As Long) As Long
Declare Function auxGetNumDevs Lib "winmm.dll" () As Long
Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Declare Function auxOutMessage Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function BackupEventLog Lib "ADVAPI32.DLL" Alias "BackupEventLogA" (ByVal hEventLog As Long, ByVal lpBackupFileName As String) As Long
Declare Function BackupRead Lib "kernel32" (ByVal hFile As Long, lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, lpContext As Any) As Long
Declare Function BackupSeek Lib "kernel32" (ByVal hFile As Long, ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, lpdwLowByteSeeked As Long, lpdwHighByteSeeked As Long, lpContext As Long) As Long
Declare Function BackupWrite Lib "kernel32" (ByVal hFile As Long, lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, lpContext As Long) As Long
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function BroadcastSystemMessage Lib "user32" (ByVal dw As Long, pdw As Long, ByVal un As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, lpDCB As DCB) As Long
Declare Function BuildCommDCBAndTimeouts Lib "kernel32" Alias "BuildCommDCBAndTimeoutsA" (ByVal lpDef As String, lpDCB As DCB, lpCommTimeouts As COMMTIMEOUTS) As Long
Declare Function CallMsgFilter Lib "user32" Alias "CallMsgFilterA" (lpMsg As msg, ByVal nCode As Long) As Long
Declare Function CallNamedPipe Lib "kernel32" Alias "CallNamedPipeA" (ByVal lpNamedPipeName As String, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, ByVal nTimeOut As Long) As Long
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function CancelDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CascadeWindows Lib "user32" (ByVal hwndParent As Long, ByVal wHow As Long, ByVal lpRect As RECT, ByVal cKids As Long, lpKids As Long) As Integer
Declare Function ChangeClipboardChain Lib "user32" (ByVal hwnd As Long, ByVal hWndNext As Long) As Long
Declare Function ChangeMenu Lib "user32" Alias "ChangeMenuA" (ByVal hMenu As Long, ByVal cmd As Long, ByVal lpszNewItem As String, ByVal cmdInsert As Long, ByVal Flags As Long) As Long
Declare Function ChangeServiceConfig Lib "ADVAPI32.DLL" Alias "ChangeServiceConfigA" (ByVal hService As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lpServiceStartName As String, ByVal lpPassword As String, ByVal lpDisplayName As String) As Long
Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Declare Function CharLowerBuff Lib "user32" Alias "CharLowerBuffA" (ByVal lpsz As String, ByVal cchLength As Long) As Long
Declare Function CharNext Lib "user32" Alias "CharNextA" (ByVal lpsz As String) As String
Declare Function CharPrev Lib "user32" Alias "CharPrevA" (ByVal lpszStart As String, ByVal lpszCurrent As String) As String
Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Declare Function CharUpperBuff Lib "user32" Alias "CharUpperBuffA" (ByVal lpsz As String, ByVal cchLength As Long) As Long
Declare Function CheckColorsInGamut Lib "gdi32" (ByVal hdc As Long, lpv As Any, lpv2 As Any, ByVal dw As Long) As Long
Declare Function CheckDlgButton Lib "user32" Alias "CheckDLGButtonA" (ByVal hDlg As Long, ByVal nIDButton As Long, ByVal wCheck As Long) As Long
Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Declare Function CheckRadioButton Lib "user32" Alias "CheckRadioButtonA" (ByVal hDlg As Long, ByVal nIDFirstButton As Long, ByVal nIDLastButton As Long, ByVal nIDCheckButton As Long) As Long
Declare Function ChildWindowFromPoint Lib "user32" (ByVal hwnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Declare Function ChildWindowFromPoint Lib "user32" (ByVal hwndParent As Long, ByVal pt As POINTAPI) As Long
Declare Function ChildWindowFromPointEx Lib "user32" (ByVal hwnd As Long, ByVal pt As POINTAPI, ByVal un As Long) As Long
Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As Long, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Declare Function Chord Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function ClearCommBreak Lib "kernel32" (ByVal nCid As Long) As Long
Declare Function ClearCommError Lib "kernel32" (ByVal hFile As Long, lpErrors As Long, lpStat As COMSTAT) As Long
Declare Function ClearEventLog Lib "ADVAPI32.DLL" Alias "ClearEventLogA" (ByVal hEventLog As Long, ByVal lpBackupFileName As String) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function CloseDesktop Lib "user32" (ByVal hDesktop As Long) As Long
Declare Function CloseDriver Lib "winmm.dll" (ByVal hDriver As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CloseEventLog Lib "ADVAPI32.DLL" (ByVal hEventLog As Long) As Long
Declare Function CloseFigure Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function CloseMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function CloseServiceHandle Lib "ADVAPI32.DLL" (ByVal hSCObject As Long) As Long
Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CloseWindowStation Lib "user32" (ByVal hWinSta As Long) As Long
Declare Function ColorMatchToTarget Lib "gdi32" (ByVal hdc As Long, ByVal hdc2 As Long, ByVal dw As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function CombineTransform Lib "gdi32" (lpxformResult As xform, lpxform1 As xform, lpxform2 As xform) As Long
Declare Function CommandLineToArgv Lib "shell32" Alias "CommandLineToArgvW" (ByVal lpCmdLine As String, pNumArgs As Integer) As Long
Declare Function CommConfigDialog Lib "kernel32" Alias "CommConfigDialogA" (ByVal lpszName As String, ByVal hwnd As Long, lpCC As COMMCONFIG) As Long
Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long
Declare Function CompareString Lib "kernel32" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
Declare Function ConfigurePort Lib "winspool.drv" Alias "ConfigurePortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pPortName As String) As Long
Declare Function ConnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function ConnectToPrinterDlg Lib "winspool.drv" (ByVal hwnd As Long, ByVal Flags As Long) As Long
Declare Function ContinueDebugEvent Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwThreadId As Long, ByVal dwContinueStatus As Long) As Long
Declare Function ControlService Lib "ADVAPI32.DLL" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Declare Function ConvertDefaultLocale Lib "kernel32" (ByVal Locale As Long) As Long
Declare Function CopyAcceleratorTable Lib "user32" Alias "CopyAcceleratorTableA" (ByVal hAccelSrc As Long, lpAccelDst As ACCEL, ByVal cAccelEntries As Long) As Long
Declare Function CopyCursor Lib "user32" (ByVal hcur As Long) As Long
Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function CopyLZFile Lib "lz32" (ByVal n1 As Long, ByVal n2 As Long) As Long
Declare Function CopyMetaFile Lib "gdi32" Alias "CopyMetaFileA" (ByVal hMF As Long, ByVal lpFileName As String) As Long
Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Declare Function CopySid Lib "ADVAPI32.DLL" (ByVal nDestinationSidLength As Long, pDestinationSid As Any, pSourceSid As Any) As Long
Declare Function CountClipboardFormats Lib "user32" () As Long
Declare Function CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableA" (lpaccl As ACCEL, ByVal cEntries As Long) As Long
Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateColorSpace Lib "gdi32" Alias "CreateColorSpaceA" (lplogcolorspace As LOGCOLORSPACE) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateConsoleScreenBuffer Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwFlags As Long, lpScreenBufferData As Any) As Long
Declare Function CreateCursor Lib "user32" (ByVal hInstance As Long, ByVal nXhotspot As Long, ByVal nYhotspot As Long, ByVal nWidth As Long, ByVal nHeight As Long, lpANDbitPlane As Any, lpXORbitPlane As Any) As Long
Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Declare Function CreateDesktop Lib "user32" Alias "CreateDesktopA" (ByVal lpszDesktop As String, ByVal lpszDevice As String, pDevmode As DEVMODE, ByVal dwFlags As Long, ByVal dwDesiredAccess As Long, lpsa As SECURITY_ATTRIBUTES) As Long
Declare Function CreateDialogIndirectParam Lib "user32" Alias "CreateDialogIndirectParamA" (ByVal hInstance As Long, lpTemplate As DLGTEMPLATE, ByVal hwndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
Declare Function CreateDialogParam Lib "user32" Alias "CreateDialogParamA" (ByVal hInstance As Long, ByVal lpName As String, ByVal hwndParent As Long, ByVal lpDialogFunc As Long, ByVal lParamInit As Long) As Long
Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function CreateDIBPatternBrush Lib "gdi32" (ByVal hPackedDIB As Long, ByVal wUsage As Long) As Long
Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function CreateDiscardableBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As String, lpRect As RECT, ByVal lpDescription As String) As Long
Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Declare Function CreateIcon Lib "user32" (ByVal hInstance As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Byte, ByVal nBitsPixel As Byte, lpANDbits As Byte, lpXORbits As Byte) As Long
Declare Function CreateIconFromResource Lib "user32" (presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As Long
Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Declare Function CreateIoCompletionPort Lib "kernel32" (ByVal FileHandle As Long, ByVal ExistingCompletionPort As Long, ByVal CompletionKey As Long, ByVal NumberOfConcurrentThreads As Long) As Long
Declare Function CreateMailslot Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function CreateMDIWindow Lib "user32" Alias "CreateMDIWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hInstance As Long, ByVal lParam As Long) As Long
Declare Function CreateMenu Lib "user32" () As Long
Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
'Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Declare Function CreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeA" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function CreatePrivateObjectSecurity Lib "ADVAPI32.DLL" (ParentDescriptor As SECURITY_DESCRIPTOR, CreatorDescriptor As SECURITY_DESCRIPTOR, NewDescriptor As SECURITY_DESCRIPTOR, ByVal IsDirectoryObject As Long, ByVal Token As Long, GenericMapping As GENERIC_MAPPING) As Long
Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function CreateProcessAsUser Lib "kernel32" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As SECURITY_ATTRIBUTES, ByVal lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, ByVal lpStartupInfo As STARTUPINFO, ByVal lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function CreateScalableFontResource Lib "gdi32" Alias "CreateScalableFontResourceA" (ByVal fHidden As Long, ByVal lpszResourceFile As String, ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) As Long
Declare Function CreateSemaphore Lib "kernel32" Alias "CreateSemaphoreA" (lpSemaphoreAttributes As SECURITY_ATTRIBUTES, ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) As Long
Declare Function CreateService Lib "ADVAPI32.DLL" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function CreateTapePartition Lib "kernel32" (ByVal hDevice As Long, ByVal dwPartitionMethod As Long, ByVal dwCount As Long, ByVal dwSize As Long) As Long
Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DdeAbandonTransaction Lib "user32" (ByVal idInst As Long, ByVal hConv As Long, ByVal idTransaction As Long) As Long
Declare Function DdeAccessData Lib "user32" Alias "DdeAccessDataA" (ByVal hData As Long, pcbDataSize As Long) As Long
Declare Function DdeAddData Lib "user32" Alias "DdeAddDataA" (ByVal hData As Long, pSrc As Byte, ByVal cb As Long, ByVal cbOff As Long) As Long
Declare Function DdeClientTransaction Lib "user32" (pData As Byte, ByVal cbData As Long, ByVal hConv As Long, ByVal hszItem As Long, ByVal wFmt As Long, ByVal wType As Long, ByVal dwTimeout As Long, pdwResult As Long) As Long
Declare Function DdeCmpStringHandles Lib "user32" (ByVal hsz1 As Long, ByVal hsz2 As Long) As Long
Declare Function DdeConnect Lib "user32" (ByVal idInst As Long, ByVal hszService As Long, ByVal hszTopic As Long, pCC As CONVCONTEXT) As Long
Declare Function DdeConnectList Lib "user32" (ByVal idInst As Long, ByVal hszService As Long, ByVal hszTopic As Long, ByVal hConvList As Long, pCC As CONVCONTEXT) As Long
Declare Function DdeCreateDataHandle Lib "user32" (ByVal idInst As Long, pSrc As Byte, ByVal cb As Long, ByVal cbOff As Long, ByVal hszItem As Long, ByVal wFmt As Long, ByVal afCmd As Long) As Long
Declare Function DdeCreateStringHandle Lib "user32" Alias "DdeCreateStringHandleA" (ByVal idInst As Long, ByVal psz As String, ByVal iCodePage As Long) As Long
Declare Function DdeDisconnect Lib "user32" (ByVal hConv As Long) As Long
Declare Function DdeDisconnectList Lib "user32" (ByVal hConvList As Long) As Long
Declare Function DdeEnableCallback Lib "user32" (ByVal idInst As Long, ByVal hConv As Long, ByVal wCmd As Long) As Long
Declare Function DdeFreeDataHandle Lib "user32" (ByVal hData As Long) As Long
Declare Function DdeFreeStringHandle Lib "user32" (ByVal idInst As Long, ByVal hsz As Long) As Long
Declare Function DdeGetData Lib "user32" Alias "DdeGetDataA" (ByVal hData As Long, pDst As Byte, ByVal cbMax As Long, ByVal cbOff As Long) As Long
Declare Function DdeGetLastError Lib "user32" (ByVal idInst As Long) As Long
Declare Function DdeImpersonateClient Lib "user32" (ByVal hConv As Long) As Long
Declare Function DdeInitialize Lib "user32" Alias "DdeInitializeA" (pidInst As Long, ByVal pfnCallback As Long, ByVal afCmd As Long, ByVal ulRes As Long) As Integer
Declare Function DdeKeepStringHandle Lib "user32" (ByVal idInst As Long, ByVal hsz As Long) As Long
Declare Function DdeNameService Lib "user32" (ByVal idInst As Long, ByVal hsz1 As Long, ByVal hsz2 As Long, ByVal afCmd As Long) As Long
Declare Function DdePostAdvise Lib "user32" (ByVal idInst As Long, ByVal hszTopic As Long, ByVal hszItem As Long) As Long
Declare Function DdeQueryConvInfo Lib "user32" (ByVal hConv As Long, ByVal idTransaction As Long, pConvInfo As CONVINFO) As Long
Declare Function DdeQueryNextServer Lib "user32" (ByVal hConvList As Long, ByVal hConvPrev As Long) As Long
Declare Function DdeQueryString Lib "user32" Alias "DdeQueryStringA" (ByVal idInst As Long, ByVal hsz As Long, ByVal psz As String, ByVal cchMax As Long, ByVal iCodePage As Long) As Long
Declare Function DdeReconnect Lib "user32" (ByVal hConv As Long) As Long
Declare Function DdeSetQualityOfService Lib "user32" (ByVal hWndClient As Long, pqosNew As SECURITY_QUALITY_OF_SERVICE, pqosPrev As SECURITY_QUALITY_OF_SERVICE) As Long
Declare Function DdeSetUserHandle Lib "user32" (ByVal hConv As Long, ByVal id As Long, ByVal hUser As Long) As Long
Declare Function DdeUnaccessData Lib "user32" Alias "DdeUnaccessDataA" (ByVal hData As Long) As Long
Declare Function DdeUninitialize Lib "user32" (ByVal idInst As Long) As Long
Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long
Declare Function DefDlgProc Lib "user32" Alias "DefDlgProcA" (ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DefDriverProc Lib "winmm.dll" (ByVal dwDriverIdentifier As Long, ByVal hdrvr As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function DefFrameProc Lib "user32" Alias "DefFrameProcA" (ByVal hwnd As Long, ByVal hWndMDIClient As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DefineDosDevice Lib "kernel32" Alias "DefineDosDeviceA" (ByVal dwFlags As Long, ByVal lpDeviceName As String, ByVal lpTargetPath As String) As Long
Declare Function DefMDIChildProc Lib "user32" Alias "DefMDIChildProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DeleteAce Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal dwAceIndex As Long) As Long
Declare Function DeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Declare Function DeleteColorSpace Lib "gdi32" (ByVal hcolorspace As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As Long) As Long
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As Long, ByVal pFormName As String) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Declare Function DeleteMonitor Lib "winspool.drv" Alias "DeleteMonitorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pMonitorName As String) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DeletePort Lib "winspool.drv" Alias "DeletePortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pPortName As String) As Long
Declare Function DeletePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function DeletePrinterConnection Lib "winspool.drv" Alias "DeletePrinterConnectionA" (ByVal pName As String) As Long
Declare Function DeletePrinterDriver Lib "winspool.drv" Alias "DeletePrinterDriverA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pDriverName As String) As Long
Declare Function DeletePrintProcessor Lib "winspool.drv" Alias "DeletePrintProcessorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pPrintProcessorName As String) As Long
Declare Function DeletePrintProvidor Lib "winspool.drv" Alias "DeletePrintProvidorA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pPrintProvidorName As String) As Long
Declare Function DeleteService Lib "ADVAPI32.DLL" (ByVal hService As Long) As Long
Declare Function DeregisterEventSource Lib "ADVAPI32.DLL" (ByVal hEventLog As Long) As Long
Declare Function DescribePixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal n As Long, ByVal un As Long, lpPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Declare Function DestroyAcceleratorTable Lib "user32" (ByVal haccel As Long) As Long
Declare Function DestroyCaret Lib "user32" () As Long
Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function DestroyPrivateObjectSecurity Lib "ADVAPI32.DLL" (ObjectDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As DEVMODE) As Long
Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function DialogBoxIndirectParam Lib "user32" Alias "DialogBoxIndirectParamA" (ByVal hInstance As Long, hDialogTemplate As DLGTEMPLATE, ByVal hwndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
Declare Function DisableThreadLibraryCalls Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function DisconnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long) As Long
Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Declare Function DlgDirList Lib "user32" Alias "DlgDirListA" (ByVal hDlg As Long, ByVal lpPathSpec As String, ByVal nIDListBox As Long, ByVal nIDStaticPath As Long, ByVal wFileType As Long) As Long
Declare Function DlgDirListComboBox Lib "user32" Alias "DlgDirListComboBoxA" (ByVal hDlg As Long, ByVal lpPathSpec As String, ByVal nIDComboBox As Long, ByVal nIDStaticPath As Long, ByVal wFileType As Long) As Long
Declare Function DlgDirSelectComboBoxEx Lib "user32" Alias "DlgDirSelectComboBoxExA" (ByVal hWndDlg As Long, ByVal lpszPath As String, ByVal cbPath As Long, ByVal idComboBox As Long) As Long
Declare Function DlgDirSelectEx Lib "user32" Alias "DlgDirSelectExA" (ByVal hWndDlg As Long, ByVal lpszPath As String, ByVal cbPath As Long, ByVal idListBox As Long) As Long
Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, pDevModeInput As DEVMODE, ByVal fMode As Long) As Long
Declare Function DoEnvironmentSubst Lib "shell32.dll" Alias "DoEnvironmentSubstA" (ByVal szString As String, ByVal cbString As Long) As Long
Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
Declare Function DPtoLP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function DragDetect Lib "user32" (ByVal hwnd As Long, ByVal pt As POINTAPI) As Long
Declare Function DragObject Lib "user32" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal un As Long, ByVal dw As Long, ByVal hCursor As Long) As Long
Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Declare Function DragQueryPoint Lib "shell32.dll" (ByVal hDrop As Long, lpPoint As POINTAPI) As Long
Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Declare Function DrawCaption Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long
Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Declare Function DrawEscape Lib "gdi32" (ByVal hdc As Long, ByVal nEscape As Long, ByVal cbInput As Long, ByVal lpszInData As String) As Long
Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Declare Function DrvGetModuleHandle Lib "winmm.dll" (ByVal hDriver As Long) As Long
Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Declare Function DuplicateIcon Lib "shell32.dll" (ByVal hInst As Long, ByVal hIcon As Long) As Long
Declare Function DuplicateToken Lib "ADVAPI32.DLL" (ByVal ExistingTokenHandle As Long, Impersonationlevel As Integer, DuplicateTokenHandle As Long) As Long
Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Declare Function EndDialog Lib "user32" (ByVal hDlg As Long, ByVal nResult As Long) As Long
Declare Function EndDoc Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long
Declare Function EnumCalendarInfo Lib "kernel32" Alias "EnumCalendarInfoA" (ByVal lpCalInfoEnumProc As Long, ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Declare Function EnumDateFormats Lib "kernel32" (ByVal lpDateFmtEnumProc As Long, ByVal Locale As Long, ByVal dwFlags As Long) As Long
Declare Function EnumDependentServices Lib "ADVAPI32.DLL" Alias "EnumDependentServicesA" (ByVal hService As Long, ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long) As Long
Declare Function EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hWinSta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Declare Function EnumEnhMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hemf As Long, ByVal lpEnhMetaFunc As Long, lpData As Any, lpRect As RECT) As Long
Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, ByVal lParam As Long) As Long
Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long
Declare Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hdc As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long
Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function EnumICMProfiles Lib "gdi32" Alias "EnumICMProfilesA" (ByVal hdc As Long, ByVal icmEnumProc As Long, ByVal lParam As Long) As Long
Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, pJob As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function EnumMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hMetafile As Long, ByVal lpMFEnumProc As Long, ByVal lParam As Long) As Long
Declare Function EnumMonitors Lib "winspool.drv" Alias "EnumMonitorsA" (ByVal pName As String, ByVal Level As Long, pMonitors As Byte, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function EnumObjects Lib "gdi32" (ByVal hdc As Long, ByVal n As Long, ByVal lpGOBJEnumProc As Long, lpVoid As Any) As Long
Declare Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, ByVal lpbPorts As Long, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function EnumPrinterDrivers Lib "winspool.drv" Alias "EnumPrinterDriversA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverInfo As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcRetruned As Long) As Long
Declare Function EnumPrinterPropertySheets Lib "winspool.drv" (hPrinter As Long, hwnd As Long, lpfnAdd As Long, ByVal lParam As Long) As Long
Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal Flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function EnumPrintProcessorDatatypes Lib "winspool.drv" Alias "EnumPrintProcessorDatatypesA" (ByVal pName As String, ByVal pPrintProcessorName As String, ByVal Level As Long, pDatatypes As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcRetruned As Long) As Long
Declare Function EnumPrintProcessors Lib "winspool.drv" Alias "EnumPrintProcessorsA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pPrintProcessorInfo As Byte, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function EnumProps Lib "user32" Alias "EnumPropsA" (ByVal hwnd As Long, ByVal lpEnumFunc As Long) As Long
Declare Function EnumPropsEx Lib "user32" Alias "EnumPropsExA" (ByVal hwnd As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumResourceLanguages Lib "kernel32" Alias "EnumResourceLanguagesA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumServicesStatus Lib "ADVAPI32.DLL" Alias "EnumServicesStatusA" (ByVal hSCManager As Long, ByVal dwServiceType As Long, ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long, lpResumeHandle As Long) As Long
Declare Function EnumSystemCodePages Lib "kernel32" (ByVal lpCodePageEnumProc As Long, ByVal dwFlags As Long) As Long
Declare Function EnumSystemLocales Lib "kernel32" (ByVal lpLocaleEnumProc As Long, ByVal dwFlags As Long) As Long
Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Declare Function EnumTimeFormats Lib "kernel32" (ByVal lpTimeFmtEnumProc As Long, ByVal Locale As Long, ByVal dwFlags As Long) As Long
Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, lParam As Any) As Long
Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EqualPrefixSid Lib "ADVAPI32.DLL" (pSid1 As Any, pSid2 As Any) As Long
Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Declare Function EqualRgn Lib "gdi32" (ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long) As Long
Declare Function EqualSid Lib "ADVAPI32.DLL" (pSid1 As Any, pSid2 As Any) As Long
Declare Function EraseTape Lib "kernel32" (ByVal hDevice As Long, ByVal dwEraseType As Long, ByVal bimmediate As Long) As Long
Declare Function Escape Lib "gdi32" (ByVal hdc As Long, ByVal nEscape As Long, ByVal nCount As Long, ByVal lpInData As String, lpOutData As Any) As Long
Declare Function EscapeCommFunction Lib "kernel32" (ByVal nCid As Long, ByVal nFunc As Long) As Long
Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function ExcludeUpdateRgn Lib "user32" (ByVal hdc As Long, ByVal hwnd As Long) As Long
Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long
Declare Function ExtCreateRegion Lib "gdi32" (lpXform As xform, ByVal nCount As Long, lpRgnData As RgnData) As Long
Declare Function ExtEscape Lib "gdi32" (ByVal hdc As Long, ByVal nEscape As Long, ByVal cbInput As Long, ByVal lpszInData As String, ByVal cbOutput As Long, ByVal lpszOutData As String) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Declare Function ExtSelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal fnMode As Long) As Long
Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDX As Long) As Long
Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SystemTime) As Long
Declare Function FillConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttribute As Long, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
Declare Function FillConsoleOutputCharacter Lib "kernel32" Alias "FillConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long
Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Declare Function FindAtom Lib "kernel32" Alias "FindAtomA" (ByVal lpString As String) As Integer
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function FindCloseChangeNotification Lib "kernel32" (ByVal hChangeHandle As Long) As Long
Declare Function FindClosePrinterChangeNotification Lib "winspool.drv" (ByVal hChange As Long) As Long
Declare Function FindEnvironmentString Lib "shell32.dll" Alias "FindEnvironmentStringA" (ByVal szEnvVar As String) As String
Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Declare Function FindFirstChangeNotification Lib "kernel32" Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindFirstFreeAce Lib "ADVAPI32.DLL" (pAcl As ACL, pAce As Long) As Long
Declare Function FindFirstPrinterChangeNotification Lib "winspool.drv" (ByVal hPrinter As Long, ByVal fdwFlags As Long, ByVal fdwOptions As Long, ByVal pPrinterNotifyOptions As String) As Long
Declare Function FindNextChangeNotification Lib "kernel32" (ByVal hChangeHandle As Long) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextPrinterChangeNotification Lib "winspool.drv" (ByVal hChange As Long, pdwChange As Long, ByVal pvReserved As String, ByVal ppPrinterNotifyInfo As Long) As Long
Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long) As Long
Declare Function FindText Lib "comdlg32.dll" Alias "FindTextA " (pFindreplace As FINDREPLACE) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function FixBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal n1 As Long, ByVal n2 As Long, lpPoint As POINTAPI) As Long
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Declare Function FlattenPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function FlushConsoleInputBuffer Lib "kernel32" (ByVal hConsoleInput As Long) As Long
Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Declare Function FlushInstructionCache Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, ByVal dwSize As Long) As Long
Declare Function FlushViewOfFile Lib "kernel32" (lpBaseAddress As Any, ByVal dwNumberOfBytesToFlush As Long) As Long
Declare Function FoldString Lib "kernel32" Alias "FoldStringA" (ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function FreeConsole Lib "kernel32" () As Long
Declare Function FreeDDElParam Lib "user32" (ByVal msg As Long, ByVal lParam As Long) As Long
Declare Function FreeEnvironmentStrings Lib "kernel32" Alias "FreeEnvironmentStringsA" (ByVal lpsz As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Declare Function GdiComment Lib "gdi32" (ByVal hdc As Long, ByVal cbSize As Long, lpData As Byte) As Long
Declare Function GdiFlush Lib "gdi32" () As Long
Declare Function GdiGetBatchLimit Lib "gdi32" () As Long
Declare Function GdiSetBatchLimit Lib "gdi32" (ByVal dwLimit As Long) As Long
Declare Function GenerateConsoleCtrlEvent Lib "kernel32" (ByVal dwCtrlEvent As Long, ByVal dwProcessGroupId As Long) As Long
Declare Function GetAce Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal dwAceIndex As Long, pAce As Any) As Long
Declare Function GetAclInformation Lib "ADVAPI32.DLL" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
Declare Function GetACP Lib "kernel32" () As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function GetArcDirection Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetAspectRatioFilterEx Lib "gdi32" (ByVal hdc As Long, lpAspectRatio As Size) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetAtomName Lib "kernel32" Alias "GetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeA" (ByVal lpApplicationName As String, lpBinaryType As Long) As Long
Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Declare Function GetBitmapDimensionEx Lib "gdi32" (ByVal hBitmap As Long, lpDimension As Size) As Long
Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetBkMode Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetBoundsRect Lib "gdi32" (ByVal hdc As Long, lprcBounds As RECT, ByVal Flags As Long) As Long
Declare Function GetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Declare Function GetCapture Lib "user32" () As Long
Declare Function GetCaretBlinkTime Lib "user32" () As Long
Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetCharABCWidths Lib "gdi32" Alias "GetCharABCWidthsA" (ByVal hdc As Long, ByVal uFirstChar As Long, ByVal uLastChar As Long, lpabc As ABC) As Long
Declare Function GetCharABCWidthsFloat Lib "gdi32" Alias "GetCharABCWidthsFloatA" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpABCF As ABCFLOAT) As Long
Declare Function GetCharacterPlacement Lib "gdi32" Alias " GetCharacterPlacementA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n1 As Long, ByVal n2 As Long, lpGcpResults As GCP_RESULTS, ByVal dw As Long) As Long
Declare Function GetCharWidth Lib "gdi32" Alias "GetCharWidthA" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, lpn As Long) As Long
Declare Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpBuffer As Long) As Long
Declare Function GetCharWidthFloat Lib "gdi32" Alias "GetCharWidthFloatA" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, pxBuffer As Double) As Long
Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long
Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetClassWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As Long
Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function GetClipboardOwner Lib "user32" () As Long
Declare Function GetClipboardViewer Lib "user32" () As Long
Declare Function GetClipBox Lib "gdi32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function GetClipCursor Lib "user32" (lprc As RECT) As Long
Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function GetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As ColorAdjustment) As Long
Declare Function GetColorSpace Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As String
Declare Function GetCommConfig Lib "kernel32" (ByVal hCommDev As Long, lpCC As COMMCONFIG, lpdwSize As Long) As Long
Declare Function GetCommMask Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long) As Long
Declare Function GetCommModemStatus Lib "kernel32" (ByVal hFile As Long, lpModemStat As Long) As Long
Declare Function GetCommProperties Lib "kernel32" (ByVal hFile As Long, lpCommProp As COMMPROP) As Long
Declare Function GetCommState Lib "kernel32" (ByVal nCid As Long, lpDCB As DCB) As Long
Declare Function GetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Declare Function GetCompressedFileSize Lib "kernel32" Alias "GetCompressedFileSizeA" (ByVal lpFileName As String, lpFileSizeHigh As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetConsoleCP Lib "kernel32" () As Long
Declare Function GetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Declare Function GetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, lpMode As Long) As Long
Declare Function GetConsoleOutputCP Lib "kernel32" () As Long
Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
Declare Function GetCPInfo Lib "kernel32" (ByVal CodePage As Long, lpCPInfo As CPINFO) As Long
Declare Function GetCurrencyFormat Lib "kernel32" Alias "GetCurrencyFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As String, lpFormat As CURRENCYFMT, ByVal lpCurrencyStr As String, ByVal cchCurrency As Long) As Long
Declare Function GetCurrentDirectory Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function GetCurrentThread Lib "kernel32" () As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function GetCursor Lib "user32" () As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SystemTime, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDCEx Lib "user32" (ByVal hwnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Declare Function GetDCOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Declare Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, lpdwSize As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function GetDeviceGammaRamp Lib "gdi32" (ByVal hdc As Long, lpv As Any) As Long
Declare Function GetDialogBaseUnits Lib "user32" () As Long
Declare Function GetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQUAD) As Long
Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Declare Function GetDlgItemInt Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpTranslated As Long, ByVal bSigned As Long) As Long
Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function GetDoubleClickTime Lib "user32" () As Long
Declare Function GetDriverModuleHandle Lib "winmm.dll" (ByVal hDriver As Long) As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Declare Function GetEnhMetaFile Lib "gdi32" Alias "GetEnhMetaFileA" (ByVal lpszMetaFile As String) As Long
Declare Function GetEnhMetaFileBits Lib "gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpbBuffer As Byte) As Long
Declare Function GetEnhMetaFileDescription Lib "gdi32" Alias "GetEnhMetaFileDescriptionA" (ByVal hemf As Long, ByVal cchBuffer As Long, ByVal lpszDescription As String) As Long
Declare Function GetEnhMetaFileHeader Lib "gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpemh As ENHMETAHEADER) As Long
Declare Function GetEnhMetaFilePaletteEntries Lib "gdi32" (ByVal hemf As Long, ByVal cEntries As Long, lppe As PALETTEENTRY) As Long
Declare Function GetEnvironmentStrings Lib "kernel32" Alias "GetEnvironmentStringsA" () As String
Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Declare Function GetExpandedName Lib "lz32.dll" Alias "GetExpandedNameA" (ByVal lpszSource As String, ByVal lpszBuffer As String) As Long
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Declare Function GetFileSecurity Lib "ADVAPI32.DLL" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long
Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Declare Function GetFocus Lib "user32" () As Long
Declare Function GetFontData Lib "gdi32" Alias "GetFontDataA" (ByVal hdc As Long, ByVal dwTable As Long, ByVal dwOffset As Long, lpvBuffer As Any, ByVal cbData As Long) As Long
Declare Function GetFontLanguageInfo Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function GetForm Lib "winspool.drv" Alias "GetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Declare Function GetGlyphOutline Lib "gdi32" Alias "GetGlyphOutlineA" (ByVal hdc As Long, ByVal uChar As Long, ByVal fuFormat As Long, lpgm As GLYPHMETRICS, ByVal cbBuffer As Long, lpBuffer As Any, lpmat2 As MAT2) As Long
Declare Function GetGraphicsMode Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetHandleInformation Lib "kernel32" (ByVal hObject As Long, lpdwFlags As Long) As Long
Declare Function GetICMProfile Lib "gdi32" Alias "GetICMProfileA" (ByVal hdc As Long, ByVal dw As Long, ByVal lpStr As String) As Long
Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Declare Function GetInputState Lib "user32" () As Long
Declare Function GetJob Lib "winspool.drv" Alias "GetJobA" (ByVal hPrinter As Long, ByVal JobId As Long, ByVal Level As Long, pJob As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Declare Function GetKBCodePage Lib "user32" () As Long
Declare Function GetKernelObjectSecurity Lib "ADVAPI32.DLL" (ByVal handle As Long, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Declare Function GetKerningPairs Lib "gdi32" Alias "GetKerningPairsA" (ByVal hdc As Long, ByVal cPairs As Long, lpkrnpair As KERNINGPAIR) As Long
Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function GetLargestConsoleWindowSize Lib "kernel32" (ByVal hConsoleOutput As Long) As COORD
Declare Function GetLastActivePopup Lib "user32" (ByVal hwndOwnder As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function GetLengthSid Lib "ADVAPI32.DLL" (pSid As Any) As Long
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Function GetLogColorSpace Lib "gdi32" Alias "GetLogColorSpaceA" (ByVal hcolorspace As Long, ByVal lplogcolorspace As LOGCOLORSPACE, ByVal dw As Long) As Long
Declare Function GetLogicalDrives Lib "kernel32" () As Long
Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Declare Function GetMessageExtraInfo Lib "user32" () As Long
Declare Function GetMessagePos Lib "user32" () As Long
Declare Function GetMessageTime Lib "user32" () As Long
Declare Function GetMetaFile Lib "gdi32" Alias "GetMetaFileA" (ByVal lpFileName As String) As Long
Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hMF As Long, ByVal nSize As Long, lpvData As Any) As Long
Declare Function GetMetaRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function GetMiterLimit Lib "gdi32" (ByVal hdc As Long, peLimit As Double) As Long
Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function GetNamedPipeHandleState Lib "kernel32" Alias "GetNamedPipeHandleStateA" (ByVal hNamedPipe As Long, lpState As Long, lpCurInstances As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long, ByVal lpUserName As String, ByVal nMaxUserNameSize As Long) As Long
Declare Function GetNamedPipeInfo Lib "kernel32" (ByVal hNamedPipe As Long, lpFlags As Long, lpOutBufferSize As Long, lpInBufferSize As Long, lpMaxInstances As Long) As Long
Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Declare Function GetNextDlgGroupItem Lib "user32" (ByVal hDlg As Long, ByVal hCtl As Long, ByVal bPrevious As Long) As Long
Declare Function GetNextDlgTabItem Lib "user32" (ByVal hDlg As Long, ByVal hCtl As Long, ByVal bPrevious As Long) As Long
Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Declare Function GetNumberFormat Lib "kernel32" Alias "GetNumberFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As String, lpFormat As NUMBERFMT, ByVal lpNumberStr As String, ByVal cchNumber As Long) As Long
Declare Function GetNumberOfConsoleInputEvents Lib "kernel32" (ByVal hConsoleInput As Long, lpNumberOfEvents As Long) As Long
Declare Function GetNumberOfConsoleMouseButtons Lib "kernel32" (lpNumberOfMouseButtons As Long) As Long
Declare Function GetNumberOfEventLogRecords Lib "ADVAPI32.DLL" (ByVal hEventLog As Long, NumberOfRecords As Long) As Long
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Declare Function GetOEMCP Lib "kernel32" () As Long
Declare Function GetOldestEventLogRecord Lib "ADVAPI32.DLL" (ByVal hEventLog As Long, OldestRecord As Long) As Long
Declare Function GetOpenClipboardWindow Lib "user32" () As Long
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetOutlineTextMetrics Lib "gdi32" Alias "GetOutlineTextMetricsA" (ByVal hdc As Long, ByVal cbData As Long, lpotm As OUTLINETEXTMETRIC) As Long
Declare Function GetOverlappedResult Lib "kernel32" (ByVal hFile As Long, lpOverlapped As OVERLAPPED, lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
Declare Function GetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetPath Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpTypes As Byte, ByVal nSize As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetPixelFormat Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetPolyFillMode Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Declare Function GetPrinterData Lib "winspool.drv" Alias "GetPrinterDataA" (ByVal hPrinter As Long, ByVal pValueName As String, pType As Long, pData As Byte, ByVal nSize As Long, pcbNeeded As Long) As Long
Declare Function GetPrinterDriver Lib "winspool.drv" Alias "GetPrinterDriverA" (ByVal hPrinter As Long, ByVal pEnvironment As String, ByVal Level As Long, pDriverInfo As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Declare Function GetPrinterDriverDirectory Lib "winspool.drv" Alias "GetPrinterDriverDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverDirectory As Byte, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Declare Function GetPrintProcessorDirectory Lib "winspool.drv" Alias "GetPrintProcessorDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, ByVal pPrintProcessorInfo As String, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Declare Function GetPriorityClipboardFormat Lib "user32" (lpPriorityList As Long, ByVal nCount As Long) As Long
Declare Function GetPrivateObjectSecurity Lib "ADVAPI32.DLL" (ObjectDescriptor As SECURITY_DESCRIPTOR, ByVal SecurityInformation As Long, ResultantDescriptor As SECURITY_DESCRIPTOR, ByVal DescriptorLength As Long, ReturnLength As Long) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Declare Function GetProcessAffinityMask Lib "kernel32" (ByVal hProcess As Long, lpProcessAffinityMask As Long, SystemAffinityMask As Long) As Long
Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function GetProcessHeaps Lib "kernel32" (ByVal NumberOfHeaps As Long, ProcessHeaps As Long) As Long
Declare Function GetProcessShutdownParameters Lib "kernel32" (lpdwLevel As Long, lpdwFlags As Long) As Long
Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Declare Function GetProcessWindowStation Lib "user32" () As Long
Declare Function GetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, lpMinimumWorkingSetSize As Long, lpMaximumWorkingSetSize As Long) As Long
Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Declare Function GetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function GetQueuedCompletionStatus Lib "kernel32" (ByVal CompletionPort As Long, lpNumberOfBytesTransferred As Long, lpCompletionKey As Long, lpOverlapped As Long, ByVal dwMilliseconds As Long) As Long
Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long
Declare Function GetRasterizerCaps Lib "gdi32" (lpraststat As RASTERIZER_STATUS, ByVal cb As Long) As Long
Declare Function GetRegionData Lib "gdi32" Alias "GetRegionDataA" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As RgnData) As Long
Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Declare Function GetROP2 Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Declare Function GetSecurityDescriptorControl Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pControl As Integer, lpdwRevision As Long) As Long
Declare Function GetSecurityDescriptorDacl Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, lpbDaclPresent As Long, pDacl As ACL, lpbDaclDefaulted As Long) As Long
Declare Function GetSecurityDescriptorGroup Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pGroup As Any, ByVal lpbGroupDefaulted As Long) As Long
Declare Function GetSecurityDescriptorLength Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function GetSecurityDescriptorOwner Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal lpbOwnerDefaulted As Long) As Long
Declare Function GetSecurityDescriptorSacl Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal lpbSaclPresent As Long, pSacl As ACL, ByVal lpbSaclDefaulted As Long) As Long
Declare Function GetServiceDisplayName Lib "ADVAPI32.DLL" Alias "GetServiceDisplayNameA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, lpcchBuffer As Long) As Long
Declare Function GetServiceKeyName Lib "ADVAPI32.DLL" Alias "GetServiceKeyNameA" (ByVal hSCManager As Long, ByVal lpDisplayName As String, ByVal lpServiceName As String, lpcchBuffer As Long) As Long
Declare Function GetShortPathName Lib "kernel32" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function GetSidIdentifierAuthority Lib "ADVAPI32.DLL" (pSid As Any) As SID_IDENTIFIER_AUTHORITY
Declare Function GetSidLengthRequired Lib "ADVAPI32.DLL" (ByVal nSubAuthorityCount As Byte) As Long
Declare Function GetSidSubAuthority Lib "ADVAPI32.DLL" (pSid As Any, ByVal nSubAuthority As Long) As Long
Declare Function GetSidSubAuthorityCount Lib "ADVAPI32.DLL" (pSid As Any) As Byte
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Declare Function GetStretchBltMode Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetStringTypeA Lib "kernel32" (ByVal lcid As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Long) As Long
Declare Function GetStringTypeEx Lib "kernel32" Alias "GetStringTypeExA" (ByVal Locale As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Integer) As Long
Declare Function GetStringTypeW Lib "kernel32" (ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Integer) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Declare Function GetSystemPaletteUse Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Declare Function GetSystemTimeAdjustment Lib "kernel32" (lpTimeAdjustment As Long, lpTimeIncrement As Long, lpTimeAdjustmentDisabled As Boolean) As Long
Declare Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long
Declare Function GetTapeParameters Lib "kernel32" (ByVal hDevice As Long, ByVal dwOperation As Long, lpdwSize As Long, lpTapeInformation As Any) As Long
Declare Function GetTapePosition Lib "kernel32" (ByVal hDevice As Long, ByVal dwPositionType As Long, lpdwPartition As Long, lpdwOffsetLow As Long, lpdwOffsetHigh As Long) As Long
Declare Function GetTapeStatus Lib "kernel32" (ByVal hDevice As Long) As Long
Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetTextAlign Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetTextCharacterExtra Lib "gdi32" Alias "GetTextCharacterExtraA" (ByVal hdc As Long) As Long
Declare Function GetTextCharset Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetTextCharsetInfo Lib "gdi32" (ByVal hdc As Long, lpSig As FONTSIGNATURE, ByVal dwFlags As Long) As Long
Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetTextExtentExPoint Lib "gdi32" Alias "GetTextExtentExPointA" (ByVal hdc As Long, ByVal lpszStr As String, ByVal cchString As Long, ByVal nMaxExtent As Long, lpnFit As Long, alpDx As Long, lpSize As Size) As Long
Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As Size) As Long
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Declare Function GetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
Declare Function GetThreadDesktop Lib "user32" (ByVal dwThread As Long) As Long
Declare Function GetThreadLocale Lib "kernel32" () As Long
Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Declare Function GetThreadSelectorEntry Lib "kernel32" (ByVal hThread As Long, ByVal dwSelector As Long, lpSelectorEntry As LDT_ENTRY) As Long
Declare Function GetThreadTimes Lib "kernel32" (ByVal hThread As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SystemTime, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Declare Function GetTokenInformation Lib "ADVAPI32.DLL" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetUpdateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function GetUpdateRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal fErase As Long) As Long
Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Declare Function GetUserName Lib "ADVAPI32.DLL" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetUserObjectInformation Lib "user32" Alias "GetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Declare Function GetUserObjectSecurity Lib "user32" (ByVal hObj As Long, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Declare Function GetVersion Lib "kernel32" () As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByVal lpVersionInformation As OSVERSIONINFO) As Long
Declare Function GetViewportExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As Size) As Long
Declare Function GetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowContextHelpId Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As Size) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Declare Function GetWinMetaFileBits Lib "gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal fnMapMode As Long, ByVal hdcRef As Long) As Long
Declare Function GetWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As xform) As Long
Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalCompact Lib "kernel32" (ByVal dwMinFree As Long) As Long
Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Declare Function GlobalFindAtom Lib "kernel32" Alias "GlobalFindAtomA" (ByVal lpString As String) As Integer
Declare Function GlobalFlags Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GlobalHandle Lib "kernel32" (wMem As Any) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnWire Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalWire Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GrayString Lib "user32" Alias "GrayStringA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpOutputFunc As Long, ByVal lpData As Long, ByVal nCount As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapCompact Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long) As Long
Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Declare Function HeapLock Lib "kernel32" (ByVal hHeap As Long) As Long
Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Declare Function HeapUnlock Lib "kernel32" (ByVal hHeap As Long) As Long
Declare Function HeapValidate Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Declare Function HiliteMenuItem Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Declare Function hwrite Lib "kernel32" Alias "_hwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal lBytes As Long) As Long
Declare Function ImmAssociateContext Lib "imm32.dll" (ByVal hwnd As Long, ByVal himc As Long) As Long
Declare Function ImmConfigureIME Lib "imm32.dll" (ByVal hkl As Long, ByVal hwnd As Long, ByVal dw As Long) As Long
Declare Function ImmCreateContext Lib "imm32.dll" () As Long
Declare Function ImmDestroyContext Lib "imm32.dll" (ByVal himc As Long) As Long
Declare Function ImmEnumRegisterWord Lib "imm32.dll" Alias "ImmEnumRegisterWordA" (ByVal hkl As Long, ByVal RegisterWordEnumProc As Long, ByVal lpszReading As String, ByVal dw As Long, ByVal lpszRegister As String, lpv As Any) As Long
Declare Function ImmEscape Lib "imm32.dll" Alias "ImmEscapeA" (ByVal hkl As Long, ByVal himc As Long, ByVal un As Long, lpv As Any) As Long
Declare Function ImmGetCandidateList Lib "imm32.dll" Alias "ImmGetCandidateListA" (ByVal himc As Long, ByVal deIndex As Long, lpCandidateList As CANDIDATELIST, ByVal dwBufLen As Long) As Long
Declare Function ImmGetCandidateListCount Lib "imm32.dll" Alias "ImmGetCandidateListCountA" (ByVal himc As Long, lpdwListCount As Long) As Long
Declare Function ImmGetCandidateWindow Lib "imm32.dll" (ByVal himc As Long, ByVal dw As Long, lpCandidateForm As CANDIDATEFORM) As Long
Declare Function ImmGetCompositionFont Lib "imm32.dll" Alias "ImmGetCompositionFontA" (ByVal himc As Long, lpLogFont As LOGFONT) As Long
Declare Function ImmGetCompositionString Lib "imm32.dll" Alias "ImmGetCompositionStringA" (ByVal himc As Long, ByVal dw As Long, lpv As Any, ByVal dw2 As Long) As Long
Declare Function ImmGetCompositionWindow Lib "imm32.dll" (ByVal himc As Long, lpCompositionForm As COMPOSITIONFORM) As Long
Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Declare Function ImmGetConversionList Lib "imm32.dll" Alias "ImmGetConversionListA" (ByVal hkl As Long, ByVal himc As Long, ByVal lpsz As String, lpCandidateList As CANDIDATELIST, ByVal dwBufLen As Long, ByVal uFlag As Long) As Long
Declare Function ImmGetConversionStatus Lib "imm32.dll" (ByVal himc As Long, lpdw As Long, lpdw2 As Long) As Long
Declare Function ImmGetDefaultIMEWnd Lib "imm32.dll" (ByVal hwnd As Long) As Long
Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Declare Function ImmGetGuideLine Lib "imm32.dll" Alias " ImmGetGuideLineA" (ByVal himc As Long, ByVal dwIndex As Long, ByVal lpStr As String, ByVal dwBufLen As Long) As Long
Declare Function ImmGetIMEFileName Lib "imm32.dll" Alias "ImmGetIMEFileNameA" (ByVal hkl As Long, ByVal lpStr As String, ByVal uBufLen As Long) As Long
Declare Function ImmGetOpenStatus Lib "imm32.dll" (ByVal himc As Long) As Long
Declare Function ImmGetProperty Lib "imm32.dll" (ByVal hkl As Long, ByVal dw As Long) As Long
Declare Function ImmGetRegisterWordStyle Lib "imm32.dll" Alias " ImmGetRegisterWordStyleA" (ByVal hkl As Long, ByVal nItem As Long, lpStyleBuf As STYLEBUF) As Long
Declare Function ImmGetStatusWindowPos Lib "imm32.dll" (ByVal himc As Long, lpPoint As POINTAPI) As Long
Declare Function ImmGetVirtualKey Lib "imm32.dll" (ByVal hwnd As Long) As Long
Declare Function ImmInstallIME Lib "imm32.dll" Alias "ImmInstallIMEA" (ByVal lpszIMEFileName As String, ByVal lpszLayoutText As String) As Long
Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Declare Function ImmIsUIMessage Lib "imm32.dll" Alias "ImmIsUIMessageA" (ByVal hwnd As Long, ByVal un As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function ImmNotifyIME Lib "imm32.dll" (ByVal himc As Long, ByVal dwAction As Long, ByVal dwIndex As Long, ByVal dwValue As Long) As Long
Declare Function ImmRegisterWord Lib "imm32.dll" Alias "ImmRegisterWordA" (ByVal hkl As Long, ByVal lpszReading As String, ByVal dw As Long, ByVal lpszRegister As String) As Long
Declare Function ImmReleaseContext Lib "imm32.dll" (ByVal hwnd As Long, ByVal himc As Long) As Long
Declare Function ImmSetCandidateWindow Lib "imm32.dll" (ByVal himc As Long, lpCandidateForm As CANDIDATEFORM) As Long
Declare Function ImmSetCompositionFont Lib "imm32.dll" Alias "ImmSetCompositionFontA" (ByVal himc As Long, lpLogFont As LOGFONT) As Long
Declare Function ImmSetCompositionString Lib "imm32.dll" Alias "ImmSetCompositionStringA" (ByVal himc As Long, ByVal dwIndex As Long, lpComp As Any, ByVal dw As Long, lpRead As Any, ByVal dw2 As Long) As Long
Declare Function ImmSetCompositionWindow Lib "imm32.dll" (ByVal himc As Long, lpCompositionForm As COMPOSITIONFORM) As Long
Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal himc As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Declare Function ImmSetOpenStatus Lib "imm32.dll" (ByVal himc As Long, ByVal b As Long) As Long
Declare Function ImmSetStatusWindowPos Lib "imm32.dll" (ByVal himc As Long, lpPoint As POINTAPI) As Long
Declare Function ImmSimulateHotKey Lib "imm32.dll" (ByVal hwnd As Long, ByVal dw As Long) As Long
Declare Function ImmUnregisterWord Lib "imm32.dll" Alias "ImmUnregisterWordA" (ByVal hkl As Long, ByVal lpszReading As String, ByVal dw As Long, ByVal lpszUnregister As String) As Long
Declare Function ImpersonateDdeClientWindow Lib "user32" (ByVal hWndClient As Long, ByVal hWndServer As Long) As Long
Declare Function ImpersonateLoggedOnUser Lib "kernel32" (ByVal hToken As Long) As Long
Declare Function ImpersonateNamedPipeClient Lib "ADVAPI32.DLL" (ByVal hNamedPipe As Long) As Long
Declare Function ImpersonateSelf Lib "ADVAPI32.DLL" (Impersonationlevel As Integer) As Long
Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Declare Function InitAtomTable Lib "kernel32" (ByVal nSize As Long) As Long
Declare Function InitializeAcl Lib "ADVAPI32.DLL" (pAcl As ACL, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
Declare Function InitializeSecurityDescriptor Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Long) As Long
Declare Function InitializeSid Lib "ADVAPI32.DLL" (Sid As Any, pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte) As Long
Declare Function InitiateSystemShutdown Lib "ADVAPI32.DLL" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Declare Function InSendMessage Lib "user32" () As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByVal lpcMenuItemInfo As MENUITEMINFO) As Long
Declare Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
Declare Function InterlockedExchange Lib "kernel32" (Target As Long, ByVal Value As Long) As Long
Declare Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Declare Function IntersectClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function InvalidateRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bErase As Long) As Long
Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function InvertRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Declare Function IsBadHugeReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Declare Function IsBadHugeWritePtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Declare Function IsBadReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Declare Function IsBadStringPtr Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As String, ByVal ucchMax As Long) As Long
Declare Function IsBadWritePtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Declare Function IsChild Lib "user32" (ByVal hwndParent As Long, ByVal hwnd As Long) As Long
Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Declare Function IsDBCSLeadByte Lib "kernel32" (ByVal bTestChar As Byte) As Long
Declare Function IsDialogMessage Lib "user32" Alias "IsDialogMessageA" (ByVal hDlg As Long, lpMsg As msg) As Long
Declare Function IsDlgButtonChecked Lib "user32" (ByVal hDlg As Long, ByVal nIDButton As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long
Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Declare Function IsValidAcl Lib "ADVAPI32.DLL" (pAcl As ACL) As Long
Declare Function IsValidCodePage Lib "kernel32" (ByVal CodePage As Long) As Long
Declare Function IsValidLocale Lib "kernel32" (ByVal Locale As Long, ByVal dwFlags As Long) As Long
Declare Function IsValidSecurityDescriptor Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function IsValidSid Lib "ADVAPI32.DLL" (pSid As Any) As Long
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowUnicode Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Declare Function joyGetNumDevs Lib "winmm.dll" Alias "joyGetNumDev" () As Long
Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long
Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Declare Function joyGetThreshold Lib "winmm.dll" (ByVal id As Long, lpuThreshold As Long) As Long
Declare Function joyReleaseCapture Lib "winmm.dll" (ByVal id As Long) As Long
Declare Function joySetCapture Lib "winmm.dll" (ByVal hwnd As Long, ByVal uID As Long, ByVal uPeriod As Long, ByVal bChanged As Long) As Long
Declare Function joySetThreshold Lib "winmm.dll" (ByVal id As Long, ByVal uThreshold As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
Declare Function lcreat Lib "kernel32" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Long) As Long
Declare Function LineDDA Lib "gdi32" (ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal lpLineDDAProc As Long, ByVal lParam As Long) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function llseek Lib "kernel32" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function LoadAccelerators Lib "user32" Alias "LoadAcceleratorsA" (ByVal hInstance As Long, ByVal lpTableName As String) As Long
Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal Flags As Long) As Long
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As String) As Long
Declare Function LoadMenuIndirect Lib "user32" Alias "LoadMenuIndirectA" (ByVal lpMenuTemplate As Long) As Long
Declare Function LoadModule Lib "kernel32" (ByVal lpModuleName As String, lpParameterBlock As Any) As Long
Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Declare Function LocalCompact Lib "kernel32" (ByVal uMinFree As Long) As Long
Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Declare Function LocalFlags Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalHandle Lib "kernel32" (wMem As Any) As Long
Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal wBytes As Long, ByVal wFlags As Long) As Long
Declare Function LocalShrink Lib "kernel32" (ByVal hMem As Long, ByVal cbNewSize As Long) As Long
Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Declare Function LockFileEx Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long, ByVal dwReserved As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Declare Function LockServiceDatabase Lib "ADVAPI32.DLL" (ByVal hSCManager As Long) As Long
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function LogonUser Lib "ADVAPI32.DLL" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Declare Function LookupAccountName Lib "ADVAPI32.DLL" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, Sid As Long, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Declare Function LookupAccountSid Lib "ADVAPI32.DLL" Alias "LookupAccountSidA" (ByVal lpSystemName As String, Sid As Any, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Declare Function LookupIconIdFromDirectory Lib "user32" (presbits As Byte, ByVal fIcon As Long) As Long
Declare Function LookupIconIdFromDirectoryEx Lib "user32" (presbits As Byte, ByVal fIcon As Boolean, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Declare Function LookupPrivilegeDisplayName Lib "ADVAPI32.DLL" Alias "LookupPrivilegeDisplayNameA" (ByVal lpSystemName As String, ByVal lpName As String, ByVal lpDisplayName As String, cbDisplayName As Long, lpLanguageID As Long) As Long
Declare Function LookupPrivilegeName Lib "ADVAPI32.DLL" Alias "LookupPrivilegeNameA" (ByVal lpSystemName As String, lpLuid As LARGE_INTEGER, ByVal lpName As String, cbName As Long) As Long
Declare Function LookupPrivilegeValue Lib "ADVAPI32.DLL" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Declare Function LPtoDP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function lread Lib "kernel32" Alias "_lread" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcmp Lib "kernel32" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Declare Function lwrite Lib "kernel32" Alias "_lwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal wBytes As Long) As Long
Declare Function LZCopy Lib "lz32.dll" (ByVal hfSource As Long, ByVal hfDest As Long) As Long
Declare Function LZInit Lib "lz32.dll" (ByVal hfSrc As Long) As Long
Declare Function LZOpenFile Lib "lz32.dll" Alias "LZOpenFileA" (ByVal lpszFile As String, lpOf As OFSTRUCT, ByVal style As Long) As Long
Declare Function LZRead Lib "lz32.dll" (ByVal hfFile As Long, ByVal lpvBuf As String, ByVal cbread As Long) As Long
Declare Function LZSeek Lib "lz32.dll" (ByVal hfFile As Long, ByVal lOffset As Long, ByVal nOrigin As Long) As Long
Declare Function LZStart Lib "lz32" () As Long
Declare Function MakeAbsoluteSD Lib "ADVAPI32.DLL" (pSelfRelativeSecurityDescriptor As SECURITY_DESCRIPTOR, pAbsoluteSecurityDescriptor As SECURITY_DESCRIPTOR, lpdwAbsoluteSecurityDescriptorSize As Long, pDacl As ACL, lpdwDaclSize As Long, pSacl As ACL, lpdwSaclSize As Long, pOwner As Any, lpdwOwnerSize As Long, pPrimaryGroup As Any, lpdwPrimaryGroupSize As Long) As Long
Declare Function MakeSelfRelativeSD Lib "ADVAPI32.DLL" (pAbsoluteSecurityDescriptor As SECURITY_DESCRIPTOR, pSelfRelativeSecurityDescriptor As SECURITY_DESCRIPTOR, lpdwBufferLength As Long) As Long
Declare Function MapDialogRect Lib "user32" (ByVal hDlg As Long, lpRect As RECT) As Long
Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Declare Function MapViewOfFileEx Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long, lpBaseAddress As Any) As Long
Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Declare Function MapVirtualKeyEx Lib "user32" Alias "MapVirtualKeyExA" (ByVal uCode As Long, ByVal uMapType As Long, ByVal dwhkl As Long) As Long
Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Declare Function MaskBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long, ByVal dwRop As Long) As Long
Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Declare Function mciGetCreatorTask Lib "winmm.dll" (ByVal wDeviceID As Long) As Long
Declare Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long
Declare Function mciGetDeviceIDFromElementID Lib "winmm.dll" Alias "mciGetDeviceIDFromElementIDA" (ByVal dwElementID As Long, ByVal lpstrType As String) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciGetYieldProc Lib "winmm" (ByVal mciId As Long, pdwYieldData As Long) As Long
Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciSetYieldProc Lib "winmm" (ByVal mciId As Long, ByVal fpYieldProc As Long, ByVal dwYieldData As Long) As Long
Declare Function MenuItemFromPoint Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal ptScreen As POINTAPI) As Long
Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Declare Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
Declare Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectA" (lpMsgBoxParams As MSGBOXPARAMS) As Long
Declare Function midiConnect Lib "winmm.dll" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
Declare Function midiDisconnect Lib "winmm.dll" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
Declare Function midiInAddBuffer Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIINCAPS, ByVal uSize As Long) As Long
Declare Function midiInGetErrorText Lib "winmm.dll" Alias "midiInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function midiInGetID Lib "winmm.dll" (ByVal hMidiIn As Long, lpuDeviceID As Long) As Long
Declare Function midiInGetNumDevs Lib "winmm.dll" () As Long
Declare Function midiInMessage Lib "winmm.dll" (ByVal hMidiIn As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Declare Function midiInOpen Lib "winmm.dll" (lphMidiIn As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiInPrepareHeader Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiInReset Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInStart Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInUnprepareHeader Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiOutCacheDrumPatches Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal uPatch As Long, lpKeyArray As Long, ByVal uFlags As Long) As Long
Declare Function midiOutCachePatches Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal uBank As Long, lpPatchArray As Long, ByVal uFlags As Long) As Long
Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function midiOutGetID Lib "winmm.dll" (ByVal hMidiOut As Long, lpuDeviceID As Long) As Long
Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Declare Function midiOutLongMsg Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiOutMessage Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiOutPrepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiOutReset Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Declare Function midiOutUnprepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiStreamClose Lib "winmm.dll" (ByVal hms As Long) As Long
Declare Function midiStreamOpen Lib "winmm.dll" (phms As Long, puDeviceID As Long, ByVal cMidi As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function midiStreamOut Lib "winmm.dll" (ByVal hms As Long, pmh As MIDIHDR, ByVal cbmh As Long) As Long
Declare Function midiStreamPause Lib "winmm.dll" (ByVal hms As Long) As Long
Declare Function midiStreamPosition Lib "winmm.dll" (ByVal hms As Long, lpmmt As MMTIME, ByVal cbmmt As Long) As Long
Declare Function midiStreamProperty Lib "winmm.dll" (ByVal hms As Long, lppropdata As Byte, ByVal dwProperty As Long) As Long
Declare Function midiStreamRestart Lib "winmm.dll" (ByVal hms As Long) As Long
Declare Function midiStreamStop Lib "winmm.dll" (ByVal hms As Long) As Long
Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As Long, ByVal pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, pumxID As Long, ByVal fdwId As Long) As Long
Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, ByVal uMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Declare Function mmioAdvance Lib "winmm.dll" (ByVal hmmio As Long, lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioCreateChunk Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioFlush Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioGetInfo Lib "winmm.dll" (ByVal hmmio As Long, lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Declare Function mmioInstallIOProcA Lib "winmm" (ByVal fccIOProc As String, ByVal pIOProc As Long, ByVal dwFlags As Long) As Long
Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As MMIOINFO, ByVal dwOpenFlags As Long) As Long
Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Declare Function mmioRename Lib "winmm.dll" Alias "mmioRenameA" (ByVal szFileName As String, ByVal SzNewFileName As String, lpmmioinfo As MMIOINFO, ByVal dwRenameFlags As Long) As Long
Declare Function mmioSeek Lib "winmm.dll" (ByVal hmmio As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function mmioSendMessage Lib "winmm.dll" (ByVal hmmio As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Declare Function mmioSetBuffer Lib "winmm.dll" (ByVal hmmio As Long, ByVal pchBuffer As String, ByVal cchBuffer As Long, ByVal uFlags As Long) As Long
Declare Function mmioSetInfo Lib "winmm.dll" (ByVal hmmio As Long, lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Declare Function mmioWrite Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Declare Function mmsystemGetVersion Lib "winmm.dll" () As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Declare Function ModifyWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As xform, ByVal iMode As Long) As Long
Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Declare Function NotifyBootConfigStatus Lib "ADVAPI32.DLL" (ByVal BootAcceptable As Long) As Long
Declare Function NotifyChangeEventLog Lib "advapi32" (ByVal hEventLog As Long, ByVal hEvent As Long) As Long
Declare Function ObjectCloseAuditAlarm Lib "ADVAPI32.DLL" Alias "ObjectCloseAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal GenerateOnClose As Long) As Long
Declare Function ObjectOpenAuditAlarm Lib "kernel32" Alias "ObjectOpenAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As Long, ByVal DesiredAccess As Long, ByVal GrantedAccess As Long, Privileges As PRIVILEGE_SET, ByVal ObjectCreation As Long, ByVal AccessGranted As Long, ByVal GenerateOnClose As Long) As Long
Declare Function ObjectPrivilegeAuditAlarm Lib "ADVAPI32.DLL" Alias "ObjectPrivilegeAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ClientToken As Long, ByVal DesiredAccess As Long, Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) As Long
Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Long) As Long
Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Declare Function OffsetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function OffsetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Declare Function OffsetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Declare Function OpenBackupEventLog Lib "ADVAPI32.DLL" Alias "OpenBackupEventLogA" (ByVal lpUNCServerName As String, ByVal lpFileName As String) As Long
Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Declare Function OpenDesktop Lib "user32" Alias "OpenDesktopA" (ByVal lpszDesktop As String, ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenDriver Lib "winmm.dll" (ByVal szDriverName As String, ByVal szSectionName As String, ByVal lParam2 As Long) As Long
Declare Function OpenEvent Lib "kernel32" Alias "OpenEventA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function OpenEventLog Lib "ADVAPI32.DLL" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function OpenInputDesktop Lib "user32" (ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function OpenProcessToken Lib "ADVAPI32.DLL" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Declare Function OpenSCManager Lib "ADVAPI32.DLL" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenSemaphore Lib "kernel32" Alias "OpenSemaphoreA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function OpenService Lib "ADVAPI32.DLL" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenThreadToken Lib "ADVAPI32.DLL" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
Declare Function OpenWindowStation Lib "user32" Alias "OpenWindowStationA" (ByVal lpszWinSta As String, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Declare Function PackDDElParam Lib "user32" (ByVal msg As Long, ByVal uiLo As Long, ByVal uiHi As Long) As Long
Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PageSetupDlg) As Long
Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long
Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hemf As Long, lpRect As RECT) As Long
Declare Function PlayEnhMetaFileRecord Lib "gdi32" (ByVal hdc As Long, lpHandletable As HANDLETABLE, lpEnhMetaRecord As ENHMETARECORD, ByVal nHandles As Long) As Long
Declare Function PlayMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hMF As Long) As Long
Declare Function PlayMetaFileRecord Lib "gdi32" (ByVal hdc As Long, lpHandletable As HANDLETABLE, lpMetaRecord As METARECORD, ByVal nHandles As Long) As Long
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long
Declare Function PolyBezier Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Declare Function PolyBezierTo Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Declare Function PolyDraw Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long
Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function PolylineTo Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Declare Function PolyPolyline Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
Declare Function PolyTextOut Lib "gdi32" Alias "PolyTextOutA" (ByVal hdc As Long, pptxt As POLYTEXT, cStrings As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function PostThreadMessage Lib "user32" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function PrepareTape Lib "kernel32" (ByVal hDevice As Long, ByVal dwOperation As Long, ByVal bimmediate As Long) As Long
Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg) As Long
Declare Function PrinterMessageBox Lib "winspool.drv" Alias "PrinterMessageBoxA" (ByVal hPrinter As Long, ByVal error As Long, ByVal hwnd As Long, ByVal pText As String, ByVal pCaption As String, ByVal dwType As Long) As Long
Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
Declare Function PrivilegeCheck Lib "ADVAPI32.DLL" (ByVal ClientToken As Long, RequiredPrivileges As PRIVILEGE_SET, ByVal pfResult As Long) As Long
Declare Function PrivilegedServiceAuditAlarm Lib "ADVAPI32.DLL" Alias "PrivilegedServiceAuditAlarmA" (ByVal SubsystemName As String, ByVal ServiceName As String, ByVal ClientToken As Long, Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) As Long
Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function PtVisible Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function PulseEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Declare Function PurgeComm Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long) As Long
Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Declare Function QueryServiceConfig Lib "ADVAPI32.DLL" Alias "QueryServiceConfigA" (ByVal hService As Long, lpServiceConfig As QUERY_SERVICE_CONFIG, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Declare Function QueryServiceLockStatus Lib "ADVAPI32.DLL" Alias "QueryServiceLockStatusA" (ByVal hSCManager As Long, lpLockStatus As QUERY_SERVICE_LOCK_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Declare Function QueryServiceObjectSecurity Lib "ADVAPI32.DLL" (ByVal hService As Long, ByVal dwSecurityInformation As Long, lpSecurityDescriptor As Any, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Declare Function QueryServiceStatus Lib "ADVAPI32.DLL" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, lpBuffer As Any, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Declare Function ReadConsoleOutput Lib "kernel32" Alias "ReadConsoleOutputA" (ByVal hConsoleOutput As Long, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpReadRegion As SMALL_RECT) As Long
Declare Function ReadConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, lpAttribute As Long, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfAttrsRead As Long) As Long
Declare Function ReadConsoleOutputCharacter Lib "kernel32" Alias "ReadConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfCharsRead As Long) As Long
Declare Function ReadEventLog Lib "ADVAPI32.DLL" Alias "ReadEventLogA" (ByVal hEventLog As Long, ByVal dwReadFlags As Long, ByVal dwRecordOffset As Long, lpBuffer As EVENTLOGRECORD, ByVal nNumberOfBytesToRead As Long, pnBytesRead As Long, pnMinNumberOfBytesNeeded As Long) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function ReadFileEx Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long
Declare Function ReadPrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pNoBytesRead As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Declare Function RectVisible Lib "gdi32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function RegCloseKey Lib "ADVAPI32.DLL" (ByVal hKey As Long) As Long
Declare Function RegConnectRegistry Lib "ADVAPI32.DLL" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Declare Function RegCreateKey Lib "ADVAPI32.DLL" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Declare Function RegCreateKeyEx Lib "ADVAPI32.DLL" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
'Declare Function RegDeleteKey Lib "ADVAPI32.DLL" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'Declare Function RegDeleteValue Lib "ADVAPI32.DLL" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'Declare Function RegEnumKey Lib "ADVAPI32.DLL" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
'Declare Function RegEnumKeyEx Lib "ADVAPI32.DLL" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
'Declare Function RegEnumValue Lib "ADVAPI32.DLL" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
'Declare Function RegFlushKey Lib "ADVAPI32.DLL" (ByVal hKey As Long) As Long
Declare Function RegGetKeySecurity Lib "ADVAPI32.DLL" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long
Declare Function RegisterClass Lib "user32" (Class As WNDCLASS) As Long
'Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
'Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
'Declare Function RegisterEventSource Lib "ADVAPI32.DLL" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
'Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
'Declare Function RegisterServiceCtrlHandler Lib "ADVAPI32.DLL" Alias "RegisterServiceCtrlHandlerA" (ByVal lpServiceName As String, ByVal lpHandlerProc As Long) As Long
'Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
'Declare Function RegLoadKey Lib "ADVAPI32.DLL" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
'Declare Function RegNotifyChangeKeyValue Lib "ADVAPI32.DLL" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
'Declare Function RegOpenKey Lib "ADVAPI32.DLL" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Declare Function RegOpenKeyEx Lib "ADVAPI32.DLL" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Declare Function RegQueryInfoKey Lib "ADVAPI32.DLL" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
'Declare Function RegQueryValue Lib "ADVAPI32.DLL" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
'Declare Function RegQueryValueEx Lib "ADVAPI32.DLL" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
'Declare Function RegReplaceKey Lib "ADVAPI32.DLL" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
'Declare Function RegRestoreKey Lib "ADVAPI32.DLL" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
'Declare Function RegSaveKey Lib "ADVAPI32.DLL" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
'Declare Function RegSetKeySecurity Lib "ADVAPI32.DLL" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
'Declare Function RegSetValue Lib "ADVAPI32.DLL" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
'Declare Function RegSetValueEx Lib "ADVAPI32.DLL" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
'Declare Function RegUnLoadKey Lib "ADVAPI32.DLL" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'Declare Function ReleaseCapture Lib "user32" () As Long
'Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
'Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
'Declare Function ReleaseSemaphore Lib "kernel32" (ByVal hSemaphore As Long, ByVal lReleaseCount As Long, lpPreviousCount As Long) As Long
'Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
'Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
'Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
'Declare Function ReplaceText Lib "comdlg32.dll" Alias "ReplaceTextA" (pFindreplace As FINDREPLACE) As Long
'Declare Function ReplyMessage Lib "user32" (ByVal lReply As Long) As Long
'Declare Function ReportEvent Lib "ADVAPI32.DLL" Alias "ReportEventA" (ByVal hEventLog As Long, ByVal wType As Long, ByVal wCategory As Long, ByVal dwEventID As Long, lpUserSid As Any, ByVal wNumStrings As Long, ByVal dwDataSize As Long, ByVal lpStrings As Long, lpRawData As Any) As Long
'Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hdc As Long, lpInitData As DEVMODE) As Long
'Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
'Declare Function ResetPrinter Lib "winspool.drv" Alias "ResetPrinterA" (ByVal hPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
'Declare Function ResizePalette Lib "gdi32" (ByVal hPalette As Long, ByVal nNumEntries As Long) As Long
'Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
'Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
'Declare Function ReuseDDElParam Lib "user32" (ByVal lParam As Long, ByVal msgIn As Long, ByVal msgOut As Long, ByVal uiLo As Long, ByVal uiHi As Long) As Long
'Declare Function RevertToSelf Lib "ADVAPI32.DLL" () As Long
'Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function ScaleViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As Size) As Long
'Declare Function ScaleWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As Size) As Long
'Declare Function ScheduleJob Lib "winspool.drv" (ByVal hPrinter As Long, ByVal JobId As Long) As Long
'Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'Declare Function ScrollConsoleScreenBuffer Lib "kernel32" Alias "ScrollConsoleScreenBufferA" (ByVal hConsoleOutput As Long, lpScrollRectangle As SMALL_RECT, lpClipRectangle As SMALL_RECT, dwDestinationOrigin As COORD, lpFill As CHAR_INFO) As Long
'Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
'Declare Function ScrollWindow Lib "user32" (ByVal hwnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long
'Declare Function ScrollWindowEx Lib "user32" (ByVal hwnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT, ByVal fuScroll As Long) As Long
'Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
'Declare Function SelectClipPath Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
'Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
'Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
'Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Declare Function SendDriverMessage Lib "winmm.dll" (ByVal hDriver As Long, ByVal message As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
'Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Declare Function SendMessageCallback Lib "user32" Alias "SendMessageCallbackA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lpResultCallBack As Long, ByVal dwData As Long) As Long
'Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
'Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Declare Function SetAbortProc Lib "gdi32" (ByVal hdc As Long, ByVal lpAbortProc As Long) As Long
'Declare Function SetAclInformation Lib "ADVAPI32.DLL" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
'Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function SetArcDirection Lib "gdi32" (ByVal hdc As Long, ByVal ArcDirection As Long) As Long
'Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
'Declare Function SetBitmapDimensionEx Lib "gdi32" (ByVal hbm As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
'Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
'Declare Function SetBoundsRect Lib "gdi32" (ByVal hdc As Long, lprcBounds As RECT, ByVal Flags As Long) As Long
'Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
'Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long
'Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Declare Function SetClassWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
'Declare Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As Long) As Long
'Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function SetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As ColorAdjustment) As Long
'Declare Function SetColorSpace Lib "gdi32" (ByVal hdc As Long, ByVal hcolorspace As Long) As Long
'Declare Function SetCommBreak Lib "kernel32" (ByVal nCid As Long) As Long
'Declare Function SetCommConfig Lib "kernel32" (ByVal hCommDev As Long, lpCC As COMMCONFIG, ByVal dwSize As Long) As Long
'Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
'Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As DCB) As Long
'Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
'Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
'Declare Function SetConsoleActiveScreenBuffer Lib "kernel32" (ByVal hConsoleOutput As Long) As Long
'Declare Function SetConsoleCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
'Declare Function SetConsoleCtrlHandler Lib "kernel32" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long
'Declare Function SetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
'Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, dwCursorPosition As COORD) As Long
'Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, ByVal dwMode As Long) As Long
'Declare Function SetConsoleOutputCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
'Declare Function SetConsoleScreenBufferSize Lib "kernel32" (ByVal hConsoleOutput As Long, dwSize As COORD) As Long
'Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
'Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
'Declare Function SetConsoleWindowInfo Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal bAbsolute As Long, lpConsoleWindow As SMALL_RECT) As Long
'Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
'Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'Declare Function SetDefaultCommConfig Lib "kernel32" Alias "SetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, ByVal dwSize As Long) As Long
'Declare Function SetDeviceGammaRamp Lib "gdi32" (ByVal hdc As Long, lpv As Any) As Long
'Declare Function SetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
'Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
'Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
'Declare Function SetDlgItemInt Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wValue As Long, ByVal bSigned As Long) As Long
'Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
'Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
'Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
'Declare Function SetEnhMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpData As Byte) As Long
'Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
'Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
'Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
'Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
'Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
'Declare Function SetFileSecurity Lib "ADVAPI32.DLL" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
'Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
'Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function SetForm Lib "winspool.drv" Alias "SetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte) As Long
'Declare Function SetGraphicsMode Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
'Declare Function SetHandleCount Lib "kernel32" (ByVal wNumber As Long) As Long
'Declare Function SetHandleInformation Lib "kernel32" (ByVal hObject As Long, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
'Declare Function SetICMMode Lib "gdi32" (ByVal hdc As Long, ByVal n As Long) As Long
'Declare Function SetICMProfile Lib "gdi32" Alias "SetICMProfileA" (ByVal hdc As Long, ByVal lpStr As String) As Long
'Declare Function SetJob Lib "winspool.drv" Alias "SetJobA" (ByVal hPrinter As Long, ByVal JobId As Long, ByVal Level As Long, pJob As Byte, ByVal Command As Long) As Long
'Declare Function SetKernelObjectSecurity Lib "ADVAPI32.DLL" (ByVal handle As Long, ByVal SecurityInformation As Long, SecurityDescriptor As SECURITY_DESCRIPTOR) As Long
'Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
'Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
'Declare Function SetLocalTime Lib "kernel32" (lpSystemTime As SystemTime) As Long
'Declare Function SetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, ByVal lReadTimeout As Long) As Long
'Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
'Declare Function SetMapperFlags Lib "gdi32" (ByVal hdc As Long, ByVal dwFlag As Long) As Long
'Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
'Declare Function SetMenuContextHelpId Lib "user32" (ByVal hMenu As Long, ByVal dw As Long) As Long
'Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
'Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
'Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
'Declare Function SetMessageExtraInfo Lib "user32" (ByVal lParam As Long) As Long
'Declare Function SetMessageQueue Lib "user32" (ByVal cMessagesMax As Long) As Long
'Declare Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Byte) As Long
'Declare Function SetMetaRgn Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function SetMiterLimit Lib "gdi32" (ByVal hdc As Long, ByVal eNewLimit As Double, peOldLimit As Double) As Long
'Declare Function SetNamedPipeHandleState Lib "kernel32" (ByVal hNamedPipe As Long, lpMode As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long) As Long
'Declare Function SetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
'Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'Declare Function SetPixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal n As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
'Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
'Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Byte, ByVal Command As Long) As Long
'Declare Function SetPrinterData Lib "winspool.drv" Alias "SetPrinterDataA" (ByVal hPrinter As Long, ByVal pValueName As String, ByVal dwType As Long, pData As Byte, ByVal cbData As Long) As Long
'Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
'Declare Function SetPrivateObjectSecurity Lib "ADVAPI32.DLL" (ByVal SecurityInformation As Long, ModificationDescriptor As SECURITY_DESCRIPTOR, ObjectsSecurityDescriptor As SECURITY_DESCRIPTOR, GenericMapping As GENERIC_MAPPING, ByVal Token As Long) As Long
'Declare Function SetProcessShutdownParameters Lib "kernel32" (ByVal dwLevel As Long, ByVal dwFlags As Long) As Long
'Declare Function SetProcessWindowStation Lib "user32" (ByVal hWinSta As Long) As Long
'Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
'Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
'Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
'Declare Function SetRectRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
'Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
'Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
'Declare Function SetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
'Declare Function SetSecurityDescriptorDacl Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bDaclPresent As Long, pDacl As ACL, ByVal bDaclDefaulted As Long) As Long
'Declare Function SetSecurityDescriptorGroup Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pGroup As Any, ByVal bGroupDefaulted As Long) As Long
'Declare Function SetSecurityDescriptorOwner Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal bOwnerDefaulted As Long) As Long
'Declare Function SetSecurityDescriptorSacl Lib "ADVAPI32.DLL" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bSaclPresent As Long, pSacl As ACL, ByVal bSaclDefaulted As Long) As Long
'Declare Function SetServiceBits Lib "advapi32" (ByVal hServiceStatus As Long, ByVal dwServiceBits As Long, ByVal bSetBitsOn As Boolean, ByVal bUpdateImmediately As Boolean) As Long
'Declare Function SetServiceObjectSecurity Lib "ADVAPI32.DLL" (ByVal hService As Long, ByVal dwSecurityInformation As Long, lpSecurityDescriptor As Any) As Long
'Declare Function SetServiceStatus Lib "ADVAPI32.DLL" (ByVal hServiceStatus As Long, lpServiceStatus As SERVICE_STATUS) As Long
'Declare Function SetStdHandle Lib "kernel32" (ByVal nStdHandle As Long, ByVal nHandle As Long) As Long
'Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
'Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
'Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Long
'Declare Function SetSystemPaletteUse Lib "gdi32" (ByVal hdc As Long, ByVal wUsage As Long) As Long
'Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long
'Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SystemTime) As Long
'Declare Function SetSystemTimeAdjustment Lib "kernel32" (ByVal dwTimeAdjustment As Long, ByVal bTimeAdjustmentDisabled As Boolean) As Long
'Declare Function SetTapeParameters Lib "kernel32" (ByVal hDevice As Long, ByVal dwOperation As Long, lpTapeInformation As Any) As Long
'Declare Function SetTapePosition Lib "kernel32" (ByVal hDevice As Long, ByVal dwPositionMethod As Long, ByVal dwPartition As Long, ByVal dwOffsetLow As Long, ByVal dwOffsetHigh As Long, ByVal bimmediate As Long) As Long
'Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
'Declare Function SetTextCharacterExtra Lib "gdi32" Alias "SetTextCharacterExtraA" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long
'Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'Declare Function SetTextJustification Lib "gdi32" (ByVal hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
'Declare Function SetThreadAffinityMask Lib "kernel32" (ByVal hThread As Long, ByVal dwThreadAffinityMask As Long) As Long
'Declare Function SetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
'Declare Function SetThreadDesktop Lib "user32" (ByVal hDesktop As Long) As Long
'Declare Function SetThreadLocale Lib "kernel32" (ByVal Locale As Long) As Long
'Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
'Declare Function SetThreadToken Lib "advapi32" (Thread As Long, ByVal Token As Long) As Long
'Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'Declare Function SetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
'Declare Function SetTokenInformation Lib "ADVAPI32.DLL" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long) As Long
'Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
'Declare Function SetupComm Lib "kernel32" (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
'Declare Function SetUserObjectInformation Lib "user32" Alias "SetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long) As Long
'Declare Function SetUserObjectSecurity Lib "user32" (ByVal hObj As Long, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR) As Long
'Declare Function SetViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
'Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
'Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
'Declare Function SetWindowContextHelpId Lib "user32" (ByVal hwnd As Long, ByVal dw As Long) As Long
'Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
'Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
'Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
'Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'Declare Function SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As Long) As Long
'Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
'Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
'Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
'Declare Function SetWinMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As Long, lpmfp As METAFILEPICT) As Long
'Declare Function SetWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As xform) As Long
'Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
'Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias " Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
'Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Declare Function SHFileOperation Lib "shell32.dll" Alias " SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
'Declare Function SHGetFileInfo Lib "shell32.dll" Alias " SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
'Declare Function SHGetNewLinkInfo Lib "shell32.dll" Alias "SHGetNewLinkInfoA" (ByVal pszLinkto As String, ByVal pszDir As String, ByVal pszName As String, pfMustCopy As Long, ByVal uFlags As Long) As Long
'Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
'Declare Function ShowOwnedPopups Lib "user32" (ByVal hwnd As Long, ByVal fShow As Long) As Long
'Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
'Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Declare Function ShowWindowAsync Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
'Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
'Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
'Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As Byte) As Long
'Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
'Declare Function StartService Lib "ADVAPI32.DLL" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
'Declare Function StartServiceCtrlDispatcher Lib "ADVAPI32.DLL" Alias "StartServiceCtrlDispatcherA" (lpServiceStartTable As SERVICE_TABLE_ENTRY) As Long
'Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
'Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function SubtractRect Lib "user32" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
'Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
'Declare Function SwapBuffers Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
'Declare Function SwitchDesktop Lib "user32" (ByVal hDesktop As Long) As Long
'Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SystemTime, lpFileTime As FILETIME) As Long
'Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SystemTime, lpLocalTime As SystemTime) As Long
'Declare Function TabbedTextOut Lib "user32" Alias "TabbedTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long
'Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
'Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'Declare Function TileWindows Lib "user32" (ByVal hwndParent As Long, ByVal wHow As Long, lpRect As RECT, ByVal cKids As Long, lpKids As Long) As Integer
'Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
'Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
'Declare Function timeGetDevCaps Lib "winmm.dll" (lpTimeCaps As TIMECAPS, ByVal uSize As Long) As Long
'Declare Function timeGetSystemTime Lib "winmm.dll" (lpTime As MMTIME, ByVal uSize As Long) As Long
'Declare Function timeGetTime Lib "winmm.dll" () As Long
'Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
'Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
'Declare Function TlsAlloc Lib "kernel32" () As Long
'Declare Function TlsFree Lib "kernel32" (ByVal dwTlsIndex As Long) As Long
'Declare Function TlsGetValue Lib "kernel32" (ByVal dwTlsIndex As Long) As Long
'Declare Function TlsSetValue Lib "kernel32" (ByVal dwTlsIndex As Long, lpTlsValue As Any) As Long
'Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
'Declare Function ToAsciiEx Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpKeyState As Byte, lpChar As Integer, ByVal uFlags As Long, ByVal dwhkl As Long) As Long
'Declare Function ToUnicode Lib "user32" (ByVal wVirtKey As Long, ByVal wScanCode As Long, lpKeyState As Byte, ByVal pwszBuff As String, ByVal cchBuff As Long, ByVal wFlags As Long) As Long
'Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
'Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hwnd As Long, lpTPMParams As TPMPARAMS) As Long
'Declare Function TransactNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
'Declare Function TranslateAccelerator Lib "user32" Alias "TranslateAcceleratorA" (ByVal hwnd As Long, ByVal hAccTable As Long, lpMsg As msg) As Long
'Declare Function TranslateCharsetInfo Lib "gdi32" (lpSrc As Long, lpcs As CHARSETINFO, ByVal dwFlags As Long) As Long
'Declare Function TranslateMDISysAccel Lib "user32" (ByVal hWndClient As Long, lpMsg As msg) As Long
'Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
'Declare Function TransmitCommChar Lib "kernel32" (ByVal nCid As Long, ByVal cChar As Byte) As Long
'Declare Function UnhandledExceptionFilter Lib "kernel32" (ExceptionInfo As EXCEPTION_POINTERS) As Long
'Declare Function UnhookWindowsHook Lib "user32" (ByVal nCode As Long, ByVal pfnFilterProc As Long) As Long
'Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
'Declare Function UnionRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
'Declare Function UnloadKeyboardLayout Lib "user32" (ByVal hkl As Long) As Long
'Declare Function UnlockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long
'Declare Function UnlockFileEx Lib "kernel32" (ByVal hFile As Long, ByVal dwReserved As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long, lpOverlapped As OVERLAPPED) As Long
'Declare Function UnlockServiceDatabase Lib "ADVAPI32.DLL" (ScLock As Any) As Long
'Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
'Declare Function UnpackDDElParam Lib "user32" (ByVal msg As Long, ByVal lParam As Long, puiLo As Long, puiHi As Long) As Long
'Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
'Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
'Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
'Declare Function UpdateColors Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
'Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function ValidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Declare Function ValidateRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
'Declare Function VerFindFile Lib "version.dll" Alias "VerFindFileA" (ByVal uFlags As Long, ByVal szFileName As String, ByVal szWinDir As String, ByVal szAppDir As String, ByVal szCurDir As String, lpuCurDirLen As Long, ByVal szDestDir As String, lpuDestDirLen As Long) As Long
'Declare Function VerInstallFile Lib "version.dll" Alias " VerInstallFileA" (ByVal uFlags As Long, ByVal szSrcFileName As String, ByVal szDestFileName As String, ByVal szSrcDir As String, ByVal szDestDir As String, ByVal szCurDir As String, ByVal szTmpFile As String, lpuTmpFileLen As Long) As Long
'Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
'Declare Function VerQueryValue Lib "version.dll" (pBlock As Any, ByVal lpSubBlock As String, ByVal lplpBuffer As Long, puLen As Long) As Long
'Declare Function VirtualAlloc Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
'Declare Function VirtualAllocEx Lib "kernel32.dll" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
'Declare Function VirtualFree Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
'Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Integer
'Declare Function VirtualLock Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long) As Long
'Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
'Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
'Declare Function VirtualQuery Lib "kernel32" (lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
'Declare Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
'Declare Function VirtualUnlock Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long) As Long
'Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
'Declare Function VkKeyScanEx Lib "user32" Alias "VkKeyScanExA" (ByVal ch As Byte, ByVal dwhkl As Long) As Integer
'Declare Function WaitCommEvent Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As OVERLAPPED) As Long
'Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
'Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
'Declare Function WaitForMultipleObjectsEx Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
'Declare Function WaitForPrinterChange Lib "winspool.drv" (ByVal hPrinter As Long, ByVal Flags As Long) As Long
'Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'Declare Function WaitForSingleObjectEx Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
'Declare Function WaitMessage Lib "user32" () As Long
'Declare Function WaitNamedPipe Lib "kernel32" Alias "WaitNamedPipeA" (ByVal lpNamedPipeName As String, ByVal nTimeOut As Long) As Long
'Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
'Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
'Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
'Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
'Declare Function waveInGetID Lib "winmm.dll" (ByVal hWaveIn As Long, lpuDeviceID As Long) As Long
'Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
'Declare Function waveInGetPosition Lib "winmm.dll" (ByVal hWaveIn As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
'Declare Function waveInMessage Lib "winmm.dll" (ByVal hWaveIn As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
'Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
'Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
'Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
'Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
'Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
'Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
'Declare Function waveOutBreakLoop Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
'Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
'Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long
'Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
'Declare Function waveOutGetID Lib "winmm.dll" (ByVal hWaveOut As Long, lpuDeviceID As Long) As Long
'Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
'Declare Function waveOutGetPitch Lib "winmm.dll" (ByVal hWaveOut As Long, lpdwPitch As Long) As Long
'Declare Function waveOutGetPlaybackRate Lib "winmm.dll" (ByVal hWaveOut As Long, lpdwRate As Long) As Long
'Declare Function waveOutGetPosition Lib "winmm.dll" (ByVal hWaveOut As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
'Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
'Declare Function waveOutMessage Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
'Declare Function waveOutOpen Lib "winmm.dll" (lphWaveOut As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
'Declare Function waveOutPause Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
'Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
'Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
'Declare Function waveOutRestart Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
'Declare Function waveOutSetPitch Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal dwPitch As Long) As Long
'Declare Function waveOutSetPlaybackRate Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal dwRate As Long) As Long
'Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
'Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
'Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
'Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
'Declare Function WidenPath Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
'Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
'Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
'Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
'Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long
'Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
'Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
'Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
'Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
'Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As Any, lpBufferSize As Long) As Long
'Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
'Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
'Declare Function WNetGetUniversalName Lib "mpr" Alias "WNetGetUniversalNameA" (ByVal lpLocalPath As String, ByVal dwInfoLevel As Long, lpBuffer As Any, lpBufferSize As Long) As Long
'Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
'Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As NETRESOURCE, lphEnum As Long) As Long
'Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
'Declare Function WriteConsoleOutput Lib "kernel32" Alias "WriteConsoleOutputA" (ByVal hConsoleOutput As Long, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpWriteRegion As SMALL_RECT) As Long
'Declare Function WriteConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, lpAttribute As Integer, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
'Declare Function WriteConsoleOutputCharacter Lib "kernel32" Alias "WriteConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long
'Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
'Declare Function WriteFileEx Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long
'Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
'Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
'Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'Declare Function WriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
'Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
'Declare Function WriteTapemark Lib "kernel32" (ByVal hDevice As Long, ByVal dwTapemarkType As Long, ByVal dwTapemarkCount As Long, ByVal bimmediate As Long) As Long
'Declare Sub DebugBreak Lib "kernel32" ()
'Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
'Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As Long, ByVal fAccept As Long)
'Declare Sub DragFinish Lib "shell32.dll" (ByVal hDrop As Long)
'Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
'Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
'Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
'Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
'Declare Sub FatalExit Lib "kernel32" (ByVal code As Long)
'Declare Sub FreeLibraryAndExitThread Lib "kernel32" (ByVal hLibModule As Long, ByVal dwExitCode As Long)
'Declare Sub FreeSid Lib "ADVAPI32.DLL" (pSid As Any)
'Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SystemTime)
'Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
'Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
'Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SystemTime)
'Declare Sub GlobalFix Lib "kernel32" (ByVal hMem As Long)
'Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
'Declare Sub GlobalUnfix Lib "kernel32" (ByVal hMem As Long)
'Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
'Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
'Declare Sub LZClose Lib "lz32.dll" (ByVal hfFile As Long)
'Declare Sub LZDone Lib "lz32" ()
'Declare Sub MapGenericMask Lib "ADVAPI32.DLL" (AccessMask As Long, GenericMapping As GENERIC_MAPPING)
'Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'Declare Sub OutputDebugStr Lib "winmm.dll" (ByVal lpszOutputString As String)
'Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
'Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
'Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)
'Declare Sub SetDebugErrorLevel Lib "user32" (ByVal dwLevel As Long)
'Declare Sub SetFileApisToANSI Lib "kernel32" ()
'Declare Sub SetFileApisToOEM Lib "kernel32" ()
'Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
'Declare Sub SetLastErrorEx Lib "user32" (ByVal dwErrCode As Long, ByVal dwType As Long)
'Declare Sub SHFreeNameMappings Lib "shell32.dll" (ByVal hNameMappings As Long)
'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Declare Sub WinExecError Lib "shell32.dll" Alias "WinExecErrorA" (ByVal hwnd As Long, ByVal error As Long, ByVal lpstrFileName As String, ByVal lpstrTitle As String)

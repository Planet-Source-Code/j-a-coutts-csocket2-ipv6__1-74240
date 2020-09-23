Attribute VB_Name = "mWinsock2"
Option Explicit
'
' ---------------------------------------------------------------------------------
' File...........: mWinsock2.bas
' Author.........: J.A. Coutts
' Created........: 02/05/11
' Modified.......: 05/05/11
' Version........: 1.0
' Website........: http://www.yellowhead.com
' Contact........: allecnarf@hotmail.com
'
'Copyright (c) 2011 by JAC Computing
'Vernon, BC, Canada
'
'Based on modSocketMaster by Emiliano Scavuzzo
'and MSocketSupport by Oleg Gdalevich
'Subclassing based on WinSubHook2 by Paul Caton
'
' Port of necessary Winsock2 declares, consts, types etc..
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
' Constants.
' ------------------------------------------------------------------------------
'
' Winsock version constants.
Public Const WINSOCK_V1_1  As Long = &H101
Public Const WINSOCK_V2_2  As Long = &H202
'
' Length of fields within the WSADATA structure.
Public Const WSADESCRIPTION_LEN  As Long = 256
Public Const WSASYS_STATUS_LEN   As Long = 128
'
' Length of string fields for IPv4 and IPv6
Public Const INET_ADDRSTRLEN As Long = 16
Public Const INET6_ADDRSTRLEN As Long = 46
Public Const AI_PASSIVE As Long = 1
'
' For socket handle errors, and bas returns from APIs.
Public Const ERROR_SUCCESS    As Long = 0
Public Const SOCKET_ERROR     As Long = -1
Public Const INVALID_SOCKET   As Long = SOCKET_ERROR
'
' Internet addresses.
Public Const INADDR_ANY          As Long = &H0
Public Const INADDR_ANY6         As Long = &H1
Public Const INADDR_LOOPBACK     As Long = &H7F000001
Public Const INADDR_BROADCAST    As Long = &HFFFFFFFF
Public Const INADDR_NONE         As Long = &HFFFFFFFF
'
' Maximum backlog when calling listen().
Public Const SOMAXCONN  As Long = 5
'
' Messages send with WSAAsyncSelect().
Public Const FD_READ       As Long = &H1
Public Const FD_WRITE      As Long = &H2
Public Const FD_OOB        As Long = &H4
Public Const FD_ACCEPT     As Long = &H8
Public Const FD_CONNECT    As Long = &H10
Public Const FD_CLOSE      As Long = &H20
'
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767
Public Const GMEM_FIXED = &H0
Public Const LOCAL_HOST_BUFF As Integer = 256
Public Const MAXGETHOSTSTRUCT = 1024
'
' Used with shutdown().
Public Const SD_RECEIVE    As Long = &H0
Public Const SD_SEND       As Long = &H1
Public Const SD_BOTH       As Long = &H2

Public Const SOL_SOCKET         As Long = 65535
Public Const SO_SNDBUF          As Long = &H1001&
Public Const SO_RCVBUF          As Long = &H1002&
Public Const SO_MAX_MSG_SIZE    As Long = &H2003
Public Const SO_BROADCAST       As Long = &H20
Public Const FIONREAD           As Long = &H4004667F
'
' Winsock error constants.
Public Const WSABASEERR          As Long = 10000
Public Const WSAEINTR            As Long = WSABASEERR + 4
Public Const WSAEBADF            As Long = WSABASEERR + 9
Public Const WSAEACCES           As Long = WSABASEERR + 13
Public Const WSAEFAULT           As Long = WSABASEERR + 14
Public Const WSAEINVAL           As Long = WSABASEERR + 22
Public Const WSAEMFILE           As Long = WSABASEERR + 24
Public Const WSAEWOULDBLOCK      As Long = WSABASEERR + 35
Public Const WSAEINPROGRESS      As Long = WSABASEERR + 36
Public Const WSAEALREADY         As Long = WSABASEERR + 37
Public Const WSAENOTSOCK         As Long = WSABASEERR + 38
Public Const WSAEDESTADDRREQ     As Long = WSABASEERR + 39
Public Const WSAEMSGSIZE         As Long = WSABASEERR + 40
Public Const WSAEPROTOTYPE       As Long = WSABASEERR + 41
Public Const WSAENOPROTOOPT      As Long = WSABASEERR + 42
Public Const WSAEPROTONOSUPPORT  As Long = WSABASEERR + 43
Public Const WSAESOCKTNOSUPPORT  As Long = WSABASEERR + 44
Public Const WSAEOPNOTSUPP       As Long = WSABASEERR + 45
Public Const WSAEPFNOSUPPORT     As Long = WSABASEERR + 46
Public Const WSAEAFNOSUPPORT     As Long = WSABASEERR + 47
Public Const WSAEADDRINUSE       As Long = WSABASEERR + 48
Public Const WSAEADDRNOTAVAIL    As Long = WSABASEERR + 49
Public Const WSAENETDOWN         As Long = WSABASEERR + 50
Public Const WSAENETUNREACH      As Long = WSABASEERR + 51
Public Const WSAENETRESET        As Long = WSABASEERR + 52
Public Const WSAECONNABORTED     As Long = WSABASEERR + 53
Public Const WSAECONNRESET       As Long = WSABASEERR + 54
Public Const WSAENOBUFS          As Long = WSABASEERR + 55
Public Const WSAEISCONN          As Long = WSABASEERR + 56
Public Const WSAENOTCONN         As Long = WSABASEERR + 57
Public Const WSAESHUTDOWN        As Long = WSABASEERR + 58
Public Const WSAETOOMANYREFS     As Long = WSABASEERR + 59
Public Const WSAETIMEDOUT        As Long = WSABASEERR + 60
Public Const WSAECONNREFUSED     As Long = WSABASEERR + 61
Public Const WSAELOOP            As Long = WSABASEERR + 62
Public Const WSAENAMETOOLONG     As Long = WSABASEERR + 63
Public Const WSAEHOSTDOWN        As Long = WSABASEERR + 64
Public Const WSAEHOSTUNREACH     As Long = WSABASEERR + 65
Public Const WSAENOTEMPTY        As Long = WSABASEERR + 66
Public Const WSAEPROCLIM         As Long = WSABASEERR + 67
Public Const WSAEUSERS           As Long = WSABASEERR + 68
Public Const WSAEDQUOT           As Long = WSABASEERR + 69
Public Const WSAESTALE           As Long = WSABASEERR + 70
Public Const WSAEREMOTE          As Long = WSABASEERR + 71
Public Const WSASYSNOTREADY      As Long = WSABASEERR + 91
Public Const WSAVERNOTSUPPORTED  As Long = WSABASEERR + 92
Public Const WSANOTINITIALISED   As Long = WSABASEERR + 93
Public Const WSAHOST_NOT_FOUND   As Long = WSABASEERR + 1001
Public Const WSATRY_AGAIN        As Long = WSABASEERR + 1002
Public Const WSANO_RECOVERY      As Long = WSABASEERR + 1003
Public Const WSANO_DATA          As Long = WSABASEERR + 1004
'
' Winsock 2 extensions.
Public Const WSA_IO_PENDING         As Long = 997
Public Const WSA_IO_INCOMPLETE      As Long = 996
Public Const WSA_INVALID_HANDLE     As Long = 6
Public Const WSA_INVALID_PARAMETER  As Long = 87
Public Const WSA_NOT_ENOUGH_MEMORY  As Long = 8
Public Const WSA_OPERATION_ABORTED  As Long = 995

Public Const WSA_WAIT_FAILED           As Long = -1
Public Const WSA_WAIT_EVENT_0          As Long = 0
Public Const WSA_WAIT_IO_COMPLETION    As Long = &HC0
Public Const WSA_WAIT_TIMEOUT          As Long = &H102
Public Const WSA_INFINITE              As Long = -1

'WINSOCK CONTROL ERROR CODES
Public Const sckOutOfMemory = 7
Public Const sckBadState = 40006
Public Const sckInvalidArg = 40014
Public Const sckUnsupported = 40018
Public Const sckInvalidOp = 40020
'
' Max size of event handle array when calling WSAWaitForMultipleEvents().
Public Const WSA_MAXIMUM_WAIT_EVENTS   As Long = 64
'
' Size of WSANETWORKEVENTS.iErrorCode[] array.
Public Const FD_MAX_EVENTS    As Long = 10
'
' Used to refer to particular elements of the WSANETWORKEVENTS.iErrorCodes[].
Public Const FD_READ_BIT                     As Long = 0
Public Const FD_WRITE_BIT                    As Long = 1
Public Const FD_OOB_BIT                      As Long = 2
Public Const FD_ACCEPT_BIT                   As Long = 3
Public Const FD_CONNECT_BIT                  As Long = 4
Public Const FD_CLOSE_BIT                    As Long = 5
Public Const FD_QOS_BIT                      As Long = 6
Public Const FD_GROUP_QOS_BIT                As Long = 7
Public Const FD_ROUTING_INTERFACE_CHANGE_BIT As Long = 8
Public Const FD_ADDRESS_LIST_CHANGE_BIT      As Long = 9
'
' ------------------------------------------------------------------------------
' Enumerations.
' ------------------------------------------------------------------------------
'
' Used with socket().
Public Enum Protocols
   IPPROTO_IP = 0
   IPPROTO_ICMP = 1
   IPPROTO_GGP = 2
   IPPROTO_TCP = 6
   IPPROTO_PUP = 12
   IPPROTO_UDP = 17
   IPPROTO_IDP = 22
   IPPROTO_ND = 77
   IPPROTO_RAW = 255
   IPPROTO_MAX = 256
End Enum
'
' Used with socket().
Public Enum SocketTypes
   SOCK_STREAM = 1
   SOCK_DGRAM = 2
   SOCK_RAW = 3
   SOCK_RDM = 4
   SOCK_SEQPACKET = 5
End Enum
'
' Used with socket().
Public Enum AddressFamilies
    AF_UNSPEC = 0
    AF_UNIX = 1
    AF_INET = 2
    AF_IMPLINK = 3
    AF_PUP = 4
    AF_CHAOS = 5
    AF_NS = 6
    AF_IPX = 6
    AF_ISO = 7
    AF_OSI = 7
    AF_ECMA = 8
    AF_DATAKIT = 9
    AF_CCITT = 10
    AF_SNA = 11
    AF_DECNET = 12
    AF_DLI = 13
    AF_LAT = 14
    AF_HYLINK = 15
    AF_APPLETALK = 16
    AF_NETBIOS = 17
    AF_MAX = 18
    AF_INET6 = 23
End Enum
'
Public Enum DestResolucion 'asynchronic host resolution destination
    destConnect = 0
    'destSendUDP = 1
End Enum
'
' ------------------------------------------------------------------------------
' Types.
' ------------------------------------------------------------------------------
'
' To initialize Winsock.
Public Type WSADATA
   wVersion                               As Integer
   wHighVersion                           As Integer
   szDescription(WSADESCRIPTION_LEN + 1)  As Byte
   szSystemstatus(WSASYS_STATUS_LEN + 1)  As Byte
   iMaxSockets                            As Integer
   iMaxUpdDg                              As Integer
   lpVendorInfo                           As Long
End Type
'
' Basic IPv4 addressing structures.
'
Private Type in_addr
   s_addr   As Long
End Type
'
Public Type sockaddr_in
    sin_family          As Integer  '2 bytes
    sin_port            As Integer  '2 bytes
    sin_addr            As in_addr  '4 bytes
    sin_zero(0 To 7)    As Byte     '8 bytes
End Type                            'Total 16 bytes
'
' Basic IPv6 addressing structures.
'
Public Type in6_addr
    s6_addr(0 To 15)      As Byte
End Type
'
Public Type sockaddr_in6
    sin6_family         As Integer  '2 bytes
    sin6_port           As Integer  '2 bytes
    sin6_flowinfo       As Long     '4 bytes
    sin6_addr           As in6_addr '16 bytes
    sin6_scope_id       As Long     '4 bytes
End Type                            'Total 28 bytes
'
Public Type sockaddr
    sa_family           As Integer  '2 bytes
    sa_data(0 To 25)    As Byte     '26 bytes
End Type                            'Total 28 bytes
'
Public Type sockaddr_storage
    sa_family_t         As Integer  '2 bytes
    sa_data(0 To 25)    As Byte     '26 bytes
End Type                            'Total 28 bytes
'
Public Type addrinfo
    ai_flags As Long
    ai_family As Long
    ai_socktype As Long
    ai_protocol As Long
    ai_addrlen As Long
    ai_canonname As Long 'strptr
    ai_addr As Long 'p sockaddr
    ai_next As Long 'p addrinfo
End Type
'
' Used with name resolution functions.
Public Type hostent
   h_name         As Long
   h_aliases      As Long
   h_addrtype     As Integer
   h_length       As Integer
   h_addr_list    As Long
End Type
'
' Used with WSAEnumNetworkEvents().
Public Type WSANETWORKEVENTS
    lNetworkEvents               As Long
    iErrorCode(FD_MAX_EVENTS)    As Integer
End Type
'
' Used when sending ICMP echos (pings).
Public Type IP_OPTION_INFORMATION
    TTL           As Byte
    Tos           As Byte
    Flags         As Byte
    OptionsSize   As Long
    OptionsData   As String * 128
End Type
'
Public Type IP_ECHO_REPLY
    Address(0 To 3)  As Byte
    Status           As Long
    RoundTripTime    As Long
    DataSize         As Integer
    Reserved         As Integer
    data             As Long
    Options          As IP_OPTION_INFORMATION
End Type
'
' ------------------------------------------------------------------------------
' APIs.
' ------------------------------------------------------------------------------
'
' DLL handling functions.
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long
'Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequired As Integer, ByRef lpWSAData As Any) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Public Declare Function WSASetLastError Lib "ws2_32.dll" (ByVal err As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'
' Resolution functions.
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Public Declare Function getpeername2 Lib "ws2_32.dll" Alias "getpeername" (ByVal s As Long, ByRef name As sockaddr, ByRef namelen As Long) As Long
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Public Declare Function getsockname2 Lib "ws2_32.dll" Alias "getsockname" (ByVal s As Long, ByRef name As sockaddr, ByRef namelen As Long) As Long
Public Declare Function GetHostByName Lib "ws2_32.dll" Alias "gethostbyname" (ByVal host_name As String) As Long
Public Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal strHostName As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function gethostbyaddr Lib "ws2_32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Public Declare Function gethostname Lib "ws2_32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
'Public Declare Function getaddrinfo Lib "ws2_32.dll" (ByVal NodeName As String, ByVal ServName As String, Hints As addrinfo, res As addrinfo) As Long
Public Declare Function getaddrinfo Lib "ws2_32.dll" (ByVal NodeName As String, ByVal ServName As String, ByVal lpHints As Long, lpResult As Long) As Long
Public Declare Function freeaddrinfo Lib "ws2_32.dll" (ByVal res As Long) As Long
'
' Conversion functions.
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal laddr As Long) As Long
Public Declare Function inet_pton Lib "ws2_32.dll" (ByVal af As Long, ByVal pszAddrString As String, ByRef pAddrBuf As Any) As Long
'Private Declare Function inet_pton Lib "ws2_32.dll" (ByVal af As Long, ByVal pszAddrString As String, ByRef pAddrBuf As Any) As Long
Public Declare Function inet_ntop Lib "ws2_32.dll" (ByVal af As Long, ByRef ppAddr As Any, ByRef pStringBuf As Any, ByVal StringBufSize As Long) As Long
'Private Declare Function inet_ntop Lib "ws2_32.dll" (ByVal af As Long, ByRef ppAddr As Any, ByRef pStringBuf As Any, ByVal StringBufSize As Long) As Long
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Integer
' Socket functions.
Public Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As AddressFamilies, ByVal stype As SocketTypes, ByVal Protocol As Protocols) As Long
Public Declare Function getsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Public Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
'
Private Declare Function Bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
'Public Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr, ByVal namelen As Long) As Long
Public Declare Function Listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function Accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As sockaddr, ByRef addrlen As Long) As Long
'Public Declare Function Connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As sockaddr, ByVal namelen As Long) As Long
Public Declare Function api_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As sockaddr_in, ByVal namelen As Long) As Long
'
Public Declare Function send Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long) As Long
Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long) As Long
Public Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long, ByRef toaddr As sockaddr_in, ByVal tolen As Long) As Long
Public Declare Function sendto2 Lib "ws2_32.dll" Alias "sendto" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long, ByRef toaddr As sockaddr, ByVal tolen As Long) As Long
Public Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long, ByRef fromaddr As sockaddr_in, ByRef fromlen As Long) As Long
Public Declare Function recvfrom2 Lib "ws2_32.dll" Alias "recvfrom" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long, ByRef fromaddr As sockaddr, ByRef fromlen As Long) As Long
'
Public Declare Function shutdown Lib "ws2_32.dll" (ByVal s As Long, ByVal how As Long) As Long
Public Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
'
' I/O model functions.
'Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Integer, ByVal lEvent As Long) As Long
Private Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
'
Public Declare Function WSACreateEvent Lib "ws2_32.dll" () As Long
Public Declare Function WSAEventSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hEventObject As Long, ByVal lNetworkEvents As Long) As Long
Public Declare Function WSAResetEvent Lib "ws2_32.dll" (ByVal hEvent As Long) As Long
Public Declare Function WSASetEvent Lib "ws2_32.dll" (ByVal hEvent As Long) As Long
Public Declare Function WSACloseEvent Lib "ws2_32.dll" (ByVal hEvent As Long) As Long
Public Declare Function WSAWaitForMultipleEvents Lib "ws2_32.dll" (ByVal cEvents As Long, ByRef lphEvents As Long, ByVal fWaitAll As Boolean, ByVal dwTimeout As Long, ByVal fAlertable As Boolean) As Long
Public Declare Function WSAEnumNetworkEvents Lib "ws2_32.dll" (ByVal s As Long, ByVal hEvent As Long, ByRef lpNetworkEvents As WSANETWORKEVENTS) As Long
Public Declare Function WSACancelAsyncRequest Lib "ws2_32.dll" (ByVal hAsyncTaskHandle As Long) As Long
'
' ICMP functions.
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Public Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean
'
' Other general Win32 APIs.
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, ByVal pSrc As Long, ByVal ByteLen As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (pDestination As Any, ByVal lByteCount As Long)
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Public Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function api_LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function api_SetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function api_KillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

'==============================================================================
'MEMBER VARIABLES
'==============================================================================

Private m_blnInitiated          As Boolean      'specify if winsock service was initiated
Private m_lngSocksQuantity      As Long         'number of instances created
Private m_colSocketsInst        As Collection   'sockets list and instance owner
Private m_colAcceptList         As Collection   'sockets in queue that need to be accepted
Private m_lngWindowHandle       As Long         'message window handle

'==============================================================================
'SUBCLASSING DECLARATIONS
'by Paul Caton
'==============================================================================
Private Declare Function api_IsWindow Lib "user32" Alias "IsWindow" (ByVal hWnd As Long) As Long
Private Declare Function api_GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function api_SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function api_GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function api_GetProcAddress Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function api_DestroyWindow Lib "user32" Alias "DestroyWindow" (ByVal hWnd As Long) As Long

Private Const PATCH_09 As Long = 119
Private Const PATCH_0C As Long = 150

Private Const GWL_WNDPROC As Long = (-4)

Private Const WM_APP As Long = 32768 '0x8000

Public Const RESOLVE_MESSAGE As Long = WM_APP
Public Const SOCKET_MESSAGE  As Long = WM_APP + 1

Private Const TIMER_TIMEOUT As Long = 200   'control timer time out, in milliseconds

Private lngMsgCntA      As Long     'TableA entry count
Private lngMsgCntB      As Long     'TableB entry count
Private lngTableA1()    As Long     'TableA1: list of async handles
Private lngTableA2()    As Long     'TableA2: list of async handles owners
Private lngTableB1()    As Long     'TableB1: list of sockets
Private lngTableB2()    As Long     'TableB2: list of sockets owners
Private hWndSub         As Long     'window handle subclassed
Private nAddrSubclass   As Long     'address of our WndProc
Private nAddrOriginal   As Long     'address of original WndProc
Private hTimer          As Long     'control timer handle
Public DbgFlg           As Boolean  'Determines if debug messages are recorded
Public Function GetErrorDescription(ByVal lngErrorCode As Long) As String
    'This function receives a number that represents an error
    'and returns the corresponding description string.
    Select Case lngErrorCode
        Case WSAEACCES
            GetErrorDescription = "Permission denied."
        Case WSAEADDRINUSE
            GetErrorDescription = "Address already in use."
        Case WSAEADDRNOTAVAIL
            GetErrorDescription = "Cannot assign requested address."
        Case WSAEAFNOSUPPORT
            GetErrorDescription = "Address family not supported by protocol family."
        Case WSAEALREADY
            GetErrorDescription = "Operation already in progress."
        Case WSAECONNABORTED
            GetErrorDescription = "Software caused connection abort."
        Case WSAECONNREFUSED
            GetErrorDescription = "Connection refused."
        Case WSAECONNRESET
            GetErrorDescription = "Connection reset by peer."
        Case WSAEDESTADDRREQ
            GetErrorDescription = "Destination address required."
        Case WSAEFAULT
            GetErrorDescription = "Bad address."
        Case WSAEHOSTUNREACH
            GetErrorDescription = "No route to host."
        Case WSAEINPROGRESS
            GetErrorDescription = "Operation now in progress."
        Case WSAEINTR
            GetErrorDescription = "Interrupted function call."
        Case WSAEINVAL
            GetErrorDescription = "Invalid argument."
        Case WSAEISCONN
            GetErrorDescription = "Socket is already connected."
        Case WSAEMFILE
            GetErrorDescription = "Too many open files."
        Case WSAEMSGSIZE
            GetErrorDescription = "Message too long."
        Case WSAENETDOWN
            GetErrorDescription = "Network is down."
        Case WSAENETRESET
            GetErrorDescription = "Network dropped connection on reset."
        Case WSAENETUNREACH
            GetErrorDescription = "Network is unreachable."
        Case WSAENOBUFS
            GetErrorDescription = "No buffer space available."
        Case WSAENOPROTOOPT
            GetErrorDescription = "Bad protocol option."
        Case WSAENOTCONN
            GetErrorDescription = "Socket is not connected."
        Case WSAENOTSOCK
            GetErrorDescription = "Socket operation on nonsocket."
        Case WSAEOPNOTSUPP
            GetErrorDescription = "Operation not supported."
        Case WSAEPFNOSUPPORT
            GetErrorDescription = "Protocol family not supported."
        Case WSAEPROCLIM
            GetErrorDescription = "Too many processes."
        Case WSAEPROTONOSUPPORT
            GetErrorDescription = "Protocol not supported."
        Case WSAEPROTOTYPE
            GetErrorDescription = "Protocol wrong type for socket."
        Case WSAESHUTDOWN
            GetErrorDescription = "Cannot send after socket shutdown."
        Case WSAESOCKTNOSUPPORT
            GetErrorDescription = "Socket type not supported."
        Case WSAETIMEDOUT
            GetErrorDescription = "Connection timed out."
        Case WSAEWOULDBLOCK
            GetErrorDescription = "Resource temporarily unavailable."
        Case WSAHOST_NOT_FOUND
            GetErrorDescription = "Host not found."
        Case WSANOTINITIALISED
            GetErrorDescription = "Successful WSAStartup not yet performed."
        Case WSANO_DATA
            GetErrorDescription = "Valid name, no data record of requested type."
        Case WSANO_RECOVERY
            GetErrorDescription = "This is a nonrecoverable error."
        Case WSASYSNOTREADY
            GetErrorDescription = "Network subsystem is unavailable."
        Case WSATRY_AGAIN
            GetErrorDescription = "Nonauthoritative host not found."
        Case WSAVERNOTSUPPORTED
            GetErrorDescription = "Winsock.dll version out of range."
        Case Else
            GetErrorDescription = "Unknown error."
    End Select
End Function

Private Function Subclass_InIDE() As Boolean
    'Return whether we're running in the IDE. Public for general utility purposes
    Debug.Assert Subclass_SetTrue(Subclass_InIDE)
End Function
Private Function Subclass_SetTrue(bValue As Boolean) As Boolean
    'Worker function for InIDE - will only be called whilst running in the IDE
    Subclass_SetTrue = True
    bValue = True
End Function

Private Function Subclass_AddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    'Return the address of the passed function in the passed dll
    Subclass_AddrFunc = api_GetProcAddress(api_GetModuleHandle(sDLL), sProc)
End Function
Private Sub Subclass_AddSocketMessage(ByVal lngSocket As Long, ByVal lngObjectPointer As Long)
    Dim Count As Long
    For Count = 1 To lngMsgCntB
        Select Case lngTableB1(Count)
            Case -1
                lngTableB1(Count) = lngSocket
                lngTableB2(Count) = lngObjectPointer
                Exit Sub
            Case lngSocket
                Debug.Print "WARNING: Socket already registered!"
                If DbgFlg Then Call LogError("WARNING: Socket already registered!")
                Exit Sub
        End Select
    Next Count
    lngMsgCntB = lngMsgCntB + 1
    ReDim Preserve lngTableB1(1 To lngMsgCntB)
    ReDim Preserve lngTableB2(1 To lngMsgCntB)
    lngTableB1(lngMsgCntB) = lngSocket
    lngTableB2(lngMsgCntB) = lngObjectPointer
    Subclass_PatchTableB
End Sub
Public Sub Subclass_ChangeOwner(ByVal lngSocket As Long, ByVal lngObjectPointer As Long)
    Dim Count As Long
    For Count = 1 To lngMsgCntB
        If lngTableB1(Count) = lngSocket Then
            lngTableB2(Count) = lngObjectPointer
            Exit Sub
        End If
    Next Count
End Sub
Private Sub Subclass_DelSocketMessage(ByVal lngSocket As Long)
    Dim Count As Long
    For Count = 1 To lngMsgCntB
        If lngTableB1(Count) = lngSocket Then
            lngTableB1(Count) = -1
            lngTableB2(Count) = -1
            Exit Sub
        End If
    Next Count
End Sub

Private Function Subclass_AddrMsgTbl(ByRef aMsgTbl() As Long) As Long
    'Return the address of the low bound of the passed table array
    On Error Resume Next                                    'The table may not be dimensioned yet so we need protection
        Subclass_AddrMsgTbl = VarPtr(aMsgTbl(1))            'Get the address of the first element of the passed message table
    On Error GoTo 0                                         'Switch off error protection
End Function

Private Sub Subclass_PatchRel(ByVal nOffset As Long, ByVal nTargetAddr As Long)
    'Patch the machine code buffer offset with the relative address to the target address
    Call CopyMemory(ByVal (nAddrSubclass + nOffset), nTargetAddr - nAddrSubclass - nOffset - 4, 4)
End Sub
Private Sub Subclass_PatchTableB()
    Const PATCH_0D As Long = 158
    Const PATCH_0E As Long = 174
    Call Subclass_PatchVal(PATCH_0C, lngMsgCntB)
    Call Subclass_PatchVal(PATCH_0D, Subclass_AddrMsgTbl(lngTableB1))
    Call Subclass_PatchVal(PATCH_0E, Subclass_AddrMsgTbl(lngTableB2))
End Sub

Private Sub Subclass_PatchVal(ByVal nOffset As Long, ByVal nValue As Long)
    'Patch the machine code buffer offset with the passed value
    Call CopyMemory(ByVal (nAddrSubclass + nOffset), nValue, 4)
End Sub
Private Function Subclass_Subclass(ByVal hWnd As Long) As Boolean
    'Set the window subclass
    Const PATCH_02 As Long = 62                                'Address of the previous WndProc
    Const PATCH_05 As Long = 82                                'Control timer handle
    Const PATCH_07 As Long = 108                               'Address of the previous WndProc
    If hWndSub = 0 Then
        Debug.Assert api_IsWindow(hWnd)                         'Invalid window handle
        hWndSub = hWnd                                          'Store the window handle
        'Get the original window proc
        nAddrOriginal = api_GetWindowLong(hWnd, GWL_WNDPROC)
        Call Subclass_PatchVal(PATCH_02, nAddrOriginal)                  'Original WndProc address for CallWindowProc, call the original WndProc
        Call Subclass_PatchVal(PATCH_07, nAddrOriginal)                  'Original WndProc address for SetWindowLong, unsubclass on IDE stop
        'Set our WndProc in place of the original
        nAddrOriginal = api_SetWindowLong(hWnd, GWL_WNDPROC, nAddrSubclass)
        If nAddrOriginal <> 0 Then
          Subclass_Subclass = True                                       'Success
        End If
    End If
    If Subclass_InIDE Then
        hTimer = api_SetTimer(0, 0, TIMER_TIMEOUT, nAddrSubclass)        'Create the control timer
        Call Subclass_PatchVal(PATCH_05, hTimer)                         'Patch the control timer handle
    End If
    Debug.Assert Subclass_Subclass
End Function
Private Sub Subclass_Terminate()
    'UnSubclass and release the allocated memory
    Call Subclass_UnSubclass                                  'UnSubclass if the Subclass thunk is active
    Call GlobalFree(nAddrSubclass)                            'Release the allocated memory
    Debug.Print "OK Freed subclass memory at: " & Hex$(nAddrSubclass)
    If DbgFlg Then Call LogError("OK Freed subclass memory at: " & Hex$(nAddrSubclass))
    nAddrSubclass = 0
    ReDim lngTableA1(1 To 1)
    ReDim lngTableA2(1 To 1)
    ReDim lngTableB1(1 To 1)
    ReDim lngTableB2(1 To 1)
End Sub
Private Function Subclass_UnSubclass() As Boolean
    'Stop subclassing the window
    If hWndSub <> 0 Then
        lngMsgCntA = 0
        lngMsgCntB = 0
        Call Subclass_PatchVal(PATCH_09, lngMsgCntA)                              'Patch the TableA entry count to ensure no further Proc callbacks
        Call Subclass_PatchVal(PATCH_0C, lngMsgCntB)                              'Patch the TableB entry count to ensure no further Proc callbacks
        'Restore the original WndProc
        Call api_SetWindowLong(hWndSub, GWL_WNDPROC, nAddrOriginal)
        If hTimer <> 0 Then
            Call api_KillTimer(0&, hTimer)            'Destroy control timer
            hTimer = 0
        End If
        hWndSub = 0                                   'Indicate the subclasser is inactive
        Subclass_UnSubclass = True                    'Success
    End If
End Function

Private Sub Subclass_Initialize()
    Const PATCH_01 As Long = 16                   'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_03 As Long = 72                   'Relative address of SetWindowsLong
    Const PATCH_04 As Long = 77                   'Relative address of WSACleanup
    Const PATCH_06 As Long = 89                   'Relative address of KillTimer
    Const PATCH_08 As Long = 113                  'Relative address of CallWindowProc
    Const FUNC_EBM As String = "EbMode"           'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL As String = "SetWindowLongA"   'SetWindowLong allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const FUNC_CWP As String = "CallWindowProcA"  'We use CallWindowProc to call the original WndProc
    Const FUNC_WCU As String = "WSACleanup"       'closesocket is called when the program is closed to release the sockets
    Const FUNC_KTM As String = "KillTimer"        'KillTimer destroys the control timer
    Const MOD_VBA5 As String = "vba5"             'Location of the EbMode function if running VB5
    Const MOD_VBA6 As String = "vba6"             'Location of the EbMode function if running VB6
    Const MOD_USER As String = "user32"           'Location of the SetWindowLong & CallWindowProc functions
    Const MOD_WS   As String = "ws2_32"           'Location of the closesocket function
    Dim i        As Long                          'Loop index
    Dim nLen     As Long                          'String lengths
    Dim sHex     As String                        'Hex code string
    Dim sCode    As String                        'Binary code string
    'Store the hex pair machine code representation in sHex
    sHex = "5850505589E55753515231C0FCEB09E8xxxxx01x85C074258B45103D0080000074543D01800000746CE8310000005A595B5FC9C21400E824000000EBF168xxxxx02x6AFCFF750CE8xxxxx03xE8xxxxx04x68xxxxx05x6A00E8xxxxx06xEBCFFF7518FF7514FF7510FF750C68xxxxx07xE8xxxxx08xC3BBxxxxx09x8B4514BFxxxxx0Ax89D9F2AF75A529CB4B8B1C9Dxxxxx0BxEB1DBBxxxxx0Cx8B4514BFxxxxx0Dx89D9F2AF758629CB4B8B1C9Dxxxxx0Ex895D088B1B8B5B1C89D85A595B5FC9FFE0"
    nLen = Len(sHex)                                          'Length of hex pair string
    'Convert the string from hex pairs to bytes and store in the ASCII string opcode buffer
    For i = 1 To nLen Step 2                                  'For each pair of hex characters
        sCode = sCode & ChrB$(Val("&H" & Mid$(sHex, i, 2)))     'Convert a pair of hex characters to a byte and append to the ASCII string
    Next i                                                    'Next pair
    nLen = LenB(sCode)                                        'Get the machine code length
    nAddrSubclass = GlobalAlloc(0, nLen)                  'Allocate fixed memory for machine code buffer
    Debug.Print "OK Subclass memory allocated at: " & Hex$(nAddrSubclass)
    If DbgFlg Then Call LogError("OK Subclass memory allocated at: " & Hex$(nAddrSubclass))
    'Copy the code to allocated memory
    Call CopyMemory(ByVal nAddrSubclass, ByVal StrPtr(sCode), nLen)
    If Subclass_InIDE Then
        'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
        Call CopyMemory(ByVal nAddrSubclass + 13, &H9090, 2)
        i = Subclass_AddrFunc(MOD_VBA6, FUNC_EBM)               'Get the address of EbMode in vba6.dll
        If i = 0 Then                                           'Found?
          i = Subclass_AddrFunc(MOD_VBA5, FUNC_EBM)             'VB5 perhaps, try vba5.dll
        End If
        Debug.Assert i                                          'Ensure the EbMode function was found
        Call Subclass_PatchRel(PATCH_01, i)                     'Patch the relative address to the EbMode api function
    End If
    Call api_LoadLibrary(MOD_WS)                              'Ensure ws_32.dll is loaded before getting WSACleanup address
    Call Subclass_PatchRel(PATCH_03, Subclass_AddrFunc(MOD_USER, FUNC_SWL))     'Address of the SetWindowLong api function
    Call Subclass_PatchRel(PATCH_04, Subclass_AddrFunc(MOD_WS, FUNC_WCU))       'Address of the WSACleanup api function
    Call Subclass_PatchRel(PATCH_06, Subclass_AddrFunc(MOD_USER, FUNC_KTM))     'Address of the KillTimer api function
    Call Subclass_PatchRel(PATCH_08, Subclass_AddrFunc(MOD_USER, FUNC_CWP))     'Address of the CallWindowProc api function
End Sub

Private Function InitiateService() As Long
    'This function initiate the winsock service calling
    'the api_WSAStartup funtion and returns resulting value.
    Dim udtWSAData As WSADATA
    Dim lngResult As Long
    lngResult = WSAStartup(&H202, udtWSAData)
    InitiateService = lngResult
End Function
Public Function StringFromPointer(ByVal lPointer As Long) As String
'Receives a string pointer and it turns it into a regular string.
    Dim strTemp As String
    Dim lRetVal As Long
    strTemp = String$(lstrlenA(ByVal lPointer), 0)
    lRetVal = lstrcpyA(ByVal strTemp, ByVal lPointer)
    If lRetVal Then StringFromPointer = strTemp
End Function

Public Sub UnregisterSocket(ByVal lngSocket As Long)
    'Removes the socket from the m_colSocketsInst collection
    'If it is the last socket in that collection, the window
    'and colection will be destroyed as well.
    Subclass_DelSocketMessage lngSocket
    On Error Resume Next
    m_colSocketsInst.Remove "S" & lngSocket
    If m_colSocketsInst.Count = 0 Then
        Set m_colSocketsInst = Nothing
        Subclass_UnSubclass
        DestroyWinsockMessageWindow
        Debug.Print "OK Destroyed socket collection"
        If DbgFlg Then Call LogError("OK Destroyed socket collection")
    End If
End Sub
Private Function DestroyWinsockMessageWindow() As Long
    'Destroy the window that is used to capture sockets messages.
    'Returns 0 if it has success.
    DestroyWinsockMessageWindow = 0
    If m_lngWindowHandle = 0 Then
        Debug.Print "WARNING lngWindowHandle is ZERO"
        If DbgFlg Then Call LogError("WARNING lngWindowHandle is ZERO")
        Exit Function
    End If
    Dim lngResult As Long
    lngResult = api_DestroyWindow(m_lngWindowHandle)
    If lngResult = 0 Then
        DestroyWinsockMessageWindow = sckOutOfMemory
        err.Raise sckOutOfMemory, "mWinsock2.DestroyWinsockMessageWindow", "Out of memory"
    Else
        Debug.Print "OK Destroyed winsock message window " & m_lngWindowHandle
        If DbgFlg Then Call LogError("OK Destroyed winsock message window" & Str$(m_lngWindowHandle))
        m_lngWindowHandle = 0
    End If
End Function

Private Function FinalizeService() As Long
    'Finish winsock service calling the function
    'api_WSACleanup and returns the result.
    Dim lngResultado As Long
    lngResultado = WSACleanup
    FinalizeService = lngResultado
End Function
Public Function HiWord(lngValue As Long) As Long
'Returns the hi word from a double word.
    If (lngValue And &H80000000) = &H80000000 Then
        HiWord = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        HiWord = (lngValue And &HFFFF0000) \ &H10000
    End If
End Function

Public Function LoWord(lngValue As Long) As Long
'Returns the low word from a double word.
    LoWord = (lngValue And &HFFFF&)
End Function

Public Function PeekB(ByVal lpdwData As Long) As Byte
    CopyMemory PeekB, ByVal lpdwData, 1
End Function
Private Function FileErrors(errVal As Integer) As Integer
'Return Value 0=Resume,              1=Resume Next,
'             2=Unrecoverable Error, 3=Unrecognized Error
Dim msgType%
Dim Msg$
Dim Response%
msgType% = 48
Select Case errVal
    Case 68
      Msg$ = "That device appears Unavailable."
      msgType% = msgType% + 4
    Case 71
      Msg$ = "Insert a Disk in the Drive"
    Case 53
      Msg$ = "Cannot Find File"
      msgType% = msgType% + 5
   Case 57
      Msg$ = "Internal Disk Error."
      msgType% = msgType% + 4
    Case 61
      Msg$ = "Disk is Full.  Continue?"
      msgType% = 35
    Case 64, 52
      Msg$ = "That Filename is Illegal!"
    Case 70
      Msg$ = "File in use by another user!"
      msgType% = msgType% + 5
    Case 76
      Msg$ = "Path does not Exist!"
      msgType% = msgType% + 2
    Case 54
      Msg$ = "Bad File Mode!"
    Case 55
      Msg$ = "File is Already Open."
    Case 62
      Msg$ = "Read Attempt Past End of File."
    Case Else
      FileErrors = 3
      Exit Function
  End Select
  Response% = MsgBox(Msg$, msgType%, "Disk Error")
  Select Case Response%
    Case 1, 4
      FileErrors = 0
    Case 5
      FileErrors = 1
    Case 2, 3
      FileErrors = 2
    Case Else
      FileErrors = 3
  End Select
End Function

Public Sub RegisterAccept(ByVal lngSocket As Long)
'Assign a temporal instance of CSocket2 to a
'socket and register this socket to the accept list.
    If m_colAcceptList Is Nothing Then
        Set m_colAcceptList = New Collection
        Debug.Print "OK Created accept collection"
        If DbgFlg Then Call LogError("OK Created accept collection")
    End If
    Dim Socket As cSocket2
    Set Socket = New cSocket2
    Socket.Accept2 lngSocket
    m_colAcceptList.Add Socket, "S" & lngSocket
End Sub

Public Sub LogError(Log$)
    Dim LogFile%
    LogFile% = OpenFile(App.Path + "\IPv6Chat.Log", 3, 0, 80)
    If LogFile% = 0 Then
        MsgBox "File Error with LogFile", 16, "ABORT PROCEDURE"
        Exit Sub
    End If
    Print #LogFile%, CStr(Now) + ": " + Log$
    Close LogFile%
End Sub

Private Function OpenFile(FileName$, Mode%, RLock%, RecordLen%) As Integer
  Const REPLACEFILE = 1, READAFILE = 2, ADDTOFILE = 3
  Const RANDOMFILE = 4, BINARYFILE = 5
  Const NOLOCK = 0, RDLOCK = 1, WRLOCK = 2, RWLOCK = 3
  Dim FileNum%
  Dim Action%
  FileNum% = FreeFile
  On Error GoTo OpenErrors
  Select Case Mode
    Case REPLACEFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Output Shared As FileNum%
            Case RDLOCK
                Open FileName For Output Lock Read As FileNum%
            Case WRLOCK
                Open FileName For Output Lock Write As FileNum%
            Case RWLOCK
                Open FileName For Output Lock Read Write As FileNum%
        End Select
    Case READAFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Input Shared As FileNum%
            Case RDLOCK
                Open FileName For Input Lock Read As FileNum%
            Case WRLOCK
                Open FileName For Input Lock Write As FileNum%
            Case RWLOCK
                Open FileName For Input Lock Read Write As FileNum%
        End Select
    Case ADDTOFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Append Shared As FileNum%
            Case RDLOCK
                Open FileName For Append Lock Read As FileNum%
            Case WRLOCK
                Open FileName For Append Lock Write As FileNum%
            Case RWLOCK
                Open FileName For Append Lock Read Write As FileNum%
        End Select
    Case RANDOMFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Random Shared As FileNum% Len = RecordLen%
            Case RDLOCK
                Open FileName For Random Lock Read As FileNum% Len = RecordLen%
            Case WRLOCK
                Open FileName For Random Lock Write As FileNum% Len = RecordLen%
            Case RWLOCK
                Open FileName For Random Lock Read Write As FileNum% Len = RecordLen%
        End Select
    Case BINARYFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Binary Shared As FileNum%
            Case RDLOCK
                Open FileName For Binary Lock Read As FileNum%
            Case WRLOCK
                Open FileName For Binary Lock Write As FileNum%
            Case RWLOCK
                Open FileName For Binary Lock Read Write As FileNum%
        End Select
    Case Else
      Exit Function
  End Select
  OpenFile = FileNum%
Exit Function
OpenErrors:
  Action% = FileErrors(err)
  Select Case Action%
    Case 0
      Resume            'Resumes at line where ERROR occured
    Case 1
        Resume Next     'Resumes at line after ERROR
    Case 2
        OpenFile = 0     'Unrecoverable ERROR-reports error, exits function with error code
        Exit Function
    Case Else
        MsgBox Error$(err) + vbCrLf + "After line " + Str$(Erl) + vbCrLf + "Program will TERMINATE!"
        'Unrecognized ERROR-reports error and terminates.
        'End
  End Select
End Function

Public Function IsSocketRegistered(ByVal lngSocket As Long) As Boolean
'Returns TRUE si the socket that is passed is registered
'in the colSocketsInst collection.
    On Error GoTo Error_Handler
    m_colSocketsInst.Item ("S" & lngSocket)
    IsSocketRegistered = True
    Exit Function
Error_Handler:
    IsSocketRegistered = False
End Function
Public Function GetAcceptClass(ByVal lngSocket As Long) As cSocket2
'Return the accept instance class from a socket.
    Set GetAcceptClass = m_colAcceptList("S" & lngSocket)
End Function

Public Function IsAcceptRegistered(ByVal lngSocket As Long) As Boolean
'Returns True is lngSocket is registered on the accept list.
    On Error GoTo Error_Handler
    m_colAcceptList.Item ("S" & lngSocket)
    IsAcceptRegistered = True
    Exit Function
Error_Handler:
    IsAcceptRegistered = False
End Function
Public Sub UnregisterAccept(ByVal lngSocket As Long)
'Unregister lngSocket from the accept list.
    m_colAcceptList.Remove "S" & lngSocket
    If m_colAcceptList.Count = 0 Then
        Set m_colAcceptList = Nothing
        Debug.Print "OK Destroyed accept collection"
        If DbgFlg Then Call LogError("OK Destroyed accept collection")
    End If
End Sub



Public Function FinalizeProcesses() As Long
    'Once we are done with the class instance we call this
    'function to discount it and finish winsock service if
    'it was the last one.
    'Returns 0 if it has success.
    FinalizeProcesses = 0
    m_lngSocksQuantity = m_lngSocksQuantity - 1
    'if the service was initiated and there's no more instances
    'of the class then we finish the service
    If m_blnInitiated And m_lngSocksQuantity = 0 Then
        If FinalizeService = SOCKET_ERROR Then
            Dim lngErrorCode As Long
            lngErrorCode = err.LastDllError
            FinalizeProcesses = lngErrorCode
            err.Raise lngErrorCode, "mWinsock2.FinalizeProcesses", GetErrorDescription(lngErrorCode)
        Else
            Debug.Print "OK Winsock service finalized"
            If DbgFlg Then Call LogError("OK Winsock service finalized")
        End If
        Subclass_Terminate
        m_blnInitiated = False
    End If
End Function


Public Function InitiateProcesses() As Long
'This function initiates the processes needed to keep
'control of sockets. Returns 0 if it has success.
    InitiateProcesses = 0
    m_lngSocksQuantity = m_lngSocksQuantity + 1
    'if the service wasn't initiated yet we do it now
    If Not m_blnInitiated Then
        Subclass_Initialize
        m_blnInitiated = True
        Dim lngResult As Long
        lngResult = InitiateService
        If lngResult = 0 Then
            Debug.Print "OK Winsock service initiated"
            If DbgFlg Then Call LogError("OK Winsock service initiated")
        Else
            Debug.Print "ERROR trying to initiate winsock service"
            If DbgFlg Then Call LogError("ERROR trying to initiate winsock service")
            err.Raise lngResult, "mWinsock2.InitiateProcesses", GetErrorDescription(lngResult)
            InitiateProcesses = lngResult
        End If
    End If
End Function
Public Function RegisterSocket(ByVal lngSocket As Long, ByVal lngObjectPointer As Long, ByVal blnEvents As Boolean) As Boolean
'Adds the socket to the m_colSocketsInst collection, and
'registers that socket with WSAAsyncSelect Winsock API
'function to receive network events for the socket.
'If this socket is the first one to be registered, the
'window and collection will be created in this function as well.
    If m_colSocketsInst Is Nothing Then
        Set m_colSocketsInst = New Collection
        Debug.Print "OK Created socket collection"
        If DbgFlg Then Call LogError("OK Created socket collection")
        If CreateWinsockMessageWindow <> 0 Then
            err.Raise sckOutOfMemory, "mWinsock2.RegisterSocket", "Out of memory"
        End If
        Subclass_Subclass (m_lngWindowHandle)
    End If
    Subclass_AddSocketMessage lngSocket, lngObjectPointer
    'Do we need to register socket events?
    If blnEvents Then
        Dim lngEvents As Long
        Dim lngResult As Long
        Dim lngErrorCode As Long
        lngEvents = FD_READ Or FD_WRITE Or FD_ACCEPT Or FD_CONNECT Or FD_CLOSE
        lngResult = WSAAsyncSelect(lngSocket, m_lngWindowHandle, SOCKET_MESSAGE, lngEvents)
        If lngResult = SOCKET_ERROR Then
            Debug.Print "ERROR trying to register events from socket " & lngSocket
            If DbgFlg Then Call LogError("ERROR trying to register events from socket " & Str$(lngSocket))
            lngErrorCode = err.LastDllError
            err.Raise lngErrorCode, "mWinsock2.RegisterSocket", GetErrorDescription(lngErrorCode)
        Else
            Debug.Print "OK Registered events from socket " & lngSocket
            If DbgFlg Then Call LogError("OK Registered events from socket" & Str$(lngSocket))
        End If
    End If
    m_colSocketsInst.Add lngObjectPointer, "S" & lngSocket
    RegisterSocket = True
End Function
Private Function CreateWinsockMessageWindow() As Long
'Create a window that is used to capture sockets messages.
'Returns 0 if it has success.
    m_lngWindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
    If m_lngWindowHandle = 0 Then
        CreateWinsockMessageWindow = sckOutOfMemory
        Exit Function
    Else
        CreateWinsockMessageWindow = 0
        Debug.Print "OK Created winsock message window " & m_lngWindowHandle
        If DbgFlg Then Call LogError("OK Created winsock message window " & Hex$(m_lngWindowHandle))
    End If
End Function

Public Function IntegerToUnsigned(Value As Integer) As Long
'The function takes a Long containing a value in the range 
'of an unsigned Integer and returns an Integer that you 
'can pass to an API that requires an unsigned Integer
    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If
End Function


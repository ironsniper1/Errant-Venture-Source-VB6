Attribute VB_Name = "Module1"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Catalyst SocketWrench 3.6
' Copyright 1995-2003, Catalyst Development Corporation
' All rights reserved
'
' This product is licensed to you pursuant to the terms of the
' Catalyst license agreement included with the original software,
' and is protected by copyright law and international treaties.
' Unauthorized reproduction or distribution may result in severe
' criminal penalties.
'
' This file contains constants for use in Visual Basic 6.0 and
' should be loaded into a code module. If you are updating an
' application written using an earlier version of this product,
' you should replace the constants with the values in this file.
'

'
' General constants used with most of the controls
'
Public Const INVALID_HANDLE = -1
Public Const CONTROL_ERRIGNORE = 0
Public Const CONTROL_ERRDISPLAY = 1

'
' SocketWrench Control Actions
'
Public Const SOCKET_OPEN = 1
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_LISTEN = 3
Public Const SOCKET_ACCEPT = 4
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_CLOSE = 7
Public Const SOCKET_DISCONNECT = 7
Public Const SOCKET_ABORT = 8
Public Const SOCKET_STARTUP = 9
Public Const SOCKET_CLEANUP = 10

'
' SocketWrench Control States
'
Public Const SOCKET_NONE = 0
Public Const SOCKET_IDLE = 1
Public Const SOCKET_LISTENING = 2
Public Const SOCKET_CONNECTING = 3
Public Const SOCKET_ACCEPTING = 4
Public Const SOCKET_RECEIVING = 5
Public Const SOCKET_SENDING = 6
Public Const SOCKET_CLOSING = 7

'
' Socket Address Families
'
Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2

'
' Socket Types
'
Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5

'
' Protocol Types
'
Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_GGP = 2
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77
Public Const IPPROTO_RAW = 255
Public Const IPPROTO_MAX = 256

'
' Well-Known Port Numbers
'
Public Const IPPORT_ANY = 0
Public Const IPPORT_ECHO = 7
Public Const IPPORT_DISCARD = 9
Public Const IPPORT_SYSTAT = 11
Public Const IPPORT_DAYTIME = 13
Public Const IPPORT_NETSTAT = 15
Public Const IPPORT_CHARGEN = 19
Public Const IPPORT_FTP = 21
Public Const IPPORT_TELNET = 23
Public Const IPPORT_SMTP = 25
Public Const IPPORT_TIMESERVER = 37
Public Const IPPORT_NAMESERVER = 42
Public Const IPPORT_WHOIS = 43
Public Const IPPORT_MTP = 57
Public Const IPPORT_TFTP = 69
Public Const IPPORT_FINGER = 79
Public Const IPPORT_HTTP = 80
Public Const IPPORT_POP3 = 110
Public Const IPPORT_NNTP = 119
Public Const IPPORT_SNMP = 161
Public Const IPPORT_EXEC = 512
Public Const IPPORT_LOGIN = 513
Public Const IPPORT_SHELL = 514
Public Const IPPORT_RESERVED = 1024
Public Const IPPORT_USERRESERVED = 5000

'
' Network Addresses
'
Public Const INADDR_ANY = "0.0.0.0"
Public Const INADDR_LOOPBACK = "127.0.0.1"
Public Const INADDR_NONE = "255.255.255.255"

'
' Shutdown Values
'
Public Const SOCKET_READ = 0
Public Const SOCKET_WRITE = 1
Public Const SOCKET_READWRITE = 2

'
' Byte Order
'
Public Const LOCAL_BYTE_ORDER = 0
Public Const NETWORK_BYTE_ORDER = 1

'
' SocketWrench Error Response
'
Public Const SOCKET_ERRIGNORE = 0
Public Const SOCKET_ERRDISPLAY = 1

'
' SocketWrench Error Codes
'
Public Const WSABASEERR = 24000
Public Const WSAEINTR = 24004
Public Const WSAEBADF = 24009
Public Const WSAEACCES = 24013
Public Const WSAEFAULT = 24014
Public Const WSAEINVAL = 24022
Public Const WSAEMFILE = 24024
Public Const WSAEWOULDBLOCK = 24035
Public Const WSAEINPROGRESS = 24036
Public Const WSAEALREADY = 24037
Public Const WSAENOTSOCK = 24038
Public Const WSAEDESTADDRREQ = 24039
Public Const WSAEMSGSIZE = 24040
Public Const WSAEPROTOTYPE = 24041
Public Const WSAENOPROTOOPT = 24042
Public Const WSAEPROTONOSUPPORT = 24043
Public Const WSAESOCKTNOSUPPORT = 24044
Public Const WSAEOPNOTSUPP = 24045
Public Const WSAEPFNOSUPPORT = 24046
Public Const WSAEAFNOSUPPORT = 24047
Public Const WSAEADDRINUSE = 24048
Public Const WSAEADDRNOTAVAIL = 24049
Public Const WSAENETDOWN = 24050
Public Const WSAENETUNREACH = 24051
Public Const WSAENETRESET = 24052
Public Const WSAECONNABORTED = 24053
Public Const WSAECONNRESET = 24054
Public Const WSAENOBUFS = 24055
Public Const WSAEISCONN = 24056
Public Const WSAENOTCONN = 24057
Public Const WSAESHUTDOWN = 24058
Public Const WSAETOOMANYREFS = 24059
Public Const WSAETIMEDOUT = 24060
Public Const WSAECONNREFUSED = 24061
Public Const WSAELOOP = 24062
Public Const WSAENAMETOOLONG = 24063
Public Const WSAEHOSTDOWN = 24064
Public Const WSAEHOSTUNREACH = 24065
Public Const WSAENOTEMPTY = 24066
Public Const WSAEPROCLIM = 24067
Public Const WSAEUSERS = 24068
Public Const WSAEDQUOT = 24069
Public Const WSAESTALE = 24070
Public Const WSAEREMOTE = 24071
Public Const WSASYSNOTREADY = 24091
Public Const WSAVERNOTSUPPORTED = 24092
Public Const WSANOTINITIALISED = 24093
Public Const WSAHOST_NOT_FOUND = 25001
Public Const WSATRY_AGAIN = 25002
Public Const WSANO_RECOVERY = 25003
Public Const WSANO_DATA = 25004
Public Const WSANO_ADDRESS = 25004

'
' RAS Control Actions
'
Public Const RAS_ACTION_CONNECT = 1
Public Const RAS_ACTION_DISCONNECT = 2
Public Const RAS_ACTION_RESET = 3

'
' RAS Control Connection States
'
Public Const RASCS_UNUSED = -1                ' The control is not in use
Public Const RASCS_OPENPORT = 0               ' The communications port is about to be opened
Public Const RASCS_PORTOPENED = 1             ' The communications port has been opened
Public Const RASCS_CONNECTDEVICE = 2          ' A device is about to be connected
Public Const RASCS_DEVICECONNECTED = 3        ' A device has been connected successfully
Public Const RASCS_ALLDEVICESCONNECTED = 4    ' All devices have been connected, physical link established
Public Const RASCS_AUTHENTICATE = 5           ' The authentication phase is starting
Public Const RASCS_AUTHNOTIFY = 6             ' An authentication event has occurred
Public Const RASCS_AUTHRETRY = 7              ' The client has requested another validation attempt
Public Const RASCS_AUTHCALLBACK = 8           ' The remote server has requested a callback number
Public Const RASCS_AUTHCHANGEPASSWORD = 9     ' The client has requested to change the password
Public Const RASCS_AUTHPROJECT = 10           ' The projection phase is starting
Public Const RASCS_AUTHLINKSPEED = 11         ' The link speed calculation phase is starting
Public Const RASCS_AUTHACK = 12               ' The authentication request has been acknowledged
Public Const RASCS_REAUTHENTICATE = 13        ' Re-authentication after a callback is starting
Public Const RASCS_AUTHENTICATED = 14         ' The client has successfully completed authentication
Public Const RASCS_PREPAREFORCALLBACK = 15    ' The link is about to disconnect in preperation for a callback
Public Const RASCS_WAITFORMODEMRESET = 16     ' The client is waiting for the modem to reset
Public Const RASCS_WAITFORCALLBACK = 17       ' The client is waiting for an incoming callback
Public Const RASCS_PROJECTED = 18             ' The projection phase has completed
Public Const RASCS_STARTAUTHENTICATION = 19   ' User authentication is being initiated
Public Const RASCS_CALLBACKCOMPLETE = 20      ' The client has been called back, resume authentication
Public Const RASCS_LOGONNETWORK = 21          ' The client is logging on to the network
Public Const RASCS_SUBENTRYCONNECTED = 22     ' A subentry has been connected (multilink connections)
Public Const RASCS_SUBENTRYDISCONNECTED = 23  ' A subentry has been disconnected (multilink connections)
Public Const RASCS_PAUSED = 4096              ' The client has paused for an interactive (terminal) session
Public Const RASCS_RETRYAUTHENTICATION = 4097 ' The client is in a re-authentication state
Public Const RASCS_CALLBACKSETBYCALLER = 4098 ' The client is in a callback state
Public Const RASCS_PASSWORDEXPIRED = 4099     ' The client is in a change password state
Public Const RASCS_CONNECTED = 8192           ' A successful connection has been established
Public Const RASCS_DISCONNECTED = 8193        ' The client has been disconnected or the connection attempt failed

'
' RAS Control Error Codes
'
' These error codes are returned by the LastError property and
' passed as an argument to the LastError event. These are the
' same codes returned by the RAS library, with 25000 added to the
' base value
'
Public Const ERROR_INVALID_PORT_HANDLE = 25601
Public Const ERROR_PORT_ALREADY_OPEN = 25602
Public Const ERROR_BUFFER_TOO_SMALL = 25603
Public Const ERROR_WRONG_INFO_SPECIFIED = 25604
Public Const ERROR_CANNOT_SET_PORT_INFO = 25605
Public Const ERROR_PORT_NOT_CONNECTED = 25606
Public Const ERROR_EVENT_INVALID = 25607
Public Const ERROR_DEVICE_DOES_NOT_EXIST = 25608
Public Const ERROR_DEVICETYPE_DOES_NOT_EXIST = 25609
Public Const ERROR_BUFFER_INVALID = 25610
Public Const ERROR_ROUTE_NOT_AVAILABLE = 25611
Public Const ERROR_ROUTE_NOT_ALLOCATED = 25612
Public Const ERROR_INVALID_COMPRESSION_SPECIFIED = 25613
Public Const ERROR_OUT_OF_BUFFERS = 25614
Public Const ERROR_PORT_NOT_FOUND = 25615
Public Const ERROR_ASYNC_REQUEST_PENDING = 25616
Public Const ERROR_ALREADY_DISCONNECTING = 25617
Public Const ERROR_PORT_NOT_OPEN = 25618
Public Const ERROR_PORT_DISCONNECTED = 25619
Public Const ERROR_NO_ENDPOINTS = 25620
Public Const ERROR_CANNOT_OPEN_PHONEBOOK = 25621
Public Const ERROR_CANNOT_LOAD_PHONEBOOK = 25622
Public Const ERROR_CANNOT_FIND_PHONEBOOK_ENTRY = 25623
Public Const ERROR_CANNOT_WRITE_PHONEBOOK = 25624
Public Const ERROR_CORRUPT_PHONEBOOK = 25625
Public Const ERROR_CANNOT_LOAD_STRING = 25626
Public Const ERROR_KEY_NOT_FOUND = 25627
Public Const ERROR_DISCONNECTION = 25628
Public Const ERROR_REMOTE_DISCONNECTION = 25629
Public Const ERROR_HARDWARE_FAILURE = 25630
Public Const ERROR_USER_DISCONNECTION = 25631
Public Const ERROR_INVALID_SIZE = 25632
Public Const ERROR_PORT_NOT_AVAILABLE = 25633
Public Const ERROR_CANNOT_PROJECT_CLIENT = 25634
Public Const ERROR_UNKNOWN = 25635
Public Const ERROR_WRONG_DEVICE_ATTACHED = 25636
Public Const ERROR_BAD_STRING = 25637
Public Const ERROR_REQUEST_TIMEOUT = 25638
Public Const ERROR_CANNOT_GET_LANA = 25639
Public Const ERROR_NETBIOS_ERROR = 25640
Public Const ERROR_SERVER_OUT_OF_RESOURCES = 25641
Public Const ERROR_NAME_EXISTS_ON_NET = 25642
Public Const ERROR_SERVER_GENERAL_NET_FAILURE = 25643
Public Const WARNING_MSG_ALIAS_NOT_ADDED = 25644
Public Const ERROR_AUTH_INTERNAL = 25645
Public Const ERROR_RESTRICTED_LOGON_HOURS = 25646
Public Const ERROR_ACCT_DISABLED = 25647
Public Const ERROR_PASSWD_EXPIRED = 25648
Public Const ERROR_NO_DIALIN_PERMISSION = 25649
Public Const ERROR_SERVER_NOT_RESPONDING = 25650
Public Const ERROR_FROM_DEVICE = 25651
Public Const ERROR_UNRECOGNIZED_RESPONSE = 25652
Public Const ERROR_MACRO_NOT_FOUND = 25653
Public Const ERROR_MACRO_NOT_DEFINED = 25654
Public Const ERROR_MESSAGE_MACRO_NOT_FOUND = 25655
Public Const ERROR_DEFAULTOFF_MACRO_NOT_FOUND = 25656
Public Const ERROR_FILE_COULD_NOT_BE_OPENED = 25657
Public Const ERROR_DEVICENAME_TOO_LONG = 25658
Public Const ERROR_DEVICENAME_NOT_FOUND = 25659
Public Const ERROR_NO_RESPONSES = 25660
Public Const ERROR_NO_COMMAND_FOUND = 25661
Public Const ERROR_WRONG_KEY_SPECIFIED = 25662
Public Const ERROR_UNKNOWN_DEVICE_TYPE = 25663
Public Const ERROR_ALLOCATING_MEMORY = 25664
Public Const ERROR_PORT_NOT_CONFIGURED = 25665
Public Const ERROR_DEVICE_NOT_READY = 25666
Public Const ERROR_READING_INI_FILE = 25667
Public Const ERROR_NO_CONNECTION = 25668
Public Const ERROR_BAD_USAGE_IN_INI_FILE = 25669
Public Const ERROR_READING_SECTIONNAME = 25670
Public Const ERROR_READING_DEVICETYPE = 25671
Public Const ERROR_READING_DEVICENAME = 25672
Public Const ERROR_READING_USAGE = 25673
Public Const ERROR_READING_MAXCONNECTBPS = 25674
Public Const ERROR_READING_MAXCARRIERBPS = 25675
Public Const ERROR_LINE_BUSY = 25676
Public Const ERROR_VOICE_ANSWER = 25677
Public Const ERROR_NO_ANSWER = 25678
Public Const ERROR_NO_CARRIER = 25679
Public Const ERROR_NO_DIALTONE = 25680
Public Const ERROR_IN_COMMAND = 25681
Public Const ERROR_WRITING_SECTIONNAME = 25682
Public Const ERROR_WRITING_DEVICETYPE = 25683
Public Const ERROR_WRITING_DEVICENAME = 25684
Public Const ERROR_WRITING_MAXCONNECTBPS = 25685
Public Const ERROR_WRITING_MAXCARRIERBPS = 25686
Public Const ERROR_WRITING_USAGE = 25687
Public Const ERROR_WRITING_DEFAULTOFF = 25688
Public Const ERROR_READING_DEFAULTOFF = 25689
Public Const ERROR_EMPTY_INI_FILE = 25690
Public Const ERROR_AUTHENTICATION_FAILURE = 25691
Public Const ERROR_PORT_OR_DEVICE = 25692
Public Const ERROR_NOT_BINARY_MACRO = 25693
Public Const ERROR_DCB_NOT_FOUND = 25694
Public Const ERROR_STATE_MACHINES_NOT_STARTED = 25695
Public Const ERROR_STATE_MACHINES_ALREADY_STARTED = 25696
Public Const ERROR_PARTIAL_RESPONSE_LOOPING = 25697
Public Const ERROR_UNKNOWN_RESPONSE_KEY = 25698
Public Const ERROR_RECV_BUF_FULL = 25699
Public Const ERROR_CMD_TOO_LONG = 25700
Public Const ERROR_UNSUPPORTED_BPS = 25701
Public Const ERROR_UNEXPECTED_RESPONSE = 25702
Public Const ERROR_INTERACTIVE_MODE = 25703
Public Const ERROR_BAD_CALLBACK_NUMBER = 25704
Public Const ERROR_INVALID_AUTH_STATE = 25705
Public Const ERROR_WRITING_INITBPS = 25706
Public Const ERROR_X25_DIAGNOSTIC = 25707
Public Const ERROR_ACCT_EXPIRED = 25708
Public Const ERROR_CHANGING_PASSWORD = 25709
Public Const ERROR_OVERRUN = 25710
Public Const ERROR_RASMAN_CANNOT_INITIALIZE = 25711
Public Const ERROR_BIPLEX_PORT_NOT_AVAILABLE = 25712
Public Const ERROR_NO_ACTIVE_ISDN_LINES = 25713
Public Const ERROR_NO_ISDN_CHANNELS_AVAILABLE = 25714
Public Const ERROR_TOO_MANY_LINE_ERRORS = 25715
Public Const ERROR_IP_CONFIGURATION = 25716
Public Const ERROR_NO_IP_ADDRESSES = 25717
Public Const ERROR_PPP_TIMEOUT = 25718
Public Const ERROR_PPP_REMOTE_TERMINATED = 25719
Public Const ERROR_PPP_NO_PROTOCOLS_CONFIGURED = 25720
Public Const ERROR_PPP_NO_RESPONSE = 25721
Public Const ERROR_PPP_INVALID_PACKET = 25722
Public Const ERROR_PHONE_NUMBER_TOO_LONG = 25723
Public Const ERROR_IPXCP_NO_DIALOUT_CONFIGURED = 25724
Public Const ERROR_IPXCP_NO_DIALIN_CONFIGURED = 25725
Public Const ERROR_IPXCP_DIALOUT_ALREADY_ACTIVE = 25726
Public Const ERROR_ACCESSING_TCPCFGDLL = 25727
Public Const ERROR_NO_IP_RAS_ADAPTER = 25728
Public Const ERROR_SLIP_REQUIRES_IP = 25729
Public Const ERROR_PROJECTION_NOT_COMPLETE = 25730
Public Const ERROR_PROTOCOL_NOT_CONFIGURED = 25731
Public Const ERROR_PPP_NOT_CONVERGING = 25732
Public Const ERROR_PPP_CP_REJECTED = 25733
Public Const ERROR_PPP_LCP_TERMINATED = 25734
Public Const ERROR_PPP_REQUIRED_ADDRESS_REJECTED = 25735
Public Const ERROR_PPP_NCP_TERMINATED = 25736
Public Const ERROR_PPP_LOOPBACK_DETECTED = 25737
Public Const ERROR_PPP_NO_ADDRESS_ASSIGNED = 25738
Public Const ERROR_CANNOT_USE_LOGON_CREDENTIALS = 25739
Public Const ERROR_TAPI_CONFIGURATION = 25740
Public Const ERROR_NO_LOCAL_ENCRYPTION = 25741
Public Const ERROR_NO_REMOTE_ENCRYPTION = 25742
Public Const ERROR_REMOTE_REQUIRES_ENCRYPTION = 25743
Public Const ERROR_IPXCP_NET_NUMBER_CONFLICT = 25744
Public Const ERROR_INVALID_SMM = 25745
Public Const ERROR_SMM_UNINITIALIZED = 25746
Public Const ERROR_NO_MAC_FOR_PORT = 25747
Public Const ERROR_SMM_TIMEOUT = 25748
Public Const ERROR_BAD_PHONE_NUMBER = 25749
Public Const ERROR_WRONG_MODULE = 25750
Public Const ERROR_INVALID_CALLBACK_NUMBER = 25751
Public Const ERROR_SCRIPT_SYNTAX = 25752
Public Const ERROR_HANGUP_FAILED = 25753
Public Const ERROR_BUNDLE_NOT_FOUND = 25754
Public Const ERROR_CANNOT_DO_CUSTOMDIAL = 25755
Public Const ERROR_DIAL_ALREADY_IN_PROGRESS = 25756
Public Const ERROR_RASAUTO_CANNOT_INITIALIZE = 25757
Public Const ERROR_CONNECTION_ALREADY_SHARED = 25758
Public Const ERROR_SHARING_CHANGE_FAILED = 25759
Public Const ERROR_SHARING_ROUTER_INSTALL = 25760
Public Const ERROR_SHARE_CONNECTION_FAILED = 25761
Public Const ERROR_SHARING_PRIVATE_INSTALL = 25762
Public Const ERROR_CANNOT_SHARE_CONNECTION = 25763
Public Const ERROR_NO_SMART_CARD_READER = 25764
Public Const ERROR_SHARING_ADDRESS_EXISTS = 25765
Public Const ERROR_NO_CERTIFICATE = 25766
Public Const ERROR_SHARING_MULTIPLE_ADDRESSES = 25767
Public Const ERROR_FAILED_TO_ENCRYPT = 25768
Public Const ERROR_BAD_ADDRESS_SPECIFIED = 25769
Public Const ERROR_CONNECTION_REJECT = 25770
Public Const ERROR_CONGESTION = 25771
Public Const ERROR_INCOMPATIBLE = 25772
Public Const ERROR_NUMBERCHANGED = 25773
Public Const ERROR_TEMPFAILURE = 25774
Public Const ERROR_BLOCKED = 25775
Public Const ERROR_DONOTDISTURB = 25776
Public Const ERROR_OUTOFORDER = 25777
Public Const ERROR_UNABLE_TO_AUTHENTICATE_SERVER = 25778
Public Const ERROR_SMART_CARD_REQUIRED = 25779
Public Const ERROR_INVALID_FUNCTION_FOR_ENTRY = 25780
Public Const ERROR_CERT_FOR_ENCRYPTION_NOT_FOUND = 25781
Public Const ERROR_SHARING_RRAS_CONFLICT = 25782
Public Const ERROR_SHARING_NO_PRIVATE_LAN = 25783
Public Const ERROR_NO_DIFF_USER_AT_LOGON = 25784
Public Const ERROR_NO_REG_CERT_AT_LOGON = 25785
Public Const ERROR_OAKLEY_NO_CERT = 25786
Public Const ERROR_OAKLEY_AUTH_FAIL = 25787
Public Const ERROR_OAKLEY_ATTRIB_FAIL = 25788
Public Const ERROR_OAKLEY_GENERAL_PROCESSING = 25789
Public Const ERROR_OAKLEY_NO_PEER_CERT = 25790
Public Const ERROR_OAKLEY_NO_POLICY = 25791
Public Const ERROR_OAKLEY_TIMED_OUT = 25792
Public Const ERROR_OAKLEY_ERROR = 25793
Public Const ERROR_UNKNOWN_FRAMED_PROTOCOL = 25794
Public Const ERROR_WRONG_TUNNEL_TYPE = 25795
Public Const ERROR_UNKNOWN_SERVICE_TYPE = 25796
Public Const ERROR_CONNECTING_DEVICE_NOT_FOUND = 25797
Public Const ERROR_NO_EAPTLS_CERTIFICATE = 25798


'
' File Encoding Constants
'
' File encoding and decoding constants for use with the DecodeFileEx
' and EncodeFileEx functions
'
Global Const FILE_ENCODE_DEFAULT = 0
Global Const FILE_ENCODE_UUCODE = 1
Global Const FILE_ENCODE_BASE64 = 2
Global Const FILE_ENCODE_COMPRESSED = 65536

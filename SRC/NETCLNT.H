/* NETCLNT.H - NetClient's Header file */

   /* Resource #defines */
#define IDBMP_CLIENT     8000
#define IDBMP_CLIENTDOWN 8001
#define IDBMP_CLIENTMONO 8003
#define IDBMP_CLIENTEGA  8006
#define IDDLG_ABOUT      200

#define IDTIME_TIMEOUT   1
#define IDTIME_FLUSH     2

#ifndef IDOK
#define IDOK 1
#endif

    // All of the property indicies
// Enumerations ofProperties in NetClient_Properties[]
// MUST BE IN SAME ORDER
enum enum_netclient_props
{
   /* Standard properties */

 IPROP_NECLIENT_CTLNAME,
 IPROP_NETCLIENT_PARENT,
 IPROP_NETCLIENT_INDEX,
 IPROP_NETCLIENT_TOP,
 IPROP_NETCLIENT_LEFT,
 IPROP_NETCLIENT_TAG,
 IPROP_NETCLIENT_HWND,
    /* Custom properties */
 IPROP_NETCLIENT_ABOUT,
 IPROP_NETCLIENT_SOCKET,
 IPROP_NETCLIENT_HOSTNAME,
 IPROP_NETCLIENT_HOSTADDR,
 IPROP_NETCLIENT_LOCALPORT,
 IPROP_NETCLIENT_REMOTEPORT,
 IPROP_NETCLIENT_RECVLINE,
 IPROP_NETCLIENT_SENDLINE,
 IPROP_NETCLIENT_LINE,
 IPROP_NETCLIENT_LINEDELIMITER,
 IPROP_NETCLIENT_RECVCOUNT,
 IPROP_NETCLIENT_SENDCOUNT,
 IPROP_NETCLIENT_RECVSIZE,
 IPROP_NETCLIENT_SENDSIZE,
 IPROP_NETCLIENT_RECVTHRESHOLD,
 IPROP_NETCLIENT_SENDTHRESHOLD,
 IPROP_NETCLIENT_CONNECT,
 IPROP_NETCLIENT_TIMEOUT,
 IPROP_NETCLIENT_ERRORNUMBER,
 IPROP_NETCLIENT_ERRORMESSAGE,
 IPROP_NETCLIENT_LOCALSERVICE,    // Change from alpha 2!
 IPROP_NETCLIENT_REMOTESERVICE,
 IPROP_NETCLIENT_VERSION,         // Change from alpha 6!
 IPROP_NETCLIENT_DEBUG,
 IPROP_NETCLIENT_HOSTALIASCOUNT,
 IPROP_NETCLIENT_HOSTALIASLIST,
 IPROP_NETCLIENT_HOSTADDRESSCNT,
 IPROP_NETCLIENT_HOSTADDRESSLIST,
 IPROP_NETCLIENT_HOST,
 IPROP_NETCLIENT_NONE1,           // Expansion for 7 properties!
 IPROP_NETCLIENT_NONE2,           // (pre VB2.0 properties that is)
 IPROP_NETCLIENT_NONE3,
 IPROP_NETCLIENT_NONE4,
 IPROP_NETCLIENT_NONE5,
 IPROP_NETCLIENT_NONE6,
 IPROP_NETCLIENT_NONE7,
     // These three are only available to VB2.00+ apps
 IPROP_NETCLIENT_RECVBLOCK,
 IPROP_NETCLIENT_SENDBLOCK,
 IPROP_NETCLIENT_BLOCK
};

typedef struct tagONERRORPARMS
{
 SHORT FAR *ErrorNumber;
 LPVOID Index;
} ONERRORPARMS;


// Enumerate the Events in NetClientEvents[]
// MUST BE IN SAME ORDER
enum enum_client_events
{
 IEVENT_NETCLIENT_ONCONNECT,
 IEVENT_NETCLIENT_ONRECV,
 IEVENT_NETCLIENT_ONSEND,
 IEVENT_NETCLIENT_ONCLOSE,
 IEVENT_NETCLIENT_ONTIMEOUT,
 IEVENT_NETCLIENT_ONERROR
};

// Errors sent back to VB (and .RC identifiers)
#define ERR_None                0

// These are sent ONLY to the OnError Event!
//#define ERR_Base                100
/*#define ERR_NoWinSock           100
#define ERR_BadVersion          101
#define ERR_NotConn             102
#define ERR_AlreadyConn         103
#define ERR_RecvNotEmpty        104
#define ERR_SendOverFlow        105
#define ERR_BadHostAddr         106
#define ERR_NoTimers            107
#define ERR_InvalidString       108
#define ERR_HostUnknown         109
#define ERR_RecvBuffer          110
#define ERR_SendBuffer          111
#define ERR_CantSet             112
#define ERR_Last                113
#define ERR_BadSocket           114
#define ERR_BufferNotEmpty      115
*/
//#define NETBASEERR              200


    /* Miscellaneous #defines for defaults and odd variables */
#define WSA_VERSION_NEEDED  0x0101

// Defaults for properties
#define DEFAULT_TIMEOUT       30000
#define DEFAULT_PORT          0
#define DEFAULT_RECVSIZE      8192
#define DEFAULT_SENDSIZE      4096
#define DEFAULT_RECVTHRESHOLD 1
#define DEFAULT_SENDTHRESHOLD 0

// No more than 20 host aliases or addresses!
#define MAX_HOSTTABLESIZE     20

typedef struct tagHOSTTABLE
{
 int iHostCount;
 MOLE mHost[MAX_HOSTTABLESIZE];
} HOSTTABLE;

    /* PROPERTIES */
typedef struct tagNETCLIENT
{ // When in Rome...
 MOLE  mHostName;
 MOLE  mHostAddr;

 SHORT sRemotePort;      // remote port to connect to
 SHORT sLocalPort;       // local port to connect to

 HSZ   hszLineDelimiter; // Delimiter for Lines
 HSZ   hszRawLineDelimiter; // wsprint() parsed LineDelimiter

 HLSTR hlRecvBlock;      // These CAN have embedded Zeros (VB2.00+ only)
 HLSTR hlSendBlock;      // Everything is sent "raw" - perfect for binary

 SHORT sRecvThreshold;   // Threshold before which OnRecv is called
 SHORT sSendThreshold;   // Low-water mark before OnSend is triggered
 SHORT sRecvSize;        // Total size of RecvBuffer
 SHORT sSendSize;        // Total size of SendBuffer
 SHORT sRecvCount;       // Current bytes waiting processing
 SHORT sSendCount;       // Current bytes waiting to be sent
 BOOL  bConnect;         // TRUE if connection exists, FALSE otherwise
 SHORT sTimeOut;         // Countdown to next timeout

 SHORT sErrorNumber;     // ErrorNumber for OnError

    // elements not used, just markers (to make VB happy)
 HSZ   hszAbout;

    // "Hidden" properties - nonpublic (OOP! Eeek!)
 SOCKET sSocket;      // Socket of any connection
 SOCKADDR_IN saHost;  // Address structure to Host

 BOOL   fCanSend;     // Flag: Enabled during FD_WRITE after WOULDBLOCK
 HANDLE hRecvBuffer;  // handle to the receive buffer
 HANDLE hSendBuffer;  // handle ro the send buffer
 HANDLE hRecvTemp;    // handle to the temporary receive space
 HANDLE hSendTemp;    // handle to the temporary send space
 LPSTR  lpRecvBuffer; // locked pointer to the receive buffer
 LPSTR  lpSendBuffer; // locked pointer to the send buffer
 LPSTR  lpRecvTemp;   // locked pointer to the temporary receive space
 LPSTR  lpSendTemp;   // locked pointer to the temporary send space

 SHORT  sRecvTempSize;// size of temporary receive space
 SHORT  sSendTempSize;// size of temporary send space
 SHORT  sTimeOutStore;// Storage for resetting TimeOut when it gets to 0

 HOSTTABLE htHostAddressList;
 HOSTTABLE htHostAliasList;
} NETCLIENT;

typedef NETCLIENT FAR * LPNETCLIENT;


#ifdef WSANet_C

    // Anti-Ugly code macro
#define VBGETOFFSET(arg) ((USHORT)&(((NETCLIENT *)0)->arg))

    // About property - open up the About dialog box
PROPINFO Property_About =
{
 "About",
 DT_SHORT | PF_fGetData | PF_fSetData | PF_fNoRuntimeW
 | PF_fGetHszMsg | PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // Socket property - the actual socket
PROPINFO Property_Socket =
{
 "Socket",
 DT_SHORT | PF_fGetData | PF_fSetMsg | PF_fNoInitDef,
 VBGETOFFSET(sSocket),
 0, 0, NULL, 0
};

    // HostName property - DNS name of HostIP
// Changing this will alter HostIP to the appropriate
// internet address in dotted string notation.
// You will not have to be connected.
PROPINFO Property_HostName =
{
 "HostName",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg | PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // HostAddr property - IP address of HostName
// Changing this will alter HostName to the appropriate
// name.
// You will not have to be connected.
PROPINFO Property_HostAddr =
{
 "HostAddr",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg | PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // LocalPort Property - TCP port on the Local host
PROPINFO Property_LocalPort =
{
 "LocalPort",
 DT_SHORT | PF_fGetData | PF_fSetMsg |
     PF_fNoInitDef,
 VBGETOFFSET(sLocalPort),
 0, 0, NULL, 0
};

    // LocalService Property - Service name for LocalPort
PROPINFO Property_LocalService =
{
 "LocalService",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg | PF_fNoInitDef,
 0,          // DUMMY! (we handle set/get with messages)
 0, 0, NULL, 0
};

    // RemotePort Property - TCP port on the remote host
PROPINFO Property_RemotePort =
{
 "RemotePort",
 DT_SHORT | PF_fGetData | PF_fSetMsg |
    PF_fSaveData | PF_fDefVal,
 VBGETOFFSET(sRemotePort),
 0, DEFAULT_PORT, NULL, 0
};

    // RemoteService Property - Service name for RemotePort
PROPINFO Property_RemoteService =
{
 "RemoteService",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg | PF_fSaveData |
    PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // RecvLine Property
PROPINFO Property_RecvLine =
{
 "RecvLine",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg |
    PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // SendLine Property - not shown in dialog
// SetMsg/SetData means that VB will set the element AND tell us
PROPINFO Property_SendLine =
{
 "SendLine",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg |
    PF_fNoInitDef | PF_fGetHszMsg,
 0,
 0, 0, NULL, 0
};

        // Line Property - not visible in design time
// This is the "default" property for this control
// Get: same as Get on RecvLine
// Set: same as Set on SendLine
// Reason: Allows 'string = NetClient1' and 'NetClient1 = "output"'
PROPINFO Property_Line =
{
 "Line",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg |
    PF_fNoInitDef | PF_fNoShow,
 0,              // Dummy, dummy!
 0, 0, NULL, 0
};

    // RecvBlock Property
PROPINFO Property_RecvBlock =
{
 "RecvBlock",
 DT_HLSTR | PF_fGetMsg | PF_fSetMsg    |
    PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // SendBlock Property - not shown in dialog
PROPINFO Property_SendBlock =
{
 "SendBlock",
 DT_HLSTR | PF_fGetData | PF_fSetMsg |
    PF_fNoInitDef | PF_fGetHszMsg,
 VBGETOFFSET(hlSendBlock),
 0, 0, NULL, 0
};

        // Block Property - not visible in design time
// This is a convenience property for this control
// Get: same as Get on RecvBlock
// Set: same as Set on SendBlock
// Reason: Allows single property to do binary input and output
PROPINFO Property_Block =
{
 "Block",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg |
    PF_fNoInitDef | PF_fNoShow,
 0,              // Dummy, dummy!
 0, 0, NULL, 0
};

    // LineDelimiter Property
PROPINFO Property_LineDelimiter =
{
 "LineDelimiter",
 DT_HSZ | PF_fGetData | PF_fSetData | PF_fSetMsg |
    PF_fSaveData,
 VBGETOFFSET(hszLineDelimiter),
 0, 0, NULL, 0
};


    // RecvCount Property - the valid bytes in RecvBuffer
PROPINFO Property_RecvCount =
{
 "RecvCount",
 DT_SHORT | PF_fGetMsg | PF_fSetMsg |
    PF_fNoInitDef,
 VBGETOFFSET(sRecvCount),
 0, 0, NULL, 0
};

    // SendCount Property - the valid bytes in SendBuffer
// not shown in dialog.
PROPINFO Property_SendCount =
{
 "SendCount",
 DT_SHORT | PF_fGetMsg | PF_fSetMsg |
    PF_fNoInitDef,
 VBGETOFFSET(sSendCount),
 0, 0, NULL, 0
};

    // RecvSize Property - the size of the RecvBuffer
PROPINFO Property_RecvSize =
{
 "RecvSize",
 DT_SHORT | PF_fGetData | PF_fSetMsg |
    PF_fSaveData | PF_fDefVal,
 VBGETOFFSET(sRecvSize),
 0, DEFAULT_RECVSIZE, NULL, 0
};

    // SendSize Property - the size of the SendBuffer
PROPINFO Property_SendSize =
{
 "SendSize",
 DT_SHORT | PF_fGetData | PF_fSetMsg |
    PF_fSaveData | PF_fDefVal,
 VBGETOFFSET(sSendSize),
 0, DEFAULT_SENDSIZE, NULL, 0
};

    // RecvThreshold Property - cutoff before Event noticed
// If data never passes this set threshold, no OnRecv is generated
PROPINFO Property_RecvThreshold =
{
 "RecvThreshold",
 DT_SHORT | PF_fSetData | PF_fGetData |
    PF_fSaveData | PF_fDefVal,
 VBGETOFFSET(sRecvThreshold),
 0, DEFAULT_RECVTHRESHOLD, NULL, 0
};

    // SendThreshold Property - cutoff before Event noticed
// If data never drops below this set threshold, no OnSend is generated
PROPINFO Property_SendThreshold =
{
 "SendThreshold",
 DT_SHORT | PF_fSetData | PF_fGetData |
    PF_fSaveData | PF_fDefVal,
 VBGETOFFSET(sSendThreshold),
 0, DEFAULT_SENDTHRESHOLD, NULL, 0
};

    // Connect Property - Initiates connection
// Remains false until connection exists, then Event
// is triggered.
PROPINFO Property_Connect =
{
 "Connect",
 DT_BOOL | PF_fSetMsg | PF_fGetData |
    PF_fNoInitDef,
 VBGETOFFSET(bConnect),
 0, FALSE, NULL, 0
};

    // TimeOut Property - timeout in milliseconds
// Do not exceed 32768 (~32 seconds)
PROPINFO Property_TimeOut =
{
 "TimeOut",
 DT_SHORT | PF_fSetMsg | PF_fGetData |
    PF_fSaveData | PF_fDefVal,
 VBGETOFFSET(sTimeOut),
 0, 0, NULL, 0
};

    // ErrorNumber Property
// Hold the last WSAGetLastError() or a NetClient specific error
PROPINFO Property_ErrorNumber =
{
 "ErrorNumber",
 DT_SHORT | PF_fSetData | PF_fGetData |
    PF_fNoInitDef,
 VBGETOFFSET(sErrorNumber),
 0, 0, NULL, 0
};

    // ErrorMessage Property
// Hold the last ErrorNumber's error message
PROPINFO Property_ErrorMessage =
{
 "ErrorMessage",
 DT_HSZ | PF_fSetMsg | PF_fGetMsg |
    PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // Version Property
// Show the Version information.
PROPINFO Property_Version =
{
 "Version",
 DT_HSZ | PF_fSetMsg | PF_fGetMsg |
    PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // Debug Property
// Show the Debug sockopt() state
PROPINFO Property_Debug =
{
 "Debug",
 DT_BOOL | PF_fSetMsg | PF_fGetMsg |
    PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // HostAliasCount Property
// Report the number of aliases for HostName
// (index into HostAliasList())
PROPINFO Property_HostAliasCount =
{
 "HostAliasCount",
 DT_SHORT | PF_fSetMsg | PF_fGetMsg |
    PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // HostAliasList Property
// References array of aliases to HostName
// (index with HostAliasCount)
PROPINFO Property_HostAliasList =
{
 "HostAliasList",
 DT_HSZ | PF_fPropArray | PF_fNoShow | PF_fSetMsg |
    PF_fGetMsg | PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // HostAddressCount Property
// Report the number of addresses for HostName
// (index into HostAddressList())
PROPINFO Property_HostAddressCount =
{
 "HostAddressCount",
 DT_SHORT | PF_fSetMsg | PF_fGetMsg |
    PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // HostAddressList Property
// References array of aliases to HostName
// (index with HostAddressCount)
PROPINFO Property_HostAddressList =
{
 "HostAddressList",
 DT_HSZ | PF_fPropArray | PF_fNoShow | PF_fSetMsg |
    PF_fGetMsg | PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // Host Property
// Acts the same as HostAddr/HostName, but tries the former
// before the latter (can handle DNS & I.P)
PROPINFO Property_Host =
{
 "Host",
 DT_HSZ | PF_fSetMsg | PF_fGetMsg | PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // All of the VB1.00 properties - in dialog order
PPROPINFO NetClient_Properties_VB1[] =
{
 PPROPINFO_STD_CTLNAME,
 PPROPINFO_STD_PARENT,
 PPROPINFO_STD_INDEX,
 PPROPINFO_STD_TOP,
 PPROPINFO_STD_LEFT,
 PPROPINFO_STD_TAG,
 PPROPINFO_STD_HWND,
 &Property_About,
 &Property_Socket,
 &Property_HostName,
 &Property_HostAddr,
 &Property_LocalPort,
 &Property_RemotePort,
 &Property_RecvLine,
 &Property_SendLine,
 &Property_Line,
 &Property_LineDelimiter,
 &Property_RecvCount,
 &Property_SendCount,
 &Property_RecvSize,
 &Property_SendSize,
 &Property_RecvThreshold,
 &Property_SendThreshold,
 &Property_Connect,
 &Property_TimeOut,
 &Property_ErrorNumber,
 &Property_ErrorMessage,
 &Property_LocalService,
 &Property_RemoteService,
 &Property_Version,
 &Property_Debug,
 &Property_HostAliasCount,
 &Property_HostAliasList,
 &Property_HostAddressCount,
 &Property_HostAddressList,
 &Property_Host,
 NULL
};

    // All of the VB2.00 properties - in dialog order
PPROPINFO NetClient_Properties_VB2[] =
{
 PPROPINFO_STD_NAME,
 PPROPINFO_STD_PARENT,
 PPROPINFO_STD_INDEX,
 PPROPINFO_STD_TOPNORUN,
 PPROPINFO_STD_LEFTNORUN,
 PPROPINFO_STD_TAG,
 PPROPINFO_STD_HWND,
 &Property_About,
 &Property_Socket,
 &Property_HostName,
 &Property_HostAddr,
 &Property_LocalPort,
 &Property_RemotePort,
 &Property_RecvLine,
 &Property_SendLine,
 &Property_Line,
 &Property_LineDelimiter,
 &Property_RecvCount,
 &Property_SendCount,
 &Property_RecvSize,
 &Property_SendSize,
 &Property_RecvThreshold,
 &Property_SendThreshold,
 &Property_Connect,
 &Property_TimeOut,
 &Property_ErrorNumber,
 &Property_ErrorMessage,
 &Property_LocalService,
 &Property_RemoteService,
 &Property_Version,
 &Property_Debug,
 &Property_HostAliasCount,
 &Property_HostAliasList,
 &Property_HostAddressCount,
 &Property_HostAddressList,
 &Property_Host,
 PPROPINFO_STD_NONE,        // Expansion room!
 PPROPINFO_STD_NONE,
 PPROPINFO_STD_NONE,
 PPROPINFO_STD_NONE,
 PPROPINFO_STD_NONE,
 PPROPINFO_STD_NONE,
 PPROPINFO_STD_NONE,
     // These three are available only to VB2.00 apps
 &Property_RecvBlock,
 &Property_SendBlock,
 &Property_Block,
 NULL
};


// Parameters to the Events.
// Currently all are (), due to random unexplained
// GPFs during testing. See "misc.h" to see what
// was trimmed.

EVENTINFO Event_OnConnect =
{
 "OnConnect",
 0,
 0,
 NULL,
 NULL
};

EVENTINFO Event_OnRecv =
{
 "OnRecv",
 0,
 0,
 NULL,
 NULL
};

    // The OnSend Event structure
EVENTINFO Event_OnSend =
{
 "OnSend",
 0,
 0,
 NULL,
 NULL
};

    // The OnClose Event structure
EVENTINFO Event_OnClose =
{
 "OnClose",
 0,
 0,
 NULL,
 NULL
};

EVENTINFO Event_OnTimeOut =
{
 "OnTimeOut",
 0,
 0,
 NULL,
 NULL
};


WORD Paramtypes_OnError[] = { ET_I2 };

EVENTINFO Event_OnError =
{
 "OnError",
 1,
 2,
 Paramtypes_OnError,
 "iErrorNumber As Integer"
};

/*
EVENTINFO Event_OnError =
{
 "OnError",
 0,
 0,
 NULL,
 NULL
};
*/
    // The NetClient Events listed in order
PEVENTINFO NetClient_Events[] =
{
 &Event_OnConnect,
 &Event_OnRecv,
 &Event_OnSend,
 &Event_OnClose,
 &Event_OnTimeOut,
 &Event_OnError,
 NULL
};

#ifdef VB100_CDK

    // The MODEL structure for the NetClient control (VB1.00)
MODEL modelNetClient_VB1 =
{
 VB100_VERSION,
 MODEL_fDesInteract | MODEL_fLoadMsg,
 (PCTLPROC)NetClientCtlProc,
 CS_VREDRAW | CS_HREDRAW,
 WS_CHILD | WS_VISIBLE,
 sizeof(NETCLIENT),
 IDBMP_CLIENT,
 "NetClient",
 "NetClient",
 NULL,
 NetClient_Properties_VB1,
 NetClient_Events,
 IPROP_NETCLIENT_LINE,
 IEVENT_NETCLIENT_ONCONNECT
};

#endif /* VB100_CDK */


#ifdef VB200_CDK

    // The MODEL structure for the NetClient control! (VB2.00)
MODEL modelNetClient_VB2 =
{
 VB200_VERSION,
 MODEL_fDesInteract | MODEL_fInvisAtRun | MODEL_fLoadMsg,
 (PCTLPROC)NetClientCtlProc,
 CS_VREDRAW | CS_HREDRAW,
 WS_CHILD | WS_VISIBLE,
 sizeof(NETCLIENT),
 IDBMP_CLIENT,
 "NetClient",
 "NetClient",
 NULL,
 NetClient_Properties_VB2,
 NetClient_Events,
 IPROP_NETCLIENT_LINE,
 IEVENT_NETCLIENT_ONCONNECT,
 IPROP_NETCLIENT_LINE
};

#endif /* VB200_CDK */


#ifdef VB300_CDK

    // The MODEL structure for the NetClient control (VB3.00+)
MODEL modelNetClient_VB3 =
{
 VB300_VERSION,
 MODEL_fDesInteract | MODEL_fInvisAtRun | MODEL_fLoadMsg,
 (PCTLPROC)NetClientCtlProc,
 CS_VREDRAW | CS_HREDRAW,
 WS_CHILD | WS_VISIBLE,
 sizeof(NETCLIENT),
 IDBMP_CLIENT,
 "NetClient",
 "NetClient",
 NULL,
 NetClient_Properties_VB2,  // This will most likely change
 NetClient_Events,
 IPROP_NETCLIENT_LINE,
 IEVENT_NETCLIENT_ONCONNECT,
 IPROP_NETCLIENT_LINE,
 WSANET_VERSION
};

#endif /* VB300_CDK */

    // All of the help "context-id's" are referenced through
    // this table
static WORD wNetClientHelpProps[] =
{
 PROPERTY_NAME,
 PROPERTY_PARENT,
 PROPERTY_INDEX,
 PROPERTY_TOP,
 PROPERTY_LEFT,
 PROPERTY_TAG,
 PROPERTY_HWND,
    /* Custom properties */
 PROPERTY_ABOUT,
 NC_PROP_SOCKET,
 NC_PROP_HOSTNAME,
 NC_PROP_HOSTADDR,
 NC_PROP_LOCALPORT,
 NC_PROP_REMOTEPORT,
 NC_PROP_RECVLINE,
 NC_PROP_SENDLINE,
 NC_PROP_LINE,
 NC_PROP_LINEDELIMITER,
 NC_PROP_RECVCOUNT,
 NC_PROP_SENDCOUNT,
 NC_PROP_RECVSIZE,
 NC_PROP_SENDSIZE,
 NC_PROP_RECVTHRESHOLD,
 NC_PROP_SENDTHRESHOLD,
 NC_PROP_CONNECT,
 NC_PROP_TIMEOUT,
 NC_PROP_ERRORNUMBER,
 NC_PROP_ERRORMESSAGE,
 NC_PROP_LOCALSERVICE,    // Change from alpha 2!
 NC_PROP_REMOTESERVICE,
 NC_PROP_VERSION,
 NC_PROP_DEBUG,
 NC_PROP_HOSTALIASCOUNT,
 NC_PROP_HOSTALIASLIST,
 NC_PROP_HOSTADDRESSCNT,
 NC_PROP_HOSTADDRESSLIST,
 NC_PROP_HOST,
 PROPERTY_UNUSED,
 PROPERTY_UNUSED,
 PROPERTY_UNUSED,
 PROPERTY_UNUSED,
 PROPERTY_UNUSED,
 PROPERTY_UNUSED,
 PROPERTY_UNUSED,
     // These three are only available to VB2.00+ apps
 NC_PROP_RECVBLOCK,
 NC_PROP_SENDBLOCK,
 NC_PROP_BLOCK,
 NULL
};

static WORD wNetClientHelpEvents[] =
{
 NC_EVNT_ONCONNECT,
 NC_EVNT_ONRECV,
 NC_EVNT_ONSEND,
 NC_EVNT_ONCLOSE,
 NC_EVNT_ONTIMEOUT,
 NC_EVNT_ONERROR,
 NULL
};

#endif /* WSANet_C */

/* End of NetClnt.H */

/* NetSrvr.H - NetServer control header file */

#ifndef NetSrvr_H
#define NetSrvr_H

#define IDBMP_SERVER     7000
#define IDBMP_SERVERDOWN 7001
#define IDBMP_SERVERMONO 7002
#define IDBMP_SERVEREGA  7003

#define DEFAULT_QUEUESIZE 5

    // All of the property indicies
// Enumerations of Properties in NetServer_Properties[]
// MUST BE IN SAME ORDER
enum enum_netserver_props
{
    /* Standard properties */
 IPROP_NETSERVER_CTLNAME,
 IPROP_NETSERVER_PARENT,
 IPROP_NETSERVER_INDEX,
 IPROP_NETSERVER_TOP,
 IPROP_NETSERVER_LEFT,
 IPROP_NETSERVER_TAG,
 IPROP_NETSERVER_HWND,

 IPROP_NETSERVER_ABOUT,
 IPROP_NETSERVER_SOCKET,
 IPROP_NETSERVER_LISTEN,
 IPROP_NETSERVER_LOCALPORT,
 IPROP_NETSERVER_LOCALSERVICE,
 IPROP_NETSERVER_ERRORNUMBER,
 IPROP_NETSERVER_ERRORMESSAGE,
 IPROP_NETSERVER_QUEUESIZE,
 IPROP_NETSERVER_VERSION,
 IPROP_NETSERVER_DEBUG
};


// Enumerate the Events in NetServerEvents[]
// MUST BE IN SAME ORDER
enum enum_server_events
{
 IEVENT_NETSERVER_ONACCEPT,
 IEVENT_NETSERVER_ONERROR
};


// Errors sent back to VB (and .RC identifiers)
#define ERR_None                0

    // Returnable Errors within VB
#define ERR_OUTOFMEMORY         7

    // NetServer's Control structure
typedef struct tagNETSERVER
{
 BOOL   bListen;
 SOCKET sSocket;
 SHORT  sLocalPort;
 SHORT  sErrorNumber;
 SHORT  sQueueSize;
} NETSERVER;

typedef NETSERVER FAR * LPNETSERVER;


#ifdef NetSrvr_C

typedef struct tagONERRORPARMS
{
 SHORT FAR *ErrorNumber;
 LPVOID Index;
} ONERRORPARMS;

#endif /* NetSrvr_C */


typedef struct tagONACCEPTPARMS
{
 SHORT FAR *RemotePort;
 HLSTR PeerAddr;
 SHORT FAR *Socket;
 LPVOID Index;
} ONACCEPTPARMS;


#ifdef WSANet_C

    // Anti-Ugly code macro
#define VBGETNETSERVEROFFSET(arg) ((USHORT)&(((NETSERVER *)0)->arg))

    // About property
PROPINFO serverProperty_About =
{
 "About",
 DT_SHORT | PF_fGetMsg | PF_fSetMsg | PF_fNoRuntimeW
 | PF_fGetHszMsg | PF_fNoInitDef,
 0,           // DUMMY! (we handle set/get with messages)
 0, 0, NULL, 0
};

    // Socket property - the actual socket
PROPINFO serverProperty_Socket =
{
 "Socket",
 DT_SHORT | PF_fGetData | PF_fSetMsg | PF_fNoInitDef,
 VBGETNETSERVEROFFSET(sSocket),
 0, 0, NULL, 0
};

    // Listen property
PROPINFO serverProperty_Listen =
{
 "Listen",
 DT_BOOL | PF_fGetData | PF_fSetMsg | PF_fNoInitDef,
 VBGETNETSERVEROFFSET(bListen),
 0, 0, NULL, 0
};

    // LocalPort Property - TCP port on the Local host
PROPINFO serverProperty_LocalPort =
{
 "LocalPort",
 DT_SHORT | PF_fGetData | PF_fSetMsg |
     PF_fNoInitDef,
 VBGETNETSERVEROFFSET(sLocalPort),
 0, 0, NULL, 0
};

    // LocalService Property - Service name for LocalPort
PROPINFO serverProperty_LocalService =
{
 "LocalService",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg | PF_fNoInitDef,
 0,            // DUMMY! (we handle set/get with messages)
 0, 0, NULL, 0
};

    // ErrorNumber Property
// Hold the last WSAGetLastError() or a NetServer specific error
PROPINFO serverProperty_ErrorNumber =
{
 "ErrorNumber",
 DT_SHORT | PF_fSetData | PF_fGetData |
    PF_fNoInitDef,
 VBGETNETSERVEROFFSET(sErrorNumber),
 0, 0, NULL, 0
};

    // ErrorMessage Property
// Hold the last ErrorNumber's error message
PROPINFO serverProperty_ErrorMessage =
{
 "ErrorMessage",
 DT_HSZ | PF_fSetMsg | PF_fGetMsg |
    PF_fNoInitDef,
 0,           // DUMMY! (we handle set/get with messages)
 0, 0, NULL, 0
};

    // QueueSize Property
PROPINFO serverProperty_QueueSize =
{
 "QueueSize",
 DT_SHORT | PF_fSetMsg | PF_fGetData | PF_fSaveData,
 VBGETNETSERVEROFFSET(sQueueSize),
 0, DEFAULT_QUEUESIZE, NULL, 0
};

    // Version Property
// Show the Version information.
PROPINFO serverProperty_Version =
{
 "Version",
 DT_HSZ | PF_fSetMsg | PF_fGetMsg |
    PF_fNoInitDef,
 0,           // DUMMY! (we handle set/get with messages)
 0, 0, NULL, 0
};

    // Debug Property
// Show the Debug sockopt() state
PROPINFO serverProperty_Debug =
{
 "Debug",
 DT_BOOL | PF_fSetMsg | PF_fGetMsg |
    PF_fNoInitDef,
 0,           // DUMMY! (we handle set/get with messages)
 0, 0, NULL, 0
};

    // All of the VB1.00 properties - in dialog order
PPROPINFO NetServer_Properties[] =
{
 PPROPINFO_STD_CTLNAME,
 PPROPINFO_STD_PARENT,
 PPROPINFO_STD_INDEX,
 PPROPINFO_STD_TOP,
 PPROPINFO_STD_LEFT,
 PPROPINFO_STD_TAG,
 PPROPINFO_STD_HWND,

 &serverProperty_About,
 &serverProperty_Socket,
 &serverProperty_Listen,
 &serverProperty_LocalPort,
 &serverProperty_LocalService,
 &serverProperty_ErrorNumber,
 &serverProperty_ErrorMessage,
 &serverProperty_QueueSize,
 &serverProperty_Version,
 &serverProperty_Debug,
 NULL
};

WORD serverParamtypes_OnError[] = { ET_I2 };

EVENTINFO serverEvent_OnError =
{
 "OnError",
 1,
 2,
 serverParamtypes_OnError,
 "iErrorNumber As Integer"
};

WORD serverParamtypes_OnAccept[] = { ET_I2 , ET_HLSTR, ET_I2 };

EVENTINFO serverEvent_OnAccept =
{
 "OnAccept",
 3,
 6,
 serverParamtypes_OnAccept,
 "Socket As Integer, PeerAddr As String, RemotePort As Integer"
};

    // The NetServer Events listed in order
PEVENTINFO NetServer_Events[] =
{
 &serverEvent_OnAccept,
 &serverEvent_OnError,
 NULL
};

    // The MODEL structure for the NetClient control (VB1.00)
MODEL modelNetServer =
{
 VB100_VERSION,
 MODEL_fDesInteract | MODEL_fLoadMsg,
 (PCTLPROC)NetServerCtlProc,
 CS_VREDRAW | CS_HREDRAW,
 WS_CHILD | WS_VISIBLE,
 sizeof(NETSERVER),
 IDBMP_SERVER,
 "NetServer",
 "NetServer",
 NULL,
 NetServer_Properties,
 NetServer_Events,
 IPROP_NETSERVER_LISTEN,
 IEVENT_NETSERVER_ONACCEPT
};

#define MODELLIST_NETSERVER (LPMODEL)&modelNetServer


static WORD wNetServerHelpProps[] =
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
 NS_PROP_SOCKET,
 NS_PROP_LISTEN,
 NS_PROP_LOCALPORT,
 NS_PROP_LOCALSERVICE,
 NS_PROP_ERRORNUMBER,
 NS_PROP_ERRORMESSAGE,
 NS_PROP_QUEUESIZE,
 NS_PROP_VERSION,
 NS_PROP_DEBUG
};

static WORD wNetServerHelpEvents[] =
{
 NS_EVNT_ONACCEPT,
 NS_EVNT_ONERROR
};

#endif /* WSANet_C */

#endif /* NetSrvr_H */

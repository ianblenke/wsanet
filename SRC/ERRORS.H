/* Errors.H - Visual Basic errors that can be returned */

#define VBERR_OUTOFMEMORY          7
#define VBERR_INVALIDPROPERTYVALUE 380
#define VBERR_BADVBINDEX           381
#define VBERR_READONLY             383
#define VBERR_WRITEONLY            394

#define WSAERR_BASE                20000

#define WSAERR_BadHostAddr         20000
#define WSAERR_AlreadyConn         20001
#define WSAERR_NoTimers            20002
#define WSAERR_RecvBuffer          20003
#define WSAERR_SendOverFlow        20004
#define WSAERR_NotConn             20005

#define IDS_WRONGVER               11

/*
    ERR_NoWinSock,     "No WinSock.DLL found on your system"
    ERR_BadVersion,    "Windows Sockets v1.1 not found"
    ERR_NotConn,       "Not connected to a remote host"
    ERR_RecvNotEmpty,  "Recieve buffer is not empty, could not reallocaate buffer"
    ERR_InvalidString, "Invalid string"
    ERR_HostUnknown,   "Unknown host"
    ERR_SendBuffer,    "Send buffer non-existant!"
    ERR_CantSet,       "You aren't allowed to set this property!"
    ERR_BufferNotEmpty,   "Buffer not empty, cannot reset 'size'"
*/
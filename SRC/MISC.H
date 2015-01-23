/* The event parameter routines were stripped due to unexplainable
   random GPFs */

//WSNETC.H

    /* Typedefs for Event Fire parameters */
/*typedef struct tagONCONNECTPARMS
{
 SHORT FAR *LocalPort;
 SHORT FAR *RemotePort;
 LPVOID Index;
} ONCONNECTPARMS;

typedef struct tagONRECVPARMS
{
 SHORT FAR *RecvCount;
 LPVOID Index;
} ONRECVPARMS;

typedef struct tagONTIMEOUTPARMS
{
 SHORT FAR *TimeOut;
 LPVOID Index;
} ONTIMEOUTPARMS;

typedef struct tagONERRORPARMS
{
 SHORT FAR *ErrorNumber;
 LPVOID Index;
} ONERRORPARMS;
*/


//INIT.C

    // Parameters to the OnConnect Event
/*WORD Paramtypes_OnConnect[] = { ET_I2, ET_I2 };

    // The OnConnect Event structure
EVENTINFO Event_OnConnect =
{
 "OnConnect",
 2,
 2,
 Paramtypes_OnConnect,
 "RemotePort As Integer, LocalPort As Integer"
}; */

    // Parameters to the OnRecv Event
/*WORD Paramtypes_OnRecv[] = { ET_I2 };

    // The OnRecv Event structure
EVENTINFO Event_OnRecv =
{
 "OnRecv",
 1,
 1,
 Paramtypes_OnRecv,
 "RecvCount As Integer"
}; */

    // Parameters to the OnTimeOut Event
/*WORD Paramtypes_OnTimeOut[] = { ET_I2 };

    // The OnRecv Event structure
EVENTINFO Event_OnTimeOut =
{
 "OnTimeOut",
 1,
 1,
 Paramtypes_OnTimeOut,
 "TimeOut As Integer"
}; */

/*WORD Paramtypes_OnError[] = { ET_I2 };

EVENTINFO Event_OnError =
{
 "OnError",
 1,
 2,
 Paramtypes_OnError,
 "ErrorNumber As Integer"
}; */

//WSNETC.DEF
/*
   WEP     @1  RESIDENTNAME
*/
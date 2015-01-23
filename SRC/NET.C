/* Net.C - Generic FIFO queue routines for WSNetC
*/
#include "WSANet.H"
#include "Version.H"
#include "NetClnt.H"
#include "NetSrvr.H"
#include <string.h>

WSADATA wsaData;

static char szError[256];

#ifdef DEBUG_BUILD
VOID DEBUGIO(LPSTR lpString)
{
 if(!lpString) return;
 wsprintf((LPSTR)szError, (LPSTR)"IO: %s\n\r", lpString);
 OutputDebugString( (LPSTR)szError );
}
#else
#define DEBUGIO( parm )
#endif /* DEBUG_BUILD */


/* BOOL netStartup(VOID);
    Purpose: To do the nitty gritty net I/O
    Given:   Nothing.
    Returns: TRUE if everything is OK
*/
BOOL netStartup(VOID)
{
 WORD wVersionRequested;

 DEBUGIO((LPSTR)"netStartup() - Begin");

    // Initialize our WinSock.DLL
 wVersionRequested=WSA_VERSION_NEEDED;
 if(WSAStartup(wVersionRequested, (LPWSADATA)&wsaData))
 {
  DEBUGIO((LPSTR)"netStartup() - BAD WINSOCK VERSION!");
  return(FALSE);
 }
 if((LOBYTE(wsaData.wVersion)!=1) ||
    (HIBYTE(wsaData.wVersion)!=1))
 {
  WSACleanup();
  DEBUGIO((LPSTR)"netStartup() - BAD WINSOCK VERSION!");
  return(FALSE);
 }
 DEBUGIO((LPSTR)"netStartup() - End");
 return(TRUE);
} /* netStartup() */


/* VOID netCleanup(VOID);
    Purpose: To cleanup after ourselves
*/
VOID netCleanup(VOID)
{
 DEBUGIO((LPSTR)"netCleanup()");
 WSACleanup();
} /* netCleanup() */


/* BOOL netCloseSocket(HWND, SOCKET);
    Purpose: To gracefully close a socket
*/
BOOL netCloseSocket(HWND hWnd, SOCKET sSocket)
{
 BOOL bDontLinger;

 DEBUGIO((LPSTR)"netCloseSocket() - Begin");

 bDontLinger = TRUE;
 if((sSocket) && (sSocket != SOCKET_ERROR))
 {
  if(hWnd) WSAAsyncSelect(sSocket, hWnd,0, 0);
  setsockopt(sSocket,
             SOL_SOCKET, SO_DONTLINGER, (char FAR *)&bDontLinger,
             sizeof(bDontLinger));
  closesocket(sSocket);
 }

 DEBUGIO((LPSTR)"netCloseSocket() - End");
 return(FALSE);
}


/* VOID netConnect(HCTL);
    Purpose:    To decentralize the connect setup code.
*/
BOOL netConnect(HCTL hCtl, SOCKET sSocket)
{
 LPNETCLIENT lpNetClient;
 int iSendTempSize = DEFAULT_SENDSIZE,
     iRecvTempSize = DEFAULT_RECVSIZE;
 int iInt = sizeof(iSendTempSize);
 HWND hWnd;

 DEBUGIO((LPSTR)"netConnect() - Begin");

 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);

 getsockopt(sSocket, SOL_SOCKET, SO_RCVBUF, (LPSTR)&iRecvTempSize,
            (int FAR *)&iInt);
 getsockopt(sSocket, SOL_SOCKET, SO_SNDBUF, (LPSTR)&iSendTempSize,
            (int FAR *)&iInt);

 iSendTempSize = max( iSendTempSize, lpNetClient->sSendSize );
 iRecvTempSize = max( iRecvTempSize, lpNetClient->sRecvSize );

    // Destroy any existing Temporary buffer space
 if(lpNetClient->hRecvTemp!=0)
 {
  GlobalUnlock(lpNetClient->hRecvTemp);
  GlobalFree(lpNetClient->hRecvTemp);
 }

 if(lpNetClient->hSendTemp!=0)
 {
  GlobalUnlock(lpNetClient->hSendTemp);
  GlobalFree(lpNetClient->hSendTemp);
 }

  // Set new size for temporary space
 lpNetClient->sRecvTempSize=(SHORT)iRecvTempSize;
 lpNetClient->sSendTempSize=(SHORT)iSendTempSize;

  // Allocate the temporary space
 lpNetClient->hRecvTemp=GlobalAlloc(GMEM_ZEROINIT | GMEM_MOVEABLE,
                                    iRecvTempSize);
 lpNetClient->hSendTemp=GlobalAlloc(GMEM_ZEROINIT | GMEM_MOVEABLE,
                                    iSendTempSize);
 lpNetClient->lpRecvTemp=GlobalLock(lpNetClient->hRecvTemp);
 lpNetClient->lpSendTemp=GlobalLock(lpNetClient->hSendTemp);

  // Lock and load the receive buffer
 lpNetClient->hRecvBuffer=GlobalAlloc(GMEM_ZEROINIT | GMEM_MOVEABLE,
                                      lpNetClient->sRecvSize);
 if(lpNetClient->hRecvBuffer)
    lpNetClient->lpRecvBuffer=GlobalLock(lpNetClient->hRecvBuffer);

  // Lock and load the send buffer
 lpNetClient->hSendBuffer=GlobalAlloc(GMEM_ZEROINIT | GMEM_MOVEABLE,
                                      lpNetClient->sSendSize);
 if(lpNetClient->hSendBuffer)
    lpNetClient->lpSendBuffer=GlobalLock(lpNetClient->hSendBuffer);

 hWnd = VBGetControlHwnd(hCtl);

 DEBUGIO((LPSTR)"netConnect() - Calling WSAAsyncSelect()");

 WSAAsyncSelect(sSocket, hWnd, CM_NETACTIVITY,
                FD_READ | FD_WRITE | FD_CLOSE);

  // Flag connection!
 lpNetClient->bConnect=TRUE;

 VBSetControlProperty(hCtl, IPROP_NETCLIENT_ERRORNUMBER, (LONG)0);

  // Let application do what it needs to do.
 if(clientFireOnConnect(hCtl))
 {
  DEBUGIO((LPSTR)"netConnect() - End - FireOnConnect() returned non-0!");
  return(TRUE);
 }

 DEBUGIO((LPSTR)"netConnect() - End");

 return(FALSE);
} /* netConnect() */


/* VOID netAccept(HCTL, SOCKET);
    Purpose:    To decentralize the server accept code.
*/
BOOL netAccept(HCTL hCtl, SOCKET sSocket)
{
 LPNETSERVER lpNetServer;
 ONACCEPTPARMS Parameters;
 SOCKET sNewSocket;
 HLSTR hlPeerAddr;
 LPSTR lpPeerAddr;
 SHORT sRemotePort;
 SOCKADDR_IN saPeer;
 int iPeer = sizeof(saPeer);
 ERR err;

 lpNetServer=(LPNETSERVER)VBDerefControl(hCtl);

 if(!sSocket) sSocket = lpNetServer->sSocket;

 sNewSocket = accept(sSocket, (struct sockaddr FAR *)&saPeer, (LPINT)&iPeer);

 if(sNewSocket == INVALID_SOCKET)
 {
  NetServerError(hCtl, WSAGetLastError());
  return(TRUE);
 }

 sRemotePort = (SHORT)htons(saPeer.sin_port);

 lpPeerAddr=inet_ntoa(saPeer.sin_addr);

 if(lpPeerAddr == NULL)
 {
  hlPeerAddr = VBCreateHlstr(NULL, 0);
 }
 else
 {
  hlPeerAddr = VBCreateHlstr(lpPeerAddr, lstrlen(lpPeerAddr));
 }

 Parameters.RemotePort = (SHORT FAR *)&sRemotePort;
 Parameters.Socket = (SHORT FAR *)&sNewSocket;
 Parameters.PeerAddr = hlPeerAddr;

 err = VBFireEvent(hCtl, IEVENT_NETSERVER_ONACCEPT,
                   (LPVOID)&Parameters);
 if(!err)
 {
  DEBUGIO((LPSTR)"netAccept() - Destroying hlPeerAddr");
  VBDestroyHlstr(hlPeerAddr);

    // If the user did a "Socket = 0" or "Socket = -1" then closed the
    // accepted connection (no NetClient was created to handle it)
  if((!sNewSocket)||(sNewSocket==SOCKET_ERROR))
  {
   netCloseSocket((HWND)NULL, sNewSocket);
   DEBUGIO((LPSTR)"netAccept() - Accepted Socket now closed");
  }
  DEBUGIO((LPSTR)"netAccept() - End");
  return(FALSE);
 }
 else
 {
  DEBUGIO((LPSTR)"netAccept() - VBFireEvent() returned non-0!");
  return(TRUE);
 }
} /* netAccept() */


/* LONG netServerAsyncProc(HWND, HCTL, WPARAM, LPARAM);
    Purpose: To handle asyncronous WinSock server transactions
*/
LONG netServerAsyncProc(HCTL hCtl, HWND hWnd, WPARAM wParam, LPARAM lParam)
{
 LPNETSERVER lpNetServer;

 lpNetServer=(LPNETSERVER)VBDerefControl(hCtl);

 DEBUGIO((LPSTR)"netServerAsyncProc() - Begin");

 if(WSAGETSELECTERROR(lParam))
 {
  NetServerError(hCtl, WSAGETSELECTERROR(lParam));
 }

 switch(WSAGETSELECTEVENT(lParam))
 {
  /*************** FD_ACCEPT ****************/
  case FD_ACCEPT:
  {
   DEBUGIO((LPSTR)"netServerAsyncProc() - FD_ACCEPT - Begin");

   if(lpNetServer->sSocket != (SOCKET)wParam)
   {
    DEBUGIO((LPSTR)"netServerAsyncProc() - FD_ACCEPT - A child socket accept()ed!");
    return(ERR_None);
   }
      // Setup buffers & such
   netAccept(hCtl,(SOCKET)wParam);

   DEBUGIO((LPSTR)"netServerAsyncProc() - FD_ACCEPT - End");
   return(ERR_None);
  }
  break; /* FD_ACCEPT */

  /*************** DEFAULT ******************/
  default:
  {
   char szError[256];

   wsprintf((LPSTR)szError, (LPSTR)"netServerAsyncProc: Unhandled WSA message #%d",
             wParam);

   DEBUGIO((LPSTR)szError);
  }
 }

 DEBUGIO((LPSTR)"netServerAsyncProc() - End");
 return(0L);
} /* netServerAsyncProc() */


/* LONG netClientAsyncProc(HWND, HCTL, WPARAM, LPARAM);
    Purpose: To handle asyncronous WinSock client transactions.
*/
LONG netClientAsyncProc(HCTL hCtl, HWND hWnd, WPARAM wParam, LPARAM lParam)
{
 LPNETCLIENT lpNetClient;

 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);

 DEBUGIO((LPSTR)"netClientAsyncProc() - Begin");

 if(WSAGETSELECTERROR(lParam))
 {
  NetClientError(hCtl, WSAGETSELECTERROR(lParam));
 }

 switch(WSAGETSELECTEVENT(lParam))
 {
  /*************** FD_CONNECT ****************/
  case FD_CONNECT:
  {
   DEBUGIO((LPSTR)"netClientAsyncProc() - FD_CONNECT - Begin");

      // Setup buffers & such
   netConnect(hCtl,(SOCKET)wParam);

   DEBUGIO((LPSTR)"netClientAsyncProc() - FD_CONNECT - End");
   return(ERR_None);
  }
  break; /* FD_CONNECT */


  /*************** FD_READ ****************/
  case FD_READ:
  {
   int iCount;

   DEBUGIO((LPSTR)"netClientAsyncProc() - FD_READ - Begin");

   if((!lpNetClient->sSocket) || (lpNetClient->sSocket == SOCKET_ERROR))
   {
    DEBUGIO((LPSTR)"netClientAsyncProc() - FD_READ - End - sSocket is invalid!");
    return(ERR_None);
   }

   if(lpNetClient->lpRecvBuffer)
   {
    iCount = netRecvData(hCtl);
   }
   else return(netReturnError( WSAERR_RecvBuffer,
                               WSAERR_RecvBuffer, NULL ));

   if(iCount==SOCKET_ERROR)
   {
    int iError;

    iError = WSAGetLastError();
    switch(iError)
    {
     case WSAECONNRESET:
      // Peer reset connection
      return(clientFireOnClose(hCtl));

     case WSAEINPROGRESS:
     case WSAEWOULDBLOCK:
      return(ERR_None);

     default:
      return(NetClientError( hCtl, WSAGetLastError()));
    }
   }

   if(iCount==0)
   {
    struct linger lLinger;

    DEBUGIO((LPSTR)"netClientAsyncProc() - FD_READ - recv() returned 0! Closing connection.");

    lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);

    if(lpNetClient->sSocket != SOCKET_ERROR)
    {
     netCloseSocket(hWnd, lpNetClient->sSocket);
     lpNetClient->bConnect=FALSE;
     lpNetClient->sSocket=(SOCKET)SOCKET_ERROR;
    }
    return(clientFireOnClose(hCtl));
   }

      // Did we receive enough to cause an event?
   if(iCount >= (lpNetClient->sRecvThreshold))
   {
    return(clientFireOnRecv(hCtl));
   }
   return(ERR_None);
  } /* FD_READ */


  /*************** FD_WRITE ***************/
  case FD_WRITE:
  {
   DEBUGIO((LPSTR)"netClientAsyncProc() - FD_WRITE - Begin");

   lpNetClient->fCanSend=TRUE;
   if(lpNetClient->sSendCount)
   {
    netSendData(hCtl);
   }

   DEBUGIO((LPSTR)"netClientAsyncProc() - FD_WRITE - End");
   return(ERR_None);
  } /* FD_WRITE */


  /*************** FD_CLOSE ***************/
  case FD_CLOSE:
  {
   struct linger lLinger;

   DEBUGIO((LPSTR)"netClientAsyncProc() - FD_CLOSE - Begin");

   if(lpNetClient->lpRecvBuffer)
   {
    netRecvData(hCtl);
   }

   if(lpNetClient->sSocket != SOCKET_ERROR)
   {
    netCloseSocket(hWnd, lpNetClient->sSocket);
    lpNetClient->bConnect=FALSE;
    lpNetClient->sSocket=(SOCKET)SOCKET_ERROR;
   }
   lpNetClient->bConnect=FALSE;

    // Let the user handle what is left over
   clientFireOnRecv(hCtl);

    // Let the user handle the closing
   clientFireOnClose(hCtl);

   DEBUGIO((LPSTR)"netClientAsyncProc() - FD_CLOSE - End");
   return(ERR_None);
  } /* FD_CLOSE */


  default:
  {
   char szError[256];

   wsprintf((LPSTR)szError, (LPSTR)"netClientAsyncProc: Unknown message #%d",
             wParam);

   DEBUGIO((LPSTR)szError);
  }
 }

 DEBUGIO((LPSTR)"netClientAsyncProc() - End");
 return(0L);
} /* netClientAsyncProc() */


/* ERR netSendData(HCTL);
    Purpose: To send data pending in the send queue to the sSocket
    Given:   The VB control handle
    Returns: ERR_None or a VB/Control error
*/
ERR netSendData(HCTL hCtl)
{
 char szError[256];
 LPSTR lpBuffer,
       lpWalker,
       lpEnd,
       lpTemp;
 int   iSent,
       iCount;
 SOCKET sSocket;
 LPNETCLIENT lpNetClient;

 DEBUGIO((LPSTR)"netSendData() - Begin");

 if(!hCtl)
 {
  DEBUGIO((LPSTR)"netSendData() - End - hCtl was 0!");
  return(ERR_None);
 }

 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);

 if(!(lpNetClient->fCanSend))
 {
  DEBUGIO((LPSTR)"netSendData() - End - Can't send data! Waiting for FD_WRITE");
  return(ERR_None);  // No FD_WRITE received yet
 }

 iCount=lpNetClient->sSendCount;
 lpBuffer=lpNetClient->lpSendBuffer;
 lpTemp=lpNetClient->lpSendTemp;
 sSocket=lpNetClient->sSocket;

 if((!sSocket)||(sSocket == SOCKET_ERROR))
 {
  DEBUGIO((LPSTR)"netSendData() - End - Socket is invalid!");
  return(ERR_None);
 }
 if(!iCount)
 {
  DEBUGIO((LPSTR)"netSendData() - End - Nothing to send!");
  return(ERR_None);
 }
 if(!lpBuffer)
 {
  DEBUGIO((LPSTR)"netSendData() - End - No Send buffer!");
  return(ERR_None);
 }

 while(iCount > 0)
 {
  iSent=send(sSocket, lpBuffer, iCount, 0);

  if(iSent==SOCKET_ERROR)
  {  // If something was blocking, always try again
   iSent=WSAGetLastError();
   switch(iSent)
   {
    case WSAEWOULDBLOCK:
    {
     MSG Msg;

     lpNetClient->fCanSend = FALSE;
     DEBUGIO((LPSTR)"netSendData() - End - Overflow (WSAEWOULDBLOCK) - Waiting for FD_WRITE");
        // Let other applications take over (somehow this is needed)
     PeekMessage(&Msg, NULL, 0, 0, PM_NOREMOVE);
     return(ERR_None);        // Mark that this happened!
    }

    case WSAEINPROGRESS:
    {
     DEBUGIO((LPSTR)"netSendData() - End - Non-fatal error (WSAEINPROGRESS)");
     return(ERR_None);        // No real error condition exists
    }

    default:
    {
     DEBUGIO((LPSTR)"netSendData() - Fatal error");
     return(NetClientError(hCtl, iSent));
    }
   }
  }

  DEBUGIO((LPSTR)"netSendData() - Sent data");

  wsprintf((LPSTR)szError,
           (LPSTR)"netSendData() - lpBuffer=%lX iSent=%X iCount=%X",
           lpBuffer, iSent, iCount);
  DEBUGIO((LPSTR)szError);

  /*lpWalker=lpBuffer+iSent;
  lpEnd=lpBuffer+iCount;
  while(lpWalker<lpEnd)
  {
   *lpBuffer++=*lpWalker++;
  }
//  lpBuffer[0] = '\0';
  */

  iCount-=iSent;
  _fmemmove(lpBuffer, (lpBuffer+iSent), (size_t)iCount);

  DEBUGIO((LPSTR)"netSendData() - Copied more to buffer");

  lpNetClient->sSendCount=(SHORT)iCount;

    // Has threshold been reached?
  if(iCount < lpNetClient->sSendThreshold)
  {
   ERR err;

   DEBUGIO((LPSTR)"netSendData() - Threshold reached!");
   clientFireOnSend(hCtl);
   if(!err)
   {    // Refresh all of our local vars.
    lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
    iCount=lpNetClient->sSendCount;
    lpBuffer=lpNetClient->lpSendBuffer;
    lpTemp=lpNetClient->lpSendTemp;
    sSocket=lpNetClient->sSocket;
    if((!sSocket) || (sSocket == SOCKET_ERROR))
    {
     DEBUGIO((LPSTR)"netSendData() - End - After firing OnSend(), Socket became invalid");
     return(ERR_None);
    }
   }
   else return(ERR_None);
  }
 }

 DEBUGIO((LPSTR)"netSendData() - End");
 return(ERR_None);
} /* netSendData() */


/* int netRecvData(HCTL);
    Purpose: To recv data pending on a socket and stuff it in the
             recv queue
    Given:   A VB control structure
    Returns: Nothing
    Does:    Magic
*/
int netRecvData(HCTL hCtl)
{
 LPSTR lpBuffer,
       lpTemp;
 int   iReceived,
       iCount,
       iSize;
 SOCKET sSocket;
 LPNETCLIENT lpNetClient;
 char szErrors[256];

 DEBUGIO((LPSTR)"netRecvData() - Begin");

 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
 if(!lpNetClient)
 {
  DEBUGIO((LPSTR)"netRecvData() - VBDerefControl() returned NULL!!! Call in the troops!");
  return(SOCKET_ERROR);
 }
 iCount=lpNetClient->sRecvCount;
 if(iCount==-1) iCount=0;
 lpBuffer=lpNetClient->lpRecvBuffer;
 lpTemp=lpBuffer+iCount;
 iSize=lpNetClient->sRecvSize;
 sSocket=lpNetClient->sSocket;

 if(!lpBuffer)
 {
    // Keep the messages coming
  DEBUGIO((LPSTR)"netRecvData() - No receive buffer!");
  recv(lpNetClient->sSocket, NULL, 0, 0);
  DEBUGIO((LPSTR)"netRecvData() - End - non-fatal");
  return(SOCKET_ERROR);
 }

 if((!sSocket)||(sSocket==SOCKET_ERROR))
 {
  DEBUGIO((LPSTR)"netRecvData() - End - Socket is not valid!");
  return(SOCKET_ERROR);
 }

 DEBUGIO((LPSTR)"netRecvData() - Debugging information");

 wsprintf((LPSTR)szErrors, (LPSTR)"netRecvData() - iCount=%d, iSize=%d",
          iCount, iSize);
 DEBUGIO((LPSTR)szErrors);

 wsprintf((LPSTR)szErrors, (LPSTR)"netRecvData() - sSocket=%d, lpTemp=%lX",
          sSocket, lpTemp);
 DEBUGIO((LPSTR)szErrors);

 if(!lpTemp) return(SOCKET_ERROR);

 DEBUGIO((LPSTR)"netRecvData() - Calling recv()");

 iReceived=recv(sSocket, lpTemp,
             iSize-iCount, 0);

 wsprintf((LPSTR)szErrors, (LPSTR)"netRecvData() - recv()==%d", iReceived);
 DEBUGIO((LPSTR)szErrors);

 DEBUGIO((LPSTR)"netRecvData() - Marker #0");

 if(iReceived==SOCKET_ERROR)
 {  // If something was blocking, always try again
  iReceived=WSAGetLastError();
  if((iReceived!=WSAEINPROGRESS) &&
     (iReceived!=WSAEWOULDBLOCK))
  {
   DEBUGIO((LPSTR)"netRecvData() - Marker #1");
   NetClientError(hCtl, iReceived);
   DEBUGIO((LPSTR)"netRecvData() - Fatal error");
   return(SOCKET_ERROR);
  }
  DEBUGIO((LPSTR)"netRecvData() - Non-fatal error");
  return(SOCKET_ERROR);
 }

 DEBUGIO((LPSTR)"netRecvData() - Marker #2");

 if(iReceived==-1)
 {
  DEBUGIO((LPSTR)"netRecvData() - iReceived==-1! Impossible!!!!");
  return(0);
 }

 if(!iReceived)
 {
  DEBUGIO((LPSTR)"netRecvData() - iReceived==0! - returning 0");
  return(0);
 }

 DEBUGIO((LPSTR)"netRecvData() - Marker #3");

 iCount+=iReceived;

 lpNetClient->sRecvCount=(SHORT)iCount;

 DEBUGIO((LPSTR)"netRecvData() - End");

 return(iCount);
} /* netRecvData() */


/* int putSendData(HCTL, MOLE);
    Purpose: To stuff data into the send queue
    Given:   A VB control structure
             A VB handled string - Not destroyed
    Returns: Send queue size.
    Does:    Magic.
*/
int putSendData(HCTL hCtl, MOLE mMole)
{
 LPSTR lpBuffer,
       lpStrEnd,
       lpBufEnd,
       lpWalker,
       lpStr,
       lpTemp;
 LPNETCLIENT lpNetClient;
 int   iSize,
       iCount;
 USHORT uRemaining,
       uStrLen;
 char szErrors[256];

 DEBUGIO((LPSTR)"putSendData() - Begin");

 if(!hCtl)
 {
  DEBUGIO((LPSTR)"putSendData() - End - hCtl is 0!");
  return(SOCKET_ERROR);
 }
 if(!mMole)
 {
  DEBUGIO((LPSTR)"putSendData() - End - mMole is 0!");
  return(SOCKET_ERROR);
 }

 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);

 iSize=lpNetClient->sSendSize;
 iCount=lpNetClient->sSendCount;
 lpBuffer=lpNetClient->lpSendBuffer;
 lpTemp=lpBuffer+iCount;

 if(!lpBuffer)
 {
  DEBUGIO((LPSTR)"putSendData() - No send buffer!");
  return(SOCKET_ERROR);
 }

 if(iCount > iSize)
 {
  DEBUGIO((LPSTR)"putSendData() - SendCount is greater than SendSize!");
  return(SOCKET_ERROR);
 }

 uRemaining=(USHORT)(iSize-iCount);

 uStrLen=VBGetHlstrLen((HLSTR)mMole);
 if(!uStrLen)
 {
  DEBUGIO((LPSTR)"putSendData() - uStrLen = 0!");
  return(0);
 }

 wsprintf((LPSTR)szErrors, (LPSTR)"putSendData() - mMole=%X, uStrLen=%X",
          mMole, uStrLen);
 DEBUGIO((LPSTR)szErrors);

    // Watch for buffer overflow
 if(uRemaining && uStrLen>uRemaining)
 {
  DEBUGIO((LPSTR)"putSendData() - Trying to make buffer room");

  netSendData(hCtl);

  DEBUGIO((LPSTR)"putSendData() - Marker #2");

  lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
  iCount=lpNetClient->sSendCount;

  DEBUGIO((LPSTR)"putSendData() - Marker #4");

  uRemaining=(USHORT)(iSize-iCount);
 }

 if(uRemaining && uStrLen>uRemaining)
 {
  DEBUGIO((LPSTR)"putSendData() - End - Overflow!!!");
  return(SOCKET_ERROR);
 }

 DEBUGIO((LPSTR)"putSendData() - Enough room, now filling");

 GetMoleName(mMole, lpTemp, iSize - iCount);

/* lpStrEnd=lpStr+uStrLen;
 lpBufEnd=lpBuffer+iSize;
 lpWalker=lpTemp;
 while((lpStr<lpStrEnd) && (lpWalker<lpBufEnd))
 {
  *lpWalker++=*lpStr++;
 }
 *lpWalker='\0';
*/

 iCount+=uStrLen;

 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
 lpNetClient->sSendCount=(SHORT)iCount;

    // If we can, flush some
 DEBUGIO((LPSTR)"putSendData() - trying to send data out of turn");
 netSendData(hCtl);

 DEBUGIO((LPSTR)"putSendData() - End");
 return(uStrLen);
} /* putSendData() */


/* HLSTR getRecvData(HCTL);
    Purpose: To get a block of data in VB string form from the recv
             queue.
    Given:   VB control handle
    Returns: A VB HLSTR holding the data.
    Note:    This routine receives all or nothing, and it will return
             embedded NULLs that were received.
*/
HLSTR getRecvData(HCTL hCtl)
{
 LPSTR lpBuffer;
 int   iCount;
 LPNETCLIENT lpNetClient;
 HLSTR hlReturn;

 DEBUGIO((LPSTR)"getRecvData() - Begin");

 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
 iCount=lpNetClient->sRecvCount;

 if(!iCount)
 {
  DEBUGIO((LPSTR)"getRecvData() - RecvBuffer is logically empty!");
  return(VBCreateHlstr(NULL, 0));
 }

 lpBuffer=lpNetClient->lpRecvBuffer;

 hlReturn=VBCreateHlstr(lpBuffer, (USHORT)iCount);

    // Did the VB string get created?
 if(hlReturn)
 {  // then update the internal count
  lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
  lpNetClient->sRecvCount=0;
 }

 DEBUGIO((LPSTR)"getRecvData() - End - Cross your fingers");
 return(hlReturn);
} /* getRecvData() */


/* HSZ getRecvLine(HCTL);
    Purpose: To get a line of data from the queue. Will only return a
             non-null string if it finds the LineDelimiter sequence.
    Given:   A VB control handle
    Returns: A VB Null terminated string - cannot/will-not include
             any nulls (must use getRecvData).
*/
HSZ getRecvLine(HCTL hCtl)
{
 LPSTR lpBuffer,
       lpDelim,
       lpEnd,
       lpBufEnd,
       lpWalker,
       lpKeep,
       lpTemp;
 int   iFound,
       iCount,
       iRemaining,
       iSize;
 HSZ   hszDelim;
 HSZ   hszReturn;
 BOOL  bEnd;
 char  cChar;
 LPNETCLIENT lpNetClient;

 DEBUGIO((LPSTR)"getRecvLine() - Begin");

 cChar='\0';
 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);

 iCount=lpNetClient->sRecvCount;

 if(!iCount)
 {
  DEBUGIO((LPSTR)"getRecvLine() - Logically empty! Returning empty HSZ");
  return(VBCreateHsz((_segment)hCtl, (LPSTR)&cChar));
 }

 iSize=lpNetClient->sRecvSize;
 lpBuffer=lpNetClient->lpRecvBuffer;

 if(!lpBuffer)
 {
  DEBUGIO((LPSTR)"getRecvLine() - No RecvBuffer!!!");
  return(VBCreateHsz((_segment)hCtl,
                     (LPSTR)&cChar));
 }

    // Lock in the delimiter string
 hszDelim=lpNetClient->hszRawLineDelimiter;
 if(hszDelim)
 {
  lpDelim=VBLockHsz(hszDelim);
  if(lpDelim)
  {
   DEBUGIO((LPSTR)"getRecvLine() - RawDelim locked Pointer is valid");
  }
 }
 else
 {
  DEBUGIO((LPSTR)"getRecvLine() - No RawDelim string!");
  lpDelim=NULL;
 }
 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);

 bEnd=FALSE;
 if(!lpDelim)
 {
  bEnd=TRUE;
 }
 else if(!*lpDelim)
 {
  bEnd=TRUE;
 }

 if(bEnd)
 {
  if(hszDelim) VBUnlockHsz(hszDelim);
  DEBUGIO((LPSTR)"getRecvLine() - End - Everything returned! (No delimiter)");
  lpBuffer[iCount]='\0';
  hszReturn=VBCreateHsz((_segment)hCtl, (LPSTR)lpBuffer);
  if(hszReturn)
  {
   lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
   lpNetClient->sRecvCount=0;
   return(hszReturn);
  }
  return(VBCreateHsz((_segment)hCtl, (LPSTR)&cChar));
 }

 lpBufEnd=lpBuffer+iCount;
 lpTemp=lpBuffer;

 DEBUGIO((LPSTR)"getRecvLine() - Searching buffer:");
//DEBUGIO(lpTemp);

 while(lpTemp<lpBufEnd)
 {
  //DEBUGIO() replacement on next line
#ifdef DEBUG_BUILD
  OutputDebugString((LPSTR)"*");
#endif // DEBUG_BUILD

      // How did a null get in here?!?!
  if(!*lpTemp)
  {
   //DEBUGIO((LPSTR)"Embedded Chr$(0) found in input");
   lpTemp++;
   continue;
  }

      // Is there a 1st character match?
  if(*lpDelim==*lpTemp)
  {
   lpKeep=lpTemp;

   //DEBUGIO((LPSTR)"First character delimiter match!");
   lpWalker=lpDelim;
      // While the delimiting string continues...
   while(*lpWalker)
   {
      // Walk with the matches, stop if not
    if(*lpWalker==*lpTemp)
    {
     lpWalker++;
     lpTemp++;
    }
    else break;
   }
      // Is the delimiter at it's end?
   if(!*lpWalker)
   {
    DEBUGIO((LPSTR)"getRecvLine() - Matched delimiter!");
    *lpKeep='\0';   // Overwrite first byte of matched delimiter
    hszReturn=VBCreateHsz((_segment)hCtl, lpBuffer);

        // If a string was created earlier, return it!
    if(hszReturn)
    {
     DEBUGIO((LPSTR)"getRecvLine() - Successful HSZ created!");

        // Calculate bytes remaining
     // 0x0000 <cr> lpKeep       lpBuffer
     // 0x0001 <lf>
     // 0x0002            lpTemp
     iFound = lpTemp-lpBuffer;
     iCount-=iFound;

        // Prevent wrap around!!!!
     if(iCount < 0)
     {
      DEBUGIO((LPSTR)"getRecvLine() - Underflow!!!!!");
      if(hszDelim) VBUnlockHsz(hszDelim);
      return(VBCreateHsz((_segment)hCtl, (LPSTR)&cChar));
     }

     lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
     lpNetClient->sRecvCount=(SHORT)iCount;

     if(iCount > 0)
     {
        // Copy used buffer down to overlap where it was
      DEBUGIO((LPSTR)"getRecvLine() - DEBUG: _fmemmove()");
      _fmemmove(lpBuffer, lpTemp, iCount);
     }

     DEBUGIO((LPSTR)"getRecvLine() - Returning string");
     VBUnlockHsz(hszDelim);
     return(hszReturn);
    }
   }
   else
   {
    lpTemp=lpKeep;
   }
  }
  lpTemp++;
 }

   // Didn't find the delimiter string
 VBUnlockHsz(hszDelim);
 cChar=0;

 DEBUGIO((LPSTR)"getRecvLine() - End - Nothing returned!");
 return(VBCreateHsz((_segment)hCtl, (LPSTR)&cChar));
} /* getRecvLine() */


// Declare Function NetClientSendBlock Lib "WSNETC.VBX" /
//      (cControl As NetClient, sString As String) As Integer
/* int CALLBACK _export NetClientSendBlock(HCTL hCtl, HLSTR);
    Purpose: To return a "block" string from the recv queue.
    Given:   The handle of the control to use
             The string to send
    Returns: The number of characters sent
*/
int CALLBACK _export NetClientSendBlock(HCTL hCtl, HLSTR hlStr)
{
 DEBUGIO((LPSTR)"NetClientSendBlock()");
 return(putSendData(hCtl, hlStr));
} /* NetClientSendBlock() */


// Declare Function NetClientRecvBlock Lib "WSNETC.VBX" /
//      (cControl As NetClient, sString As String, iNumber As Integer) As Integer
/* HLSTR CALLBACK _export NetClientRecvBlock(int);
    Purpose: To return a "block" string from the recv queue.
    Given:   A NetClient control
             A VB String
             The number of characters to grab, or 0 for everything
    Returns: 0 if everything went ok, an error if not
*/
int CALLBACK _export NetClientRecvBlock(HCTL hCtl, HLSTR hlStr, int iNumber)
{
 LPNETCLIENT lpNetClient;
 LPSTR lpBuffer;
 int iCount;
 HLSTR hlTemp;
 ERR err;

 DEBUGIO((LPSTR)"NetClientRecvBlock() - Begin");

 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
 iCount=lpNetClient->sRecvCount;

 if(iNumber==0) iNumber=iCount;
 if(iNumber>iCount) iNumber=iCount;

 hlTemp = hlStr;

 if(!iCount)
 {
  DEBUGIO((LPSTR)"NetClientRecvBlock() - RecvBuffer is logically empty!");

  VBSetHlstr((LPHLSTR)&hlTemp, NULL, 0);
  return(0);
 }

 lpBuffer=lpNetClient->lpRecvBuffer;

 err = VBSetHlstr((LPHLSTR)&hlTemp, (LPVOID)lpBuffer, (SHORT)iNumber);
 if(!err)
 {
  lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
  lpNetClient->sRecvCount=(SHORT)(iCount - iNumber);
 }

 DEBUGIO((LPSTR)"NetClientRecvBlock() - End - Cross your fingers ;)");
 return(err);
} /* NetClientRecvBlock() */


/* ERR netReturnError(ERR, int, LPSTR);
    Purpose: To set the text for a VB error, and return
*/
ERR netReturnError(ERR eError, int iResource, LPSTR lpMessage)
{
 static szString[128];

 DEBUGIO((LPSTR)"netReturnError()");
    // Resource #0 means use default text for error
 if(lpMessage == NULL)
 {
  if(iResource == 0) return(eError);

  if(!LoadString(hModDLL, iResource, (LPSTR)szString,
           sizeof(szString)))
  {
   wsprintf((LPSTR)szString, (LPSTR)"Unknown Error #%d: Resource string #%d missing",
                eError, iResource);
  }
  return(VBSetErrorMessage(eError, (LPSTR)szString));
 }
 else
  return(VBSetErrorMessage(eError, lpMessage));
} /* netReturnError() */


/* HSZ GetMoleHsz(HCTL, MOLE);
    Purpose: To convert an MOLE string to a VB HSZ string.
    Note:    This is a little out of place, but this module is
             common to all of the others.
*/
HSZ GetMoleHsz(HCTL hCtl, MOLE mMole)
{
 static szBuffer[256];
 HSZ hszTemp;
 LPSTR lpTemp;
 char cChar;
 USHORT uLength;

 DEBUGIO((LPSTR)"GetMoleHsz() - Begin");

 if(mMole)
 {
  uLength = VBGetHlstrLen((HLSTR)mMole);

#ifdef DEBUG_BUILD
  wsprintf((LPSTR)szBuffer, (LPSTR)"uLength=%X", uLength);
  DEBUGIO((LPSTR)szBuffer);
#endif

  lpTemp = VBDerefHlstr((HLSTR)mMole);
  if(lpTemp)
  {
   lstrcpyn((LPSTR)szBuffer, lpTemp, uLength);
   szBuffer[uLength+1] = '\0';
   DEBUGIO((LPSTR)szBuffer);
   hszTemp=VBCreateHsz((_segment)hCtl, (LPSTR)szBuffer);
   DEBUGIO((LPSTR)"GetMoleHsz() - End - Built valid HSZ");
   return(hszTemp);
  }
 }

 cChar = '\0';
 hszTemp=VBCreateHsz((_segment)hCtl, (LPSTR)&cChar);
 DEBUGIO((LPSTR)"GetMoleHsz() - End - Returned an empty string!");
 return(hszTemp);
} /* GetMoleHsz() */


/* VOID GetMoleName(MOLE, LPSTR, int);
    Purpose: To convert an MOLE string to a C string.
    Note:    This is a little out of place, but this module is
             common to all of the others.
*/
VOID GetMoleName(MOLE mMole, LPSTR lpString, int iSize)
{
 static szBuffer[256];
 LPSTR lpTemp;
 size_t stCount;

 //DEBUGIO((LPSTR)"GetMoleName() - Begin");

 if(!lpString) return;
 if(!iSize) return;

 if(mMole)
 {
  stCount = (size_t)VBGetHlstrLen((HLSTR)mMole);
  if(stCount > (size_t)iSize)
  {
   DEBUGIO((LPSTR)"GetMoleName() - I don't have enough room to hold it all!");
   stCount = (size_t)iSize;
  }
  lpTemp = VBDerefHlstr((HLSTR)mMole);
  if(lpTemp)
  {
   _fmemmove(lpString, lpTemp, stCount);
   //DEBUGIO((LPSTR)"GetMoleName() - End - Success!");
   return;
  }
 }

 *lpString = '\0';
 DEBUGIO((LPSTR)"GetMoleName() - End - Returned an empty string!");
 return;
} /* GetMoleName() */


/* VOID SetMole(MOLE, LPSTR);
    Purpose: To set the handed MOLE to the handed LPSTR.
*/
VOID SetMole(LPMOLE pmMole, LPSTR lpString)
{
 HLOCAL hLocal;
 LPSTR lpTemp;

 DEBUGIO((LPSTR)"SetMole() - Begin");

 if(!lpString)
 {
  DEBUGIO((LPSTR)"SetMole() - End - lpString was NULL");
  return;
 }

 if(VBSetHlstr((LPHLSTR)pmMole, lpString, lstrlen(lpString) + 1))
 {
  DEBUGIO((LPSTR)"SetMole() - VBSetHlstr() failed!");
 }
 else
 {
  DEBUGIO((LPSTR)"SetMole() - VBSetHlstr() was successful!");
 }

 DEBUGIO((LPSTR)"SetMole() - End");
 return;
}   /* SetMole() */


/* MOLE AddMole(LPSTR);
    Purpose: To create a new MOLE for an LPSTR.
*/
MOLE AddMole(LPSTR lpString)
{
 MOLE mTemp;

 //DEBUGIO((LPSTR)"AddMole() - Begin");

 mTemp = CreateMole(lpString , (USHORT)(lstrlen(lpString) + 1));

 //DEBUGIO((LPSTR)"AddMole() - End");
 return(mTemp);
}   /* AddMole() */


/* MOLE CreateMole(LPSTR, int);
    Purpose: To allow set length Mole's
*/
MOLE CreateMole(LPSTR lpString, USHORT uLength)
{
 HLSTR hlStr;

 DEBUGIO((LPSTR)"CreateMole() - Begin");

 if(!lpString)
 {
  DEBUGIO((LPSTR)"CreateMole() - End - lpString was NULL");
  return((MOLE)0);
 }

 hlStr = VBCreateHlstr(lpString, uLength);

 if(hlStr)
 {
  DEBUGIO((LPSTR)"CreateMole() - VBCreateHlstr() Succeeded!");
 }
 else
 {
  DEBUGIO((LPSTR)"CreateMole() - VBCreateHlstr() Failed!");
 }

 DEBUGIO((LPSTR)"CreateMole() - End");
 return((MOLE)hlStr);
} /* CreateMole() */


/* VOID DeleteMole(LPMOLE);
    Purpose: To delete the handed MOLE
*/
VOID DeleteMole(MOLE mMole)
{
 //DEBUGIO((LPSTR)"DeleteMole() - Begin");

 if(mMole)
 {
  VBDestroyHlstr((HLSTR)mMole);
  DEBUGIO((LPSTR)"DeleteMole() - End");
  return;
 }
 //DEBUGIO((LPSTR)"DeleteMole() - End - Mole was 0!");
} /* DeleteMole() */


/* END of Net.C */

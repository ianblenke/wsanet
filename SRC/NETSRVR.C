/* NetSrvr.C - WinSock NetServer Control for VB/VC++
  - WinSock V1.1 required
  - Visual Basic 1.0 required, 2.0+ preferred (tested only on 3.0)

   Author: Ian Blenke

In this document, VBX refers to the actual Visual Basic control distributed
with and based on this WSANET project, and author refers to Ian Blenke. The
author cannot be held responsible for damages, expressed or implied, for the
use of this software. No commercial use can be made of this product without
the consent of the author. No profit of any kind can be made on the sale or
distribution of this program. If you wish to distribute this program with
other samples of WinSock programming, you must first contact the author so
that he can keep accurate records of its usage; no charge may be made for
this project's availability other than the cost of the physical medium used
to store it on and any processing costs. If you writany programs based on
this source code, including parts of this source code, or in some way derived
from this source code, you may not sell them for any profit without the
written consent of the author. If you incorporate this VBX into a public
dain program althe author requires is a notification that "The NetClient
control was written by Ian Blenke" and some form of notification that his name
was used in the public domain software distribution. No profit of any kind,
shareware/commercial or otherwise, may be made for software based on this VBX
without the written consent of the author. This does not represent a contract
on the part of the author. If any issues cannot clearly be resolved by reading
this text, immediately contact the author.

The VBX control distributed with this project must be distributed with your
programs free of charge. You cannot charge money for any program based on the
VBX without the written consent of the author.

Notes:

If you have any bug reports and/or source patches... By all means, tell me! I
would be glad to help keep this code up to date. Do not, however, modify this
source in any way and re-distribute it without the author's knowledge. Follow
the above Copyright, and all will be well.

I don't like such agreements, but in today's world of lawyers and lawbreakers,
I have little other choice. Enjoy!

(This is meant to scare code "scalpers". Those that are actually learning C
 programming for Visual Basic controls and WinSock applications are free to
 adapt any of the ideas of this code. Remember, new code should be better
 code, so "adapt" this, do NOT copy it.)
*/

#define NetSrvr_C
#include "WSANet.H"
#include "NetSrvr.H"

    // Debugging macro's/functions
#ifdef DEBUG_BUILD
static char szErrors[256];

VOID DEBUGNETS(LPSTR lpString)
{
 wsprintf((LPSTR)szErrors, (LPSTR)"NetServer: %s\n\r", lpString);
 OutputDebugString( (LPSTR)szErrors );
}
#else
#define DEBUGNETS( parm )
#endif /* DEBUG_BUILD */


/* LONG CALLBACK _export NetServerCtlProc(HCTL, HWND, UINT, WPARAM, LPARAM);
    Purpose: To handle the VBX messages
*/
LONG CALLBACK _export NetServerCtlProc(HCTL hCtl, HWND hWnd, UINT Msg,
                                       WPARAM wParam, LPARAM lParam)
{
 LPNETSERVER lpNetServer;

 lpNetServer=(LPNETSERVER)VBDerefControl(hCtl);

 switch(Msg)
 {

#ifdef DESIGN_TIME
    // The MODEL want's to know our default size
  case VBM_GETDEFSIZE:
  {
   BITMAP bmp;
   HBITMAP hBitmap;

   DEBUGNETS("VBM_GETDEFSIZE[2.0] Begin\n\r");
   hBitmap=LoadBitmap(hModDLL, MAKEINTRESOURCE(IDBMP_SERVER));
   if(hBitmap)
   {
    GetObject(hBitmap, sizeof(BITMAP), (LPSTR)&bmp);
    DeleteObject(hBitmap);
    DEBUGNETS((LPSTR)"VBM_GETDEFSIZE[2.0] End\n\r");
    return(MAKELONG(bmp.bmWidth, bmp.bmHeight));
   }
   DEBUGNETS((LPSTR)"VBM_GETDEFSIZE[2.0] Allowing DefVBProc to handle\n\r");
  }
  break;

    // The re-size message - only in design time
  case WM_SIZE:
  {
   BITMAP bmp;
   HBITMAP hBitmap;
   RECT rect;
   POINT pos;

   DEBUGNETS((LPSTR)"WM_SIZE Start");
   GetWindowRect(hWnd, (RECT FAR *)&rect);
   pos.x=rect.left;
   pos.y=rect.top;
   ScreenToClient((HWND)GetWindowWord(hWnd, GWW_HWNDPARENT),
                  (POINT FAR *)&pos);

   hBitmap=LoadBitmap(hModDLL, MAKEINTRESOURCE(IDBMP_SERVER));
   if(hBitmap)
   {
    GetObject(hBitmap, sizeof(BITMAP), (LPSTR)&bmp);
    MoveWindow(hWnd, pos.x, pos.y, bmp.bmWidth, bmp.bmHeight,
               SWP_NOZORDER | SWP_NOMOVE);
    DeleteObject(hBitmap);
    DEBUGNETS((LPSTR)"WM_SIZE End");
    return(0);
   }
   DEBUGNETS((LPSTR)"WM_SIZE End - Bitmap load failure!");
  }
  break;


    // The paint message - only design time!
  case WM_PAINT:
  {
   PAINTSTRUCT ps;
   BITMAP bmp;
   RECT rect;
   HBITMAP hBitmap,
           hOldBitmap;
   HDC  hdcMem;

   DEBUGNETS((LPSTR)"WM_PAINT Start");
   if(VBGetMode()==MODE_RUN) break;

   GetWindowRect(hWnd, (RECT FAR *)&rect);
   BeginPaint(hWnd, (PAINTSTRUCT FAR *)&ps);
   hBitmap=LoadBitmap(hModDLL, MAKEINTRESOURCE(IDBMP_SERVER));
   if(hBitmap)
   {
    GetObject(hBitmap, sizeof(BITMAP), (LPSTR)&bmp);
    hdcMem = CreateCompatibleDC(ps.hdc);
    if(hdcMem)
    {
     hOldBitmap=SelectObject(hdcMem, hBitmap);
     BitBlt(ps.hdc, 0, 0, bmp.bmWidth, bmp.bmHeight, hdcMem,
            0, 0, SRCCOPY);
     SelectObject(hdcMem, hOldBitmap);
     DeleteDC(hdcMem);
    }
    DeleteObject(hBitmap);
   }
   EndPaint(hWnd, (PAINTSTRUCT FAR *)&ps);
   DEBUGNETS((LPSTR)"WM_PAINT End");
  }
  break;
#endif // DESIGN_TIME


    // The control is closing - free everything.
  case WM_NCDESTROY:
  {
   DEBUGNETS((LPSTR)"WM_NCDESTROY Begin");

   if(lpNetServer->sSocket != SOCKET_ERROR)
   {
    netCloseSocket(hWnd, lpNetServer->sSocket);
    lpNetServer->bListen=FALSE;
    lpNetServer->sSocket=(SOCKET)SOCKET_ERROR;
   }

   lpNetServer->sLocalPort = 0;
   lpNetServer->sErrorNumber = 0;
   lpNetServer->sQueueSize = 0;
   DEBUGNETS((LPSTR)"WM_NCDESTROY End");
  }
  break; // WM_NCDESTROY


  // A New control has been created/loaded
  case VBM_CREATED:
  {
   DEBUGNETS((LPSTR)"VBM_CREATED Begin");

   lpNetServer->sLocalPort = 0;
   lpNetServer->sSocket = SOCKET_ERROR;
   lpNetServer->sErrorNumber = 0;
   lpNetServer->sQueueSize = 0;
   lpNetServer->bListen = FALSE;

   if(VBGetMode()==MODE_RUN)
   {
    ShowWindow(hWnd, SW_HIDE);
   }
   else
   {
    ShowWindow(hWnd, SW_SHOWNA);
   }

   DEBUGNETS((LPSTR)"VBM_CREATED End");
   return(ERR_None);
  } // VBM_CREATED


#ifdef DESIGN_TIME
    // Here is how to implement a dialog box property!
  case VBM_INITPROPPOPUP:
  {
   switch(wParam)
   {
    case IPROP_NETSERVER_ABOUT:
    {
     DEBUGNETS((LPSTR)"VBM_INITPROPPOPUP: About - Begin");

     if(bDialogInUse) return(NULL);
     bDialogInUse = TRUE;
     hDialogHwnd=CreateWindow(CLASS_ABOUTPOPUP, NULL,
            WS_POPUP,
            0, 0, 0, 0,
            NULL, NULL,
            hModDLL, NULL);
     DEBUGNETS((LPSTR)"VBM_INITPROPPOPUP: About - End");
     return(hDialogHwnd);
    }
    break;
   }
  }
  break;
#endif // DESIGN_TIME


/**********************************************************************
 ***                        SET PROPERTY                            ***
 **********************************************************************/
  case VBM_SETPROPERTY:
  {
   ERR err;
#ifdef DEBUG_BUILD
   char szError[128];

   wsprintf((LPSTR)szError, (LPSTR)"VBM_SETPROPERTY: #%d - Begin",
             wParam);
   DEBUGNETS((LPSTR)szError);
#endif /* DEBUG_BUILD */

   switch(wParam)
   {

    /****************** SET: Socket *********************/
    // Can't set!!!
    case IPROP_NETSERVER_SOCKET:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"Socket"));


    /****************** SET: ErrorMessage ****************/
    // Can't set!!!
    case IPROP_NETSERVER_ERRORMESSAGE:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"ErrorMessage"));


    /****************** SET: Version ********************/
    // Can't set!!!
    case IPROP_NETSERVER_VERSION:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"Version"));


    /****************** SET: Listen *********************/
    // User is trying to listen() for incoming connections
    // This is the most complex part of the program
    case IPROP_NETSERVER_LISTEN:
    {
     SOCKET sSocket;
     LONG lNonBlocking = TRUE;
     SOCKADDR_IN saLocal;
     int iError;
     BOOL bReuseAddr = TRUE;

     DEBUGNETS((LPSTR)"SET: Listen - Begin");

     if((lpNetServer->bListen) && *(BOOL *)&lParam)
     {
      DEBUGNETS((LPSTR)"SET: Listen - End - Already listening");
      return(netReturnError(WSAERR_AlreadyConn, WSAERR_AlreadyConn, NULL));
     }

     if(!(lpNetServer->bListen) && !*(BOOL *)&lParam)
     {
      DEBUGNETS((LPSTR)"SET: Listen - End - Pointless!");
      return(ERR_None);
     }

     if((lpNetServer->bListen) && !*(BOOL *)&lParam)
     {
      netCloseSocket(hWnd, lpNetServer->sSocket);
      lpNetServer->bListen=FALSE;
      lpNetServer->sSocket=(SOCKET)SOCKET_ERROR;
      DEBUGNETS((LPSTR)"SET: Listen - End - Stopped listening!");
      return(ERR_None);
     }

        // In the odd chance that a socket remains that is
        // listening, close it!
     if(lpNetServer->sSocket != SOCKET_ERROR)
     {
      netCloseSocket(hWnd, lpNetServer->sSocket);
      lpNetServer->sSocket=(SOCKET)SOCKET_ERROR;
     }

        // Create a new socket
     sSocket = socket(PF_INET, SOCK_STREAM, 0);

     if( sSocket==INVALID_SOCKET )
     {
      return(NetServerError(hCtl, WSAGetLastError()));
     }

        // Set up asyncronous messages. Remember, child
        // accept()ed sockets will have this same mask.
     if(WSAAsyncSelect(sSocket, hWnd, CM_NETACTIVITY,
          FD_ACCEPT | FD_CONNECT | FD_READ | FD_WRITE | FD_CLOSE)
        == SOCKET_ERROR)
     {
      NetServerError(hCtl, WSAGetLastError());
      closesocket(sSocket);
      return(ERR_None);
     }

        // Save the socket
     lpNetServer->sSocket=(SOCKET)sSocket;

        // If Listen = 2 then DON'T USE REUSE_ADDR!
     if(*(int *)&lParam != 2)
     {
      setsockopt(sSocket, SOL_SOCKET, SO_REUSEADDR,
                (char FAR *)&bReuseAddr, sizeof(BOOL));
     }
        // Bind the socket
     DEBUGNETS((LPSTR)"SET: Listen - bind() - Begin");
     saLocal.sin_family=AF_INET;
     saLocal.sin_addr.s_addr=INADDR_ANY;
     saLocal.sin_port=htons(lpNetServer->sLocalPort);

     if(bind(sSocket, (LPSOCKADDR)&saLocal,
        sizeof(saLocal))==SOCKET_ERROR)
     {
      NetServerError(hCtl, WSAGetLastError());
      closesocket(sSocket);
      lpNetServer->sSocket=(SOCKET)SOCKET_ERROR;
      return(ERR_None);
     }
     DEBUGNETS((LPSTR)"SET: Listen - bind() - End");

     DEBUGNETS((LPSTR)"SET: Listen - listen() - Begin");
        // Listen for incoming
     if(listen(sSocket, lpNetServer->sQueueSize)==SOCKET_ERROR)
     {
      NetServerError(hCtl, WSAGetLastError());
      closesocket(sSocket);
      lpNetServer->sSocket=(SOCKET)SOCKET_ERROR;
      return(ERR_None);
     }
     DEBUGNETS((LPSTR)"SET: Listen - listen() - End");

     lpNetServer->bListen = TRUE;

     DEBUGNETS((LPSTR)"SET: Listen - End");
     return(ERR_None);
    } // IPROP_NETSERVER_LISTEN


    /****************** SET: QueueSize *****************/
    // Only set if we aren't listening
    case IPROP_NETSERVER_QUEUESIZE:
    {
     DEBUGNETS((LPSTR)"SET: QueueSize - Begin");
     if(!lpNetServer->bListen)
        lpNetServer->sQueueSize=*(SHORT *)&lParam;
     DEBUGNETS((LPSTR)"SET: QueueSize - End");
     return(ERR_None);
    } // IPROP_NETSERVER_QUEUESIZE


    /****************** SET: LocalPort ******************/
    // Only set if we aren't listening
    case IPROP_NETSERVER_LOCALPORT:
    {
     DEBUGNETS((LPSTR)"SET: LocalPort - Begin");
     if(!lpNetServer->bListen)
        lpNetServer->sLocalPort=*(SHORT *)&lParam;
     DEBUGNETS((LPSTR)"SET: LocalPort - End");
     return(ERR_None);
    } // IPROP_NETSERVER_LOCALPORT


    /****************** SET: LocalService ***************/
    // Handle Setting of LocalService
    case IPROP_NETSERVER_LOCALSERVICE:
    {
     LPSTR lpLocalService;
     LPSERVENT lpseServEnt;

     DEBUGNETS((LPSTR)"SET: LocalService - Begin");
     if(!lParam)
     {
      lpNetServer->sLocalPort = 0;
      return( netReturnError( VBERR_INVALIDPROPERTYVALUE, 0, NULL));
     }

     if(!lpNetServer->bListen)
     {
      lpLocalService=*(LPSTR *)&lParam;

        // Get service by name, but don't do families... yet?
      lpseServEnt = getservbyname( lpLocalService, NULL);

      if(lpseServEnt)
      {
        // Copy out the port number
       lpNetServer->sLocalPort = ntohs(lpseServEnt->s_port);
      }
      else
      {
       lpNetServer->sLocalPort = 0;
       DEBUGNETS((LPSTR)"SET: LocalService - getservbyname() returned NULL!");
       return(NetServerError( hCtl, WSAGetLastError()));
      }
     }
     DEBUGNETS((LPSTR)"SET: LocalService - End");
     return(ERR_None);
    } // IPROP_NETSERVER_LOCALSERVICE


    /****************** SET: Debug **********************/
    // Set SO_DEBUG
    case IPROP_NETSERVER_DEBUG:
    {
     BOOL fDebug = *(BOOL *)&lParam;

     DEBUGNETS((LPSTR)"SET: Debug - Begin");

     if((lpNetServer->sSocket != 0) &&
        (lpNetServer->sSocket != SOCKET_ERROR))
     {
      if(setsockopt(lpNetServer->sSocket, SOL_SOCKET, SO_DEBUG,
                    (LPSTR)&fDebug, sizeof(fDebug))
               != SOCKET_ERROR)
      {
       DEBUGNETS((LPSTR)"SET: Debug - End - Ok");
       return(ERR_None);
      }
      else return(NetServerError(hCtl, WSAGetLastError()));
     }
     return(netReturnError(WSAERR_NotConn, WSAERR_NotConn, NULL));
    }
    break;
   }
  }
  break;


/**********************************************************************
 ***                        GET PROPERTY                            ***
 **********************************************************************/
  case VBM_GETPROPERTY:
  {
#ifdef DEBUG_BUILD
   char szError[128];

   wsprintf((LPSTR)szError, (LPSTR)"VBM_GETPROPERTY: #%d: Begin",
            wParam);
   DEBUGNETS((LPSTR)szError);
#endif /* DEBUG_BUILD */

   switch(wParam)
   {

    /****************** GET: LocalService ***************/
    case IPROP_NETSERVER_LOCALSERVICE:
    {
     HSZ hszTemp;
     LPSERVENT lpseServEnt;
     BOOL bClear = FALSE;

     DEBUGNETS((LPSTR)"GET: LocalService - Begin");

     lpseServEnt = getservbyport( htons(lpNetServer->sLocalPort), NULL);
     if(lpseServEnt)
     {
      if(lpseServEnt->s_name)
      {
       hszTemp = VBCreateHsz((_segment)hCtl, lpseServEnt->s_name);
      }
      else bClear = TRUE;
     }
     else
     {
      bClear = TRUE;
      NetServerError(hCtl, WSAGetLastError());
     }

     if(bClear)
     {
      char cChar = '\0';

      hszTemp = VBCreateHsz((_segment)hCtl, (LPSTR)&cChar);
     }

     *(*(LPHSZ *)&lParam) = hszTemp;

     DEBUGNETS((LPSTR)"GET: LocalService - End");
     return(ERR_None);
    } // IPROP_NETSERVER_LOCALSERVICE


    /****************** GET: ErrorMessage ***************/
    case IPROP_NETSERVER_ERRORMESSAGE:
    {
     char szString[256];
     HSZ  hszTemp;
     char cChar;
     int iError;

     DEBUGNETS((LPSTR)"GET: ErrorMessage - Begin");

     iError=lpNetServer->sErrorNumber;
     if(!LoadString(hModDLL, iError, (LPSTR)szString,
                sizeof(szString)))
     {
      wsprintf((LPSTR)szString, (LPSTR)"Unknown Error #%d",
               iError);
     }

#ifdef DESIGN_TIME
     DEBUGNETS((LPSTR)szString);
#endif

     // Set the error message text

     cChar='\0';
     hszTemp=VBCreateHsz((_segment)hCtl, (LPSTR)szString);
     DEBUGNETS((LPSTR)"GET: ErrorMessage - Created new error string");
     if(!hszTemp) hszTemp=VBCreateHsz((_segment)hCtl, (LPSTR)&cChar);

     *(*(LPHSZ *)&lParam) = hszTemp;

     DEBUGNETS((LPSTR)"GET: ErrorMessage - End");
     return(ERR_None);
    } // IPROP_NETSERVER_ERRORMESSAGE


    /****************** GET: Version ********************/
    // Get NetServer version string...
    case IPROP_NETSERVER_VERSION:
    {
     HSZ hszTemp;
     char cChar = '\0';

     DEBUGNETS((LPSTR)"GET: Version - Begin");

     hszTemp=VBCreateHsz((_segment)hCtl, (LPSTR)WSANET_VERSION_TEXT);
     if(!hszTemp) hszTemp=VBCreateHsz((_segment)hCtl, (LPSTR)&cChar);

     *(*(LPHSZ *)&lParam) = hszTemp;
     DEBUGNETS((LPSTR)WSANET_VERSION_TEXT);

     DEBUGNETS((LPSTR)"GET: Version - End");
     return(ERR_None);
    } // IPROP_NETSERVER_VERSION


    /****************** GET: Debug **********************/
    // Get SO_DEBUG
    case IPROP_NETSERVER_DEBUG:
    {
     BOOL fDebug;
     int iSize = sizeof(fDebug);

     DEBUGNETS((LPSTR)"GET: Debug - Begin");

     if((lpNetServer->sSocket != 0) &&
        (lpNetServer->sSocket != SOCKET_ERROR))
     {
      if(getsockopt(lpNetServer->sSocket, SOL_SOCKET, SO_DEBUG,
                    (LPSTR)&fDebug, (int FAR *)&iSize)
           != SOCKET_ERROR)
      {
       *(*(SHORT FAR **)&lParam)=(SHORT)fDebug;
       DEBUGNETS((LPSTR)"GET: Debug - End - SO_DEBUG returned");
       return(ERR_None);
      }
     }
     *(*(SHORT FAR **)&lParam)=0;
     DEBUGNETS((LPSTR)"GET: Debug - End - False returned");
     return(ERR_None);
    } // IPROP_NETSERVER_DEBUG


    /****************** GET: DEFAULT handler! ***********/
    default:
    {
     DEBUGNETS((LPSTR)"GET: Property not handled");
    }
    break;
   }
   DEBUGNETS((LPSTR)"VBM_GETPROPERTY: End");
  }
  break;


/**********************************************************************
 ***                        GET PROPERTY HSZ                        ***
 **********************************************************************/
#ifdef DESIGN_TIME
    // Set the About propery to display "About NetServer"
  case VBM_GETPROPERTYHSZ:
  {
   DEBUGNETS((LPSTR)"VBM_GETPROPERTYHSZ Begin");
   switch(wParam)
   {
    case IPROP_NETSERVER_ABOUT:
    {
     DEBUGNETS((LPSTR)"GETPROPERTYHSZ: About");
     *(*(LPHSZ *)&lParam) = VBCreateHsz((_segment) hCtl, (LPSTR)STRING_ABOUTTEXT);
     return(ERR_None);
    }
   }
   DEBUGNETS((LPSTR)"VBM_GETPROPERTYHSZ End");
  }
  break;

/**********************************************************************
 ***                   CONTEXT SENSITIVE HELP                       ***
 **********************************************************************/
  case VBM_HELP:
  {
#ifdef DEBUG_BUILD
   char szError[128];

   DEBUGNETS((LPSTR)"VBM_HELP: Begin");
   wsprintf((LPSTR)szError, (LPSTR)"VBM_HELP: Prop. #%d",
             wParam );
   DEBUGNETS((LPSTR)szError);
#endif // DEBUG_BULD

   if(bHelpStdPropEvt(wParam, (LPMODEL)lParam))
        break;
   else return(DisplayHelpTopic(TOPIC_CONTROL_NETSERVER, wParam,
                                (LPMODEL)lParam));
   DEBUGNETS((LPSTR)"VBM_HELP: Standard handler allowed");
  }
  break;
#endif // DESIGN_TIME


    // **** The network activity message ****
  case CM_NETACTIVITY:
  {
   DEBUGNETS((LPSTR)"CM_NETACTIVITY");

   return(netServerAsyncProc(hCtl, hWnd, wParam, lParam));
  }

 }
 return(VBDefControlProc(hCtl, hWnd, (USHORT)Msg, (USHORT)wParam,
        (LONG)lParam));
} /* NetServerCtlProc() */


/* ERR NetServerError(HCTL hCtl, int);
    Purpose: To set a VB error level and string to any errors
*/
ERR NetServerError(HCTL hCtl, int iError)
{
 char szString[256];
 LPNETSERVER lpNetServer;

 wsprintf((LPSTR)szString, (LPSTR)"NetServerError(%d) - Begin", iError);
 DEBUGNETS((LPSTR)szString);

 if(!iError) return(ERR_None);

 if(iError == WSANO_DATA) return(ERR_None);

 lpNetServer=(LPNETSERVER)VBDerefControl(hCtl);
 lpNetServer->sErrorNumber=(SHORT)iError;

 DEBUGNETS((LPSTR)"NetServerError() - Fire OnError()");

 return(serverFireOnError(hCtl));
} /* NetServerError() */


/* ERR serverFireOnError(HCTL);
    Purpose: To fire the OnError event.
    Given:   The control handle
    Returns: Any error condition.
*/
ERR serverFireOnError(HCTL hCtl)
{
 ONERRORPARMS Parameters;
 LPNETSERVER lpNetServer;
 SHORT       sErrorNumber;

    // Dereference the control so we can look inside
    // Becomes invalid after the first VBLockHsz() below!
 DEBUGNETS((LPSTR)"serverFireOnError() - Begin");
 lpNetServer=(LPNETSERVER)VBDerefControl(hCtl);

 sErrorNumber=lpNetServer->sErrorNumber;

    // Set the error number
 Parameters.ErrorNumber=(SHORT FAR *)&sErrorNumber;

 return(VBFireEvent(hCtl, IEVENT_NETSERVER_ONERROR, (LPVOID)&Parameters));
} /* serverFireOnError() */




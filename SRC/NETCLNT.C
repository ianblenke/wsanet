/* NetClnt.C - WinSock NetClient Control for VB/VC++
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

#define NetClnt_C
#include "WSANet.H"
#include "NetClnt.H"

    // Specific globals to NetClnt.C
static char szLocalHostName[256];       // Name of the machine running us

#ifdef DEBUG_BUILD
static char szErrors[256];

VOID DEBUGNETC(LPSTR lpString)
{
 wsprintf((LPSTR)szErrors, (LPSTR)"NetClient: %s\n\r", lpString);
 OutputDebugString( (LPSTR)szErrors );
}
#else
#define DEBUGNETC( parm )
#endif /* DEBUG_BUILD */


/* LONG CALLBACK _export NetClientCtlProc(HCTL, HWND, UINT, WPARAM, LPARAM);
    Purpose: To handle the VBX messages
*/
LONG CALLBACK _export NetClientCtlProc(HCTL hCtl, HWND hWnd, UINT Msg,
                                       WPARAM wParam, LPARAM lParam)
{
 LPNETCLIENT lpNetClient;

 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);

 switch(Msg)
 {

#ifdef DESIGN_TIME
    // The MODEL want's to know our default size
  case VBM_GETDEFSIZE:
  {
   BITMAP bmp;
   HBITMAP hBitmap;

   DEBUGNETC("VBM_GETDEFSIZE[2.0] Begin\n\r");
   hBitmap=LoadBitmap(hModDLL, MAKEINTRESOURCE(IDBMP_CLIENT));
   if(hBitmap)
   {
    GetObject(hBitmap, sizeof(BITMAP), (LPSTR)&bmp);
    DeleteObject(hBitmap);
    DEBUGNETC((LPSTR)"VBM_GETDEFSIZE[2.0] End\n\r");
    return(MAKELONG(bmp.bmWidth, bmp.bmHeight));
   }
   DEBUGNETC((LPSTR)"VBM_GETDEFSIZE[2.0] Allowing DefVBProc to handle\n\r");
  }
  break;

    // The re-size message - only in design time
  case WM_SIZE:
  {
   BITMAP bmp;
   HBITMAP hBitmap;
   RECT rect;
   POINT pos;

   DEBUGNETC((LPSTR)"WM_SIZE Start");
   GetWindowRect(hWnd, (RECT FAR *)&rect);
   pos.x=rect.left;
   pos.y=rect.top;
   ScreenToClient((HWND)GetWindowWord(hWnd, GWW_HWNDPARENT),
                  (POINT FAR *)&pos);

   hBitmap=LoadBitmap(hModDLL, MAKEINTRESOURCE(IDBMP_CLIENT));
   if(hBitmap)
   {
    GetObject(hBitmap, sizeof(BITMAP), (LPSTR)&bmp);
    MoveWindow(hWnd, pos.x, pos.y, bmp.bmWidth, bmp.bmHeight,
               SWP_NOZORDER | SWP_NOMOVE);
    DeleteObject(hBitmap);
    DEBUGNETC((LPSTR)"WM_SIZE End");
    return(0);
   }
   DEBUGNETC((LPSTR)"WM_SIZE End - Bitmap load failure!");
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

   DEBUGNETC((LPSTR)"WM_PAINT Start");
   if(VBGetMode()==MODE_RUN) break;

   GetWindowRect(hWnd, (RECT FAR *)&rect);
   BeginPaint(hWnd, (PAINTSTRUCT FAR *)&ps);
   hBitmap=LoadBitmap(hModDLL, MAKEINTRESOURCE(IDBMP_CLIENT));
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
   DEBUGNETC((LPSTR)"WM_PAINT End");
  }
  break;
#endif // DESIGN_TIME


    // The control is closing - free everything.
  case WM_NCDESTROY:
  {
   DEBUGNETC((LPSTR)"WM_NCDESTROY Begin");

    // If there is still a LineDelimiter string, destroy it
   if(lpNetClient->hszLineDelimiter)
   {
    VBDestroyHsz(lpNetClient->hszLineDelimiter);
    lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
    lpNetClient->hszLineDelimiter=NULL;
   }

    // If there is still a RawLineDelimiter string, destroy it
   if(lpNetClient->hszRawLineDelimiter)
   {
    VBDestroyHsz(lpNetClient->hszRawLineDelimiter);
    lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
    lpNetClient->hszRawLineDelimiter=NULL;
   }

    // If there is still a SendBlock string, destroy it
   if(lpNetClient->hlSendBlock)
   {
    VBDestroyHlstr(lpNetClient->hlSendBlock);
    lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
    lpNetClient->hlSendBlock=0;
   }

    // If there is still a RecvBlock string, destroy it
   if(lpNetClient->hlRecvBlock)
   {
    VBDestroyHlstr(lpNetClient->hlRecvBlock);
    lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
    lpNetClient->hlRecvBlock=NULL;
   }

     // If there is still a RecvBuffer area, destroy it
   if(lpNetClient->hRecvBuffer)
   {
    GlobalUnlock(lpNetClient->hRecvBuffer);
    GlobalFree(lpNetClient->hRecvBuffer);
    lpNetClient->hRecvBuffer=0;
    lpNetClient->lpRecvBuffer=NULL;
    lpNetClient->sRecvCount=0;
   }

     // If there is still a SendBuffer area, destroy it
   if(lpNetClient->hSendBuffer)
   {
    GlobalUnlock(lpNetClient->hSendBuffer);
    GlobalFree(lpNetClient->hSendBuffer);
    lpNetClient->hSendBuffer=0;
    lpNetClient->lpSendBuffer=NULL;
    lpNetClient->sSendCount=0;
   }

     // If there is still a RecvTemp area, destroy it
   if(lpNetClient->hRecvTemp)
   {
    GlobalUnlock(lpNetClient->hRecvTemp);
    GlobalFree(lpNetClient->hRecvTemp);
    lpNetClient->lpRecvTemp=NULL;
    lpNetClient->hRecvTemp=0;
    lpNetClient->sRecvTempSize=0;
   }

     // If there is still a SendTemp area, destroy it
   if(lpNetClient->hSendTemp)
   {
    GlobalUnlock(lpNetClient->hSendTemp);
    GlobalFree(lpNetClient->hSendTemp);
    lpNetClient->hSendTemp=0;
    lpNetClient->lpSendTemp=NULL;
    lpNetClient->sSendTempSize=0;
   }

    // If the HostAddresses() list was full, delete it
   if(lpNetClient->htHostAddressList.iHostCount)
   {
    int iCounter;
    MOLE mTemp;

    for(iCounter=0;iCounter<(lpNetClient->htHostAddressList.iHostCount);
        iCounter++)
    {
     mTemp = lpNetClient->htHostAddressList.mHost[iCounter];
     DeleteMole(mTemp);
     lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
     lpNetClient->htHostAddressList.mHost[iCounter] = 0;
    }
    lpNetClient->htHostAddressList.iHostCount = 0;
   }

    // If the HostAliases() list was full, delete it
   if(lpNetClient->htHostAliasList.iHostCount)
   {
    int iCounter;
    MOLE mTemp;

    for(iCounter=0;iCounter<(lpNetClient->htHostAliasList.iHostCount);
        iCounter++)
    {
     mTemp = lpNetClient->htHostAliasList.mHost[iCounter];
     DeleteMole(mTemp);
     lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
     lpNetClient->htHostAliasList.mHost[iCounter] = 0;
    }
    lpNetClient->htHostAliasList.iHostCount = 0;
   }

     // If there is still a live connection, hang up
   if(lpNetClient->sSocket!=SOCKET_ERROR)
   {
    netCloseSocket(hWnd, lpNetClient->sSocket);
    lpNetClient->bConnect=FALSE;
    lpNetClient->sSocket=(SOCKET)SOCKET_ERROR;
   }
   DEBUGNETC((LPSTR)"WM_NCDESTROY End");
  }
  break; // WM_NCDESTROY


    // Another control was created, initialize the pieces
  case VBM_CREATED:
  {
   LPHOSTENT lpheHost;
   int iError;
   char cChar;
   HSZ hszTempOne,
       hszTempTwo;
   int iCounter;
   LPSTR lpTemp;
   LPSTR FAR *lppTemp;
   MOLE mTemp;

   DEBUGNETC((LPSTR)"VBM_CREATED Begin");

    // Haven't received anything YET! ;)
   lpNetClient->sRecvCount=0;
   lpNetClient->sSendCount=0;

    // These are allocated upon a FD_CONNECT
    // and destroyed by a WM_NCDESTROY
   lpNetClient->sRecvTempSize=0;
   lpNetClient->sSendTempSize=0;
   lpNetClient->hRecvTemp=0;
   lpNetClient->hSendTemp=0;
   lpNetClient->lpRecvTemp=NULL;
   lpNetClient->lpSendTemp=NULL;

    // Alpha #3 feature, Flag waiting for FD_WRITE
   lpNetClient->fCanSend=FALSE;

    // Alpha #6 feature, Multiple aliases and addresses
   lpNetClient->htHostAddressList.iHostCount = 0;
   lpNetClient->htHostAliasList.iHostCount = 0;

    // HostName/HostAddr MOLE initialization (Alpha #6)
   lpNetClient->mHostName = 0;
   lpNetClient->mHostAddr = 0;

    // Setup a "null string" to point to
   cChar='\0';

    // If there was no default line delimiter, make a null one
   if(!lpNetClient->hszLineDelimiter)
   {
    hszTempOne=VBCreateHsz((_segment)hCtl, (LPSTR)&cChar);
    hszTempTwo=VBCreateHsz((_segment)hCtl, (LPSTR)&cChar);

    lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
    lpNetClient->hszRawLineDelimiter=hszTempOne;
    lpNetClient->hszLineDelimiter=hszTempTwo;

    DEBUGNETC((LPSTR)"VBM_CREATED Created empty LineDelimiter strings");
   }
   else
   {
    char szBuffer[256];
    HSZ  hszTemp;
    LPSTR lpDelim;
    LPSTR lpRaw;

    DEBUGNETC((LPSTR)"VBM_CREATED Building RawDelimiter string");
        // There is a default one: Make our raw string
    hszTemp=lpNetClient->hszLineDelimiter;
    lpDelim=VBDerefHsz(hszTemp);

    lpRaw=(LPSTR)szBuffer;
    while(*lpDelim)
    {
     if( (*lpDelim=='\\') && *(lpDelim+1))
     {
      lpDelim++;
      switch(*lpDelim)
      {
       case 'n':
       *lpRaw++='\n';
        lpDelim++;
       break;

       case 't':
       *lpRaw++='\t';
        lpDelim++;
       break;

       case 'r':
       *lpRaw++='\r';
        lpDelim++;
       break;

       default:
       *lpRaw++=*lpDelim++;
       break;
      }
     }
     else *lpRaw++=*lpDelim++;
    }
    *lpRaw='\0';

    hszTempOne=VBCreateHsz((_segment)hCtl, (LPSTR)szBuffer);
    if(!hszTempOne)
    {
     hszTempOne=VBCreateHsz((_segment)hCtl, (LPSTR)&cChar);
    }

    lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
    lpNetClient->hszRawLineDelimiter=hszTempOne;
   }

    // This pretty much be-rids us of MODEL_fInvisAtRun
   if(VBGetMode()==MODE_RUN)
   {
    ShowWindow(hWnd, SW_HIDE);
   }
   else
   {
    ShowWindow(hWnd, SW_SHOWNA);
   }

   lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);

   DEBUGNETC((LPSTR)"VBM_CREATED Marker #7");

    // Initialize the timeout
   lpNetClient->sTimeOutStore=lpNetClient->sTimeOut;

       // Initialize the other properties
   lpNetClient->sSocket = (SOCKET)SOCKET_ERROR;
   lpNetClient->bConnect= FALSE;
   lpNetClient->sErrorNumber = 0;

   if(!lpNetClient->sSendSize)
        lpNetClient->sSendSize=DEFAULT_SENDSIZE;
   if(!lpNetClient->sRecvSize)
        lpNetClient->sRecvSize=DEFAULT_SENDSIZE;

   lpNetClient->sErrorNumber=0;

   DEBUGNETC((LPSTR)"VBM_CREATED Marker #7.5");

       // Use the local host name if nothing was loaded
   iError = gethostname((LPSTR)szLocalHostName,
                          sizeof(szLocalHostName));

   DEBUGNETC((LPSTR)"VBM_CREATED Marker #8");

       // The local host name wasn't returned!
   if(iError==SOCKET_ERROR)
   {   // Something REAL bad happend! Abort!

    DEBUGNETC((LPSTR)"VBM_CREATED ERROR: Couldn't retrieve local host's name!");

    return(NetClientError( hCtl, WSAGetLastError()));
   }

   lpNetClient->sErrorNumber=0;

   DEBUGNETC((LPSTR)"VBM_CREATED Marker #10");

   lpheHost = gethostbyname((LPSTR)szLocalHostName);

        // The local host name MUST be valid!
   if(!lpheHost)
   {   // This SHOULD have worked!!!

    DEBUGNETC((LPSTR)"VBM_CREATED: NULL returned by gethostbyname()");

    return(NetClientError( hCtl, WSAGetLastError()));
   }

   DEBUGNETC((LPSTR)"VBM_CREATED gethostbyname() returned a HOSTENT pointer");

       // Create a new HostName
   mTemp=AddMole(lpheHost->h_name);
   lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
   lpNetClient->mHostName=mTemp;

       // Create a new HostAddr
   mTemp=AddMole(inet_ntoa(*(struct in_addr FAR *)(lpheHost->h_addr)));
   lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
   lpNetClient->mHostAddr=mTemp;

       // Save the internet address for later (Connect)
   lpNetClient->saHost.sin_addr=*(struct in_addr FAR *)lpheHost->h_addr;

     // Rebuild the HostAliasList property array
   lppTemp=lpheHost->h_aliases;
   iCounter=0;
   if(lppTemp!=NULL)
   {
    while(*lppTemp!=NULL)
    {
     mTemp=AddMole(*lppTemp);
     lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
     if(!mTemp) break;
     lpNetClient->htHostAliasList.mHost[iCounter]=mTemp;
     iCounter++;
     *lppTemp++;
    }
   }
   lpNetClient->htHostAliasList.iHostCount=iCounter;

       // Rebuild the HostAddressList property array
   lppTemp=lpheHost->h_addr_list;
   iCounter=0;
   if(lppTemp!=NULL)
   {
    while(*lppTemp!=NULL)
    {
     lpTemp=inet_ntoa(*(struct in_addr FAR *)(*lppTemp));
     if(!lpTemp) break;
     mTemp=AddMole(lpTemp);
     lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
     if(!mTemp) break;
     lpNetClient->htHostAddressList.mHost[iCounter]=mTemp;
     iCounter++;
     *lppTemp++;
    }
   }
   lpNetClient->htHostAddressList.iHostCount=iCounter;

   DEBUGNETC((LPSTR)"VBM_CREATE End - 1");

   return(ERR_None);
  } // VBM_CREATED


#ifdef DESIGN_TIME
    // Here is how to implement a dialog box property!
  case VBM_INITPROPPOPUP:
  {
   switch(wParam)
   {
    case IPROP_NETCLIENT_ABOUT:
    {
     DEBUGNETC((LPSTR)"VBM_INITPROPPOPUP: About - Begin");

     if(bDialogInUse) return(NULL);
     bDialogInUse = TRUE;
     hDialogHwnd=CreateWindow(CLASS_ABOUTPOPUP, NULL,
            WS_POPUP,
            0, 0, 0, 0,
            NULL, NULL,
            hModDLL, NULL);
     DEBUGNETC((LPSTR)"VBM_INITPROPPOPUP: About - End");
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
   DEBUGNETC((LPSTR)szError);
#endif /* DEBUG_BUILD */

   switch(wParam)
   {

    /****************** SET: RecvSize *******************/
    // User is trying to reallocate the RecvBuffer
    case IPROP_NETCLIENT_RECVSIZE:
    {
     DEBUGNETC((LPSTR)"SET: RecvSize - Begin");
     if(lpNetClient->bConnect)      // Are we connected?
     {
        // Is the buffer empty?
      if(lpNetClient->sRecvCount)
      {
        // Give the user a chance to resolve his ways...
       err = clientFireOnRecv(hCtl);
       if(err) return(err);

       // VB routines were probably called - dereference again
       lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
        // Is it empty now?
       if(lpNetClient->sRecvCount)
       {
        DEBUGNETC((LPSTR)"SET: RecvSize - Buffer not empty!!!");
        return(ERR_None);     // Give up
       }
      }
     }

     lpNetClient->sRecvSize = *(SHORT *)&lParam;

        // Unlock and free the old buffer
     if(lpNetClient->hRecvBuffer)
     {
      GlobalUnlock(lpNetClient->hRecvBuffer);
      GlobalFree(lpNetClient->hRecvBuffer);
     }

        // Lock and load the receive buffer
     lpNetClient->hRecvBuffer=GlobalAlloc(GMEM_ZEROINIT | GMEM_MOVEABLE,
                                          lpNetClient->sRecvSize);
     lpNetClient->lpRecvBuffer=GlobalLock(lpNetClient->hRecvBuffer);
     DEBUGNETC((LPSTR)"SET: RecvSize - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_SENDSIZE


    /****************** SET: SendSize *******************/
    // User is trying to reallocate the SendBuffer
    case IPROP_NETCLIENT_SENDSIZE:
    {
     DEBUGNETC((LPSTR)"SET: SendSize - Begin");
     if(lpNetClient->sSendCount)
     {
        // Try flushing the SendBuffer to the net - might not work!
      DEBUGNETC((LPSTR)"SET: SendSize - Buffer not empty! ");
      netSendData(hCtl);
      if(lpNetClient->sSendCount)
        return(ERR_None);  // Couldn't flush it out, won't free it!
     }

     lpNetClient->sSendSize = *(SHORT *)&lParam;

        // Unlock and free the old buffer
     if(lpNetClient->hSendBuffer)
     {
      GlobalUnlock(lpNetClient->hSendBuffer);
      GlobalFree(lpNetClient->hSendBuffer);
     }

        // Lock and load the send buffer
     lpNetClient->hSendBuffer=GlobalAlloc(GMEM_ZEROINIT | GMEM_MOVEABLE,
                                        lpNetClient->sSendSize);
     lpNetClient->lpSendBuffer=GlobalLock(lpNetClient->hSendBuffer);

     DEBUGNETC((LPSTR)"SET: SendSize - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_SENDSIZE


    /****************** SET: HostName *******************/
    // Try to set the HostName - We only allow valid ones!
    // If an error happens, the appropriate Windows Sockets error is
    // returned
    case IPROP_NETCLIENT_HOSTNAME:
    {
     LPSTR lpHostName;
     LPHOSTENT lpheHost;
     MOLE mTemp;
     LPSTR lpTemp;
     LPSTR FAR *lppTemp;
     int iCounter;

     DEBUGNETC((LPSTR)"SET: HostName - Begin");

       // Free any previous HostName
     if(lpNetClient->mHostName)
     {
      mTemp = lpNetClient->mHostName;
      DeleteMole(mTemp);
      lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
      lpNetClient->mHostName = 0;
     }
       // Free any previous HostAddr
     if(lpNetClient->mHostAddr)
     {
      mTemp = lpNetClient->mHostAddr;
      DeleteMole(mTemp);
      lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
      lpNetClient->mHostAddr = 0;
     }

     if(!lParam)
     {
        // 'Twas an invalid string - tell the app
      DEBUGNETC((LPSTR)"SET: HostName - End - lParam was NULL!!!!");

      return(netReturnError(VBERR_INVALIDPROPERTYVALUE, 0, NULL));
     }

        // lParam points to a locked string during VBM_SETPROPERTY on a HSZ
     lpHostName=*(LPSTR *)&lParam;

        // Try looking up the name
     lpheHost=gethostbyname(lpHostName);

        // Is the pointer Valid? And the name?
     if(lpheHost && lpheHost->h_name)
     {
      DEBUGNETC((LPSTR)"SET: HostName - gethostbyname() returned a HOSTENT pointer!!");

        // Create a new HostName
      mTemp=AddMole(lpheHost->h_name);
      if(!mTemp)
      {
       DEBUGNETC((LPSTR)"SET: HostName - End - AddMole() returned 0! (HostName)");
       return(ERR_None);
      }
      lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
      lpNetClient->mHostName=mTemp;

        // Create a new HostAddr
      mTemp=AddMole(inet_ntoa(*(struct in_addr FAR *)(lpheHost->h_addr)));
      if(!mTemp)
      {
       mTemp = lpNetClient->mHostName;
       DeleteMole(mTemp);
       DEBUGNETC((LPSTR)"SET: HostName - End - AddMole() returned 0! (HostAddr)");
       return(ERR_None);
      }
      lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
      lpNetClient->mHostAddr=mTemp;

        // Save the internet address for later (Connect)
      lpNetClient->saHost.sin_addr=*(struct in_addr FAR *)lpheHost->h_addr;

        // Delete the HostAliasList property array
      if(lpNetClient->htHostAliasList.iHostCount)
      {
       for(iCounter=0; (iCounter < (lpNetClient->htHostAliasList.iHostCount)) &&
                       (iCounter < MAX_HOSTTABLESIZE); iCounter++)
       {
        mTemp = lpNetClient->htHostAliasList.mHost[iCounter];
        DeleteMole(mTemp);
        lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
        lpNetClient->htHostAliasList.mHost[iCounter] = 0;
       }
       lpNetClient->htHostAliasList.iHostCount=0;
      }

        // Rebuild the HostAliasList property array
      lppTemp=lpheHost->h_aliases;
      iCounter=0;
      if(lppTemp!=NULL)
      {
       while(*lppTemp!=NULL)
       {
        mTemp=AddMole(*lppTemp);
        lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
        if(!mTemp) break;
        lpNetClient->htHostAliasList.mHost[iCounter]=mTemp;
        iCounter++;
        *lppTemp++;
       }
      }
      lpNetClient->htHostAliasList.iHostCount=iCounter;

        // Delete the HostAddressList property array
      if(lpNetClient->htHostAddressList.iHostCount)
      {
       for(iCounter=0; (iCounter < (lpNetClient->htHostAddressList.iHostCount)) &&
                       (iCounter < MAX_HOSTTABLESIZE); iCounter++)
       {
        mTemp = lpNetClient->htHostAddressList.mHost[iCounter];
        DeleteMole(mTemp);
        lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
        lpNetClient->htHostAddressList.mHost[iCounter] = 0;
       }
       lpNetClient->htHostAddressList.iHostCount=0;
      }

        // Rebuild the HostAddressList property array
      lppTemp=lpheHost->h_addr_list;
      iCounter=0;
      if(lppTemp!=NULL)
      {
       while(*lppTemp!=NULL)
       {
        lpTemp=inet_ntoa(*(struct in_addr FAR *)(*lppTemp));
        if(!lpTemp) break;
        mTemp=AddMole(lpTemp);
        lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
        if(!mTemp) break;
        lpNetClient->htHostAddressList.mHost[iCounter]=mTemp;
        iCounter++;
        *lppTemp++;
       }
      }
      lpNetClient->htHostAddressList.iHostCount=iCounter;

      DEBUGNETC((LPSTR)"SET: HostName - End - 1");

      return(ERR_None);
     }
     else
     {
      DEBUGNETC((LPSTR)"SET: HostName - End - NULL returned by gethostbyname() ");

      return(NetClientError( hCtl, WSAGetLastError()));
     }
    } // IPROP_NETCLIENT_HOSTNAME


    /****************** SET: HostAddr *******************/
    // Do a reverse name query - we only accept valid IP addresses
    case IPROP_NETCLIENT_HOSTADDR:
    {
     LPSTR lpHostAddr;
     LPHOSTENT lpheHost;
     IN_ADDR inHost;
     LONG lHost;
     MOLE mTemp;
     LPSTR lpTemp;
     LPSTR FAR *lppTemp;
     int iCounter;

     DEBUGNETC((LPSTR)"SET: HostAddr - Begin");

       // Free any previous HostName
     if(lpNetClient->mHostName)
     {
      mTemp = lpNetClient->mHostName;
      DeleteMole(mTemp);
      lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
      lpNetClient->mHostName = 0;
     }

       // Free any previous HostAddr
     if(lpNetClient->mHostAddr)
     {
      mTemp = lpNetClient->mHostAddr;
      DeleteMole(mTemp);
      lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
      lpNetClient->mHostAddr = 0;
     }

       // No string entered?!?!
     if(!lParam)
     {
      DEBUGNETC((LPSTR)"SET: HostAddr - End - lParam was NULL!");

      return(netReturnError(VBERR_INVALIDPROPERTYVALUE, 0, NULL ));
     }

     lpHostAddr=*(LPSTR *)&lParam;
     DEBUGNETC((LPSTR)lpHostAddr);

        // Change string to a long (really an in_addr)
     lHost=inet_addr(lpHostAddr);
     if(lHost==INADDR_NONE)
     {
      DEBUGNETC((LPSTR)"SET: HostAddr - End - Bad Host address format");
      return(netReturnError( WSAERR_BadHostAddr, WSAERR_BadHostAddr, NULL ));
     }

        // Assign the in_addr to be the return from inet_addr()
     inHost.s_addr=(unsigned)lHost;
     lpNetClient->saHost.sin_addr=inHost;

        // Query the database
     lpheHost=gethostbyaddr((LPSTR)&lHost, 4, PF_INET);

        // Is the pointer Valid? And the name?
     if(lpheHost && lpheHost->h_name)
     {
      DEBUGNETC((LPSTR)"SET: HostAddr - gethostbyaddr() returned a HOSTENT pointer!!");

        // Create a new HostName
      mTemp=AddMole(lpheHost->h_name);
      lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
      if(!mTemp)
      {
       DEBUGNETC((LPSTR)"SET: HostAddr - End - AddMole() -1- failed!");
       return(ERR_None);
      }
      lpNetClient->mHostName=mTemp;

        // Create a new HostAddr
      mTemp=AddMole(inet_ntoa(*(struct in_addr FAR *)(lpheHost->h_addr)));
      lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
      if(!mTemp)
      {
        // If one fails, destroy the other
       DEBUGNETC((LPSTR)"SET: HostAddr - End - AddMole() -2- failed!");
       if(lpNetClient->mHostName)
       {
        mTemp = lpNetClient->mHostName;
        DeleteMole(mTemp);
        lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
        lpNetClient->mHostName = 0;
       }
       return(ERR_None);
      }
      lpNetClient->mHostAddr=mTemp;

        // Save the internet address for later (Connect)
      lpNetClient->saHost.sin_addr=*(struct in_addr FAR *)lpheHost->h_addr;

        // Delete the HostAliasList property array
      if(lpNetClient->htHostAliasList.iHostCount)
      {
       for(iCounter=0; (iCounter < (lpNetClient->htHostAliasList.iHostCount)) &&
                       (iCounter < MAX_HOSTTABLESIZE); iCounter++)
       {
        mTemp = lpNetClient->htHostAliasList.mHost[iCounter];
        DeleteMole(mTemp);
        lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
        lpNetClient->htHostAliasList.mHost[iCounter] = 0;
       }
       lpNetClient->htHostAliasList.iHostCount=0;
      }

        // Rebuild the HostAliasList property array
      lppTemp=lpheHost->h_aliases;
      iCounter=0;
      if(lppTemp!=NULL)
      {
       while(*lppTemp!=NULL)
       {
        mTemp=AddMole(*lppTemp);
        lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
        if(!mTemp) break;
        lpNetClient->htHostAliasList.mHost[iCounter]=mTemp;
        iCounter++;
        *lppTemp++;
       }
      }
      lpNetClient->htHostAliasList.iHostCount=iCounter;

        // Delete the HostAddressList property array
      if(lpNetClient->htHostAddressList.iHostCount)
      {
       for(iCounter=0; (iCounter < (lpNetClient->htHostAddressList.iHostCount)) &&
                       (iCounter < MAX_HOSTTABLESIZE); iCounter++)
       {
        mTemp = lpNetClient->htHostAddressList.mHost[iCounter];
        DeleteMole(mTemp);
        lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
        lpNetClient->htHostAddressList.mHost[iCounter]=0;
       }
       lpNetClient->htHostAddressList.iHostCount=0;
      }

        // Rebuild the HostAddressList property array
      lppTemp=lpheHost->h_addr_list;
      iCounter=0;
      if(lppTemp!=NULL)
      {
       while(*lppTemp!=NULL)
       {
        lpTemp=inet_ntoa(*(struct in_addr FAR *)(*lppTemp));
        if(!lpTemp) break;
        mTemp=AddMole(lpTemp);
        lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
        if(!mTemp) break;
        lpNetClient->htHostAddressList.mHost[iCounter]=mTemp;
        iCounter++;
        *lppTemp++;
       }
      }
      lpNetClient->htHostAddressList.iHostCount=iCounter;

      DEBUGNETC((LPSTR)"SET: HostAddr - End - 1");

      return(ERR_None);
     }
     else
     {
      DEBUGNETC((LPSTR)"SET: HostAddr - End - NULL returned by gethostbyname() ");

      return(NetClientError( hCtl, WSAGetLastError()));
     }
    } // IPROP_NETCLIENT_HOSTADDR


    /****************** SET: LineDelimiter **************/
    // Setting the Line terminator
    case IPROP_NETCLIENT_LINEDELIMITER:
    {
     char szBuffer[256];
     HSZ  hszTemp;
     LPSTR lpRaw, lpDelim;

     DEBUGNETC((LPSTR)"SET: LineDelimiter - Begin");
         // Set the RawLineDelimiter appropriately
     lpDelim=*(LPSTR *)&lParam;
     lpRaw=(LPSTR)szBuffer;
     while(*lpDelim)
     {
      if((*lpDelim=='\\')&&*(lpDelim+1))
      {
       lpDelim++;
       switch(*lpDelim)
       {                  // No, these aren't needed.. but they sure are
        case 'n':         // nice for C addicts. You can fill the
        *lpRaw++='\n';    // LineDelimiter property with "\r\n" for example.
        lpDelim++;
        break;

        case 't':
        *lpRaw++='\t';
        lpDelim++;
        break;

        case 'r':
        *lpRaw++='\r';
        lpDelim++;
        break;

        default:
        *lpRaw++=*lpDelim++;
        break;
       }
      }
      else *lpRaw++=*lpDelim++;
     }
     *lpRaw='\0';

     if(lpNetClient->hszRawLineDelimiter)
            VBDestroyHsz(lpNetClient->hszRawLineDelimiter);

     hszTemp=VBCreateHsz((_segment)hCtl, (LPSTR)szBuffer);
     lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
     lpNetClient->hszRawLineDelimiter=hszTemp;

     DEBUGNETC((LPSTR)"SET: LineDelimiter - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_LINEDELIMITER


    /****************** SET: RecvLine *******************/
    // Can't set!!!
    case IPROP_NETCLIENT_RECVLINE:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"RecvLine"));


    /****************** SET: RecvBlock ******************/
    // Can't set!!!
    case IPROP_NETCLIENT_RECVBLOCK:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"RecvBlock"));


    /****************** SET: RecvCount ******************/
    // Can't set!!!
    case IPROP_NETCLIENT_RECVCOUNT:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"RecvCount"));


    /****************** SET: SendCount ******************/
    // Can't set!!!
    case IPROP_NETCLIENT_SENDCOUNT:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"RecvLine"));


    /****************** SET: ErrorMessage ****************/
    // Can't set!!!
    case IPROP_NETCLIENT_ERRORMESSAGE:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"ErrorMessage"));


    /****************** SET: Version ********************/
    // Can't set!!!
    case IPROP_NETCLIENT_VERSION:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"Version"));


    /****************** SET: HostAliasList **************/
    // Can't set!!!
    case IPROP_NETCLIENT_HOSTALIASLIST:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"HostAliasList"));


    /****************** SET: HostAliasCount *************/
    // Can't set!!!
    case IPROP_NETCLIENT_HOSTALIASCOUNT:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"HostAliasCount"));


    /****************** SET: HostAddressList ************/
    // Can't set!!!
    case IPROP_NETCLIENT_HOSTADDRESSLIST:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"HostAddressList"));


    /****************** SET: HostAddressCount ************/
    // Can't set!!!
    case IPROP_NETCLIENT_HOSTADDRESSCNT:
        return(netReturnError(VBERR_READONLY, 0, (LPSTR)"HostAddressCount"));


    /****************** SET: LocalPort ******************/
    // Only set if we aren't connected
    case IPROP_NETCLIENT_LOCALPORT:
    {
     DEBUGNETC((LPSTR)"SET: LocalPort - Begin");
     if(!lpNetClient->bConnect)
        lpNetClient->sLocalPort=*(SHORT *)&lParam;
     DEBUGNETC((LPSTR)"SET: LocalPort - Begin");
     return(ERR_None);
    } // IPROP_NETCLIENT_LOCALPORT


    /****************** SET: RemotePort *****************/
    // Only set if we aren't connected
    case IPROP_NETCLIENT_REMOTEPORT:
    {
     DEBUGNETC((LPSTR)"SET: RemotePort - Begin");
     if(!lpNetClient->bConnect)
        lpNetClient->sRemotePort=*(SHORT *)&lParam;
     DEBUGNETC((LPSTR)"SET: RemotePort - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_REMOTEPORT


    /****************** SET: LocalService ***************/
    // Handle Setting of LocalService
    case IPROP_NETCLIENT_LOCALSERVICE:
    {
     LPSTR lpLocalService;
     LPSERVENT lpseServEnt;

     DEBUGNETC((LPSTR)"SET: LocalService - Begin");
     if(!lParam)
     {
      lpNetClient->sLocalPort = 0;
      return(netReturnError( VBERR_INVALIDPROPERTYVALUE, 0, NULL ));
     }

     if(!lpNetClient->bConnect)
     {
      lpLocalService=*(LPSTR *)&lParam;

        // Get service by name, but don't do families... yet?
      lpseServEnt = getservbyname( lpLocalService, NULL);

      if(lpseServEnt)
      {
        // Copy out the port number
       lpNetClient->sLocalPort = ntohs(lpseServEnt->s_port);
      }
      else
      {
       lpNetClient->sLocalPort = 0;
       DEBUGNETC((LPSTR)"SET: LocalService - getservbyname() returned NULL!");
       return(NetClientError( hCtl, WSAGetLastError()));
      }
     }
     DEBUGNETC((LPSTR)"SET: LocalService - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_LOCALSERVICE


    /****************** SET: RemoteService **************/
    // Handle Setting of RemoteService
    case IPROP_NETCLIENT_REMOTESERVICE:
    {
     LPSTR lpRemoteService;
     LPSERVENT lpseServEnt;

     DEBUGNETC((LPSTR)"SET: RemoteService - Begin");
     if(!lParam)
     {
      lpNetClient->sRemotePort = 0;
      return(netReturnError( VBERR_INVALIDPROPERTYVALUE, 0, NULL ));
     }

     if(!lpNetClient->bConnect)
     {
      lpRemoteService=*(LPSTR *)&lParam;

        // Get service by name, but don't do families... yet?
      lpseServEnt = getservbyname( lpRemoteService, NULL);

      if(lpseServEnt)
      {
        // Copy out the port number
       lpNetClient->sRemotePort = ntohs(lpseServEnt->s_port);
      }
      else
      {
       lpNetClient->sRemotePort = 0;
       DEBUGNETC((LPSTR)"SET: RemoteService - getservbyname() returned NULL!");
       return(NetClientError( hCtl, WSAGetLastError()));
      }
     }

     DEBUGNETC((LPSTR)"SET: RemoteService - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_REMOTESERVICE


    /****************** SET: Connect ********************/
    // Try connecting to a remote host
    case IPROP_NETCLIENT_CONNECT:
    {
     SOCKET sSocket;
     LONG lNonBlocking = TRUE;
     SOCKADDR_IN saLocal;
     int iError;

     DEBUGNETC((LPSTR)"SET: Connect - Begin");

        // If we are alread connected, set an error condition.
     if((lpNetClient->bConnect) && *(BOOL *)&lParam)
     {
      DEBUGNETC((LPSTR)"SET: Connect - End - Already connected");
      return(netReturnError( WSAERR_AlreadyConn, WSAERR_AlreadyConn, NULL ));
     }

     if(!(lpNetClient->bConnect) && !*(BOOL *)&lParam)
     {
      DEBUGNETC((LPSTR)"SET: Connect - End - Pointless!");
      return(ERR_None);        // Ignore no connection close attempts
     }

        // Catch the connected close attempt
     if((lpNetClient->bConnect) && !*(BOOL *)&lParam)
     {
      netCloseSocket(hWnd, lpNetClient->sSocket);
      lpNetClient->bConnect=FALSE;
      lpNetClient->sSocket=(SOCKET)SOCKET_ERROR;
      DEBUGNETC((LPSTR)"SET: Connect - End - Closed connection!");
      return(ERR_None);
     }

        // In the odd case of a socket without connection, close the
        // old socket (like a non-responsive host, and a user tries
        // to connect again while we wait for FD_CONNECT).
     if(lpNetClient->sSocket != SOCKET_ERROR)
     {
      netCloseSocket(hWnd, lpNetClient->sSocket);
      lpNetClient->sSocket=(SOCKET)SOCKET_ERROR;
     }

        // Create a new socket
     sSocket = socket(PF_INET, SOCK_STREAM, 0);

        // Make sure it's valid
     if( sSocket==INVALID_SOCKET )
     {
      NetClientError( hCtl, WSAGetLastError());
      return(ERR_None);
     }

     WSAAsyncSelect(sSocket, hWnd, CM_NETACTIVITY,
                    FD_CONNECT | FD_READ | FD_WRITE | FD_CLOSE);

        // Save the socket for requests
     lpNetClient->sSocket=(SOCKET)sSocket;

         // Set the family and port
     lpNetClient->saHost.sin_family=AF_INET;
     lpNetClient->saHost.sin_port=htons(lpNetClient->sRemotePort);

         // Bind the socket to our local side - Un-Neccisary
     DEBUGNETC((LPSTR)"BIND: Begin");
     saLocal.sin_family=AF_INET;
     saLocal.sin_addr.s_addr=INADDR_ANY;
     saLocal.sin_port=htons(lpNetClient->sLocalPort);

     if(bind(sSocket, (LPSOCKADDR)&saLocal,
         sizeof(saLocal))==SOCKET_ERROR)
     {
      NetClientError( hCtl, WSAGetLastError());
      netCloseSocket( hWnd, sSocket);
      lpNetClient->sSocket=(SOCKET)SOCKET_ERROR;
      return(ERR_None);
     }
     DEBUGNETC((LPSTR)"BIND: End");   /**/

        // Start the connection to the remote host
     if(connect(sSocket, (LPSOCKADDR)&(lpNetClient->saHost),
             sizeof(lpNetClient->saHost)) == SOCKET_ERROR)
     {
        // Make sure the error we are getting is the WouldBlock warning...
      iError = WSAGetLastError();
      if(iError == WSAEWOULDBLOCK)
      {
       DEBUGNETC((LPSTR)"SET: Connect - End - All ok");
       return(ERR_None);   // Alright!
      }

      netCloseSocket(hWnd, sSocket);
      lpNetClient->sSocket=(SOCKET)SOCKET_ERROR;

        // Ah oh, something else bad went wrong... tell the user
      DEBUGNETC((LPSTR)"SET: Connect - End - connect() error");
      return(NetClientError( hCtl, iError));
     }

        // Must be a LocalHost connect (i.e. loopback)
     DEBUGNETC((LPSTR)"SET: Connect - End - Loopback?!?!");

        // Trick it!
     PostMessage(hWnd, CM_NETACTIVITY, sSocket,
                 WSAMAKESELECTREPLY(FD_CONNECT, 0));
     return(ERR_None);
    }
    break; // IPROP_NETCLIENT_CONNECT


    /****************** SET: TimeOut ********************/
    // Set the TimeOut property
    case IPROP_NETCLIENT_TIMEOUT:
    {
     DEBUGNETC((LPSTR)"SET: TimeOut - Begin");
        // If setting it to 0, mark for WM_TIMER that it shan't go on
     if(lParam==0)
     {
      lpNetClient->sTimeOut=lpNetClient->sTimeOutStore=0;
      return(ERR_None);
     }

        // Set the new TimeOut value
     lpNetClient->sTimeOutStore=*(SHORT *)&lParam;

     if(lpNetClient->sTimeOut)
     {
      lpNetClient->sTimeOut=lpNetClient->sTimeOutStore;
      return(ERR_None);
     }

     if(!SetTimer(hWnd, IDTIME_TIMEOUT, 1000, NULL))
     {
      return(netReturnError( WSAERR_NoTimers, WSAERR_NoTimers, 0 ));
     }

     lpNetClient->sTimeOut=lpNetClient->sTimeOutStore;

     DEBUGNETC((LPSTR)"SET: TimeOut - End");
     return(ERR_None);
    }
    break; // IPROP_NETCLIENT_TIMEOUT


    /****************** SET: Line ***********************/
    /****************** SET: SendLine *******************/
    // Send a line
    case IPROP_NETCLIENT_LINE:
    case IPROP_NETCLIENT_SENDLINE:
    {
     LPSTR lpLine;
     MOLE mLine;

     DEBUGNETC((LPSTR)"SET: SendLine - Begin");

     if(!lParam)
     {
      DEBUGNETC((LPSTR)"SET: SendLine - No lParam!");
      return(netReturnError( VBERR_INVALIDPROPERTYVALUE, 0, NULL ));
     }

     lpLine=*(LPSTR *)&lParam;

        // Put some data in the SendLine property
     DEBUGNETC((LPSTR)"SET: SendLine - Marker #1");

        // Repackage the HSZ string as a HLSTR
     mLine = CreateMole(lpLine, (USHORT)lstrlen(lpLine));

     if(!mLine)
     {
      DEBUGNETC((LPSTR)"SET: SendLine - End - CreateMole() returned 0!");
      return(netReturnError(VBERR_OUTOFMEMORY, 0, NULL));
     }

        // Send the string as a block
     DEBUGNETC((LPSTR)"SET: SendLine - calling putSendData()");

     if(putSendData(hCtl, mLine) == SOCKET_ERROR)
     {
      DeleteMole(mLine);
      return(netReturnError(WSAERR_SendOverFlow, WSAERR_SendOverFlow, NULL ));
     }

        // Unlock and destroy the strings
     DeleteMole(mLine);
     DEBUGNETC((LPSTR)"SET: SendLine - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_SENDLINE/IPROP_NETCLIENT_LINE


    /****************** SET: Block **********************/
    /****************** SET: SendBlock ******************/
    // Send a block
    case IPROP_NETCLIENT_BLOCK:
    case IPROP_NETCLIENT_SENDBLOCK:
    {
     HLSTR hlLine;

     DEBUGNETC((LPSTR)"SET: SendBlock - Begin");

        // We MUST copy the HLSTR to another one (that we own)
        //  before we can even USE IT! Seems that HLSTR must be
        //  stored in VB's local address space.
     hlLine=VBCreateHlstr(NULL, 0);
     VBSetHlstr((HLSTR FAR *)&hlLine, (LPVOID)(*(HLSTR *)&lParam), 1);

        // Send the data passed as a block
     putSendData(hCtl, (MOLE)hlLine);

     VBDestroyHlstr(hlLine);

     DEBUGNETC((LPSTR)"SET: SendBlock - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_SENDBLOCK/IPROP_NETCLIENT_BLOCK


    /****************** SET: Socket *********************/
    // Allow the user to set the socket parameter
    case IPROP_NETCLIENT_SOCKET:
    {  //"xxx.xxx.xxx.xxx"
     char szHostAddr[16];
     LPSTR lpHostAddr;
     SOCKET sSocket = *(SOCKET *)&lParam;
     SOCKADDR_IN saHost;
     int iSize = sizeof(saHost);

     DEBUGNETC((LPSTR)"SET: Socket - Start");

     if((sSocket!=0) && (sSocket!=SOCKET_ERROR))
     {
      if((lpNetClient->sSocket != 0) &&
         (lpNetClient->sSocket != SOCKET_ERROR))
      {
       netCloseSocket(hWnd, lpNetClient->sSocket);
      }
      lpNetClient->sSocket=sSocket;

      if(getpeername(sSocket, (LPSOCKADDR)&saHost, (LPINT)&iSize) ==
         SOCKET_ERROR)
        return(NetClientError(hCtl, WSAGetLastError()));

      DEBUGNETC((LPSTR)"SET: Socket - getpeername() was successful");

      lpHostAddr=inet_ntoa(saHost.sin_addr);

      if(lpHostAddr == NULL)
        return(NetClientError(hCtl, WSAGetLastError()));

      DEBUGNETC((LPSTR)"SET: Socket - inet_ntoa() was successful");

      if(lstrlen(lpHostAddr) < 16)
       lstrcpy((LPSTR)szHostAddr, lpHostAddr);
      else
      {
       DEBUGNETC((LPSTR)"SET: Socket - host address is more than 15 chars?!?!");
       return(ERR_None);
      }

      lpNetClient->bConnect = FALSE;

      VBSetControlProperty(hCtl, IPROP_NETCLIENT_HOSTADDR, (LONG)(LPSTR)szHostAddr);

      lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);

      DEBUGNETC((LPSTR)"SET: Socket - set HostAddr.. properties");

      lpNetClient->sRemotePort=(SHORT)htons(saHost.sin_port);

      if(getsockname(sSocket, (LPSOCKADDR)&saHost, (LPINT)&iSize) ==
         SOCKET_ERROR)
        return(NetClientError(hCtl, WSAGetLastError()));

      DEBUGNETC((LPSTR)"SET: Socket - getsockname() was successful");

      lpNetClient->sLocalPort=(SHORT)htons(saHost.sin_port);

      netConnect(hCtl, sSocket);

      DEBUGNETC((LPSTR)"SET: Socket - End");
      return(ERR_None);
     }
     else
     {  // User obviously wants to close the connection, if any.
      return(VBSetControlProperty(hCtl, IPROP_NETCLIENT_CONNECT, (LONG)FALSE));
     }
    }
    break; // IPROP_NETCLIENT_SOCKET


    /****************** SET: Debug **********************/
    // Set SO_DEBUG
    case IPROP_NETCLIENT_DEBUG:
    {
     BOOL fDebug = *(BOOL *)&lParam;

     DEBUGNETC((LPSTR)"SET: Debug - Begin");

     if((lpNetClient->sSocket != 0) &&
        (lpNetClient->sSocket != SOCKET_ERROR))
     {
      if(setsockopt(lpNetClient->sSocket, SOL_SOCKET, SO_DEBUG,
                    (LPSTR)&fDebug, sizeof(fDebug))
               != SOCKET_ERROR)
      {
       DEBUGNETC((LPSTR)"SET: Debug - End - Ok");
       return(ERR_None);
      }
      else return(NetClientError(hCtl, WSAGetLastError()));
     }
     return(netReturnError(WSAERR_NotConn, WSAERR_NotConn, NULL));
    }
    break; // IPROP_NETCLIENT_DEBUG


    /****************** SET: Host ***********************/
    // Try HostAddr and then HostName
    case IPROP_NETCLIENT_HOST:
    {
     ERR err;

     DEBUGNETC((LPSTR)"SET: Host - Begin");

     lpNetClient->sErrorNumber = 0;
     err=VBSetControlProperty(hCtl, IPROP_NETCLIENT_HOSTNAME, lParam);
     lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
     if(lpNetClient)
     {
      if(!(lpNetClient->mHostName))
      {
       DEBUGNETC((LPSTR)"SET: Host - HostName failed... Trying HostAddr");

       err=VBSetControlProperty(hCtl, IPROP_NETCLIENT_HOSTADDR, lParam);
      }
      else
      {
       DEBUGNETC((LPSTR)"SET: Host - HostName was successful!");
      }
     }

     DEBUGNETC((LPSTR)"SET: Host - End");
     return(err);
    }
    break; // IPROP_NETCLIENT_HOST

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
   DEBUGNETC((LPSTR)szError);
#endif /* DEBUG_BUILD */

   switch(wParam)
   {

    /****************** GET: Line     *******************/
    /****************** GET: RecvLine *******************/
    // Receive a line
    case IPROP_NETCLIENT_LINE:
    case IPROP_NETCLIENT_RECVLINE:
    {
     DEBUGNETC((LPSTR)"GET: RecvLine - Start");

     *(*(LPHSZ *)&lParam) = getRecvLine(hCtl);

     DEBUGNETC((LPSTR)"GET: RecvLine - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_LINE/IPROP_NETCLIENT_RECVLINE


    /****************** GET: Block **********************/
    /****************** GET: RecvBlock ******************/
    // Receive a block
    case IPROP_NETCLIENT_BLOCK:
    case IPROP_NETCLIENT_RECVBLOCK:
    {
     HLSTR hlStr;

     DEBUGNETC((LPSTR)"GET: RecvBlock - Start");

     hlStr = getRecvData(hCtl);
     if(hlStr==NULL)
     {
      hlStr=VBCreateHlstr(NULL, 0);
     }
     *(*(LPHLSTR *)&lParam) = hlStr;

     DEBUGNETC((LPSTR)"GET: RecvBlock - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_BLOCK/IPROP_NETCLIENT_RECVBLOCK


    /****************** GET: SendLine *******************/
    // Can't get!!!
    case IPROP_NETCLIENT_SENDLINE:
        return(netReturnError(VBERR_WRITEONLY, 0, (LPSTR)"SendLine"));


    /****************** GET: SendBlock ******************/
    // Can't get!!!
    case IPROP_NETCLIENT_SENDBLOCK:
        return(netReturnError(VBERR_WRITEONLY, 0, (LPSTR)"SendBlock"));


    /****************** GET: LocalService ***************/
    case IPROP_NETCLIENT_LOCALSERVICE:
    {
     HSZ hszTemp;
     LPSERVENT lpseServEnt;
     BOOL bClear = FALSE;

     DEBUGNETC((LPSTR)"GET: LocalService - Begin");

     lpseServEnt = getservbyport( htons(lpNetClient->sLocalPort), NULL);
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
      NetClientError(hCtl, WSAGetLastError());
     }

     if(bClear)
     {
      char cChar = '\0';

      hszTemp = VBCreateHsz((_segment)hCtl, (LPSTR)&cChar);
     }

     *(*(LPHSZ *)&lParam) = hszTemp;

     DEBUGNETC((LPSTR)"GET: LocalService - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_LOCALSERVICE


    /****************** GET: RemoteService **************/
    case IPROP_NETCLIENT_REMOTESERVICE:
    {
     HSZ hszTemp;
     LPSERVENT lpseServEnt;
     BOOL bClear = FALSE;

     DEBUGNETC((LPSTR)"GET: RemoteService - Begin");
     lpseServEnt = getservbyport( htons(lpNetClient->sRemotePort), NULL);
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
      NetClientError(hCtl, WSAGetLastError());
     }

     if(bClear)
     {
      char cChar = '\0';

      hszTemp = VBCreateHsz((_segment)hCtl, (LPSTR)&cChar);
     }

     *(*(LPHSZ *)&lParam) = hszTemp;

     DEBUGNETC((LPSTR)"GET: RemoteService - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_REMOTESERVICE


    /****************** GET: RecvCount ******************/
    // Calculate the current RecvBuffer size
    case IPROP_NETCLIENT_RECVCOUNT:
    {
     DEBUGNETC((LPSTR)"GET: RecvCount - Start");

      *(*(SHORT FAR **)&lParam)=lpNetClient->sRecvCount;

     DEBUGNETC((LPSTR)"GET: RecvCount - End");
     return(ERR_None);
    } // IPROP_NETCLINET_RECVCOUNT


    /****************** GET: SendCount ******************/
    // Calculate the current SendBuffer size
    case IPROP_NETCLIENT_SENDCOUNT:
    {
      DEBUGNETC((LPSTR)"GET: SendCount - Start");

      *(*(SHORT FAR **)&lParam)=lpNetClient->sSendCount;

      DEBUGNETC((LPSTR)"GET: SendCount - End");
      return(ERR_None);
    } // IPROP_NETCLINET_RECVCOUNT


    /****************** GET: ErrorMessage ***************/
    case IPROP_NETCLIENT_ERRORMESSAGE:
    {
     char szString[256];
     HSZ  hszTemp;
     char cChar;
     int iError;

     DEBUGNETC((LPSTR)"GET: ErrorMessage - Begin");

     iError=lpNetClient->sErrorNumber;
     if(!LoadString(hModDLL, iError, (LPSTR)szString,
                sizeof(szString)))
     {
      wsprintf((LPSTR)szString, (LPSTR)"Unknown Error #%d",
               iError);
     }

#ifdef DESIGN_TIME
     // MessageBeep(MB_ICONEXCLAMATION);
     DEBUGNETC((LPSTR)szString);
#endif

     // Set the error message text

     cChar='\0';
     hszTemp=VBCreateHsz((_segment)hCtl, (LPSTR)szString);
     DEBUGNETC((LPSTR)"GET: ErrorMessage - Created new error string");
     if(!hszTemp) hszTemp=VBCreateHsz((_segment)hCtl, (LPSTR)&cChar);

     *(*(LPHSZ *)&lParam) = hszTemp;

     DEBUGNETC((LPSTR)"GET: ErrorMessage - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_ERRORMESSAGE


    /****************** GET: Version ********************/
    // Get NetClient version string...
    case IPROP_NETCLIENT_VERSION:
    {
     HSZ hszTemp;
     char cChar = '\0';

     DEBUGNETC((LPSTR)"GET: Version - Begin");

     hszTemp=VBCreateHsz((_segment)hCtl, (LPSTR)WSANET_VERSION_TEXT);
     if(!hszTemp) hszTemp=VBCreateHsz((_segment)hCtl, (LPSTR)&cChar);

     *(*(LPHSZ *)&lParam) = hszTemp;
     DEBUGNETC((LPSTR)WSANET_VERSION_TEXT);

     DEBUGNETC((LPSTR)"GET: Version - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_VERSION


    /****************** GET: Debug **********************/
    // Get SO_DEBUG
    case IPROP_NETCLIENT_DEBUG:
    {
     BOOL fDebug;
     int iSize = sizeof(fDebug);

     DEBUGNETC((LPSTR)"GET: Debug - Begin");

     if((lpNetClient->sSocket != 0) &&
        (lpNetClient->sSocket != SOCKET_ERROR))
     {
      if(getsockopt(lpNetClient->sSocket, SOL_SOCKET, SO_DEBUG,
                    (LPSTR)&fDebug, (int FAR *)&iSize)
           != SOCKET_ERROR)
      {
       *(*(SHORT FAR **)&lParam)=(SHORT)fDebug;
       DEBUGNETC((LPSTR)"GET: Debug - End - SO_DEBUG returned");
       return(ERR_None);
      }
     }
     *(*(SHORT FAR **)&lParam)=0;
     DEBUGNETC((LPSTR)"GET: Debug - End - False returned");
     return(ERR_None);
    } // IPROP_NETCLIENT_DEBUG


    /****************** GET: Host *************************/
    // Get HostName.
    case IPROP_NETCLIENT_HOST:
    /****************** GET: HostName *******************/
    // Get HostName from its MOLE
    case IPROP_NETCLIENT_HOSTNAME:
    {
     HSZ hszTemp;

     DEBUGNETC((LPSTR)"GET: HostName - Begin");

     hszTemp = GetMoleHsz(hCtl,lpNetClient->mHostName);
     if(!hszTemp)
     {
      DEBUGNETC((LPSTR)"GET: HostName - End - Critical error!");
      return(netReturnError(VBERR_OUTOFMEMORY, 0, NULL));
     }

     *(*(LPHSZ *)&lParam) = hszTemp;

     DEBUGNETC((LPSTR)"GET: HostName - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_HOSTNAME


    /****************** GET: HostAddr *******************/
    // Get HostAddr from its MOLE
    case IPROP_NETCLIENT_HOSTADDR:
    {
     HSZ hszTemp;

     DEBUGNETC((LPSTR)"GET: HostAddr - Begin");

     hszTemp = GetMoleHsz(hCtl,lpNetClient->mHostAddr);
     if(!hszTemp)
     {
      DEBUGNETC((LPSTR)"GET: HostAddr - End - Critical error!");
      return(netReturnError(VBERR_OUTOFMEMORY, 0, NULL));
     }

     *(*(LPHSZ *)&lParam) = hszTemp;

     DEBUGNETC((LPSTR)"GET: HostAddr - End");
     return(ERR_None);
    } // IPROP_NETCLIENT_HOSTADDR


    /****************** GET: HostAliasCount *************/
    // Get the current number of elements in
    //  HostAliasList
    case IPROP_NETCLIENT_HOSTALIASCOUNT:
    {
      DEBUGNETC((LPSTR)"GET: HostAliasCount - Start");

      *(*(SHORT FAR **)&lParam)=lpNetClient->htHostAliasList.iHostCount;

      DEBUGNETC((LPSTR)"GET: HostAliasCount - End");
      return(ERR_None);
    } // IPROP_NETCLINET_HOSTALIASCOUNT


    /****************** GET: HostAddressCount ***********/
    // Get the current number of elements in
    //  HostAddressList
    case IPROP_NETCLIENT_HOSTADDRESSCNT:
    {
      DEBUGNETC((LPSTR)"GET: HostAddressCount - Start");

      *(*(SHORT FAR **)&lParam)=lpNetClient->htHostAddressList.iHostCount;

      DEBUGNETC((LPSTR)"GET: HostAddressCount - End");
      return(ERR_None);
    } // IPROP_NETCLIENT_HOSTADDRESSCNT


    /****************** GET: HostAliasList() ************/
    // Get a element of HostAliasList
    case IPROP_NETCLIENT_HOSTALIASLIST:
    {
      LONG lIndex;
      MOLE mTemp;
      HSZ hszTemp;
      LPDATASTRUCT lpdsElement = (LPDATASTRUCT)lParam;

      DEBUGNETC((LPSTR)"GET: HostAliasList - Start");

      lIndex = lpdsElement->index[0].data;
      if((lIndex < 0) ||
         (lIndex > lpNetClient->htHostAliasList.iHostCount))
      {
       return(netReturnError(VBERR_BADVBINDEX, 0, NULL));
      }

      mTemp = lpNetClient->htHostAliasList.mHost[lIndex];

      hszTemp = GetMoleHsz(hCtl,mTemp);

      if(!hszTemp) return(netReturnError(VBERR_OUTOFMEMORY, 0, NULL));

      lpdsElement->data=(LONG)hszTemp;

      DEBUGNETC((LPSTR)"GET: HostAliasList - End");
      return(ERR_None);
    } // IPROP_NETCLINET_HOSTALIASLIST


    /****************** GET: HostAddressList() ************/
    // Get a element of HostAddressList
    case IPROP_NETCLIENT_HOSTADDRESSLIST:
    {
      LONG lIndex;
      MOLE mTemp;
      HSZ hszTemp;
      LPDATASTRUCT lpdsElement = (LPDATASTRUCT)lParam;

      DEBUGNETC((LPSTR)"GET: HostAddressList - Start");

      lIndex = lpdsElement->index[0].data;
      if((lIndex < 0) ||
         (lIndex > lpNetClient->htHostAddressList.iHostCount))
      {
       return(netReturnError(VBERR_BADVBINDEX, 0, NULL));
      }

      mTemp = lpNetClient->htHostAddressList.mHost[lIndex];

      hszTemp = GetMoleHsz(hCtl,mTemp);

      if(!hszTemp) return(netReturnError(VBERR_OUTOFMEMORY, 0, NULL));

      lpdsElement->data=(LONG)hszTemp;

      DEBUGNETC((LPSTR)"GET: HostAddressList - End");
      return(ERR_None);
    } // IPROP_NETCLIENT_HOSTALIASLIST


    /****************** GET: DEFAULT handler! ***********/
    default:
    {
     DEBUGNETC((LPSTR)"GET: Property not handled");
    }
    break;
   }
   DEBUGNETC((LPSTR)"VBM_GETPROPERTY: End");
  }
  break;


/**********************************************************************
 ***                        GET PROPERTY HSZ                        ***
 **********************************************************************/
#ifdef DESIGN_TIME
    // Set the About propery to display "About NetClient"
  case VBM_GETPROPERTYHSZ:
  {
   DEBUGNETC((LPSTR)"VBM_GETPROPERTYHSZ Begin");
   switch(wParam)
   {
    case IPROP_NETCLIENT_ABOUT:
    {
     DEBUGNETC((LPSTR)"GETPROPERTYHSZ: About");
     *(*(LPHSZ *)&lParam) = VBCreateHsz((_segment) hCtl, (LPSTR)STRING_ABOUTTEXT);
     return(ERR_None);
    }

    case IPROP_NETCLIENT_SENDBLOCK:
    case IPROP_NETCLIENT_SENDLINE:
    {
     DEBUGNETC((LPSTR)"GETPROPERTYHSZ: Send(Block/Line)");
     *(*(LPHSZ *)&lParam) = VBCreateHsz((_segment) hCtl, (LPSTR)STRING_WRITEONLY);
     return(ERR_None);
    }
   }
   DEBUGNETC((LPSTR)"VBM_GETPROPERTYHSZ End");
  }
  break;


/**********************************************************************
 ***                   CONTEXT SENSITIVE HELP                       ***
 **********************************************************************/
  case VBM_HELP:
  {
#ifdef DEBUG_BUILD
   char szError[128];

   DEBUGNETC((LPSTR)"VBM_HELP: Begin");
   wsprintf((LPSTR)szError, (LPSTR)"VBM_HELP: Prop. #%d",
             wParam );
   DEBUGNETC((LPSTR)szError);
#endif // DEBUG_BULD

   if(bHelpStdPropEvt(wParam, (LPMODEL)lParam))
        break;
   else return(DisplayHelpTopic(TOPIC_CONTROL_NETCLIENT, wParam,
                                (LPMODEL)lParam));
   DEBUGNETC((LPSTR)"VBM_HELP: Standard handler allowed");
  }
  break;
#endif // DESIGN_TIME


    // **** The network activity message ****
  case CM_NETACTIVITY:
  {
   DEBUGNETC((LPSTR)"CM_NETACTIVITY");

   return(netClientAsyncProc(hCtl, hWnd, wParam, lParam));
  }

    // TimeOut handler.
  case WM_TIMER:
  {
   DEBUGNETC((LPSTR)"WM_TIMER: Begin");
   if(lpNetClient->sTimeOut) (lpNetClient->sTimeOut)--;
   if(lpNetClient->sTimeOut <= 0)
   {
    lpNetClient->sTimeOut=0;
    KillTimer(hWnd, wParam);
    clientFireOnTimeOut(hCtl);
   }
   DEBUGNETC((LPSTR)"WM_TIMER: End");
   return(ERR_None);
  }
  break;

 }
 return(VBDefControlProc(hCtl, hWnd, (USHORT)Msg, (USHORT)wParam,
        (LONG)lParam));
} /* NetClientCtlProc() */


/* LONG CALLBACK _export AboutWndProc(HWND, UINT, WPARAM, LPARAM);
    Purpose: To display a dialog for a popup window
*/
LONG CALLBACK _export AboutWndProc(HWND hWnd, UINT Msg,
                                   WPARAM wParam, LPARAM lParam)
{
 switch(Msg)
 {
  case WM_DESTROY:
  {
   bDialogInUse=FALSE;
  }
  break;

  case WM_SHOWWINDOW:
  {
   if (wParam)
   {
    PostMessage(hWnd, CM_OPENABOUTDLG, 0, 0L);
    return(0L);
   }
  }
  break;

  case CM_OPENABOUTDLG:
  {
   VBDialogBoxParam(hModDLL, "AboutDlg",
                    (FARPROC)AboutDlgProc , 0L);
   return(0L);
  }
  break;
 }
 return(DefWindowProc(hWnd, Msg, wParam, lParam));
} /* AboutWndProc() */


/* BOOL CALLBACK _export AboutDlgProc(HWND, UINT, WPARAM, LPARAM);
    Purpose: To handle the About dialog
*/
BOOL CALLBACK _export AboutDlgProc(HWND hWndDlg, UINT Msg,
                                   WPARAM wParam, LPARAM lParam)
{
#ifdef USE_VCPP
 lParam = lParam;
#endif /* USE_VCPP */
 switch(Msg)
 {
  case WM_INITDIALOG:
  {
   return(TRUE);
  }
  break;

  case WM_COMMAND:
  {
   switch(wParam)
   {
    case IDOK:
    {
     EndDialog(hWndDlg, TRUE);
     return(TRUE);
    }
    break;
   }
  }
  break;
 }
 return(FALSE);
} /* AboutDlgProc() */


/* ERR NetClientError(HCTL hCtl, int);
    Purpose: To set a VB error level and string to any errors
*/
ERR NetClientError(HCTL hCtl, int iError)
{
 char szString[256];
 LPNETCLIENT lpNetClient;

 wsprintf((LPSTR)szString, (LPSTR)"NetClientError(%d) - Begin", iError);
 DEBUGNETC((LPSTR)szString);

 if(!iError) return(ERR_None);

 if(iError == WSANO_DATA) return(ERR_None);

 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);
 lpNetClient->sErrorNumber=(SHORT)iError;

 DEBUGNETC((LPSTR)"NetClientError() - Fire OnError()");

 return(clientFireOnError(hCtl));
} /* NetClientError() */


/* ERR clientFireOnConnect(HCTL);
    Purpose: To fire the OnRecv event.
    Given:   The control handle
    Returns: Any error condition
*/
ERR clientFireOnConnect(HCTL hCtl)
{
 DEBUGNETC((LPSTR)"clientFireOnConnect()");
 return(VBFireEvent(hCtl, IEVENT_NETCLIENT_ONCONNECT, NULL));
} /* clientFireOnConnect() */


/* ERR clientFireOnRecv(HCTL);
    Purpose: To fire the OnRecv event.
    Given:   The control handle
    Returns: Any error condition
*/
ERR clientFireOnRecv(HCTL hCtl)
{
 DEBUGNETC((LPSTR)"clientFireOnRecv()");
 return(VBFireEvent(hCtl, IEVENT_NETCLIENT_ONRECV, NULL));
} /* clientFireOnRecv() */


/*ERR clientFireOnSend(HCTL);
    Purpose: To fire the OnSend event.
    Given:   The control handle
    Returns: Any error condition
*/
ERR clientFireOnSend(HCTL hCtl)
{
 DEBUGNETC((LPSTR)"clientFireOnSend()");
 return(VBFireEvent(hCtl, IEVENT_NETCLIENT_ONSEND, NULL));
} /* clientFireOnSend() */


/* ERR clientFireOnClose(HCTL);
    Purpose: To fire the OnClose event.
    Given:   The control handle
    Returns: Any error condition
*/
ERR clientFireOnClose(HCTL hCtl)
{
 LPNETCLIENT lpNetClient;
 ERR err;

    // Dereference the control so we can look inside
 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);

    // Make sure our app has the chance to handle ALL data
 if(lpNetClient->sRecvCount)
 {
  DEBUGNETC((LPSTR)"clientFireOnClose() - calling clientFireOnRecv()");
  err=clientFireOnRecv(hCtl);
  if(err) return(err);
 }
 return(VBFireEvent(hCtl, IEVENT_NETCLIENT_ONCLOSE, NULL));
} /* clientFireOnClose() */


/* ERR clientFireOnTimeOut(HCTL);
    Purpose: To fire the OnTimeOut event.
    Given:   The control handle
    Returns: Any error condition
*/
ERR clientFireOnTimeOut(HCTL hCtl)
{
 DEBUGNETC((LPSTR)"clientFireOnTimeOut()");
 return(VBFireEvent(hCtl, IEVENT_NETCLIENT_ONTIMEOUT, NULL));
} /* clientFireOnTimeOut() */


/* ERR clientFireOnError(HCTL);
    Purpose: To fire the OnError event.
    Given:   The control handle
    Returns: Any error condition.
*/
ERR clientFireOnError(HCTL hCtl)
{
 ONERRORPARMS Parameters;
 LPNETCLIENT lpNetClient;
 SHORT       sErrorNumber;

    // Dereference the control so we can look inside
    // Becomes invalid after the first VBLockHsz() below!
 DEBUGNETC((LPSTR)"clientFireOnError() - Begin");
 lpNetClient=(LPNETCLIENT)VBDerefControl(hCtl);

 sErrorNumber=lpNetClient->sErrorNumber;

    // Set the error number
 Parameters.ErrorNumber=(SHORT FAR *)&sErrorNumber;

 return(VBFireEvent(hCtl, IEVENT_NETCLIENT_ONERROR, (LPVOID)&Parameters));
} /* clientFireOnError() */

/* End of NetClnt.C */


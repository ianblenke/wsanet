/* WSANet.H */
#ifndef WSANet_H
#define WSANet_H

#include <Windows.h>
#include <VBAPI.H>      // You need the CDK to compile WSNetC.VBX
#include <WinSock.h>    // WinSock.h is included with this distribution
#include "Topics.h"
#include "Version.h"
#include "Errors.h"

    /* Conditional compilation #defines */
#define DESIGN_TIME     // Undefine this to take out all design time code
#define FULL_DEBUG      // Undefine this to take all of the debugging messages

#define DEBUG_BUILD     // Undefine this for the debugging version

#ifdef DEBUG_BUILD
#else
#define DEBUGIT( param )
#endif /* DEBUG_BUILD */

//#define VB100_CDK // define this if you have the VB 1.00 CDK
//#define VB200_CDK // define this if you have the VB 2.00 CDK
#define VB300_CDK   // define this if you have the VB 3.00 CDK


//#define USE_VCPP  // define this if you are using VC++ and want to avoid
                    // it's overactive warnings. ;)

/* Each one has the features of the previous */
#ifdef VB300_CDK
#define VB200_CDK
#define VB100_CDK
#else
#ifdef VB200_CDK
#define VB100_CDK
#else
#ifndef VB100_CDK
#define VB100_CDK
#endif /* VB100_CDK */
#endif /* VB200_CDK */
#endif /* VB300_CDK */


    /* Common #define(s) */

#define CM_NETACTIVITY   WM_USER+1

#ifdef DESIGN_TIME
#define CM_OPENABOUTDLG  WM_USER+2

#define CLASS_ABOUTPOPUP "AboutWSetClass"
#define CONTROLNAME      "WSANet VBX"
#define WSANET_HELPFILE  "wsanet.hlp"

#define STRING_ABOUTTEXT "About WSANet.VBX"
#define STRING_WRITEONLY "*Write only*"
#define STRING_WRONGTITL "WSANET.VBX Run-time problem"

#endif // DESIGN_TIME


    /* PROTOTYPEs */
typedef HLSTR MOLE;
typedef HLSTR FAR *LPMOLE;

#ifdef __cplusplus
extern "C" {            /* Assume C declarations for C++ */
#endif  /* __cplusplus */

#ifdef DEBUG_BUILD
VOID DEBUGIT(LPSTR);
#endif /* DEBUG_BUILD */

// Functions in WSANet.C
BOOL  DisplayHelpTopic(WORD, WPARAM, LPMODEL);
BOOL  bHelpStdPropEvt(WPARAM, LPMODEL);

// Functions in NetSrvr.C
LONG  CALLBACK _export NetServerCtlProc(HCTL, HWND, UINT, WPARAM, LPARAM);
ERR   NetServerError(HCTL, int);
ERR   serverFireOnError(HCTL);

// Functions in NetClnt.C
LONG  CALLBACK _export NetClientCtlProc(HCTL, HWND, UINT, WPARAM, LPARAM);
int   CALLBACK _export NetClientSendBlock(HCTL, HLSTR);
int   CALLBACK _export NetClientRecvBlock(HCTL, HLSTR, int);
BOOL  CALLBACK _export AboutDlgProc(HWND, UINT, WPARAM, LPARAM);
LONG  CALLBACK _export AboutWndProc(HWND, UINT, WPARAM, LPARAM);
ERR   NetClientError(HCTL, int);
ERR   clientFireOnConnect(HCTL);
ERR   clientFireOnRecv(HCTL);
ERR   clientFireOnSend(HCTL);
ERR   clientFireOnClose(HCTL);
ERR   clientFireOnTimeOut(HCTL);
ERR   clientFireOnError(HCTL);

// Functions in Net.C
BOOL  netStartup(VOID);
VOID  netCleanup(VOID);
BOOL  netCloseSocket(HWND, SOCKET);
BOOL  netConnect(HCTL, SOCKET);
BOOL  netAccept(HCTL, SOCKET);
LONG  netClientAsyncProc(HCTL, HWND, WPARAM, LPARAM);
LONG  netServerAsyncProc(HCTL, HWND, WPARAM, LPARAM);
ERR   netSendData(HCTL);
int   netRecvData(HCTL);
int   putSendData(HCTL, MOLE);
HLSTR getRecvData(HCTL);
HSZ   getRecvLine(HCTL);
ERR   netReturnError(ERR, int, LPSTR);
HSZ   GetMoleHsz(HCTL, MOLE);
VOID  GetMoleName(MOLE, LPSTR, int);
VOID  SetMole(LPMOLE, LPSTR);
MOLE  AddMole(LPSTR);
MOLE  CreateMole(LPSTR, USHORT);
VOID  DeleteMole(MOLE);

// Functions in Ini.C
LONG  CALLBACK _export IniCtlProc(HCTL, HWND, UINT, WPARAM, LPARAM);


// Functions in WSANet.C
int   CALLBACK LibMain(HANDLE, WORD, WORD, LPSTR);
BOOL  CALLBACK _export VBINITCC(USHORT, BOOL);
VOID  CALLBACK _export VBTERMCC(VOID);

#ifdef __cplusplus
 }           /* Assume C declarations for C++ */
#endif  /* __cplusplus */

    /* Type definitions */
#ifndef LPHSZ            // Define LPHSZ for FAR HSZ *
typedef HSZ FAR * LPHSZ; // Why doesn't VBAPI.H have this
//typedef HLSTR FAR * LPHLSTR;
#endif // LPHSZ

    /* Globals */
#ifdef NetClnt_C
HANDLE hModDLL;
BOOL   bDialogInUse = FALSE;       // Don't pop up more than 1 About
HWND   hDialogHwnd  = 0;           // Invisible popup window handle
#else
extern HANDLE hModDLL;
extern BOOL   bDialogInUse;
extern HWND   hDialogHwnd;
#endif

#endif /* WSANet_H */

/* End of WSANet.H */

/* WSANet.C - Initialization for WSANET.VBX */

#define WSANet_C

#include "WSANet.H"
#include "Version.H"
#include "NetClnt.H"    // NetClient control Model
#include "NetSrvr.H"    // NetServer control Model
#include "Ini.H"        // Ini control Model

static const char szEmbedded[] = "WSANET.VBX (c)1993 Ian C. Blenke";

// Specific globals to INIT.C
static WORD wVbxUsers      = 0;
static BOOL bDevTimeInited = FALSE;

/* BOOL bHelpStdPropEvt(WPARAM, LPMODEL);
    Purpose: To return TRUE if the property is standard, and
             FALSE if it is custom
*/
BOOL bHelpStdPropEvt(WPARAM wParam, LPMODEL lpModel)
{
 switch(LOBYTE(wParam))
 {
  case VBHELP_PROP:
    if(lpModel->npproplist[HIBYTE(wParam)] >= PPROPINFO_STD_LAST)
        return(TRUE);
  break;

  case VBHELP_EVT:
    if(lpModel->npeventlist[HIBYTE(wParam)] >= PEVENTINFO_STD_LAST)
        return(TRUE);
  break;
 }
 return(FALSE);
} /* bHelpStdPropEvt() */


/* BOOL DisplayHelpTopic(WORD, WPARAM, LPMODEL);
    Purpose: To display context sensitive help
*/
BOOL DisplayHelpTopic(WORD wControl, WPARAM wParam, LPMODEL lpModel)
{
 WORD *wHelpProperties = NULL;
 WORD *wHelpEvents = NULL;

#ifdef USE_VCPP
 lpModel = lpModel;
#endif /* USE_VCPP */

 switch(wControl)
 {
  case TOPIC_CONTROL_NETCLIENT:
    wHelpProperties = wNetClientHelpProps;
    wHelpEvents = wNetClientHelpEvents;
  break;

  case TOPIC_CONTROL_INI:
    wHelpProperties = wIniHelpProps;
    wHelpEvents = NULL;
  break;

  case TOPIC_CONTROL_NETSERVER:
    wHelpProperties = wNetServerHelpProps;
    wHelpEvents = wNetServerHelpEvents;
  break;

  default:
    return(ERR_None);
 }

 switch(LOBYTE(wParam))
 {
  case VBHELP_PROP:
     if(wHelpProperties)
     {
      if(WinHelp(NULL, WSANET_HELPFILE, HELP_CONTEXT, (DWORD)(wHelpProperties[HIBYTE(wParam)])))
         return(ERR_None);
     }
  break;

  case VBHELP_EVT:
     if(wHelpEvents)
     {
      if(WinHelp(NULL, WSANET_HELPFILE, HELP_CONTEXT, (DWORD)(wHelpEvents[HIBYTE(wParam)])))
         return(ERR_None);
     }
  break;

  case VBHELP_CTL:
     if(WinHelp(NULL, WSANET_HELPFILE, HELP_CONTEXT, (DWORD)wControl))
        return(ERR_None);
 }
 return(ERR_None);
} // DisplayHelpTopic()


    // List of all the controls this VBX supports
LPMODEL modellistControls[] =
{
 NULL,                              // This is changed in code
 MODELLIST_INI,
 MODELLIST_NETSERVER,
 NULL
};

    // Model information structure
MODELINFO modelinfoControls =
{
 VB100_VERSION,                     // This is changed in code
 (LPMODEL FAR *)&modellistControls
};


/* int CALLBACK LibMain(HANDLE, WORD, WORD, LPSTR);
    Purpose: Handles the startup of the VBX (DLL)
    Note:    Quite unstable in here.
*/
int CALLBACK LibMain(HANDLE hModule, WORD wDataSeg, WORD cbHeapSize,
                     LPSTR lpszCmdLine)
{
 #ifdef USE_VCPP
  lpszCmdLine=lplszCmdLine;
  cbHeapSize=cbHeapSize;
  wDataSeg=wDataSeg;
 #endif /* USE_VCPP */

 // Remember the Module (Instance) of the VBX
 hModDLL  = hModule;

 return(1);
} /* LibMain() */


/* BOOL CALLBACK _export VBINITCC(USHORT, BOOL);
    Purpose: To handle the Visual Basic entrance point for the VBX
*/
BOOL CALLBACK _export VBINITCC(USHORT usVersion, BOOL bRuntime)
{
    // Keep count of the number of users
 ++wVbxUsers;

#ifdef DESIGN_TIME
    // Register popup class if this is from the development environment.
 if (!bRuntime && !bDevTimeInited)
 {
  WNDCLASS class;

    // Show the beginning debug banner
#ifdef DEBUG_BUILD
  OutputDebugString((LPSTR)WSANET_BANNER);
#endif /* DEBUG_BUILD */

  class.style         = 0; // (Invisible)
  class.lpfnWndProc   = (WNDPROC)AboutWndProc;
  class.cbClsExtra    = 0;
  class.cbWndExtra    = 0;
  class.hInstance     = hModDLL;
  class.hIcon         = NULL;
  class.hCursor       = NULL;
  class.hbrBackground = NULL;
  class.lpszMenuName  = NULL;
  class.lpszClassName = CLASS_ABOUTPOPUP;

  if (!RegisterClass(&class))
      return FALSE;

    // We successfully initialized the stuff we need at dev time
  bDevTimeInited = TRUE;
 }
#else
 if(!bRunTime) return(FALSE);
#endif //DESIGN_TIME


    // Hack to allow context sensitive help in VB2.0 and newer
    // in the single model NetServer and INI controls.
 if(usVersion>=VB200_VERSION)
 {
  modelNetServer.usVersion = VB200_VERSION;
  modelIni.usVersion = VB200_VERSION;
 }
 else
 {
  modelNetServer.usVersion = VB100_VERSION;
  modelIni.usVersion = VB100_VERSION;
 }

    // Initialize the network
 if(!netStartup())
 {
  char szTemp[256];

  if(LoadString(hModDLL, IDS_WRONGVER, (LPSTR)szTemp, sizeof(szTemp)))
  {
   MessageBox(NULL, (LPSTR)szTemp,
             (LPSTR)STRING_WRONGTITL, MB_ICONSTOP | MB_OK);
  }
  return(FALSE);
 }

 if(!VBRegisterModel(hModDLL, &modelIni))
 {
  OutputDebugString("Failure registering Ini control!\n");
  return(FALSE);
 }

 if(!VBRegisterModel(hModDLL, &modelNetServer))
 {
  OutputDebugString("Failure registering NetServer control!\n");
  return(FALSE);
 }

#ifdef VB100_CDK
 if(usVersion==VB100_VERSION) return(VBRegisterModel(hModDLL, &modelNetClient_VB1));
#ifdef VB200_CDK
 if(usVersion==VB200_VERSION) return(VBRegisterModel(hModDLL, &modelNetClient_VB2));
#ifdef VB300_CDK
 if(usVersion>=VB300_VERSION) return(VBRegisterModel(hModDLL, &modelNetClient_VB3));
#endif /* VB300_CDK */
#endif /* VB200_CDK */
#endif /* VB100_CDK */
} /* VBINITCC() */


/* LPMODELINFO CALLBACK VBGetModelInfo(USHORT usVersion);
    Purpose: To export the model information without actually
             registering it (Read VB3.0 Professional book
             #1, page 274)
*/
LPMODELINFO FAR PASCAL _export VBGetModelInfo(USHORT usVersion)
{
#ifdef VB100_CDK
 if(usVersion==VB100_VERSION)
 {
  modellistControls[0] = (LPMODEL)&modelNetClient_VB1;
  modelinfoControls.usVersion = usVersion;
  return(&modelinfoControls);
 }
#ifdef VB200_CDK
 if(usVersion==VB200_VERSION)
 {
  modellistControls[0] = (LPMODEL)&modelNetClient_VB2;
  modelinfoControls.usVersion = usVersion;
  return(&modelinfoControls);
 }
#ifdef VB300_CDK
 if(usVersion>=VB300_VERSION)
 {
  modellistControls[0] = (LPMODEL)&modelNetClient_VB3;
  modelinfoControls.usVersion = VB300_VERSION;
  return(&modelinfoControls);
 }
#endif /* VB300_CDK */
#endif /* VB200_CDK */
#endif /* VB100_CDK */
}/* VBGetModelInfo() */


/* VOID CALLBACK _export VBTERMCC(VOID);
    Purpose: To handle the Visual Basic exit point from the VBX
*/
VOID CALLBACK _export VBTERMCC(VOID)
{
 --wVbxUsers;

#ifdef DESIGN_TIME
 if (!wVbxUsers & bDevTimeInited)
 {
  // Free any resources created for Dev environment
  UnregisterClass(CLASS_ABOUTPOPUP, hModDLL);
#ifdef VB200_CDK
  WinHelp(NULL, WSANET_HELPFILE, HELP_QUIT, NULL);
#endif /* VB200_CDK */
 }
#endif //DESIGN_TIME

 netCleanup();
 return;
} /* VBTERMCC() */

/* End of WSANet.C */

/* Ini.C - Ini Control */

#define INI_C
#include "WSANet.H"
#include "Ini.H"

#ifdef DEBUG_BUILD
static char szErrors[256];

VOID DEBUGINI(LPSTR lpString)
{
 wsprintf((LPSTR)szErrors, (LPSTR)"Ini: %s\n\r", lpString);
 OutputDebugString( (LPSTR)szErrors );
}
#else
#define DEBUGINI( parm )
#endif /* DEBUG_BUILD */

    // Static variables (to save the poor application's stack)
// Only used when you Get or Set Value
static char szSection[MAX_SECTION_SIZE];
static char szEntry[MAX_ENTRY_SIZE];
static char szDefault[MAX_DEFAULT_SIZE];
static char szFilename[MAX_FILENAME_SIZE];
static char szValue[MAX_VALUE_SIZE];

/* LONG CALLBACK _export IniCtlProc(HCTL, HWND, UINT, WPARAM, LPARAM);
    Purpose: To handle the VBX messages
*/
LONG CALLBACK _export IniCtlProc(HCTL hCtl, HWND hWnd, UINT Msg,
                                       WPARAM wParam, LPARAM lParam)
{
 LPINI lpIni;

 lpIni=(LPINI)VBDerefControl(hCtl);

 switch(Msg)
 {

#ifdef DESIGN_TIME
    // The MODEL want's to know our default size
  case VBM_GETDEFSIZE:
  {
   BITMAP bmp;
   HBITMAP hBitmap;

   DEBUGINI("VBM_GETDEFSIZE[2.0] Begin\n\r");
   hBitmap=LoadBitmap(hModDLL, MAKEINTRESOURCE(IDBMP_INI));
   if(hBitmap)
   {
    GetObject(hBitmap, sizeof(BITMAP), (LPSTR)&bmp);
    DeleteObject(hBitmap);
    DEBUGINI((LPSTR)"VBM_GETDEFSIZE[2.0] End\n\r");
    return(MAKELONG(bmp.bmWidth, bmp.bmHeight));
   }
   DEBUGINI((LPSTR)"VBM_GETDEFSIZE[2.0] Allowing DefVBProc to handle\n\r");
  }
  break;

    // The re-size message - only in design time
  case WM_SIZE:
  {
   BITMAP bmp;
   HBITMAP hBitmap;
   RECT rect;
   POINT pos;

   DEBUGINI((LPSTR)"WM_SIZE Start");
   GetWindowRect(hWnd, (RECT FAR *)&rect);
   pos.x=rect.left;
   pos.y=rect.top;
   ScreenToClient((HWND)GetWindowWord(hWnd, GWW_HWNDPARENT),
                  (POINT FAR *)&pos);

   hBitmap=LoadBitmap(hModDLL, MAKEINTRESOURCE(IDBMP_INI));
   if(hBitmap)
   {
    GetObject(hBitmap, sizeof(BITMAP), (LPSTR)&bmp);
    MoveWindow(hWnd, pos.x, pos.y, bmp.bmWidth, bmp.bmHeight,
               SWP_NOZORDER | SWP_NOMOVE);
    DeleteObject(hBitmap);
    DEBUGINI((LPSTR)"WM_SIZE End");
    return(0);
   }
   DEBUGINI((LPSTR)"WM_SIZE End - Bitmap load failure!");
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

   DEBUGINI((LPSTR)"WM_PAINT Start");
   if(VBGetMode()==MODE_RUN) break;

   GetWindowRect(hWnd, (RECT FAR *)&rect);
   BeginPaint(hWnd, (PAINTSTRUCT FAR *)&ps);
   hBitmap=LoadBitmap(hModDLL, MAKEINTRESOURCE(IDBMP_INI));
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
   DEBUGINI((LPSTR)"WM_PAINT End");
  }
  break;
#endif // DESIGN_TIME


    // The control is closing - free everything.
  case WM_NCDESTROY:
  {
   MOLE mTemp;

   DEBUGINI((LPSTR)"WM_NCDESTROY Begin");

   if(lpIni->mSection)
   {
    mTemp = lpIni->mSection;
    DeleteMole(mTemp);
    lpIni=(LPINI)VBDerefControl(hCtl);
    lpIni->mSection = 0;
   }

   if(lpIni->mEntry)
   {
    mTemp = lpIni->mEntry;
    DeleteMole(mTemp);
    lpIni=(LPINI)VBDerefControl(hCtl);
    lpIni->mEntry = 0;
   }

   if(lpIni->mDefault)
   {
    mTemp = lpIni->mDefault;
    DeleteMole(mTemp);
    lpIni=(LPINI)VBDerefControl(hCtl);
    lpIni->mDefault = 0;
   }

   if(lpIni->mFilename)
   {
    mTemp = lpIni->mFilename;
    DeleteMole(mTemp);
    lpIni=(LPINI)VBDerefControl(hCtl);
    lpIni->mFilename = 0;
   }

   DEBUGINI((LPSTR)"WM_NCDESTROY End");
  }
  break; // WM_NCDESTROY


  case VBM_CREATED:
  {
   DEBUGINI((LPSTR)"VBM_CREATED Begin");

   lpIni->mSection = 0;
   lpIni->mEntry = 0;
   lpIni->mDefault = 0;
   lpIni->mFilename = 0;

   if(VBGetMode()==MODE_RUN)
   {
    ShowWindow(hWnd, SW_HIDE);
   }
   else
   {
    ShowWindow(hWnd, SW_SHOWNA);
   }

   DEBUGINI((LPSTR)"VBM_CREATED End");
   return(ERR_None);
  }     // VBM_CREATED


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
   DEBUGINI((LPSTR)szError);
#endif /* DEBUG_BUILD */

   switch(wParam)
   {

    /****************** SET: Section ********************/
    // User is setting the [Section] string
    case IPROP_INI_SECTION:
    {
     MOLE mTemp;

     DEBUGINI((LPSTR)"SET: Section - Begin");

     mTemp = lpIni->mSection;
     SetMole((LPMOLE)&(mTemp), (LPSTR)lParam);
     lpIni=(LPINI)VBDerefControl(hCtl);
     lpIni->mSection = mTemp;

     DEBUGINI((LPSTR)"SET: Section - End");
     return(ERR_None);
    }   // IPROP_INI_SECTION


    /****************** SET: Entry **********************/
    // User is setting the Entry= string
    case IPROP_INI_ENTRY:
    {
     MOLE mTemp;

     DEBUGINI((LPSTR)"SET: Entry - Begin");

     mTemp = lpIni->mEntry;
     SetMole((LPMOLE)&(mTemp), (LPSTR)lParam);
     lpIni=(LPINI)VBDerefControl(hCtl);
     lpIni->mEntry = mTemp;

     DEBUGINI((LPSTR)"SET: Entry - End");
     return(ERR_None);
    }   // IPROP_INI_ENTRY


    /****************** SET: Default ********************/
    // User is setting the Default string
    case IPROP_INI_DEFAULT:
    {
     MOLE mTemp;

     DEBUGINI((LPSTR)"SET: Default - Begin");

     mTemp = lpIni->mDefault;
     SetMole((LPMOLE)&(mTemp), (LPSTR)lParam);
     lpIni=(LPINI)VBDerefControl(hCtl);
     lpIni->mDefault = mTemp;

     DEBUGINI((LPSTR)"SET: Default - End");
     return(ERR_None);
    }   // IPROP_INI_DEFAULT


    /****************** SET: Filename *******************/
    // User is setting the Filename string
    case IPROP_INI_FILENAME:
    {
     MOLE mTemp;

     DEBUGINI((LPSTR)"SET: Filename - Begin");

     mTemp = lpIni->mFilename;
     SetMole((LPMOLE)&(mTemp), (LPSTR)lParam);
     lpIni=(LPINI)VBDerefControl(hCtl);
     lpIni->mFilename = mTemp;

     DEBUGINI((LPSTR)"SET: Filename - End");
     return(ERR_None);
    }   // IPROP_INI_FILENAME


    /****************** SET: Value **********************/
    // User is setting the Value string
    case IPROP_INI_VALUE:
    {
     LPSTR lpValue;

     DEBUGINI((LPSTR)"SET: Value - Begin");

     iniRebuildStrings(hCtl, lpIni);

     lpValue = (LPSTR)lParam;

     if(!WritePrivateProfileString((LPSTR)szSection, (LPSTR)szEntry,
                                  (LPSTR)lpValue, (LPSTR)szFilename))
     {
      DEBUGINI((LPSTR)"SET: Value - End - WritePrivateProfileString() returned FALSE!!");
      return(ERR_None);
     }

     DEBUGINI((LPSTR)"SET: Value - End");
     return(ERR_None);
    }   // IPROP_INI_VALUE

    default:
    {
     DEBUGINI((LPSTR)"VBM_SETPROPERTY Allowing default actions");
    }
   } // switch(wParam)
  } // VBM_SETPROPERTY

/**********************************************************************
 ***                        GET PROPERTY                            ***
 **********************************************************************/
  case VBM_GETPROPERTY:
  {
#ifdef DEBUG_BUILD
   char szError[128];

   wsprintf((LPSTR)szError, (LPSTR)"VBM_GETPROPERTY: #%d: Begin",
            wParam);
   DEBUGINI((LPSTR)szError);
#endif /* DEBUG_BUILD */

   switch(wParam)
   {

    /****************** GET: Section ********************/
    // Get Section property
    case IPROP_INI_SECTION:
    {
     HSZ hszTemp;

     DEBUGINI((LPSTR)"GET: Section - Start");

     if(!lParam) return(ERR_OUTOFMEMORY);

     hszTemp = GetMoleHsz(hCtl, lpIni->mSection);
     if(!hszTemp) return(ERR_OUTOFMEMORY);

     *(*(LPHSZ *)&lParam) = hszTemp;

     DEBUGINI((LPSTR)"GET: Section - End");
     return(ERR_None);
    }   // IPROP_INI_SECTION


    /****************** GET: Entry **********************/
    // Get Entry property
    case IPROP_INI_ENTRY:
    {
     HSZ hszTemp;

     DEBUGINI((LPSTR)"GET: Entry - Start");

     if(!lParam) return(ERR_OUTOFMEMORY);

     hszTemp = GetMoleHsz(hCtl, lpIni->mEntry);
     if(!hszTemp) return(ERR_OUTOFMEMORY);

     *(*(LPHSZ *)&lParam) = hszTemp;

     DEBUGINI((LPSTR)"GET: Entry - End");
     return(ERR_None);
    }   // IPROP_INI_ENTRY


    /****************** GET: Default ********************/
    // Get Default property
    case IPROP_INI_DEFAULT:
    {
     HSZ hszTemp;

     DEBUGINI((LPSTR)"GET: Default - Start");

     if(!lParam) return(ERR_OUTOFMEMORY);

     hszTemp = GetMoleHsz(hCtl, lpIni->mDefault);
     if(!hszTemp) return(ERR_OUTOFMEMORY);

     *(*(LPHSZ *)&lParam) = hszTemp;

     DEBUGINI((LPSTR)"GET: Default - End");
     return(ERR_None);
    }   // IPROP_INI_DEFAULT


    /****************** GET: Filename *******************/
    // Get Filename property
    case IPROP_INI_FILENAME:
    {
     HSZ hszTemp;

     DEBUGINI((LPSTR)"GET: Filename - Start");

     if(!lParam) return(ERR_OUTOFMEMORY);

     hszTemp = GetMoleHsz(hCtl, lpIni->mFilename);
     if(!hszTemp) return(ERR_OUTOFMEMORY);

     *(*(LPHSZ *)&lParam) = hszTemp;

     DEBUGINI((LPSTR)"GET: Filename - End");
     return(ERR_None);
    }   // IPROP_INI_FILENAME


    /****************** GET: Value **********************/
    // Get Value from .INI file
    case IPROP_INI_VALUE:
    {
     HSZ hszTemp;

     DEBUGINI((LPSTR)"GET: Value - Start");

     if(!lParam) return(ERR_OUTOFMEMORY);

     iniRebuildStrings(hCtl, lpIni);

     if(!GetPrivateProfileString((LPSTR)szSection, (LPSTR)szEntry,
                                 (LPSTR)szDefault, (LPSTR)szValue,
                                 sizeof(szValue), (LPSTR)szFilename))
     {
      DEBUGINI((LPSTR)"GET: Value - End - GetPrivateProfileString() returned 0!");
      hszTemp = GetMoleHsz(hCtl, 0);
     }
     else
     {
      hszTemp = VBCreateHsz((_segment)hCtl, (LPSTR)szValue);
     }
     if(!hszTemp) return(ERR_OUTOFMEMORY);

     *(*(LPHSZ *)&lParam) = hszTemp;

     DEBUGINI((LPSTR)"GET: Value - End");
     return(ERR_None);
    }   // IPROP_INI_VALUE

    default:
    {
     DEBUGINI((LPSTR)"VBM_GETPROPERTY Allowing default actions");
    }
   } // switch(wParam)
  } // VBM_GETPROPERTY
  break;

#ifdef DESIGN_TIME
/**********************************************************************
 ***                   CONTEXT SENSITIVE HELP                       ***
 **********************************************************************/
  case VBM_HELP:
  {
   DEBUGINI((LPSTR)"VBM_HELP: Begin");

   if(bHelpStdPropEvt(wParam, (LPMODEL)lParam))
        break;
   else return(DisplayHelpTopic(TOPIC_CONTROL_INI, wParam,
                                (LPMODEL)lParam));
   DEBUGINI((LPSTR)"VBM_HELP: Standard handler allowed");
  }
  break;
#endif // DESIGN_TIME

/**********************************************************************
 ***                        METHOD SUPPORT                          ***
 **********************************************************************/
  case VBM_METHOD:
  {
   DEBUGINI((LPSTR)"VBM_METHOD: Begin");

   switch(wParam)
   {
    /****************** METHOD: AddItem *****************/
    // Add an item to List
    case METH_ADDITEM:
    {
     DEBUGINI((LPSTR)"METHOD: AddItem - Begin");

     DEBUGINI((LPSTR)"METHOD: AddItem - End");
    }
    break; /* METH_ADDITEM */

    /****************** METHOD: RemoveItem **************/
    // Remove an item from the list
    case METH_REMOVEITEM:
    {
     DEBUGINI((LPSTR)"METHOD: RemoveItem - Begin");

     DEBUGINI((LPSTR)"METHOD: RemoveItem - End");
    }
    break; /* METH_REMOVEITEM */

    /****************** METHOD: Clear *******************/
    // Clear a [Section] or a group of entries=
    case METH_CLEAR:
    {
     DEBUGINI((LPSTR)"METHOD: Clear - Begin");

     DEBUGINI((LPSTR)"METHOD: Clear - End");
    }
    break; /* METH_CLEAR */
   }
   DEBUGINI((LPSTR)"VBM_METHOD: End - Default method used");
  }
  break;

 } // switch(Msg)
 return(VBDefControlProc(hCtl, hWnd, (USHORT)Msg, (USHORT)wParam,
        (LONG)lParam));
} /* IniCtlProc() */


/* VOID iniRebuildStrings(HCTL, LPINI);
    Purpose: To setup the static strings for usage
*/
VOID iniRebuildStrings(HCTL hCtl, LPINI lpIni)
{
 MOLE mTemp;

 DEBUGINI((LPSTR)"iniRebuildStrings() - Begin");

 *szSection = '\0';
 *szEntry = '\0';
 *szDefault = '\0';
 *szValue = '\0';
 lstrcpy((LPSTR)szFilename, (LPSTR)DEFAULT_FILENAME);

 if(!lpIni) return;

 if(lpIni->mSection)
 {
  mTemp = lpIni->mSection;
  GetMoleName(mTemp, (LPSTR)szSection, sizeof(szSection));
  lpIni=(LPINI)VBDerefControl(hCtl);
 }

 if(lpIni->mEntry)
 {
  mTemp = lpIni->mEntry;
  GetMoleName(mTemp, (LPSTR)szEntry, sizeof(szEntry));
  lpIni=(LPINI)VBDerefControl(hCtl);
 }

 if(lpIni->mDefault)
 {
  mTemp = lpIni->mDefault;
  GetMoleName(mTemp, (LPSTR)szDefault, sizeof(szDefault));
  lpIni=(LPINI)VBDerefControl(hCtl);
 }

 if(lpIni->mFilename)
 {
  mTemp = lpIni->mFilename;
  GetMoleName(mTemp, (LPSTR)szFilename, sizeof(szFilename));
  lpIni=(LPINI)VBDerefControl(hCtl);
 }

 DEBUGINI((LPSTR)"iniRebuildStrings() - End");

 return;
} /* iniRebuildStrings() */

/* End of Ini.C */


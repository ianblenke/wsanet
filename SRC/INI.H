/* INI.H - Ini Control header file */

    // Resource ID's for the button BMP's
#define IDBMP_INI        9000
#define IDBMP_INIDOWN    9001
#define IDBMP_INIMONO    9003
#define IDBMP_INIEGA     9006

    // Sizes for static strings
#define MAX_SECTION_SIZE  80
#define MAX_ENTRY_SIZE    80
#define MAX_DEFAULT_SIZE  256
#define MAX_FILENAME_SIZE 144
#define MAX_VALUE_SIZE    256

    // Default values
#define DEFAULT_FILENAME  "win.ini"

    // All of the property indicies
// Enumerations of Properties in Ini_Properties[]
// MUST BE IN SAME ORDER
enum enum_INI_props
{
    /* Standard properties */
 IPROP_INI_CTLNAME,
 IPROP_INI_PARENT,
 IPROP_INI_INDEX,
 IPROP_INI_TOP,
 IPROP_INI_LEFT,
 IPROP_INI_TAG,
 IPROP_INI_SECTION,
 IPROP_INI_ENTRY,
 IPROP_INI_DEFAULT,
 IPROP_INI_VALUE,
 IPROP_INI_FILENAME
};

// Errors sent back to VB (and .RC identifiers)
#define ERR_None                0

    // Returnable Errors within VB
#define ERR_OUTOFMEMORY         7

    // Ini's Control structure
typedef struct tagINI
{
 MOLE mSection;
 MOLE mEntry;
 MOLE mDefault;
 MOLE mFilename;
} INI;

typedef INI FAR * LPINI;


#ifdef WSANet_C

    // Section property
// This is the value [Section] in an .INI file
PROPINFO iniProperty_Section =
{
 "Section",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg | PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // Entry property
// This is the value Entry= in an .INI file
PROPINFO iniProperty_Entry =
{
 "Entry",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg | PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // Default property
// This sets the default value for Value if no such
// Entry= in [Section] in Filename.ini is found.
PROPINFO iniProperty_Default =
{
 "Default",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg | PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // Value property
// This is the Value of the Entry= in [Section] in Filename.ini
PROPINFO iniProperty_Value =
{
 "Value",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg | PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // Filename property
// This should hold the filename of the .INI file.
PROPINFO iniProperty_Filename =
{
 "Filename",
 DT_HSZ | PF_fGetMsg | PF_fSetMsg | PF_fNoInitDef,
 0,
 0, 0, NULL, 0
};

    // The Ini control's property information structure
PPROPINFO Ini_Properties[] =
{
 PPROPINFO_STD_CTLNAME,
 PPROPINFO_STD_PARENT,
 PPROPINFO_STD_INDEX,
 PPROPINFO_STD_TOP,
 PPROPINFO_STD_LEFT,
 PPROPINFO_STD_TAG,
 &iniProperty_Section,
 &iniProperty_Entry,
 &iniProperty_Default,
 &iniProperty_Value,
 &iniProperty_Filename,
 NULL
};

PEVENTINFO Ini_Events[] =
{
 NULL
};

MODEL modelIni =
{
 VB100_VERSION,
 MODEL_fDesInteract | MODEL_fLoadMsg,
 (PCTLPROC)IniCtlProc,
 CS_VREDRAW | CS_HREDRAW,
 WS_CHILD | WS_VISIBLE,
 sizeof(INI),
 IDBMP_INI,
 "Ini",
 "Ini",
 NULL,
 Ini_Properties,
 Ini_Events,
 IPROP_INI_VALUE,
 0,
     // VB2.0+ specific
 IPROP_INI_VALUE,
     // VB3.0+ specific
 INI_VERSION
};

#endif /* WSANet_C */

#define MODELLIST_INI (LPMODEL)&modelIni

#ifdef WSANet_C

    // All of the help "context-id's" are referenced through
    // this table
static WORD wIniHelpProps[] =
{
 PROPERTY_NAME,
 PROPERTY_PARENT,
 PROPERTY_INDEX,
 PROPERTY_TOP,
 PROPERTY_LEFT,
 PROPERTY_TAG,
 INI_PROP_SECTION,
 INI_PROP_ENTRY,
 INI_PROP_DEFAULT,
 INI_PROP_VALUE,
 INI_PROP_FILENAME,
 NULL
};

#endif /* WSANet_C */

    // Function Prototypes
VOID iniRebuildStrings(HCTL, LPINI);

/* End of Ini.H */

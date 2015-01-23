# Microsoft Visual C++ generated build script - Do not modify

PROJ = WSNETC
DEBUG = 1
PROGTYPE = 2
CALLER =
ARGS =
DLLS =
D_RCDEFINES = -d_DEBUG
R_RCDEFINES = -dNDEBUG
ORIGIN = MSVC
ORIGIN_VER = 1.00
PROJPATH = C:\WSNETC\
USEMFC = 0
CC = cl
CPP = cl
CXX = cl
CCREATEPCHFLAG =
CPPCREATEPCHFLAG =
CUSEPCHFLAG =
CPPUSEPCHFLAG =
FIRSTC = INIT.C
FIRSTCPP =
RC = rc
CFLAGS_D_VEXE = /nologo /G2 /Gc /Zp1 /Zi /ALw /Od /D "_DEBUG" /FR /GD
/Fd"WSNETC.PDB"
CFLAGS_R_VEXE = /nologo /Gs /G2 /Gc /Zp1 /W3 /ALw /Ox /D "NDEBUG" /FR /GD
LFLAGS_D_VEXE = /NOLOGO /NOD /NOE /PACKC:61440 /ALIGN:16 /ONERROR:NOEXE /CO
LFLAGS_R_VEXE = /NOLOGO /NOD /NOE /PACKC:61440 /ALIGN:16 /ONERROR:NOEXE
LIBS_D_VEXE = oldnames libw ldllcew commdlg.lib olecli.lib olesvr.lib
shell.lib
LIBS_R_VEXE = oldnames libw ldllcew commdlg.lib olecli.lib olesvr.lib
shell.lib
RCFLAGS = /nologo
RESFLAGS = /nologo
RUNFLAGS =
DEFFILE = WSNETC.DEF
OBJS_EXT = c:\VCP\CDK\LIBENTRY.OBJ
LIBS_EXT = WSNETC.LIB ..\WINSOCK\WINSOCK.LIB c:\VCP\CDK\VBAPI.LIB
!if "$(DEBUG)" == "1"
CFLAGS = $(CFLAGS_D_VEXE)
LFLAGS = $(LFLAGS_D_VEXE)
LIBS = $(LIBS_D_VEXE)
MAPFILE = nul
RCDEFINES = $(D_RCDEFINES)
!else
CFLAGS = $(CFLAGS_R_VEXE)
LFLAGS = $(LFLAGS_R_VEXE)
LIBS = $(LIBS_R_VEXE)
MAPFILE = nul
RCDEFINES = $(R_RCDEFINES)
!endif
!if [if exist MSVC.BND del MSVC.BND]
!endif
SBRS = INIT.SBR \
        NET.SBR \
        WSNETC.SBR


INIT_DEP = c:\wsnetc\wsnetc.h \
    /vb/cdk/vbapi.h \
    c:\wsnetc\include/winsock.h \
    c:\wsnetc\topics.h \
    c:\wsnetc\version.h


NET_DEP = c:\wsnetc\wsnetc.h \
    /vb/cdk/vbapi.h \
    c:\wsnetc\include/winsock.h \
    c:\wsnetc\topics.h \
    c:\wsnetc\version.h


WSNETC_DEP = c:\wsnetc\wsnetc.h \
    /vb/cdk/vbapi.h \
    c:\wsnetc\include/winsock.h \
    c:\wsnetc\topics.h


WSNETC_DEP =

WSNETC_RCDEP = c:\wsnetc\wsnetc.h \
    /vb/cdk/vbapi.h \
    c:\wsnetc\include/winsock.h \
    c:\wsnetc\topics.h \
    c:\wsnetc\version.h \
    c:\wsnetc\netcu.bmp \
    c:\wsnetc\netcd.bmp \
    c:\wsnetc\netmu.bmp \
    c:\wsnetc\neteu.bmp \
    c:\wsnetc\about.dlg


WINSOCK_DEP =

LIBENTRY_DEP =

VBAPI_DEP =

all:    $(PROJ).VBX $(PROJ).BSC

INIT.OBJ:   INIT.C $(INIT_DEP)
    $(CC) $(CFLAGS) $(CCREATEPCHFLAG) /c INIT.C

NET.OBJ:    NET.C $(NET_DEP)
    $(CC) $(CFLAGS) $(CUSEPCHFLAG) /c NET.C

WSNETC.OBJ: WSNETC.C $(WSNETC_DEP)
    $(CC) $(CFLAGS) $(CUSEPCHFLAG) /c WSNETC.C

WSNETC.RES: WSNETC.RC $(WSNETC_RCDEP)
    $(RC) $(RCFLAGS) $(RCDEFINES) -r WSNETC.RC


$(PROJ).VBX::   WSNETC.RES

$(PROJ).VBX::   INIT.OBJ NET.OBJ WSNETC.OBJ $(OBJS_EXT) $(DEFFILE)
    echo >NUL @<<$(PROJ).CRF
INIT.OBJ +
NET.OBJ +
WSNETC.OBJ +
$(OBJS_EXT)
$(PROJ).VBX
$(MAPFILE)
c:\msvc\lib\+
WSNETC.LIB+
..\WINSOCK\WINSOCK.LIB+
c:\VCP\CDK\VBAPI.LIB+
$(LIBS)
$(DEFFILE);
<<
    link $(LFLAGS) @$(PROJ).CRF
    $(RC) $(RESFLAGS) WSNETC.RES $@
    @copy $(PROJ).CRF MSVC.BND

$(PROJ).VBX::   WSNETC.RES
    if not exist MSVC.BND   $(RC) $(RESFLAGS) WSNETC.RES $@

run: $(PROJ).VBX
    $(PROJ) $(RUNFLAGS)


$(PROJ).BSC: $(SBRS)
    bscmake @<<
/o$@ $(SBRS)
<<

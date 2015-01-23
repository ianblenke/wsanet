@echo off
rem Dos 6.0 batch file
c:\dos\choice Debug or Release? (D,R) /N /C:RD /T:D,10
if errorlevel 2 goto DEBUG
:RELEASE
echo Building resources for RELEASE.VBX
set filebase=wsanet
set target=release
goto DOIT
:DEBUG
echo Building resources for WSANET.VBX (Debugging)
set filebase=wsanet
set target=wsanet
:DOIT
c:
cd \wsanet\src
copy c:\qcwin\bin\rcpp.* > NUL:
c:\qcwin\bin\rc -dUSE_31 -iC:\QCWIN\INCLUDE;C:\VB\CDK; -x -foC:\WSANET\SRC\%filebase%.RES -r C:\WSANET\SRC\WSANET.RC 
if not exist c:\WSANET\SRC\WSANET.RES goto ERROR_NORES
if not exist c:\WSANET\SRC\WSANET.DLL goto ERROR_NODLL
c:\qcwin\bin\rc /t c:\wsanet\src\%filebase%.res c:\wsanet\src\%filebase%.dll 
cd \wsanet\src
copy %filebase%.dll %target%.vbx > NUL:
copy %target%.vbx c:\windows\system > NUL:
del %filebase%.dll
del rcpp.* > NUL:
del *.res > NUL:
del *.mdt > NUL:
cd \wsanet\src
goto END

:ERROR_NORES
echo .RES file was not built! - Aborting
goto END

:ERROR_NODLL
echo .DLL file was not available! - Aborting

:END
c:\dos\choice Delay... hit Y to continue? /N /C:YN /T:Y,5
echo End.


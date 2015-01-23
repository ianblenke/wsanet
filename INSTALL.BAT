@echo off
echo WSANET.VBX v1.08alpha Installation
echo ----------------------------------
echo    WARNING: This will copy VBRUN300.DLL, THREED.VBX, and CMDIALOG.VBX
echo             over even newer (?) versions! Please CTL-BREAK now to
echo             abort now and manually install if you have a version of
echo             Visual Basic newer than VB3.0pro
pause
xcopy *.ini c:\windows
xcopy *.vbx c:\windows\system
rem xcopy runtime\*.* c:\windows\system
echo ----------------------------------
echo Copy complete. You may now add *.EXE to a ProgMan group.
echo ----------------------------------



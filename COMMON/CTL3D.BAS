Option Explicit

' Ctl3d Function declarations
Declare Function GetModuleHandle Lib "Kernel" (ByVal ModuleName As String) As Integer
Declare Function Ctl3dAutoSubclass Lib "Ctl3D.DLL" (ByVal hInst As Integer) As Integer
Declare Function Ctl3dRegister Lib "Ctl3D.DLL" (ByVal hInst As Integer) As Integer
Declare Function Ctl3dUnregister Lib "Ctl3D.DLL" (ByVal hInst As Integer) As Integer
Declare Function Ctl3dSubclassDlg Lib "Ctl3D.DLL" (ByVal hWndDlg As Integer, ByVal wgrBit As Integer) As Integer
Declare Function Ctl3dSubclassDlgEx Lib "Ctl3D.DLL" (ByVal hWndDlg As Integer, ByVal wgrBit As Long) As Integer
Declare Function Ctl3dDlgFramePaint Lib "Ctl3D.DLL" (ByVal hWndDlg As Integer, ByVal uMsg As Integer, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function Ctl3dSubclassCtl Lib "Ctl3D.DLL" (ByVal hWnd As Integer) As Integer

' Ctl3dSubclassDlg3d flags
Global Const CTL3D_BUTTONS = 1
Global Const CTL3D_LISTBOXES = 2
Global Const CTL3D_EDITS = 4
Global Const CTL3D_COMBOS = 8
Global Const CTL3D_STATICTEXTS = 16
Global Const CTL3D_STATICFRAMES = 32

Global Const CTL3D_NODLGWINDOW = &H10000
Global Const CTL3D_ALL = &HFFFF

Global Const WM_USER = &H400
Global Const WM_DLGBORDER = WM_USER + 3567

' WM_DLGBORDER *(int FAR *)lParam return codes
Global Const CTL3D_NOBORDER = 0
Global Const CTL3D_BORDER = 1

Global Const WM_DLGSUBCLASS = WM_USER + 3568

' WM_DLGSUBCLASS *(int FAR *)lParam return codes
Global Const CTL3D_NOSUBCLASS = 0
Global Const Ctl3d_Subclass = 1

' Resource ID for 3dcheck.bmp (for .lib version of ctl3d)
Global Const CTL3D_3DCHECK = 26567

Sub Ctl3d_Load ()
Dim hInst As Integer
Dim iRet As Integer
Dim sAppName As String

    sAppName = App.EXEName
    hInst = GetModuleHandle(sAppName)
    iRet = Ctl3dRegister(hInst)
    iRet = Ctl3dAutoSubclass(hInst)

End Sub

Sub Ctl3d_Unload ()
Dim hInst As Integer
Dim iRet As Integer
Dim sAppName As String

    sAppName = App.EXEName
    hInst = GetModuleHandle(sAppName)
    iRet = Ctl3dUnregister(hInst)

End Sub


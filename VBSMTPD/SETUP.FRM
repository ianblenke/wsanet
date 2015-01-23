VERSION 2.00
Begin Form Setup 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Setup VB-SMTPD"
   ClientHeight    =   3300
   ClientLeft      =   1560
   ClientTop       =   1830
   ClientWidth     =   3750
   ControlBox      =   0   'False
   Height          =   3675
   Left            =   1515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   3750
   Top             =   1500
   Width           =   3840
   Begin CommonDialog CMDialog 
      Left            =   120
      Top             =   2880
   End
   Begin SSFrame FrameOptions 
      Caption         =   "Logging options"
      Font3D          =   3  'Inset w/light shading
      Height          =   1875
      Left            =   120
      TabIndex        =   4
      Top             =   90
      Width           =   3495
      Begin SSCheck ChkLogSend 
         Caption         =   "Log sent lines"
         Font3D          =   0  'None
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin SSCheck ChkLogSummary 
         Caption         =   "Log incoming message summaries"
         Font3D          =   0  'None
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   3255
      End
      Begin SSCheck ChkLogWarning 
         Caption         =   "Log all warnings"
         Font3D          =   0  'None
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   3255
      End
      Begin SSCheck ChkLogRecv 
         Caption         =   "Log received lines."
         Font3D          =   0  'None
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   3255
      End
      Begin SSCheck ChkLogError 
         Caption         =   "Log all errors"
         Font3D          =   0  'None
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   3255
      End
      Begin SSCheck ChkLogDebug 
         Caption         =   "Log debugging information"
         Font3D          =   0  'None
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   3255
      End
   End
   Begin CommandButton BtnOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2880
      Width           =   915
   End
   Begin CommandButton BtnFilename 
      Caption         =   "&Reader: None enabled"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   3255
   End
   Begin SSCheck ChkNewRead 
      Caption         =   "New reader for each message"
      Font3D          =   0  'None
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   2895
   End
End
Option Explicit

Dim msFilename As String

Sub BtnFilename_Click ()
    
    On Error Resume Next

    CMDialog.CancelError = True
    CMDialog.DefaultExt = "*.exe;*.com;*.bat;*.pif"
    CMDialog.DialogTitle = "VB-SMTPD - Reader app"
    CMDialog.Filename = msFilename
    CMDialog.Filter = "Executables (*.exe;*.com;*.bat;*.pif)|*.exe;*.com;*.bat;*.pif|Everything (*.*)|*.*"
    CMDialog.FilterIndex = 1
    CMDialog.Flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
    CMDialog.Action = 1

    If Err <> 32755 Then
    
        msFilename = CMDialog.Filename
        BtnFilename.Caption = "&Reader: " & msFilename

        ChkNewRead.Enabled = True

        Main!Ini.Entry = "Reader"
        Main!Ini.Value = msFilename
    
    End If

End Sub

Sub BtnOk_Click ()

    mSetupSave

    Unload Setup

End Sub

Sub Form_Load ()

    Main!Ini.Section = "Windows"
    Setup.Top = gfMISCIniGetInt("Setup.Top", (Main.Top))
    Setup.Left = gfMISCIniGetInt("Setup.Left", (Main.Left))
    
    mSetupLoad

End Sub

Sub Form_QueryUnload (Cancel As Integer, UnloadMode As Integer)

    Main!Ini.Section = "Windows"
    gsMISCIniPutInt "Setup.Top", (Setup.Top)
    gsMISCIniPutInt "Setup.Left", (Setup.Left)

End Sub

Sub mCheckButton (iFlags As Integer, iSwitch As Integer, cChkControl As SSCheck)

    If iFlags And iSwitch Then
        cChkControl = True
    Else
        cChkControl = False
    End If
    
End Sub

Sub mSetupLoad ()
Dim iFlags As Integer
Dim sTemp As String

    Main!Ini.Section = "Configuration"

    ' Get the global log flags
    iFlags = gSMTPLogGetFlags()

    ' Mark/Unmark each item
    mCheckButton iFlags, LOG_ERROR, ChkLogError
    mCheckButton iFlags, LOG_WARNING, ChkLogWarning
    mCheckButton iFlags, LOG_SEND, ChkLogSend
    mCheckButton iFlags, LOG_RECV, ChkLogRecv
    mCheckButton iFlags, LOG_SUMMARY, ChkLogSummary
    mCheckButton iFlags, LOG_DEBUG, ChkLogDebug

    Main!Ini.Entry = "UniqueReader"
    sTemp = Main!Ini.Value
    sTemp = UCase$(sTemp)
    If sTemp = "1" Or Left$(sTemp, 1) = "T" Or Right$(sTemp, 1) = "N" Then
        ChkNewRead = False
    Else
        ChkNewRead = True
    End If

    Main!Ini.Entry = "Reader"
    If Main!Ini.Value = "" Then
        ChkNewRead.Enabled = False
        BtnFilename.Caption = "&Reader: None enabled"
    Else
        ChkNewRead.Enabled = True
        msFilename = Main!Ini.Value
        BtnFilename.Caption = "&Reader: " & msFilename
    End If

End Sub

Sub mSetupSave ()

    Main!Ini.Section = "Configuration"

    ' Clear ALL flags, we need to only set ones that
    ' the user kept selected
    gSMTPLogSetFlags LOG_NOTHING

    If ChkLogError Then
        gSMTPLogSetFlags LOG_ERROR
    End If

    If ChkLogWarning Then
        gSMTPLogSetFlags LOG_WARNING
    End If

    If ChkLogSend Then
        gSMTPLogSetFlags LOG_SEND
    End If

    If ChkLogRecv Then
        gSMTPLogSetFlags LOG_RECV
    End If

    If ChkLogSummary Then
        gSMTPLogSetFlags LOG_SUMMARY
    End If

    If ChkLogDebug Then
        gSMTPLogSetFlags LOG_DEBUG
    End If
    
    Main!Ini.Entry = "UniqueReader"
    If ChkNewRead Then
        Main!Ini.Value = "False"
    Else
        Main!Ini.Value = "True"
    End If

End Sub


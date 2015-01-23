VERSION 2.00
Begin Form Main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "VB-SMTPD"
   ClientHeight    =   3375
   ClientLeft      =   1035
   ClientTop       =   1740
   ClientWidth     =   5430
   Height          =   4035
   Icon            =   VBSMTPD.FRX:0000
   Left            =   990
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   5430
   Top             =   1125
   Width           =   5520
   Begin CommonDialog CMDialog 
      Left            =   4920
      Top             =   2520
   End
   Begin CommandButton BtnFilename 
      Caption         =   "No mailbox configured."
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4695
   End
   Begin SSFrame FrameListHosts 
      Alignment       =   2  'Center
      Font3D          =   2  'Raised w/heavy shading
      Height          =   1905
      Left            =   120
      ShadowStyle     =   1  'Raised
      TabIndex        =   3
      Top             =   960
      Width           =   4695
      Begin ListBox ListHosts 
         Height          =   1590
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   4515
      End
   End
   Begin SSPanel StatusBar 
      Align           =   2  'Align Bottom
      Alignment       =   0  'Left Justify - TOP
      AutoSize        =   3  'AutoSize Child To Panel
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      BorderWidth     =   1
      Caption         =   "Ok."
      Font3D          =   3  'Inset w/light shading
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   3075
      Width           =   5430
   End
   Begin SSPanel PanelHostName 
      Align           =   1  'Align Top
      Alignment       =   6  'Center - TOP
      BackColor       =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "PanelHostName"
      Font3D          =   4  'Inset w/heavy shading
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5430
      Begin Label LabelHostAddr 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "LabelHostAddr"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   5460
      End
   End
   Begin NetClient NetClient 
      Index           =   0
      Left            =   4950
      LineDelimiter   =   ""
      RecvSize        =   4096
      RecvThreshold   =   0
      RemoteService   =   ""
      Top             =   2070
   End
   Begin Ini Ini 
      Left            =   4950
      Top             =   1620
   End
   Begin NetServer NetServer 
      Left            =   4950
      QueueSize       =   0
      Top             =   1170
   End
   Begin Menu MnuSetup 
      Caption         =   "&Setup"
   End
   Begin Menu MnuAbout 
      Caption         =   "&About"
   End
End
Option Explicit

Sub BtnFilename_Click ()

    On Error Resume Next

    CMDialog.CancelError = True
    CMDialog.DefaultExt = ".mbx"
    CMDialog.DialogTitle = "VB-SMTPD - Mailbox file"
    CMDialog.Filename = msMailFile
    CMDialog.Filter = "Mailbox files (*.mbx)|*.mbx|Everything (*.*)|*.*"
    CMDialog.FilterIndex = 1
    CMDialog.Flags = OFN_PATHMUSTEXIST
    CMDialog.Action = 1

    If Err <> 32755 Then
    
        msMailFile = CMDialog.Filename
        BtnFilename.Caption = "&Mailbox: " & msMailFile

        Ini.Entry = "Mailbox"
        Ini.Value = msMailFile

    End If
End Sub

Sub Form_Load ()

    'Ctl3d_Load

    If NetClient(0).HostName <> "" Then
        PanelHostName = NetClient(0).HostName
        LabelHostAddr = NetClient(0).HostAddr
    Else
        PanelHostName = "LocalHost"
        LabelHostAddr = "unable to retreive local address"
    End If

    Ini.Default = ""
    Ini.Filename = "vbsmtpd.ini"
    Ini.Section = "Configuration"
    Ini.Entry = "Mailbox"
    msMailFile = Ini.Value
    If msMailFile = "" Then
        BtnFilename.Caption = "&Mailbox: None - Popup"
    Else
        BtnFilename.Caption = "&Mailbox: " & msMailFile
    End If

    gSMTPStartup NetServer

    Ini.Section = "Windows"
    Main.Top = gfMISCIniGetInt("Main.Top", (Main.Top))
    Main.Left = gfMISCIniGetInt("Main.Left", (Main.Left))
    Main.Width = gfMISCIniGetInt("Main.Width", (Main.Width))
    Main.Height = gfMISCIniGetInt("Main.Height", (Main.Height))

    Main.Show

End Sub

Sub Form_QueryUnload (Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Ini.Section = "Windows"
        gsMISCIniPutInt "Main.Top", (Main.Top)
        gsMISCIniPutInt "Main.Left", (Main.Left)
        gsMISCIniPutInt "Main.Width", (Main.Width)
        gsMISCIniPutInt "Main.Height", (Main.Height)
    End If

    gSMTPCleanup
    'Ctl3d_UnLoad

End Sub

Sub Form_Resize ()

    If Main.WindowState = 1 Then Exit Sub

    LabelHostAddr.Width = ScaleWidth
    BtnFilename.Width = Abs(ScaleWidth - 240)

    FrameListHosts.Width = Abs(ScaleWidth - 240)
    FrameListHosts.Height = Abs(ScaleHeight - FrameListHosts.Top - StatusBar.Height - 120)

    ListHosts.Width = Abs(FrameListHosts.Width - 240)
    ListHosts.Height = Abs(FrameListHosts.Height - 240)

End Sub

Sub MnuAbout_Click ()

    gsMISCAboutLoad "About VB-SMTPD", "VB-SMTPD v1.1", "This is a multi-connection SMTP server (normally port 25). You should configure a mailbox to hold incoming mail."

End Sub

Sub MnuSetup_Click ()
    
    Load Setup
    Setup.Show

End Sub

Sub NetClient_OnError (Index As Integer, iErrorNumber As Integer)
Dim sTemp As String

    sTemp = NetClient(Index).ErrorMessage
    gSMTPError -1, iErrorNumber, sTemp

End Sub

Sub NetClient_OnRecv (Index As Integer)

    gSMTPReceive Index

End Sub

Sub NetServer_OnAccept (Socket As Integer, PeerAddr As String, RemotePort As Integer)

    gSMTPAccept Socket, PeerAddr, RemotePort

End Sub

Sub NetServer_OnError (iErrorNumber As Integer)
Dim sTemp As String

    sTemp = NetServer.ErrorMessage
    gSMTPError -1, iErrorNumber, sTemp

End Sub


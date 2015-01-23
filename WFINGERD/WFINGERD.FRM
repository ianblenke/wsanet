VERSION 2.00
Begin Form Main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "WFingerD - Finger Daemon"
   ClientHeight    =   4620
   ClientLeft      =   960
   ClientTop       =   1680
   ClientWidth     =   6090
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "MS Sans Serif"
   FontSize        =   9.75
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   5280
   Icon            =   WFINGERD.FRX:0000
   Left            =   915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   6090
   Top             =   1065
   Width           =   6180
   Begin CommandButton PanelFilename 
      Caption         =   "PanelFilename"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   5295
   End
   Begin SSPanel PanelLocalHost 
      Align           =   1  'Align Top
      Alignment       =   6  'Center - TOP
      BackColor       =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "PanelLocalHost"
      Font3D          =   4  'Inset w/heavy shading
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6090
      Begin Label LabelLocalAddr 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "LabelLocalAddr"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   6015
      End
   End
   Begin CommonDialog CMDialog 
      Left            =   5520
      Top             =   3000
   End
   Begin Ini Ini 
      Left            =   5520
      Top             =   1320
   End
   Begin NetClient NetClient 
      Left            =   5520
      LineDelimiter   =   ""
      RecvThreshold   =   0
      RemoteService   =   ""
      SendSize        =   8192
      Top             =   1800
   End
   Begin SSPanel StatusBar 
      Align           =   2  'Align Bottom
      Alignment       =   1  'Left Justify - MIDDLE
      AutoSize        =   3  'AutoSize Child To Panel
      BackColor       =   &H00C0C0C0&
      BevelInner      =   1  'Inset
      BorderWidth     =   1
      Caption         =   "Ok."
      Font3D          =   3  'Inset w/light shading
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   4320
      Width           =   6090
   End
   Begin SSFrame FrameListHosts 
      Caption         =   "Finger Clients"
      Font3D          =   2  'Raised w/heavy shading
      Height          =   3135
      Left            =   120
      ShadowStyle     =   1  'Raised
      TabIndex        =   1
      Top             =   1080
      Width           =   5295
      Begin ListBox ListHosts 
         Height          =   2760
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5055
      End
   End
   Begin NetServer NetServer 
      Left            =   5520
      QueueSize       =   0
      Top             =   2400
   End
   Begin Menu MenuAbout 
      Caption         =   "&About"
   End
End
Option Explicit

Dim msFingerFile As String
Dim miFile As Integer
Dim mlLength As Long

Const OFN_FILEMUSTEXIST = &H1000&
Const OFN_PATHMUSTEXIST = &H800&
Const OFN_READONLY = &H1&

Const MAX_SIZE = 16384
Const CHUNK_SIZE = 8192

Sub Form_Load ()
    
    On Error Resume Next

    'Ctl3d_Load

    PanelLocalHost = NetClient.HostName
    LabelLocalAddr = NetClient.HostAddr
    
    If PanelLocalHost = "" Then
        PanelLocalHost = "LocalHost"
        LabelLocalAddr = "unknown address"
    End If

    NetServer.LocalService = "finger"
    If NetServer.LocalService = "" Then
        NetServer.LocalPort = 79
    End If

    NetClient.RecvSize = MAX_SIZE
    NetClient.SendSize = MAX_SIZE

    NetServer.Listen = True

    ListHosts.Clear

    ' Read our .INI file for a default msFingerFile
    Ini.Default = ""
    Ini.Filename = "wfingerd.ini"
    Ini.Section = "Configuration"
    Ini.Entry = "FingerFile"
    If Ini.Value = "" Then
        Ini.Value = ""
        msFingerFile = ""
    Else
        msFingerFile = Ini.Value
    End If
        
    PanelFilename.Caption = msFingerFile
    
    Ini.Section = "Windows"
    Main.Top = gfMISCIniGetInt("Main.Top", (Main.Top))
    Main.Left = gfMISCIniGetInt("Main.Left", (Main.Left))
    Main.Width = gfMISCIniGetInt("Main.Width", (Main.Width))
    Main.Height = gfMISCIniGetInt("Main.Height", (Main.Height))

End Sub

Sub Form_QueryUnload (Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Ini.Section = "Windows"
        gsMISCIniPutInt "Main.Top", (Main.Top)
        gsMISCIniPutInt "Main.Left", (Main.Left)
        gsMISCIniPutInt "Main.Width", (Main.Width)
        gsMISCIniPutInt "Main.Height", (Main.Height)
    End If

End Sub

Sub Form_Resize ()

    If Main.WindowState = 1 Then Exit Sub

    LabelLocalAddr.Width = ScaleWidth

    PanelFilename.Width = Abs(ScaleWidth - 240)

    FrameListHosts.Width = Abs(ScaleWidth - 240)
    FrameListHosts.Height = Abs(ScaleHeight - FrameListHosts.Top - StatusBar.Height - 120)
    
    ListHosts.Width = Abs(FrameListHosts.Width - 240)
    ListHosts.Height = Abs(FrameListHosts.Height - 360)

End Sub

Sub Form_Unload (Cancel As Integer)
    
    NetServer.Listen = False
    
    'Ctl3d_UnLoad

End Sub

Sub MenuAbout_Click ()

    gsMISCAboutLoad "About WFingerD", "WFingerD v1.1", "This daemon application services TCP Finger clients (typically port 79). You can select any file you wish to be your machine's finger response."

End Sub

Sub NetClient_OnError (iErrorNumber As Integer)
    
    StatusBar = NetClient.ErrorMessage
    
    Debug.Print "NetClient:" & NetClient.ErrorMessage

    Close #miFile
    NetClient.Connect = False

End Sub

Sub NetClient_OnRecv ()
Dim sTemp As String
Dim iPlace As Integer

    On Error GoTo ErrorHandler

    If NetClient.RecvCount = 0 Then Exit Sub
    sTemp = NetClient

    While Right$(sTemp, 1) = Chr$(13) Or Right$(sTemp, 1) = Chr$(10)
        sTemp = Left$(sTemp, Len(sTemp) - 1)
    Wend

    NetClient.SendThreshold = 1
    ListHosts.AddItem NetClient.HostName & "(" & NetClient.HostAddr & ")" & Chr$(9) & Str$(NetClient.RemotePort) & Chr$(9) & sTemp

    If msFingerFile = "" Then
        NetClient = "Oops, I've forgotten what I'm supposed to tell you... Try again later!"
        NetClient.Connect = False
        Exit Sub
    End If
        
    miFile = FreeFile
    Open msFingerFile For Input As #miFile
    
    mlLength = LOF(miFile) - 1
    
    SendFingerFile

    Exit Sub

ErrorHandler:
    Select Case Err
        Case 62:
            Resume Next
        Case 53:
            NetClient = "Tell my owner that his WFingerD file is missing!" & Chr$(13) & Chr$(10)
            NetClient.Connect = False
            Exit Sub
    End Select
    NetClient = Chr$(13) & Chr$(10) & "Oops, Tell my owner that I had an error" & Str$(Err) & " (" & Error$(Err) & ") on line " & Str$(Erl) & " in WFingerD. Thanks!" & Chr$(13) & Chr$(10)
    If miFile <> 0 Then Close #miFile
    NetClient.Connect = False
    Exit Sub

End Sub

Sub NetClient_OnSend ()
Dim sTemp As String

    SendFingerFile

End Sub

Sub NetServer_OnAccept (Socket As Integer, PeerAddr As String, RemotePort As Integer)
    
    On Error Resume Next

    If NetClient.Connect = True Then
        Socket = 0
        Exit Sub
    End If

    NetClient.Socket = Socket
    
    NetClient.Debug = True

End Sub

Sub NetServer_OnError (iErrorNumber As Integer)

    StatusBar = NetServer.ErrorMessage
    
    Debug.Print "NetServer:" & NetServer.ErrorMessage

End Sub

Sub PanelFilename_Click ()

    On Error Resume Next

    CMDialog.CancelError = True
    CMDialog.DefaultExt = ".txt"
    CMDialog.DialogTitle = "WFingerD - Load Finger info"
    CMDialog.Filename = msFingerFile
    CMDialog.Filter = "Text (*.txt)|*.txt|Everything (*.*)|*.*"
    CMDialog.FilterIndex = 1
    CMDialog.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_READONLY
    CMDialog.Action = 1
    
    If Err <> 32755 Then
       
       msFingerFile = CMDialog.Filename
       PanelFilename.Caption = msFingerFile
       
       Ini.Filename = "wfingerd.ini"
       Ini.Section = "Configuration"
       Ini.Entry = "FingerFile"
       Ini.Value = msFingerFile

    End If
        
End Sub

Sub SendFingerFile ()
Dim sTemp As String

    On Error GoTo ErrorHandler2

    If Not EOF(miFile) Then
        If mlLength > CHUNK_SIZE Then
            sTemp = Input$(CHUNK_SIZE, miFile)
            NetClient = sTemp
            mlLength = mlLength - CHUNK_SIZE
        Else
            sTemp = Input$(mlLength, miFile)
            If sTemp = "" Then
                Close #miFile
                NetClient.Connect = False
                Exit Sub
            End If
            NetClient = sTemp
        End If
    Else
        Close #miFile
        NetClient.Connect = False
    End If

    Exit Sub

ErrorHandler2:
    Select Case Err
        Case 62:
            Resume Next
    End Select
    NetClient = Chr$(13) + Chr$(10) + "Oops, Tell my owner that I had an error" + Str$(Err) + " (" + Error$(Err) + ") on line " + Str$(Erl) + " in WFingerD. Thanks!" + Chr$(13) + Chr$(10)
    If miFile <> 0 Then Close #miFile
    NetClient.Connect = False
    Exit Sub

End Sub


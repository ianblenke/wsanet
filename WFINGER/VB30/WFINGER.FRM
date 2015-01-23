VERSION 2.00
Begin Form Main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "WFinger - Test of WSANET"
   ClientHeight    =   5265
   ClientLeft      =   990
   ClientTop       =   1680
   ClientWidth     =   6930
   Height          =   5925
   Icon            =   WFINGER.FRX:0000
   Left            =   945
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   6930
   Top             =   1065
   Width           =   7020
   Begin Ini Ini 
      Left            =   6480
      Top             =   1920
   End
   Begin SSPanel StatusBar 
      Align           =   2  'Align Bottom
      Alignment       =   1  'Left Justify - MIDDLE
      AutoSize        =   3  'AutoSize Child To Panel
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ok."
      Font3D          =   3  'Inset w/light shading
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5010
      Width           =   6930
   End
   Begin SSFrame FrameTextDisplay 
      Alignment       =   2  'Center
      Font3D          =   3  'Inset w/light shading
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   6255
      Begin TextBox TextDisplay 
         Height          =   3495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "Empty."
         Top             =   240
         Width           =   6015
      End
   End
   Begin NetClient NetClient 
      Left            =   6480
      LineDelimiter   =   ""
      RecvSize        =   4096
      RecvThreshold   =   0
      Top             =   1425
   End
   Begin SSPanel PanelLocalName 
      Align           =   1  'Align Top
      Alignment       =   6  'Center - TOP
      BackColor       =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "LocalHost"
      Font3D          =   1  'Raised w/light shading
      Height          =   855
      Left            =   0
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Top             =   0
      Width           =   6930
      Begin SSPanel PanelUsername 
         Alignment       =   6  'Center - TOP
         AutoSize        =   3  'AutoSize Child To Panel
         BackColor       =   &H00C0C0C0&
         Font3D          =   1  'Raised w/light shading
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1215
         Begin TextBox TextUserName 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   225
            Left            =   15
            MaxLength       =   64
            TabIndex        =   9
            Text            =   "iblenke"
            Top             =   15
            Width           =   1185
         End
      End
      Begin SSPanel PanelHost 
         Alignment       =   6  'Center - TOP
         AutoSize        =   3  'AutoSize Child To Panel
         BackColor       =   &H00C0C0C0&
         Font3D          =   1  'Raised w/light shading
         Height          =   255
         Left            =   1740
         TabIndex        =   3
         Top             =   480
         Width           =   5040
         Begin TextBox TextHost 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   225
            Left            =   15
            MaxLength       =   128
            TabIndex        =   4
            Text            =   "rhino.ess.harris.com"
            Top             =   15
            Width           =   5010
         End
      End
      Begin Label LabelLocalAddr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6735
      End
      Begin Label LabelAt 
         BackColor       =   &H00C0C0C0&
         Caption         =   "@"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   1425
         TabIndex        =   2
         Top             =   480
         Width           =   255
      End
   End
   Begin Menu MenuAbout 
      Caption         =   "&About"
   End
End
Dim CRLF As String

Sub CancelFinger ()

    NetClient.Connect = False

    TextHost.Enabled = True
    TextUserName.Enabled = True

    Main.MousePointer = 0
    PanelLocalName.MousePointer = 2
    StatusBar.MousePointer = 0

End Sub

Sub FingerUser ()
Dim sDebugString As String

    On Error Resume Next

    StatusBar.Caption = "Ok."
    
    TextDisplay.Text = ""

    ' WSNetC code
    NetClient.HostName = TextHost
    If NetClient.HostName = "" Then
        NetClient.HostAddr = TextHost
        If NetClient.HostAddr = "" Then
            StatusBar = "Sorry, " & TextHost & " could not be resolved."
            Exit Sub
        End If
    End If
    NetClient.Connect = True
    
    ' We are completing a query, ignore user
    TextDisplay.SetFocus
    TextUserName.Enabled = False
    TextHost.Enabled = False

    StatusBar.Caption = "Connecting to Host - Click this message to Abort"
    
    Main.MousePointer = 11
    PanelLocalName.MousePointer = 12
    StatusBar.MousePointer = 2

End Sub

Sub Form_Load ()
    
    On Error Resume Next

    If NetClient.HostName <> "" Then
        PanelLocalName = NetClient.HostName
        LabelLocalAddr = NetClient.HostAddr
    Else
        PanelLocalName = "LocalHost"
        LabelLocalAddr = "unknown address"
    End If
    
    CRLF = Chr$(13) + Chr$(10)

    ' We want to grab everything in "chunks"
    NetClient.LineDelimiter = ""
    NetClient.SendSize = 512
    NetClient.RecvSize = 8192

    ' Use the "finger" service (or port 79)
    NetClient.RemoteService = "finger"
    If NetClient.RemotePort = 0 Then
        NetClient.RemotePort = 79
    End If

    ' Handle saved user info
    Ini.Filename = "wfinger.ini"
    Ini.Section = "Configuration"
    
    Ini.Entry = "DefaultUserName"
    If Ini.Value = "" Then
        Ini.Value = "iblenke"
    End If
    TextUserName.Text = Ini.Value

    Ini.Entry = "DefaultHostName"
    If Ini.Value = "" Then
        Ini.Value = "rhino.ess.harris.com"
    End If
    TextHost.Text = Ini.Value

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

End Sub

Sub Form_Resize ()
    
    If Main.WindowState = 1 Then Exit Sub
    
    LabelLocalAddr.Width = Abs(ScaleWidth - 240)

    FrameTextDisplay.Width = Abs(ScaleWidth - 240)
    FrameTextDisplay.Height = Abs(ScaleHeight - FrameTextDisplay.Top - StatusBar.Height - 120)

    TextDisplay.Width = Abs(FrameTextDisplay.Width - 240)
    TextDisplay.Height = Abs(FrameTextDisplay.Height - 360)

    PanelHost.Width = Abs(ScaleWidth - PanelHost.Left - 120)

End Sub

Sub Form_Unload (Cancel As Integer)
    
    Ini.Entry = "DefaultHostName"
    Ini.Value = TextHost.Text

    Ini.Entry = "DefaultUserName"
    Ini.Value = TextUserName.Text

End Sub

Sub LabelLocalAddr_DblClick ()
    
    PanelLocalName_DblClick

End Sub

Sub MenuAbout_Click ()
    
    gsMISCAboutLoad "WFinger", "WFinger v1.0", "This client can query remote hosts through the Finger service (typically port 79). You can usually leave the 'User name' field empty for a full listing of people logged in."

End Sub

Sub NetClient_OnClose ()
    
    On Error Resume Next
    
    CancelFinger

    StatusBar.Caption = "Connection closed."
    
End Sub

Sub NetClient_OnConnect ()
    
    On Error Resume Next

    FrameTextDisplay.Caption = TextUserName & "@" & TextHost
    StatusBar.Caption = "Connected to port" & Str$(NetClient.RemotePort) & ". Click this message to Abort"
    NetClient.SendLine = TextUserName.Text & CRLF

End Sub

Sub NetClient_OnError (iErrorNumber As Integer)
Dim ErrorMessage As String
Dim ErrorNumberText As String

    On Error Resume Next

    ErrorMessage = NetClient.ErrorMessage
    ErrorNumberText = Str$(iErrorNumber)

    StatusBar.Caption = "Error #" & ErrorNumberText & ": " & ErrorMessage
    
    ' Disconnect for good measure
    CancelFinger

End Sub

Sub NetClient_OnRecv ()
Dim sInput As String

    On Error Resume Next

    sInput = NetClient.RecvLine
    
    ' Loop until buffer is empty (this doesn't block)
    Do While Not (sInput = "")
        TextDisplay.Text = TextDisplay.Text + sInput
        sInput = NetClient.RecvLine
    Loop
End Sub

Sub PanelLocalName_DblClick ()
    
    PanelLocalName.BevelOuter = 1
    PanelLocalName.Font3D = 2
    FingerUser
    PanelLocalName.Font3D = 1
    PanelLocalName.BevelOuter = 2

End Sub

Sub StatusBar_DblClick ()

    If Not NetClient.Connect Then Exit Sub

    StatusBar.BevelOuter = 1
    CancelFinger
    StatusBar.BevelOuter = 2

    StatusBar.Caption = "Finger aborted."

End Sub

Sub TextHost_GotFocus ()
    PanelHost.BevelOuter = 1
    TextHost.SelStart = 0
    TextHost.SelLength = 128
End Sub

Sub TextHost_KeyPress (KeyAscii As Integer)

    If KeyAscii = 13 Then
        FingerUser
        KeyAscii = 0
    End If
End Sub

Sub TextHost_LostFocus ()
    PanelHost.BevelOuter = 2

End Sub

Sub TextUserName_GotFocus ()
    PanelUsername.BevelOuter = 1
    TextUserName.SelStart = 0
    TextUserName.SelLength = 64
End Sub

Sub TextUserName_KeyPress (KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        FingerUser
        KeyAscii = 0
    End If
End Sub

Sub TextUsername_LostFocus ()
    PanelUsername.BevelOuter = 2
End Sub


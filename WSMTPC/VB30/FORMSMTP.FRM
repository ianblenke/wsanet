VERSION 2.00
Begin Form Main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SMTP Test Client"
   ClientHeight    =   4260
   ClientLeft      =   1035
   ClientTop       =   1665
   ClientWidth     =   8475
   Height          =   4920
   Icon            =   FORMSMTP.FRX:0000
   Left            =   990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   8475
   Top             =   1050
   Width           =   8565
   Begin Ini Ini 
      Left            =   7800
      Top             =   1800
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
      TabIndex        =   9
      Top             =   4005
      Width           =   8475
   End
   Begin SSFrame FrameProgress 
      Font3D          =   0  'None
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   7575
      Begin ListBox ListProgress 
         Height          =   2370
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7335
      End
   End
   Begin NetClient NetClient 
      Left            =   7800
      LineDelimiter   =   ""
      RecvSize        =   4096
      RecvThreshold   =   0
      Top             =   1320
   End
   Begin SSPanel PanelLocalName 
      Align           =   1  'Align Top
      Alignment       =   6  'Center - TOP
      BackColor       =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "PanelLocalHost"
      Font3D          =   4  'Inset w/heavy shading
      Height          =   855
      Left            =   0
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   0
      Width           =   8475
      Begin SSPanel PanelUserName 
         Alignment       =   6  'Center - TOP
         AutoSize        =   3  'AutoSize Child To Panel
         BackColor       =   &H00C0C0C0&
         Font3D          =   1  'Raised w/light shading
         Height          =   255
         Left            =   105
         TabIndex        =   4
         Top             =   480
         Width           =   1620
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
            TabIndex        =   5
            Text            =   "iblenke"
            Top             =   15
            Width           =   1590
         End
      End
      Begin SSPanel PanelHost 
         Alignment       =   6  'Center - TOP
         AutoSize        =   3  'AutoSize Child To Panel
         BackColor       =   &H00C0C0C0&
         Font3D          =   1  'Raised w/light shading
         Height          =   255
         Left            =   2115
         TabIndex        =   2
         Top             =   480
         Width           =   2985
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
            TabIndex        =   3
            Text            =   "rhino.ess.harris.com"
            Top             =   15
            Width           =   2955
         End
      End
      Begin Label LabelLocalAddr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LabelLocalAddr"
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
         Width           =   8175
      End
      Begin Label LabelAt 
         BackColor       =   &H00C0C0C0&
         Caption         =   "@"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   375
         Left            =   1785
         TabIndex        =   1
         Top             =   480
         Width           =   375
      End
   End
   Begin Menu MenuAbout 
      Caption         =   "&About"
   End
End
Dim LocalHostName As String
Dim CRLF As String

Const STATE_INACTIVE = 0
Const STATE_HELO = 1
Const STATE_VRFY = 2
Const STATE_MAILFROM = 3
Const STATE_RCPTTO = 4
Const STATE_DATA = 5
Const STATE_SENDBODY = 6
Const STATE_QUIT = 7

Dim SMTPState As Integer

Sub Form_Load ()
    
    If NetClient.HostName = "" Then
        PanelLocalName = "LocalHost"
        PanelLocalAddr = "unknown address"
    Else
        PanelLocalName = NetClient.HostName
        LabelLocalAddr = NetClient.HostAddr
    End If
    
    LocalHostName = PanelLocalName
     
    CRLF = Chr$(13) + Chr$(10)
    NetClient.LineDelimiter = CRLF
    
    ' Use SMTP service (port 25)
    NetClient.RemoteService = "smtp"
    If NetClient.RemotePort = 0 Then
        NetClient.RemotePort = 25
    End If

    Ini.Filename = "wsmtpc.ini"
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
    PanelHost.Width = Abs(ScaleWidth - PanelHost.Left - 120)

    TextHost.Width = Abs(ScaleWidth - TextHost.Left - 120)

    FrameProgress.Width = Abs(ScaleWidth - 240)
    FrameProgress.Height = Abs(ScaleHeight - FrameProgress.Top - StatusBar.Height - 120)
    
    ListProgress.Width = Abs(FrameProgress.Width - 240)
    ListProgress.Height = Abs(FrameProgress.Height - 360)

End Sub

Sub LabelLocalAddr_DblClick ()
    
    PanelLocalName_DblClick

End Sub

Sub MenuAbout_Click ()

    Ini.Section = "Windows"
    gsMISCAboutLoad "About WSMTPC", "WSMTPC v1.1", "This client is merely an example of a state-based application. This program merely sends a single line 'message' to a user on an SMTP host, and shows the SMTP conversation."

End Sub

Sub NetClient_OnClose ()
    
    On Error Resume Next

    SMTPState = STATE_INACTIVE
    
    SMTPCancel

    StatusBar.Caption = "Ok."

End Sub

Sub NetClient_OnConnect ()

    On Error Resume Next
    
    StatusBar.Caption = "Connected to " & NetClient.HostName & " - Click this message to Abort."
    
    SMTPState = STATE_HELO

End Sub

Sub NetClient_OnError (ErrorNumber As Integer)
Dim sTemp As String

    On Error Resume Next

    ' Can't read past end of connection!
    If ErrorNumber = 10038 Then
        If NetClient.Socket = -1 Then
            Exit Sub
        End If
    End If
    
    sTemp = NetClient.ErrorMessage
    ReportProgress sTemp
    
    SMTPCancel

End Sub

Sub NetClient_OnRecv ()
Dim sLine As String
Dim iReturn As Integer

    On Error Resume Next

    Do
        sLine = NetClient.RecvLine
        If sLine = "" Then Exit Do
        ReportProgress ">" + sLine

        ' Handle SMTP multi-line responses
        While Mid$(sLine, 4, 1) = "-"
            sLine = NetClient.RecvLine
        Wend

        SMTPParse sLine
        DoEvents
    Loop While NetClient.RecvCount > 0
    
End Sub

Sub PanelLocalName_DblClick ()

    PanelLocalName.BevelOuter = 1
    PanelLocalName.Font3D = 2
    SMTPStart
    PanelLocalName.Font3D = 1
    PanelLocalName.BevelOuter = 2
    
End Sub

Sub ReportProgress (sMessage As String)
    ListProgress.AddItem sMessage
End Sub

Sub SMTPCancel ()

    NetClient.Connect = False

    Main.MousePointer = 0
    PanelLocalName.MousePointer = 2
    StatusBar.MousePointer = 0

    TextUserName.Enabled = True
    TextHost.Enabled = True
    
End Sub

Sub SMTPParse (sLine As String)
Dim iReturn As Integer
Dim sTemp As String

    On Error Resume Next

    iReturn = Val(Left$(sLine, 3))

    Select Case iReturn
        Case 200 To 299:    ' Ok reply
            Select Case SMTPState
                Case STATE_HELO:
                    sLine = "HELO " & LocalHostName
                    NetClient = sLine & CRLF
                    ReportProgress sLine
                    SMTPState = STATE_VRFY
                Case STATE_VRFY
                    sLine = "VRFY <" & TextUserName & "@" & TextHost & " > "
                    NetClient = sLine & CRLF
                    ReportProgress sLine
                    SMTPState = STATE_MAILFROM
                Case STATE_MAILFROM:
                    sLine = "MAIL FROM: <postmaster" & "@" & LocalHostName & ">"
                    NetClient = sLine & CRLF
                    ReportProgress sLine
                    SMTPState = STATE_RCPTTO
                Case STATE_RCPTTO:
                    sLine = "RCPT TO: <" & TextUserName & "@" & TextHost & ">"
                    NetClient = sLine & CRLF
                    ReportProgress sLine
                    SMTPState = STATE_DATA
                Case STATE_DATA
                    sLine = "DATA"
                    NetClient = sLine & CRLF
                    ReportProgress sLine
                    SMTPState = STATE_SENDBODY
                Case STATE_QUIT
                    sLine = "QUIT"
                    NetClient = sLine & CRLF
                    ReportProgress sLine
                    SMTPState = STATE_INACTIVE
                Case STATE_INACTIVE
                    SMTPCancel
                    StatusBar.Caption = "Ok."
                End Select

        Case 300 To 399:    ' Informational reply
            Select Case SMTPState
                Case STATE_SENDBODY:
                    sTemp = "From: <postmaster@" & LocalHostName & ">"
                    ReportProgress sTemp
                    sLine = sTemp & CRLF
                    sTemp = "To: <" & TextUserName & "@" & TextHost & ">"
                    ReportProgress sTemp
                    sLine = sLine & sTemp & CRLF
                    sTemp = "Subject: WSMTPC Test Message"
                    ReportProgress sTemp
                    sLine = sLine & sTemp & CRLF & CRLF
                    sTemp = "This is a test of the SMTP client that comes with WSANET."
                    ReportProgress sTemp
                    sLine = sLine & sTemp & CRLF
                    sTemp = "."
                    ReportProgress sTemp
                    sLine = sLine & sTemp & CRLF
                    NetClient = sLine
                    SMTPState = STATE_QUIT
            End Select
            
        Case 500 To 599:    ' Error reply
            ReportProgress "500 Error! Abort!"
            Select Case SMTPState
                Case STATE_SENDBODY:
                    NetClient = "." + CRLF
                    ReportProgress "."
                    SMTPState = STATE_QUIT
                Case Else
                    NetClient = "QUIT" + CRLF
                    ReportProgress "QUIT"
            End Select
        
        Case Else
                ReportProgress "Unknown reply #" & Str$(iReturn) & "0!"
    End Select
    
End Sub

Sub SMTPStart ()
    
    On Error Resume Next

    NetClient.HostName = TextHost.Text
    If NetClient.HostName = "" Then
        NetClient.HostAddr = TextHost.Text
        If NetClient.HostAddr = "" Then
            ListProgress.AddItem "Host " & TextHost.Text & " unknown."
            Exit Sub
        End If
    End If
    
    NetClient.Connect = True
    
    ListProgress.Clear
    StatusBar.Caption = "Connecting to " & TextHost.Text & " - Click this message to Abort."
    
    Main.MousePointer = 11
    PanelLocalName.MousePointer = 12
    StatusBar.MousePointer = 2

    SMTPState = STATE_HELO

End Sub

Sub StatusBar_DblClick ()

    If NetClient.Connect Then Exit Sub
    
    StatusBar.BevelOuter = 1
    SMTPCancel
    StatusBar.BevelOuter = 2

    StatusBar.Caption = "Connection Aborted."

End Sub

Sub TextHost_GotFocus ()
    
    PanelHost.BevelOuter = 1
    TextHost.SelStart = 0
    TextHost.SelLength = 128

End Sub

Sub TextHost_KeyPress (KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        SMTPStart

        KeyAscii = 0
    End If

End Sub

Sub TextHost_LostFocus ()
    PanelHost.BevelOuter = 2

End Sub

Sub TextUserName_GotFocus ()
    
    PanelUserName.BevelOuter = 1
    TextUserName.SelStart = 0
    TextUserName.SelLength = 64

End Sub

Sub TextUserName_KeyPress (KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SMTPStart

        KeyAscii = 0
    End If

End Sub

Sub TextUserName_LostFocus ()
    
    PanelUserName.BevelOuter = 2

End Sub


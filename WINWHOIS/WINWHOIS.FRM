VERSION 2.00
Begin Form Main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "WinWhois"
   ClientHeight    =   3900
   ClientLeft      =   1005
   ClientTop       =   1695
   ClientWidth     =   7365
   Height          =   4560
   Icon            =   WINWHOIS.FRX:0000
   Left            =   960
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   7365
   Top             =   1080
   Width           =   7455
   Begin SSPanel PanelLocalName 
      Align           =   1  'Align Top
      Alignment       =   6  'Center - TOP
      BackColor       =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "LocalHost"
      Font3D          =   4  'Inset w/heavy shading
      Height          =   495
      Left            =   0
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   0
      Width           =   7365
      Begin Label LabelLocalAddr 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "LocalAddr"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   7095
      End
   End
   Begin Ini Ini 
      Left            =   120
      Top             =   0
   End
   Begin NetClient NetClient 
      Left            =   600
      LineDelimiter   =   ""
      RecvSize        =   4096
      RecvThreshold   =   0
      RemoteService   =   ""
      Top             =   0
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
      TabIndex        =   7
      Top             =   3645
      Width           =   7365
   End
   Begin SSFrame FrameResponses 
      Alignment       =   2  'Center
      Font3D          =   2  'Raised w/heavy shading
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   7095
      Begin ListBox ListResponses 
         Height          =   1785
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6855
      End
   End
   Begin ComboBox ComboServers 
      BackColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   4560
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin SSPanel PanelQuery 
      Alignment       =   6  'Center - TOP
      AutoSize        =   3  'AutoSize Child To Panel
      BackColor       =   &H00C0C0C0&
      Font3D          =   4  'Inset w/heavy shading
      Height          =   300
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   3735
      Begin TextBox TextQuery 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   15
         MaxLength       =   128
         TabIndex        =   4
         Text            =   "?"
         Top             =   15
         Width           =   3705
      End
   End
   Begin Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Whois:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin Menu MenuAbout 
      Caption         =   "&About"
   End
End
Option Explicit

Dim igServersCount As Integer

Sub CancelQuery ()

    On Error Resume Next

    TextQuery.Enabled = True
    ComboServers.Enabled = True

    Main.MousePointer = 0
    PanelLocalName.MousePointer = 2
    StatusBar.MousePointer = 0
    
    NetClient.Connect = False

    StatusBar.Caption = "Query canceled."

End Sub

Sub ComboServers_Click ()
Dim sHost As String

    If ComboServers.ListIndex = -1 Then Exit Sub
    sHost = ComboServers.List(ComboServers.ListIndex)
    NetClient.HostName = sHost
    If NetClient.HostName = "" Then
        NetClient.HostAddr = sHost
    End If
    
End Sub

Sub ComboServers_KeyPress (KeyAscii As Integer)

    If KeyAscii = 13 Then
        NetClient.HostName = ComboServers.Text
        If NetClient.HostName = "" Then
            NetClient.HostAddr = ComboServers.Text
            If NetClient.HostAddr = "" Then
                StatusBar = "Unknown host: " + ComboServers.Text
                ComboServers.ListIndex = ComboServers.NewIndex
                Exit Sub
            End If
        End If
        ComboServers.AddItem ComboServers.Text
        ComboServers.ListIndex = ComboServers.NewIndex
        
        Ini.Filename = "winwhois.ini"
        Ini.Section = "Servers"
        Ini.Entry = "Server" + Right$(Str$(igServersCount), Len(Str$(igServersCount)) - 1)
        Ini.Value = ComboServers.Text

        igServersCount = igServersCount + 1

        KeyAscii = 0
    End If

End Sub

Sub Form_Load ()
Dim iIndex As Integer
Dim sEntry As String
Dim sValue As String

    On Error Resume Next

    If NetClient.HostName <> "" Then
        PanelLocalName = NetClient.HostName
        LabelLocalAddr = NetClient.HostAddr
    Else
        PanelLocalName = "LocalHost"
        LabelLocalAddr = "Unknown Address"
    End If

    ' Use "whois" service, or port #43
    NetClient.RemoteService = "whois"
    If NetClient.RemoteService = "" Then
        NetClient.RemotePort = 43
    End If

    ' Use Unix Chr$(10) Newline characters for delimiters
    NetClient.LineDelimiter = "\n"

    ComboServers.Clear
    Ini.Filename = "winwhois.ini"
    Ini.Section = "Servers"
    Ini.Default = ""
    iIndex = 0
     
    Ini.Entry = "Server" + Right$(Str$(iIndex), Len(Str$(iIndex)) - 1)

    sValue = Ini.Value
    While sValue <> ""
        ComboServers.AddItem sValue
        iIndex = iIndex + 1
        Ini.Entry = "Server" + Right$(Str$(iIndex), Len(Str$(iIndex)) - 1)
        sValue = Ini.Value
    Wend

    ' Automatically select the previously used one!
    Ini.Entry = "Selected"
    Ini.Default = "0"
    ComboServers.ListIndex = Val(Ini.Value)
    
    If iIndex > 0 Then ComboServers.ListIndex = 0
    
    igServersCount = iIndex

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
    FrameResponses.Width = Abs(ScaleWidth - 240)
    ListResponses.Width = Abs(ScaleWidth - 480)
    ComboServers.Left = Abs(ScaleWidth - ComboServers.Width - 120)
    PanelQuery.Width = Abs(ScaleWidth - ComboServers.Width - 240 - PanelQuery.Left)
    
    FrameResponses.Height = Abs(ScaleHeight - StatusBar.Height - FrameResponses.Top - 120)
    ListResponses.Height = Abs(FrameResponses.Height - 360)

End Sub

Sub Form_Unload (Cancel As Integer)

    ' Automatically select the previously used one!
    Ini.Entry = "Selected"
    Ini.Value = Right$(Str$(ComboServers.ListIndex), Len(ComboServers.ListIndex) - 1)
    
End Sub

Sub LabelLocalAddr_DblClick ()
    
    PanelLocalName_DblClick

End Sub

Sub MenuAbout_Click ()

    gsMISCAboutLoad "About WinWhois", "WinWhois v1.1", "This client application queries Whois servers (typically port 43). SERVERS.TXT includes a list of all publically known Whois hosts."

End Sub

Sub NetClient_OnClose ()
    
    On Error Resume Next

    TextQuery.Enabled = True
    ComboServers.Enabled = True

    CancelQuery

    StatusBar = "Ok."

End Sub

Sub NetClient_OnConnect ()

    On Error Resume Next

    StatusBar.Caption = "Connected, Sending request - Click this message to Abort."

    NetClient.SendLine = TextQuery.Text + Chr$(13) + Chr$(10)

End Sub

Sub NetClient_OnError (iErrorNumber As Integer)
    
    On Error Resume Next

    CancelQuery

    StatusBar.Caption = "Error" + Str$(iErrorNumber) + ": " + NetClient.ErrorMessage
    
End Sub

Sub NetClient_OnRecv ()
Dim sLine As String

    On Error Resume Next

    sLine = NetClient.RecvLine
    While sLine <> ""
        Do While Right$(sLine, 1) = Chr$(13)
            sLine = Left$(sLine, Len(sLine) - 1)
            If sLine = "" Then Exit Do
        Loop
        ListResponses.AddItem sLine
        sLine = NetClient.RecvLine
    Wend

End Sub

Sub PanelLocalName_DblClick ()
    
    PanelLocalName.BevelOuter = 1
    PanelLocalName.Font3D = 2
    WhoisQuery
    PanelLocalName.Font3D = 1
    PanelLocalName.BevelOuter = 2

End Sub

Sub StatusBar_Click ()

    StatusBar.BevelOuter = 1
    
    CancelQuery
    
    StatusBar.BevelOuter = 2

End Sub

Sub TextQuery_GotFocus ()
    PanelQuery.BevelOuter = 1
    TextQuery.SelStart = 0
    TextQuery.SelLength = 128
End Sub

Sub TextQuery_KeyPress (KeyAscii As Integer)

    If KeyAscii = 13 Then

        WhoisQuery

        KeyAscii = 0
    End If
End Sub

Sub TextQuery_LostFocus ()
    PanelQuery.BevelOuter = 2
End Sub

Sub WhoisQuery ()
    
    On Error Resume Next

    If NetClient.HostName = "" Then
        StatusBar.Caption = "No valid Whois server specified!"
        Exit Sub
    End If
        
    ListResponses.Clear
    ListResponses.SetFocus

    StatusBar.Caption = "Initiating Query - Click this message to Abort."
    
    TextQuery.Enabled = False
    ComboServers.Enabled = False
    
    Main.MousePointer = 11
    PanelLocalName.MousePointer = 12
    StatusBar.MousePointer = 2
    
    NetClient.Connect = True

End Sub


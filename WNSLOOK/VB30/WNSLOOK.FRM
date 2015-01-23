VERSION 2.00
Begin Form Main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "WiNslookup"
   ClientHeight    =   3465
   ClientLeft      =   975
   ClientTop       =   1680
   ClientWidth     =   5775
   Height          =   4125
   Icon            =   WNSLOOK.FRX:0000
   Left            =   930
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   5775
   Top             =   1065
   Width           =   5865
   Begin Ini Ini 
      Left            =   5280
      Top             =   1320
   End
   Begin SSFrame FrameAliases 
      Alignment       =   2  'Center
      Caption         =   "Aliases"
      Font3D          =   1  'Raised w/light shading
      Height          =   1695
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   3135
      Begin ListBox ListAliases 
         Height          =   1395
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
   End
   Begin SSFrame FrameAddresses 
      Alignment       =   2  'Center
      Caption         =   "Addresses"
      Font3D          =   1  'Raised w/light shading
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
      Begin ListBox ListAddresses 
         Height          =   1395
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin SSPanel PanelStatus 
      Align           =   2  'Align Bottom
      Alignment       =   1  'Left Justify - MIDDLE
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ok."
      Font3D          =   3  'Inset w/light shading
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3210
      Width           =   5775
   End
   Begin SSPanel PanelName 
      Align           =   1  'Align Top
      Alignment       =   6  'Center - TOP
      BackColor       =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "PanelName"
      Font3D          =   4  'Inset w/heavy shading
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5775
      Begin Label LabelAddr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LabelAddr"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5535
      End
   End
   Begin SSPanel PanelHost 
      AutoSize        =   3  'AutoSize Child To Panel
      BackColor       =   &H00C0C0C0&
      Font3D          =   0  'None
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   5055
      Begin TextBox TextHost 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   15
         MaxLength       =   128
         TabIndex        =   1
         Top             =   15
         Width           =   5025
      End
   End
   Begin NetClient NetClient 
      Left            =   5280
      LineDelimiter   =   ""
      RecvSize        =   4096
      RecvThreshold   =   0
      Top             =   840
   End
   Begin Label LabelPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter I.P. or DNS name:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin Menu MenuAbout 
      Caption         =   "&About"
   End
End
Option Explicit

Const ERR_INVALIDHOSTADDR = 106

Sub Form_Load ()
    If NetClient.HostName <> "" Then
        PanelName = NetClient.HostName
        LabelAddr = NetClient.HostAddr
    Else
        PanelName = "LocalHost"
        LabelAddr = "name unknown"
    End If

    Ini.Filename = "wnslook.ini"
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
Dim iTemp As Integer

    If Main.WindowState = 1 Then Exit Sub
    
    LabelAddr.Width = ScaleWidth

    PanelHost.Width = Abs(ScaleWidth - 240)
    iTemp = Abs(ScaleWidth - FrameAliases.Left - 120)
    ListAliases.Width = Abs(iTemp - 240)
    FrameAliases.Width = iTemp

    iTemp = Abs(ScaleHeight - FrameAliases.Top - PanelStatus.Height - 120)
    FrameAddresses.Height = iTemp
    FrameAliases.Height = iTemp
    ListAddresses.Height = Abs(iTemp - 360)
    ListAliases.Height = Abs(iTemp - 360)

End Sub

Sub MenuAbout_Click ()

    gsMISCAboutLoad "WiNSLookup", "WiNSLookup v1.1", "This application translated I.P addresses to DNS names (and vise versa). No sockets are created, your WinSock.DLL handles the Name Server queries."

End Sub

Sub NetClient_OnError (iErrorNumber As Integer)
    If iErrorNumber = ERR_INVALIDHOSTADDR Then Exit Sub
    PanelStatus = NetClient.ErrorMessage
End Sub

Sub TextHost_GotFocus ()
    PanelHost.BevelOuter = 1
    TextHost.SelStart = 0
    TextHost.SelLength = 128
End Sub

Sub TextHost_KeyPress (KeyAscii As Integer)
Dim iIndex As Integer

    On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        TextHost.Enabled = False
        ListAddresses.Clear
        ListAliases.Clear
        PanelStatus = "Querying HostAddr"
        NetClient.HostAddr = TextHost
        If NetClient.HostAddr = "" Then
            PanelStatus = "Querying HostName"
            NetClient.HostName = TextHost
            If NetClient.HostName = "" Then
                PanelStatus = "Sorry, I couldn't find that address"
                TextHost.Enabled = True
                Exit Sub
            End If
        End If

        ListAddresses.AddItem NetClient.HostAddr
        If NetClient.HostAddressCount > 1 Then
            For iIndex = 1 To NetClient.HostAddressCount - 1
                ListAddresses.AddItem NetClient.HostAddressList(iIndex)
            Next
        End If
        ListAliases.AddItem NetClient.HostName
        If NetClient.HostAliasCount > 0 Then
            For iIndex = 1 To NetClient.HostAliasCount
                ListAliases.AddItem NetClient.HostAliasList(iIndex)
            Next
        End If
        TextHost.Enabled = True
        TextHost.SetFocus
        TextHost.SelStart = 0
        TextHost.SelLength = 128
        PanelStatus = "Ok."
    End If
End Sub

Sub TextHost_LostFocus ()
    PanelHost.BevelOuter = 2
End Sub


VERSION 2.00
Begin Form Main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "WiNslookup - Sample of WSANET.VBX"
   Height          =   3645
   Icon            =   WNSLOOK.FRX:0000
   Left            =   1065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   5280
   Top             =   1560
   Width           =   5370
   Begin ListBox ListAliases 
      Height          =   1368
      Left            =   1680
      TabIndex        =   4
      Top             =   1440
      Width           =   3492
   End
   Begin ListBox ListAddresses 
      Height          =   1368
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1452
   End
   Begin TextBox TextHost 
      Height          =   288
      Left            =   120
      MaxLength       =   128
      TabIndex        =   1
      Top             =   840
      Width           =   5052
   End
   Begin NetClient NetClient 
      Left            =   4560
      LineDelimiter   =   ""
      RecvSize        =   4096
      RecvThreshold   =   0
      Top             =   360
   End
   Begin Label LabelStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Ok."
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   5052
   End
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
      Height          =   252
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   5292
   End
   Begin Label LabelAliases 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aliases"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   1680
      TabIndex        =   6
      Top             =   1200
      Width           =   3492
   End
   Begin Label LabelAddresses 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Addresses"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1452
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
      Height          =   192
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2052
   End
   Begin Label LabelHost 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LabelHost"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   264
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5292
   End
End
Option Explicit


Sub Form_Load ()
    
    On Error Resume Next

    If NetClient.HostName <> "" Then
        LabelHost = NetClient.HostName
        LabelAddr = NetClient.HostAddr
    Else
        LabelHost = "LocalHost"
        LabelAddr = "name unknown"
    End If
End Sub

Sub Form_Resize ()
Dim iTemp As Integer

    On Error Resume Next

    If Main.WindowState = 1 Then Exit Sub
    
    LabelHost.Width = ScaleWidth
    LabelAddr.Width = ScaleWidth

    TextHost.Width = Abs(ScaleWidth - 240)
    iTemp = Abs(ScaleWidth - ListAliases.Left - 120)
    ListAliases.Width = iTemp
    LabelAliases.Width = iTemp

    LabelStatus = Abs(ScaleWidth - 240)
    
    iTemp = Abs(ScaleHeight - ListAliases.Top - 492)
    ListAddresses.Height = iTemp
    ListAliases.Height = iTemp

    LabelStatus.Top = Abs(ListAddresses.Top + iTemp + 120)
End Sub

Sub NetClient_OnError (iErrorNumber As Integer)
    
    On Error Resume Next

    LabelStatus = NetClient.ErrorMessage
End Sub

Sub TextHost_KeyPress (KeyAscii As Integer)
Dim iIndex As Integer

    On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        TextHost.Enabled = False
        ListAddresses.Clear
        ListAliases.Clear
        LabelStatus = "Querying HostAddr"
        NetClient.HostAddr = TextHost
        If NetClient.HostAddr = "" Then
            LabelStatus = "Querying HostName"
            NetClient.HostName = TextHost
            If NetClient.HostName = "" Then
                LabelStatus = "Sorry, I couldn't find that address"
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
    End If
End Sub


VERSION 2.00
Begin Form Main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "WFinger - Test of WSNetC"
   Height          =   5865
   Icon            =   WFINGER.FRX:0000
   Left            =   930
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   6930
   Top             =   1080
   Width           =   7020
   Begin TextBox TextHost 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "rhino.ess.harris.com"
      Top             =   495
      Width           =   5010
   End
   Begin TextBox TextUsername 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Text            =   "iblenke"
      Top             =   495
      Width           =   1185
   End
   Begin CommandButton BtnFinger 
      Caption         =   "&Finger"
      Height          =   375
      Left            =   165
      TabIndex        =   1
      Top             =   840
      Width           =   1200
   End
   Begin NetClient NetClient 
      Left            =   6420
      LineDelimiter   =   ""
      RecvSize        =   4096
      RecvThreshold   =   0
      Top             =   1380
   End
   Begin TextBox TextDisplay 
      Height          =   4095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Empty."
      Top             =   1320
      Width           =   6375
   End
   Begin Label LabelHost 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "LabelHost"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   390
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   6780
   End
   Begin Label LabelErrors 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "No Errors."
      Height          =   375
      Left            =   1815
      TabIndex        =   5
      Top             =   855
      Width           =   4935
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "@"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   510
      Width           =   255
   End
End
Dim CRLF As String

Sub BtnFinger_Click ()

    On Error Resume Next

    LabelErrors.Caption = "No Errors."
    
    TextDisplay.Text = ""

    ' WSNetC code
    NetClient.HostName = TextHost
    If NetClient.HostName = "" Then
        NetClient.HostAddr = TextHost
        If NetClient.HostName = "" Then
            MsgBox "Sorry, " + TextHost + " could not be resolved.", 48
            Exit Sub
        End If
    End If
    NetClient.Connect = True
    
    ' We are completing a query, ignore user
    TextDisplay.SetFocus
    BtnFinger.Enabled = False
    Main.MousePointer = 11
End Sub

Sub Form_Load ()
    CRLF = Chr$(13) + Chr$(10)

    LabelHost = NetClient.HostName

    ' Use "finger" service (or port 79)
    NetClient.RemoteService = "finger"
    If NetClient.RemotePort = 0 Then
        NetClient.RemotePort = 79
    End If
End Sub

Sub Form_Resize ()
    
    If Main.WindowState = 1 Then Exit Sub

    TextDisplay.Width = ScaleWidth
    TextDisplay.Height = Abs(ScaleHeight - TextDisplay.Top)

    TextHost.Width = Abs(ScaleWidth - TextHost.Left - 120)
    LabelErrors.Width = Abs(ScaleWidth - LabelErrors.Left - 120)

End Sub

Sub NetClient_OnClose ()
    
    On Error Resume Next

    NetClient.Connect = False

    LabelErrors.Caption = "Connection closed."
    BtnFinger.Enabled = True
    Main.MousePointer = 0

End Sub

Sub NetClient_OnConnect ()
    
    On Error Resume Next
    
    LabelErrors.Caption = "Connected to port " + Str$(NetClient.RemotePort)
    NetClient.SendLine = TextUsername.Text + CRLF

End Sub

Sub NetClient_OnError (iErrorNumber As Integer)
Dim ErrorMessage As String
Dim ErrorNumberText As String

    On Error Resume Next

    ErrorMessage = NetClient.ErrorMessage
    ErrorNumberText = Str$(iErrorNumber)

    LabelErrors.Caption = "Error #" + ErrorNumberText + ": " + ErrorMessage
    
    ' Disconnect for good measure
    NetClient.Connect = False
    BtnFinger.Enabled = True
    Main.MousePointer = 0

End Sub

Sub NetClient_OnRecv ()
Dim sInput As String

    On Error Resume Next

    ' Loop until buffer is empty (this doesn't block)
    Do While Not NetClient.RecvCount = 0
        sInput = NetClient.RecvLine
        TextDisplay.Text = TextDisplay.Text + sInput
        DoEvents
    Loop

End Sub


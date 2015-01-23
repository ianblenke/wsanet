VERSION 2.00
Begin Form FormSMTPClient 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SMTP Client - Test of WSNetC"
   Height          =   4545
   Icon            =   FORMSMTP.FRX:0000
   Left            =   960
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   7890
   Top             =   1065
   Width           =   7980
   Begin NetClient NetClient 
      Left            =   7080
      LineDelimiter   =   ""
      RecvSize        =   4096
      RecvThreshold   =   0
      Top             =   360
   End
   Begin CommandButton BtnSend 
      Caption         =   "&Send Test"
      Height          =   330
      Left            =   5460
      TabIndex        =   3
      Top             =   390
      Width           =   1455
   End
   Begin TextBox TextUser 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   360
      Left            =   225
      TabIndex        =   2
      Text            =   "iblenke"
      Top             =   375
      Width           =   1590
   End
   Begin TextBox TextHost 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   360
      Left            =   2265
      TabIndex        =   1
      Text            =   "rhino.ess.harris.com"
      Top             =   375
      Width           =   2955
   End
   Begin ListBox ListProgress 
      Height          =   3150
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   7815
   End
   Begin Label LabelHost 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Localhost"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   15
      TabIndex        =   5
      Top             =   0
      Width           =   8310
   End
   Begin Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   1875
      TabIndex        =   4
      Top             =   420
      Width           =   375
   End
End
Dim LocalHostName As String
Dim CRLF As String

Const STATE_INACTIVE = 0
Const STATE_HELO = 1
Const STATE_MAILFROM = 2
Const STATE_RCPTTO = 3
Const STATE_DATA = 4
Const STATE_SENDBODY = 5
Const STATE_QUIT = 6

Dim SMTPState As Integer

Sub BtnSend_Click ()
    
    On Error Resume Next

    NetClient.HostName = TextHost.Text
    If NetClient.HostName = "" Then
        ListProgress.AddItem "Host " + TextHost.Text + " unknown."
        Exit Sub
    End If
    
    ' Use SMTP service (port 25)
    NetClient.RemoteService = "smtp"
    If NetClient.RemotePort = 0 Then
        NetClient.RemotePort = 25
    End If

    CRLF = Chr$(13) + Chr$(10)
    NetClient.LineDelimiter = CRLF

    NetClient.Connect = True
    
    ListProgress.Clear
    ListProgress.AddItem "Connecting to " + TextHost.Text + " via SMTP Port 25"

    SMTPState = STATE_HELO

    BtnSend.Enabled = False

End Sub

Sub Form_Load ()
    
    On Error Resume Next

    LocalHostName = NetClient.HostName
    LabelHost = LocalHostName
End Sub

Sub Form_Resize ()

    On Error Resume Next

    If FormSMTPClient.WindowState = 1 Then Exit Sub

    ListProgress.Width = ScaleWidth
    ListProgress.Height = Abs(ScaleHeight - ListProgress.Top)

    LabelHost.Width = ScaleWidth
End Sub

Sub NetClient_OnClose ()
    
    On Error Resume Next

    SMTPState = STATE_INACTIVE
    BtnSend.Enabled = True

End Sub

Sub NetClient_OnConnect ()

    On Error Resume Next

    ReportProgress "Connected to " + NetClient.HostName
    
    SMTPState = STATE_MAILFROM

End Sub

Sub NetClient_OnError (ErrorNumber As Integer)
Dim sTemp As String

    On Error Resume Next

    sTemp = NetClient.ErrorMessage
    ReportProgress sTemp
End Sub

Sub NetClient_OnRecv ()
Dim sLine As String
Dim iReturn As Integer

    On Error Resume Next

    sLine = NetClient
    
    Do While sLine <> ""
        ReportProgress ">" + sLine

        'Strip out SMTP multi-line replies
        While Mid$(sLine, 4, 1) = "-"
            sLine = NetClient
        Wend

        ' Get the SMTP reply number
        iReturn = Val(Left$(sLine, 3))

        Select Case iReturn
            Case 200 To 299:    ' Ok replies
                Select Case SMTPState
                    Case STATE_HELO:
                        sLine = "HELO " + LocalHostName + CRLF
                        NetClient = sLine
                        ReportProgress Left$(sLine, Len(sLine) - 2)
                    Case STATE_MAILFROM:
                        sLine = "MAIL FROM: <" + TextUser + "@" + LocalHostName + ">" + CRLF
                        NetClient = sLine
                        ReportProgress Left$(sLine, Len(sLine) - 2)
                        SMTPState = STATE_RCPTTO
                    Case STATE_RCPTTO:
                        sLine = "RCPT TO: <" + TextUser + "@" + TextHost + ">" + CRLF
                        NetClient = sLine
                        ReportProgress Left$(sLine, Len(sLine) - 2)
                        SMTPState = STATE_DATA
                    Case STATE_DATA
                        sLine = "DATA" + CRLF
                        NetClient = sLine
                        ReportProgress Left$(sLine, Len(sLine) - 2)
                        SMTPState = STATE_SENDBODY
                    Case STATE_QUIT
                        sLine = "QUIT" + CRLF
                        NetClient = sLine
                        ReportProgress Left$(sLine, Len(sLine) - 2)
                        SMTPState = STATE_INACTIVE
                    Case STATE_HELO:
                        NetClient.Connect = False
                        BtnSend.Enabled = True
                End Select
            Case 300 To 399:    ' Informational replies
                Select Case SMTPState
                    Case STATE_SENDBODY:
                        sLine = "This is a test of the SMTP client that comes with WSANET."
                        ReportProgress sLine
                        sLine = sLine + CRLF + "." + CRLF
                        NetClient = sLine
                        ReportProgress "."
                        SMTPState = STATE_QUIT
                End Select
            Case 500 To 599:
                ReportProgress "500 error! Abort! Abort!"
                Select Case SMTPState
                    Case STATE_SENDBODY:
                        NetClient = "." + CRLF
                        SMTPState = STATE_QUIT
                    Case Else
                        NetClient = "QUIT" + CRLF
                End Select
            Case Else
                ReportProgress "Unknown reply #" + Str$(iReturn) + "0!"
        End Select
        DoEvents
        sLine = NetClient
    Loop
    
End Sub

Sub ReportProgress (sMessage As String)
    ListProgress.AddItem sMessage
End Sub


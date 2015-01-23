Option Explicit
'                           VB-SMTPD
' An SMTP daemon state server written with the WSANET.VBX control.

Const VERSION_BANNER = "1.0/VB-SMTPSRV"

' Define our SMTP state structure
Type tSMTPSTATE
    iState As Integer
    ' Command state information
    sHelo As String
    sMailFrom As String
    sRcptTo As String
    iMailType As Integer
    ' Message state information
    iFileHandle As Integer
    sFilename As String
End Type

' A connection can only be in 1 of 3 states:
Const SMTP_CLOSEDSTATE = 0
Const SMTP_COMMANDSTATE = 1
Const SMTP_DATASTATE = 2

' Mark any odd handshakes!
Const MAIL_MAIL = 0
Const MAIL_SEND = 1
Const MAIL_SOML = 2
Const MAIL_SAML = 3
Const MAIL_EHLO = 4
Const MAIL_8BITMIME = 8

' Hold the actual dynamically created controls &
' the number OF them.
Global giIndex As Integer
Global gSMTPState() As tSMTPSTATE

' These are filled in at the start
Global gLocalHostName As String
Global CRLF As String

' These are used when logging
Global Const LOG_NOTHING = 0
Global Const LOG_ERROR = 1
Global Const LOG_WARNING = 2
Global Const LOG_SEND = 4
Global Const LOG_RECV = 8
Global Const LOG_SUMMARY = 16
Global Const LOG_NORMAL = 32
Global Const LOG_DEBUG = 64

Global giLogFlags As Integer
Global giLogSize As Integer

Global msMailFile As String

'WinAPI Declarations
Declare Function GetTempFileName Lib "Kernel" (ByVal cDriveLetter As Integer, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer
Declare Function GetTempDrive Lib "Kernel" (ByVal cDriveLetter As Integer) As Integer

Global Const OFN_FILEMUSTEXIST = &H1000&
Global Const OFN_PATHMUSTEXIST = &H800&
Global Const OFN_READONLY = &H1&

Private Function Date2HEX (d#) As String
Dim seconds#

    ' Stolen from a VB magazine (which I can't remember the name of)
    seconds# = (d# - 1) * 86400

    If seconds# > 4294967294# Then
        Date2HEX = ""
    ElseIf seconds# < 2147483647# Then
        Date2HEX = Hex$(seconds# + 1)
    Else
        Date2HEX = Hex$(seconds# - 4294967295#)
    End If


End Function

Sub gSMTPAccept (Socket As Integer, PeerName As String, RemotePort As Integer)
Dim sDateTime As String
Dim Index As Integer
Dim cNetClient As NetClient

    On Error GoTo AcceptError

    gSMTPLog LOG_DEBUG, "gSMTPAccept(" & Str$(Socket) & ", " & PeerName & "," & Str$(RemotePort) & ")"

    gSMTPSocketOpen Index

    If Index = 0 Then
        gSMTPLog LOG_ERROR, "gSMTPSocketOpen() Failed!"
        Socket = 0
        Exit Sub
    End If

    Main!NetClient(Index).Socket = Socket
    
    ' If for some reason, there is an error, say goodbye!
    If Err <> 0 Then
        gSMTPSend Index, 421, "Service unavailable. Closing transmission channel."
        gSMTPLog LOG_ERROR, "Connection from [" & PeerName & "] refused!"
        gSMTPLog LOG_DEBUG, "Err =" & Str$(Err) & ": " & Error$
        Socket = 0
        Exit Sub
    End If

    gSMTPLog LOG_NORMAL, "Connection from:" & Main!NetClient(Index).HostName & "(" & PeerName & ")"

    Main!NetClient(Index).LineDelimiter = CRLF
    
    gSMTPState(Index).iState = SMTP_COMMANDSTATE
    gSMTPState(Index).sHelo = ""
    gSMTPState(Index).sMailFrom = ""
    gSMTPState(Index).sRcptTo = ""

    'Tue, 14 Dec 93 10:14:13 PST
    sDateTime = Format$(Now, "ddd, d mmm yy hh:nn:ss")

    gSMTPSend giIndex, 220, gLocalHostName & " SendMail v1.00/VB-WSMTPD ready at " & sDateTime

    Exit Sub

AcceptError:
    gSMTPError Erl, Err, Error$
    Resume Next

End Sub

Sub gSMTPCleanup ()
    
    ' Close all connections
    While giIndex > 0
        gSMTPSend giIndex, 421, "Service shutting down"
        gSMTPSocketClose giIndex
    Wend

    Main!Ini.Section = "Configuration"

    gSMTPLogPutFlag "LogError", LOG_ERROR
    gSMTPLogPutFlag "LogWarning", LOG_WARNING
    gSMTPLogPutFlag "LogSend", LOG_SEND
    gSMTPLogPutFlag "LogRecv", LOG_RECV
    gSMTPLogPutFlag "LogSummary", LOG_SUMMARY
    gSMTPLogPutFlag "LogNormal", LOG_NORMAL
    gSMTPLogPutFlag "LogDebug", LOG_DEBUG

End Sub

Sub gSMTPClose (Index As Integer)
Dim cNetClient As NetClient

    gSMTPLog LOG_WARNING, "Lost connection to " & Main!NetClient(Index).HostName & "!"
    
    gSMTPSocketClose Index

End Sub

Function gSMTPDoHelo (Index As Integer, sParameters As String) As String
    
    gSMTPState(Index).sHelo = sParameters
    If sParameters <> "" Then
        If sParameters <> Main!NetClient(Index).HostName Then
            gSMTPDoHelo = "Hello " & Main!NetClient(Index).HostName & ", why do you call yourself " & sParameters & "?"
        Else
            gSMTPDoHelo = "Hello " & sParameters & ", glad to meet you."
        End If
    Else
        gSMTPDoHelo = "Thanks for the help, bub!"
    End If

End Function

Function gSMTPDoHelp (Index As Integer, sParameters As String) As String
Dim sHelp As String
Dim sTemp As String
Dim sSection As String
Dim iLine As Integer
Dim sLine As String

    
    gSMTPLog LOG_WARNING, "SMTP session requested HELP on '" & sParameters & "'"
    
    If sParameters = "" Then sParameters = "HELP"
    
    ' Try the [Help] section mappings first
    Main!Ini.Section = "Help"
    Main!Ini.Entry = sParameters

    If Main!Ini.Value = "" Then
        gSMTPLog LOG_DEBUG, "Requested [HELP] has no mapping '" & sParameters & "= {}'"
        sSection = sParameters
    Else
        sSection = Main!Ini.Value
    End If

    ' Set the section of the Help topic to use
    Main!Ini.Section = sSection

    ' Loop until no more entries are found
    iLine = 1
    sHelp = ""
    Do
        sLine = Str$(iLine)
        sLine = Right$(sLine, Len(sLine) - 1)
        Main!Ini.Entry = "Line" & sLine
        sLine = Main!Ini.Value
        If sLine <> "" Then
            sHelp = sHelp & sLine & "+"
        End If
        iLine = iLine + 1
    Loop While sLine <> ""

    If sHelp <> "" Then
        gSMTPDoHelp = Left$(sHelp, Len(sHelp) - 1)
    Else
        gSMTPDoHelp = sHelp
    End If

End Function

Sub gSMTPDoMail (Index As Integer, sParameters As String)
Dim iPlace As Integer
Dim sTemp As String
Dim sExtended As String
    
    sTemp = UCase$(sParameters)

    If Left$(sTemp, 5) = "FROM:" Then
        If gSMTPState(Index).sMailFrom = "" Then
            sTemp = Right$(sParameters, Len(sParameters) - 5)
            ' If E/SMTP is in action
            If gSMTPState(Index).iMailType And MAIL_EHLO Then
                ' Check for more after the address
                iPlace = InStr(sTemp, "> ")
                If iPlace > 0 Then
                    sExtended = Trim$(Right$(sTemp, Len(sTemp) - iPlace - 1))
                    While sExtended <> ""
                        gSMTPDoMailExt Index, sExtended
                    Wend
                End If
            End If
            
            gSMTPState(Index).sMailFrom = sTemp
            gSMTPSend Index, 250, sParameters & " Sender Ok."
        Else
            gSMTPSend Index, 503, sTemp & "Sender already specified!"
        End If
    Else
        gSMTPSend Index, 550, "MAIL '" & sParameters & "' Format bad."
    End If

End Sub

Sub gSMTPDoMailExt (Index As Integer, sExtended As String)
Dim iPlace As String
Dim sKeyWord As String
Dim sTemp As String


    ' Don't do any E/SMTP commands yet

    sExtended = ""
    Exit Sub
    
    ' Put it into parseable form
'    sExtended = UCase$(sExtended)
    
    ' Recurse down until the end is found
'    iPlace = InStr(sExtended, ",")
'    If iPlace > 0 Then
'        Do While iPlace > 0
'            sTemp = Right$(sExtended, Len(sExtended) - iPlace - 1)
'            gSMTPDoMailExt Index, sTemp
'            sExtended = Left$(sExtended, iPlace - 1)
'            iPlace = InStr(sExtended, ",")
'        Loop
'    End If

    ' Rip out the extended portion
'    iPlace = InStr(sTemp, "=")
'    If iPlace > 0 Then
        ' Now check the left side
'        sTemp = Trim$(Left$(sExtended, iPlace - 1))
'        Select Case sTemp
'            Case "BODY"
'                sTemp = Trim$(Right$(sExtended, Len(sExtended) - iPlace - 1))
'                If sTemp = "8BITMIME" Then
'                    gSMTPState(Index).iMailType = gSMTPState(Index).iMailType Or MAIL_8BITMIME
'                End If
'             Case "SIZE"
                ' Not Implemented yet
'            Case Else
'        End Select

'    sExtended = ""

End Sub

Sub gSMTPDoRcpt (Index As Integer, sParameters As String)
Dim iPlace As Integer
Dim sTemp As String
    
    sTemp = UCase$(sParameters)

    If Left$(sTemp, 3) = "TO:" Then
        sTemp = Right$(sParameters, Len(sParameters) - 3)
        If gSMTPState(Index).sRcptTo = "" Then
            gSMTPState(Index).sRcptTo = sTemp
        Else
            gSMTPState(Index).sRcptTo = gSMTPState(Index).sRcptTo & "," & sTemp
        End If
        gSMTPSend Index, 250, sTemp & "Recipient Ok."
    Else
        gSMTPSend Index, 501, "Syntax error in parameters: 'RCPT " & sParameters & "'"
    End If

End Sub

Sub gSMTPError (iLine As Integer, iNumber As Integer, sString As String)
' This subroutine logs all SMTP and VB errors

    Select Case iLine
        Case 0:
            If iNumber = 0 Then
                gSMTPLog LOG_WARNING, sString
            Else
                gSMTPLog LOG_WARNING, "VB Warning" + Str$(iNumber) + ": " & sString
            End If
        
        Case -1:
            gSMTPLog LOG_WARNING, "WSA Error:" & Str$(iNumber) & " " & sString

        Case Else:
            gSMTPLog LOG_ERROR, "VB Error:" & Str$(iNumber) & " on line" & Str$(iLine) & ": " & sString
    
    End Select

End Sub

Sub gSMTPFileClose (Index As Integer)
Dim iFileHandle As Integer
Dim sTemp As String
Dim i%
Dim sReader As String
Dim hIn As Integer
Dim hOut As Integer
Dim lSize As Long
Dim iRead As Integer

    On Error GoTo FileCloseError

    iFileHandle = gSMTPState(Index).iFileHandle
    Close #iFileHandle
    gSMTPState(Index).iFileHandle = 0
    
    If msMailFile <> "" Then
        hIn = FreeFile
        Open gSMTPState(Index).sFilename For Input As #hIn
        hOut = FreeFile
        Open msMailFile For Append As #hOut
        lSize = LOF(hIn)
        Do While Not EOF(hIn)
            If lSize = 0 Then Exit Do
            If lSize > 16384 Then
                iRead = 16384
            Else
                iRead = lSize
            End If
            
            sTemp = Input$(iRead, hIn)
            Print #hOut, sTemp;
            lSize = lSize - iRead
        Loop
        Close #hIn
        Close #hOut

    End If
    
    Main!Ini.Section = "Configuration"
    Main!Ini.Entry = "Reader"
    sReader = Main!Ini.Value
    If sReader <> "" Then
        Main!Ini.Entry = "UniqueReader"
        sTemp = Main!Ini.Value
        If sTemp = "1" Or Left$(sTemp, 1) = "T" Or Right$(sTemp, 1) = "N" Then
            If msMailFile = "" Then
                sTemp = sReader
            Else
                sTemp = sReader & " " & msMailFile
            End If
            i% = Shell(sTemp)
        Else
            sTemp = sReader & " " & gSMTPState(Index).sFilename
            i% = Shell(sTemp, 1)
        End If
    End If

FileCloseError:
    Resume Next
End Sub

Sub gSMTPFileOpen (Index As Integer)
Dim hFile As Integer
Dim sIdentifier As String
Dim sHostName As String
Dim sHostAddr As String
Dim sMsgId As String
Dim sDate As String
Dim sTempFile As String
Dim iTemp%
Dim sHelo As String
Dim sRcptTo As String
Dim sMailFrom As String
Dim sMsgLine As String
Dim sTemp As String

    On Error GoTo FileOpenError

    hFile = FreeFile
    If Err Then Exit Sub
    
    ' Wed, 15 Dec 93 21:48:00
    sDate = Format$(Now, "ddd, d mmm yy hh:nn:ss")
    
    ' Create a temporary file based on the date AND time
    sMsgId = Date2HEX(Now)
    iTemp% = GetTempDrive(iTemp%)
    sTempFile = String$(144, 0)
    iTemp% = GetTempFileName(iTemp%, sMsgId, 0, sTempFile)
    iTemp% = InStr(sTempFile, Chr$(0))
    If iTemp% <> 0 Then
        sTempFile = Left$(sTempFile, iTemp% - 1)
    End If

    ' Store the filehandle and filename for later
    gSMTPState(Index).iFileHandle = hFile
    gSMTPState(Index).sFilename = sTempFile
    
    ' Who are we talking to?
    sHostName = Main!NetClient(Index).HostName
    sHostAddr = Main!NetClient(Index).HostAddr

    ' Get the handshake info back to build the header
    sHelo = gSMTPState(Index).sHelo
    sRcptTo = gSMTPState(Index).sRcptTo
    sMailFrom = gSMTPState(Index).sMailFrom

    ' Open the tempory file for output
    Open sTempFile For Output As #hFile
    
    ' Make up a version string depending on the handshaking
    If gSMTPState(Index).iMailType And MAIL_8BITMIME Then
        sIdentifier = "(" & VERSION_BANNER & "/8BITMIME)"
    Else
        sIdentifier = "(" & VERSION_BANNER & ")"
    End If

    Print #hFile,
    
    sTemp = Mid$(sMailFrom, 2, Len(sMailFrom) - 2)
    sMsgLine = "From " & sTemp & " " & sDate
    gSMTPLog LOG_SUMMARY, sMsgLine
    Print #hFile, sMsgLine
    
    sMsgLine = "X-Envelope-To: " & sRcptTo
    gSMTPLog LOG_SUMMARY, sMsgLine
    Print #hFile, sMsgLine
    
    sMsgLine = "Return-Path: " & sMailFrom
    gSMTPLog LOG_SUMMARY, sMsgLine
    Print #hFile, sMsgLine
    
    sMsgLine = "Received: from " & sHostName & " [" & sHostAddr & "] by " & gLocalHostName
    gSMTPLog LOG_SUMMARY, sMsgLine
    Print #hFile, sMsgLine
    
    sMsgLine = Chr$(9) & "id " & sMsgId & " " & sDate
    gSMTPLog LOG_SUMMARY, sMsgLine
    Print #hFile, sMsgLine
    
    Exit Sub

FileOpenError:
    gSMTPError Erl, Err, Error$
    Resume Next

End Sub

Sub gSMTPFileWrite (Index As Integer, sString As String)
Dim iFileHandle As Integer

    iFileHandle = gSMTPState(Index).iFileHandle
    Print #iFileHandle, sString

End Sub

Sub gSMTPLog (iLevel As Integer, sString As String)
' This routine logs incoming conversations

    Debug.Print "Level" & Str$(iLevel) & ": " & sString

    Select Case iLevel
        Case LOG_NOTHING:
        
        Case LOG_ERROR:
            If giLogFlags And LOG_ERROR Then
                gSMTPLogWrite "Error: " & sString
            End If
        
        Case LOG_WARNING:
            If giLogFlags And LOG_WARNING Then
                gSMTPLogWrite "Warning: " & sString
            End If
        
        Case LOG_SEND:
            If giLogFlags And LOG_SEND Then
                gSMTPLogWrite "Send: " & sString
            End If
        
        Case LOG_RECV:
            If giLogFlags And LOG_RECV Then
                gSMTPLogWrite "Recv: " & sString
            End If
        
        Case LOG_SUMMARY:
            If giLogFlags And LOG_SUMMARY Then
                gSMTPLogWrite "Summary: " & sString
            End If

        Case LOG_NORMAL:
            gSMTPLogWrite sString

        Case LOG_DEBUG:
            If giLogFlags And LOG_DEBUG Then
                gSMTPLogWrite "Debug: " & sString
            End If
    End Select

End Sub

Sub gSMTPLogChkFlag (sEntry As String, iValue As Integer)
Dim sTemp As String

    Main!Ini.Entry = sEntry
    sTemp = Main!Ini.Value
    If sTemp = "" Then
        gSMTPLogSetFlags iValue
        Exit Sub
    End If
    
    If (Left$(sTemp, 1) = "1") Or (Left$(sTemp, 1) = "T") Then
        gSMTPLogSetFlags iValue
    End If
End Sub

Function gSMTPLogGetFlags () As Integer
    
    gSMTPLogGetFlags = giLogFlags

End Function

Sub gSMTPLogPutFlag (sEntry As String, iValue As Integer)

    Main!Ini.Entry = sEntry
    
    If giLogFlags And iValue Then
        Main!Ini.Value = "True"
    Else
        Main!Ini.Value = "False"
    End If

End Sub

Sub gSMTPLogSetFlags (iLevel As Integer)

    If iLevel = LOG_NOTHING Then
        giLogFlags = LOG_NOTHING
    Else
        giLogFlags = giLogFlags Or iLevel
    End If

End Sub

Sub gSMTPLogWrite (sString As String)
Dim iTemp As Integer

    On Error GoTo LogWriteError

    Main!ListHosts.AddItem sString
    
    ' Auto scroll the window
    iTemp = Main!ListHosts.ListCount
    If iTemp = (giLogSize + 2) Then
        Main!ListHosts.RemoveItem 0
        Main!ListHosts.ListIndex = giLogSize
    Else
        If iTemp > 0 Then Main!ListHosts.ListIndex = iTemp - 1
    End If

    Exit Sub

LogWriteError:
    Main!ListHosts.Clear
    gSMTPLog LOG_WARNING, "Window full, listbox cleared"
    Resume Next

End Sub

Sub gSMTPParse (Index As Integer, sCommand As String, sParameters As String)
Dim sHelo As String
    
    Select Case sCommand
        Case "EHLO":
            ' Extended SMTP (E/SMTP) identifier
            sHelo = gSMTPDoHelo(Index, sParameters)
            sHelo = sHelo & " Extended SMTP:+HELP+VERB+ONEX+MULT+EXPN+TICK+XWIN8"
            gSMTPSend Index, 250, sHelo
            gSMTPState(Index).iMailType = gSMTPState(Index).iMailType Or MAIL_EHLO

        Case "HELO":
            ' The remote machine says HELlO
            sHelo = gSMTPDoHelo(Index, sParameters)
            gSMTPSend Index, 250, sHelo

        Case "SEND":
            ' Send to screen
            gSMTPState(Index).iMailType = gSMTPState(Index).iMailType Or MAIL_SEND
            gSMTPDoMail Index, sParameters

        Case "SAML":
            ' Send to screen and mailbox
            gSMTPState(Index).iMailType = gSMTPState(Index).iMailType Or MAIL_SAML
            gSMTPDoMail Index, sParameters
        
        Case "SOML":
            ' Send to screen or mailbox
            gSMTPState(Index).iMailType = gSMTPState(Index).iMailType Or MAIL_SOML
            gSMTPDoMail Index, sParameters
        
        Case "MAIL":
            ' Send to mailbox
            gSMTPState(Index).iMailType = gSMTPState(Index).iMailType Or MAIL_MAIL
            gSMTPDoMail Index, sParameters
        
        Case "RCPT":
            ' Identify recipient
            gSMTPDoRcpt Index, sParameters

        Case "DATA":
            ' Start receiving the message
            If gSMTPState(Index).sMailFrom = "" Then
                gSMTPSend Index, 550, "You haven't told me who sent the message!"
                Exit Sub
            End If

            If gSMTPState(Index).sRcptTo = "" Then
                gSMTPSend Index, 550, "Who am I supposed to give the message to!?!?"
                Exit Sub
            End If

            If gSMTPState(Index).sHelo = "" Then
                gSMTPState(Index).sHelo = "UNIDENTIFIED"
            End If

            gSMTPFileOpen Index
            If Err <> 0 Then
                gSMTPSend Index, 550, "Local Error: " + Error$
            Else
                gSMTPSend Index, 354, "Enter message, terminate with '.' alone on a line."
                gSMTPState(Index).iState = SMTP_DATASTATE
            End If

        Case "QUIT":
            gSMTPSend Index, 221, gLocalHostName & " closing connection. Goodbye!"
            Main!NetClient(Index).Connect = False
            gSMTPSocketClose Index

        Case "RSET":
            ' Reset the "state" of the connection
            gSMTPState(Index).iState = SMTP_COMMANDSTATE
            gSMTPState(Index).sHelo = ""
            gSMTPState(Index).sMailFrom = ""
            gSMTPState(Index).sRcptTo = ""
            gSMTPState(Index).iMailType = gSMTPState(Index).iMailType And MAIL_EHLO
            gSMTPSend Index, 250, "Reset state."

        Case "VERB":
            gSMTPSend Index, 550, "I know that command, but do not implement it."
        
        Case "ONEX":
            gSMTPSend Index, 250, "Yes, I can do one messages at a time."
        
        Case "MULT":
            gSMTPSend Index, 250, "Yes, I can do multiple messages."
        
        Case "VRFY":
            gSMTPSend Index, 250, "They're ok, I guess."

        Case "EXPN":
            gSMTPSend Index, 250, "They're ok, I guess."
        
        Case "XWIN3":
            gSMTPSend Index, 250, "Wow! Another Windows addict!"
        
        Case "TURN":
            gSMTPSend Index, 502, "I can't act as a client, yet..."
        
        Case "NOOP":
            gSMTPSend Index, 250, "Ok. *twiddling thumbs*"

        Case "TICK":
            gSMTPSend Index, 250, "Ok."

        Case "SHOWQ":
            gSMTPSend Index, 550, "Hacker alert! Lame attack!"
        
        Case "DEBUG":
            gSMTPSend Index, 550, "Who you calling buggy?"

        Case "HELP":
            Dim sHelp As String

            sHelp = gSMTPDoHelp(Index, sParameters)
            If sHelp = "" Then sHelp = "Sorry, I have no help on " & sParameters & "."
            sHelp = sHelp & "+End of HELP."
            gSMTPSend Index, 214, sHelp

        Case Else:
            gSMTPLog LOG_WARNING, "Can't parse this: '" & sCommand & " " & sParameters & "'"
            gSMTPSend Index, 500, "Syntax error, command unrecognized."
    
    End Select
End Sub

Sub gSMTPReceive (Index As Integer)
Dim sTemp As String
Dim sCommand As String
Dim sParameters As String
Dim iPlace As Integer

  'gSMTPLog LOG_DEBUG, "gSMTPReceive(" & Str$(Index) & " )"

  On Error GoTo ReceiveError
  
  If Main!NetClient(Index).RecvCount = 0 Then Exit Sub
    
  Do
    If gSMTPState(Index).iState = SMTP_COMMANDSTATE Then
        Do
            sTemp = Main!NetClient(Index)
            If sTemp = "" Then
                If Main!NetClient(Index).RecvCount = 0 Then
                    gSMTPSend Index, 500, "Try HELP if you're lost."
                End If
                Exit Sub
            End If
            gSMTPLog LOG_RECV, sTemp
            iPlace = InStr(sTemp, " ")
            If iPlace <> 0 Then
                sCommand = Left$(sTemp, iPlace - 1)
                sParameters = Right$(sTemp, Len(sTemp) - iPlace)
            Else
                sCommand = sTemp
                sParameters = ""
            End If
            sCommand = UCase$(sCommand)
            gSMTPParse Index, sCommand, sParameters
            If Not gSMTPSocketIsValid(Index) Then Exit Do
            If gSMTPState(Index).iState <> SMTP_COMMANDSTATE Then Exit Do
            sTemp = ""
        Loop While (Main!NetClient(Index).RecvCount > 0)
    Else
        ' Message body!
        Do
            sTemp = Main!NetClient(Index)
'            If sTemp = "" Then Exit Do
            gSMTPLog LOG_RECV, sTemp
            Select Case sTemp
                Case ".."
                    gSMTPFileWrite Index, "."

                Case "."
                    gSMTPState(Index).iState = SMTP_COMMANDSTATE
                    gSMTPFileClose Index
                    If Err <> 0 Then
                        gSMTPSend Index, 550, "Error: " + Error$
                    Else
                        gSMTPSend Index, 250, "Ok."
                    End If
                    Exit Do

                Case Else
                    gSMTPFileWrite Index, sTemp
            End Select
            If Not gSMTPSocketIsValid(Index) Then Exit Do
        Loop While (Main!NetClient(Index).RecvCount > 0)
    End If
  
    DoEvents
    If Not gSMTPSocketIsValid(Index) Then Exit Do
  Loop While (Main!NetClient(Index).RecvCount > 0)

  Exit Sub

ReceiveError:
    gSMTPError Erl, Err, Error$
    Resume Next

End Sub

Sub gSMTPSend (Index As Integer, iCode As Integer, sString As String)
Dim sCode As String
Dim sLine As String
Dim iPlace As Integer
Dim iReplace As Integer

    On Error GoTo SendError
    
    ' This is the code that sends SMTP response lines
    sCode = Str$(iCode)
    sCode = Right$(sCode, Len(sCode) - 1)
    
    '
    ' Replace "_" at the beginning of any line to be sent
    If Left$(sString, 1) = "_" Then
        iReplace = 1
        Do While Right$(Left$(sString, iReplace), 1) = "_"
            iReplace = iReplace + 1
        Loop
        Mid$(sString, 1) = String$(iReplace, " ")
    End If

    iPlace = InStr(sString, "+")

    ' Multi-line responses
    Do While iPlace <> 0
        sLine = Left$(sString, iPlace - 1)
        ' Replace "_" at the beginning of any line to be sent
        If Left$(sLine, 1) = "_" Then
            iReplace = 1
            Do While Right$(Left$(sLine, iReplace), 1) = "_"
                iReplace = iReplace + 1
            Loop
            Mid$(sLine, 1) = String$(iReplace - 1, " ")
        End If
        sLine = sCode & "-" & sLine
        
        Main!NetClient(Index) = sLine & CRLF
        gSMTPLog LOG_SEND, sLine
        sString = Right$(sString, Len(sString) - iPlace)
        iPlace = InStr(sString, "+")
    Loop
    sLine = sCode & " " & sString
    Main!NetClient(Index) = sLine & CRLF
    gSMTPLog LOG_SEND, sLine
    
    ' Be Nice
    DoEvents
    Exit Sub

SendError:
    gSMTPError Erl, Err, Error$
    Resume Next

End Sub

Sub gSMTPSocketClose (Index As Integer)
' This routine is called to close a dynamic connection

' *DIRTY* But it has to be done to close a communications
' channel. Replace this with your form and control array.

    Unload Main!NetClient(Index)

    ' If this is the "last" control in the array, then
    ' decrement the "number" count, otherwise just mark
    ' it as available
    If Index = giIndex Then
        gSMTPState(giIndex).iState = SMTP_CLOSEDSTATE

        While giIndex > 0
            If gSMTPState(giIndex).iState = SMTP_CLOSEDSTATE Then
                giIndex = giIndex - 1
                ReDim Preserve gSMTPState(giIndex) As tSMTPSTATE
            End If
        Wend

    Else
     
        gSMTPState(Index).iState = SMTP_CLOSEDSTATE
                        
    End If

End Sub

Function gSMTPSocketIsValid (Index As Integer) As Integer

    If Index > giIndex Then
        gSMTPSocketIsValid = False
        Exit Function
    End If

    If gSMTPState(Index).iState = SMTP_CLOSEDSTATE Then
        gSMTPSocketIsValid = False
        Exit Function
    End If

    gSMTPSocketIsValid = True
End Function

Sub gSMTPSocketOpen (Index As Integer)
Dim iIndex As Integer
' *DIRTY* But it has to be done to allocate a communications
' channel. Replace these with your form and control array.

    'Make sure the element exists!
    ReDim Preserve gSMTPState(giIndex) As tSMTPSTATE
    
    If giIndex > 0 Then
        For iIndex = 1 To giIndex
            If gSMTPState(iIndex).iState = SMTP_CLOSEDSTATE Then
                gSMTPState(iIndex).iState = SMTP_COMMANDSTATE
                Load Main!NetClient(iIndex)
                Index = iIndex
                Exit Sub
            End If
        Next
    End If

    giIndex = giIndex + 1

    Load Main!NetClient(giIndex)

    ReDim Preserve gSMTPState(giIndex) As tSMTPSTATE
                
    gSMTPState(giIndex).iState = SMTP_COMMANDSTATE
    Index = giIndex

End Sub

Sub gSMTPStartup (cNetServer As NetServer)
Dim sFlags As String

    On Error GoTo InitError
    
    Main!Ini.Section = "Configuration"
    gSMTPLogSetFlags LOG_NOTHING

    gSMTPLogChkFlag "LogError", LOG_ERROR
    gSMTPLogChkFlag "LogWarning", LOG_WARNING
    gSMTPLogChkFlag "LogSend", LOG_SEND
    gSMTPLogChkFlag "LogRecv", LOG_RECV
    gSMTPLogChkFlag "LogSummary", LOG_SUMMARY
    gSMTPLogChkFlag "LogNormal", LOG_NORMAL
    gSMTPLogChkFlag "LogDebug", LOG_DEBUG

    giIndex = 0
    giLogSize = 100

    Main!Ini.Entry = "LogSize"
    If Main!Ini.Value = "" Then
        Main!Ini.Value = Str$(giLogSize)
    Else
        giLogSize = Val(Main!Ini.Value)
    End If
    
    gSMTPLog LOG_NORMAL, "Log size set to" & Str$(giLogSize) & " lines."

    CRLF = Chr$(13) + Chr$(10)
    
    ' *DIRTY* But it has to be done to get the
    ' local host's name. Replace this with your
    ' form and control array.
    gLocalHostName = Main!NetClient(0).HostName

    cNetServer.LocalService = "smtp"
    If cNetServer.LocalPort = 0 Then
        cNetServer.LocalPort = 25
    End If
'    cNetServer.LocalPort = 2015
   
    cNetServer.QueueSize = 5
    cNetServer.Listen = True

    Exit Sub

InitError:
    gSMTPError Erl, Err, Error$
    Resume Next

End Sub

Sub gSMTPTimeOut (Index As Integer)

End Sub

Private Function HEX2Date (h$) As Double
Dim seconds#

    ' Stolen from a VB magazine (which I can't remember the name of)
    seconds# = Val("&H" + h$)
    If Len(h$) Then
        If seconds# > 0 Then
            HEX2Date = (seconds# - 1) / 86400 + 1
        Else
            HEX2Date = (seconds# + 4294967295#) / 86400 + 1
        End If
    End If

End Function


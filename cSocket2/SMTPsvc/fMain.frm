VERSION 5.00
Object = "{072B0A5C-D6FC-496F-B587-7587679E1D60}#1.0#0"; "CSocket.ocx"
Begin VB.Form fMain 
   AutoRedraw      =   -1  'True
   Caption         =   "SMTPSvc"
   ClientHeight    =   1245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3750
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin CSocketOCX.Socket wscFwd 
      Left            =   840
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   2400
      Top             =   120
   End
   Begin VB.Timer tmrProperties 
      Interval        =   100
      Left            =   1680
      Top             =   120
   End
   Begin SMTPSvc.WinsockArray WinsockArray1 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const gAppName As String = "SMTPServer"
Const gsDelimiter As String = "|"
Const str220 As String = "220 "
Const str221 As String = "221 "
Const str250 As String = "250 "
Const str250plus As String = "250-"
Const str354 As String = "354 "
Const str500 As String = "500 "
Const str501 As String = "501 "
Const str502 As String = "502 "
Const str550 As String = "550 "
Const str553 As String = "553 "
Const str554 As String = "554 "
Public strGreet As String
Public FlgDataEvents As String
Public FlgAccept As String
Public LastTime As Double
Public FlgStatus As String
Public UseIPv6 As Boolean
Public TimerInterval As Integer
Dim m_strRemoteHost As String
Dim m_strMagic As String
Public FwdSocketConnected As Boolean
Public lblChngFlg As Boolean
Dim ParseArray() As String
Dim strArray(25) As String

Public Function QuickConnect() As Boolean
    strGreet = GetSettings("Greeting")
    If Len(strGreet) = 0 Then
        strGreet = "554 This server does not accept mail!"
    End If
    FlgDataEvents = Val(GetSettings("ShowLog"))
    If Len(FlgDataEvents) = 0 Then FlgDataEvents = "1"
    FlgAccept = Val(GetSettings("AcceptMail"))
    If Len(FlgAccept) = 0 Then FlgAccept = "0"
    m_strRemoteHost = GetSettings("Remote")
    If Len(m_strRemoteHost) > 0 Then
        If InStr(m_strRemoteHost, ":") Then UseIPv6 = True
    End If
    m_strMagic = GetSettings("Magic")
    QuickConnect = True
End Function
Public Sub SaveSettings(sKey As String, sValue As String)
    'SaveSetting gAppName, "Settings", sKey, sValue
    modReg.SaveSettingEx HKEY_LOCAL_MACHINE, "Software\VB Settings\" & gAppName & "\Settings", sKey, sValue
End Sub
Public Function GetSettings(sKey As String) As String
    'GetSettings = GetSetting(gAppName, "Settings", sKey, "")
    GetSettings = modReg.GetSettingEx(HKEY_LOCAL_MACHINE, "Software\VB Settings\" & gAppName & "\Settings", sKey)
End Function
Function FileErrors(errVal As Long) As Integer
    Dim Msg$, msgType%, Response%
'Return Value 0=Resume,              1=Resume Next,
'             2=Unrecoverable Error, 3=Unrecognized Error
msgType% = 48
Select Case errVal
    Case 68
      Msg$ = "That device appears Unavailable."
      msgType% = msgType% + 4
    Case 71
      Msg$ = "Insert a Disk in the Drive"
    Case 53
      Msg$ = "Cannot Find File"
      msgType% = msgType% + 5
   Case 57
      Msg$ = "Internal Disk Error."
      msgType% = msgType% + 4
    Case 61
      Msg$ = "Disk is Full.  Continue?"
      msgType% = 35
    Case 64, 52
      Msg$ = "That Filename is Illegal!"
    Case 70
      Msg$ = "File in use by another user!"
      msgType% = msgType% + 5
    Case 76
      Msg$ = "Path does not Exist!"
      msgType% = msgType% + 2
    Case 54
      Msg$ = "Bad File Mode!"
    Case 55
      Msg$ = "File is Already Open."
    Case 62
      Msg$ = "Read Attempt Past End of File."
    Case Else
      FileErrors = 3
      Exit Function
  End Select
  FileErrors = 2
End Function

Function OpenFile(Filename$, Mode%, RLock%, RecordLen%) As Integer
    Const REPLACEFILE = 1, READAFILE = 2, ADDTOFILE = 3
    Const RANDOMFILE = 4, BINARYFILE = 5
    Const NOLOCK = 0, RDLOCK = 1, WRLOCK = 2, RWLOCK = 3
    Dim ErrorCode As Long
    Dim FileNum%, Action%, LockFlg%
    LockFlg% = 0
    FileNum% = FreeFile
    On Error GoTo OpenErrors
    Select Case Mode
        Case REPLACEFILE
            Select Case RLock%
                Case NOLOCK
                    Open Filename For Output Shared As FileNum%
                Case RDLOCK
                    Open Filename For Output Lock Read As FileNum%
                Case WRLOCK
                    Open Filename For Output Lock Write As FileNum%
                Case RWLOCK
                    Open Filename For Output Lock Read Write As FileNum%
            End Select
        Case READAFILE
            Select Case RLock%
                Case NOLOCK
                    Open Filename For Input Shared As FileNum%
                Case RDLOCK
                    Open Filename For Input Lock Read As FileNum%
                Case WRLOCK
                    Open Filename For Input Lock Write As FileNum%
                Case RWLOCK
                    Open Filename For Input Lock Read Write As FileNum%
            End Select
        Case ADDTOFILE
            Select Case RLock%
                Case NOLOCK
                    Open Filename For Append Shared As FileNum%
                Case RDLOCK
                    Open Filename For Append Lock Read As FileNum%
                Case WRLOCK
                    Open Filename For Append Lock Write As FileNum%
                Case RWLOCK
                    Open Filename For Append Lock Read Write As FileNum%
            End Select
        Case RANDOMFILE
            Select Case RLock%
                Case NOLOCK
                    Open Filename For Random Shared As FileNum% Len = RecordLen%
                Case RDLOCK
                    Open Filename For Random Lock Read As FileNum% Len = RecordLen%
                Case WRLOCK
                    Open Filename For Random Lock Write As FileNum% Len = RecordLen%
                Case RWLOCK
                    Open Filename For Random Lock Read Write As FileNum% Len = RecordLen%
            End Select
        Case BINARYFILE
            Select Case RLock%
                Case NOLOCK
                    Open Filename For Binary Shared As FileNum%
                Case RDLOCK
                    Open Filename For Binary Lock Read As FileNum%
                Case WRLOCK
                    Open Filename For Binary Lock Write As FileNum%
                Case RWLOCK
                    Open Filename For Binary Lock Read Write As FileNum%
            End Select
        Case Else
          Exit Function
    End Select
    OpenFile = FileNum%
    Exit Function
OpenErrors:
    If Err = 70 Then  'File is locked, try 3 times
        LockFlg% = LockFlg% + 1
        Debug.Print "File Locked!"
        If LockFlg > 3 Then GoTo OpenErrCont
        Resume
    End If
OpenErrCont:
    ErrorCode = Err
    Action% = FileErrors(ErrorCode)
    Select Case Action%
        Case 0
          Resume            'Resumes at line where ERROR occured
        Case 1
            Resume Next     'Resumes at line after ERROR
        Case 2
            OpenFile = 0     'Unrecoverable ERROR-reports error, exits function with error code
            Exit Function
        Case Else
            MsgBox Error$(Err) + vbCrLf + "After line " + Str$(Erl) + vbCrLf + "Program will TERMINATE!"
            'Unrecognized ERROR-reports error and terminates.
            End
    End Select
End Function
Sub LogAccess(Log$, Optional NoTime As Boolean)
    Dim LogFile%
    Dim Filename As String
    Filename = App.Path & "\logs\" & Format(Now, "yymmdd") & ".log"
    On Error Resume Next 'If file does not exist default to overwrite
    If Fix(Now) > Fix(FileDateTime(Filename)) Then
        'Overwrite file
        LogFile% = OpenFile(Filename, 1, 0, 80)
    Else
        'Add to file
        LogFile% = OpenFile(Filename, 3, 0, 80)
    End If
    If LogFile% = 0 Then
        Exit Sub
    End If
    If NoTime Then
        Print #LogFile%, Log$
    Else
        Print #LogFile%, Log$ & gsDelimiter & Format(Now, "hh:mm:ss")
    End If
    Close LogFile%
    On Error GoTo 0
End Sub
Private Function ParseString(lngIndex As Long, A$, Sep$) As String
    Dim M%, N%, N1%
    ReDim ParseArray(100)
    A$ = Trim$(Replace(A$, ":", " "))
    M% = 1: N1% = 0
    Do Until M% = 0
        M% = N% + 1
        N% = InStr(M%, A$, Sep$)
        If N% = 0 Then
            ParseArray(N1%) = Trim$(Mid$(A$, M%))
            N1% = N1% + 1
            Exit Do
        End If
        ParseArray(N1%) = Trim$(Mid$(A$, M%, N% - M%))
        N1% = N1% + 1
        Do Until Mid$(A$, N% + 1, 1) <> Sep$
            N% = N% + 1
        Loop
    Loop
    ReDim Preserve ParseArray(N1% - 1)
    Select Case UCase(ParseArray(0))
        Case "HELO"
            If UBound(ParseArray) < 1 Then
                ParseString = str501 + "helo rquires domain address." + vbCrLf
            Else
                ParseString = str250 + "Hello " + ParseArray(1) + " from " + WinsockArray1.RemoteHostIP(lngIndex) + ", pleased to meet you." + vbCrLf
                'ParseString = str554 + "yellowhead.com does not accept mail [RFC5321]." + vbCrLf
            End If
        Case "EHLO"
            If UBound(ParseArray) < 1 Then
                ParseString = str501 + "ehlo requires domain address." + vbCrLf
            Else
                ParseString = str250plus + "Hello " + ParseArray(1) + " from " + WinsockArray1.RemoteHostIP(lngIndex) + ", pleased to meet you." + vbCrLf _
                  + str250plus + "8BITMIME" + vbCrLf _
                  + str250plus + "SIZE 4000000" + vbCrLf _
                  + str250plus + "DSN" + vbCrLf _
                  + str250plus + "ONEX" + vbCrLf _
                  + str250plus + "ETRN" + vbCrLf _
                  + str250plus + "XUSR" + vbCrLf _
                  + str250 + "HELP" + vbCrLf
                'ParseString = str554 + "yellowhead.com does not accept mail [RFC5321]." + vbCrLf
            End If
        Case "AUTH"
                ParseString = str502 + "Command not implemented." + vbCrLf
        Case "MAIL"
            If UBound(ParseArray) > 1 Then
                If Len(ParseArray(2)) = 0 Then
                    ParseString = str553 + "Domain name required." + vbCrLf
                Else
                    'ParseString = str250 + "Sender " + ParseArray(2) + " Ok!" + vbCrLf
                    ParseString = str553 + "Sender " + ParseArray(2) + " is Invalid!" + vbCrLf
                End If
            Else
                ParseString = str553 + "Domain name required." + vbCrLf
            End If
        Case "RCPT"
            If UBound(ParseArray) > 1 Then
                If Len(ParseArray(2)) = 0 Then
                    ParseString = str553 + "Domain name required." + vbCrLf
                Else
                    If FlgAccept Then
                        ParseString = str250 + "Recipient " + ParseArray(2) + " Ok!" + vbCrLf
                    Else
                        Call WinsockArray1.Send(lngIndex, str550 + WinsockArray1.RemoteHostIP(lngIndex) + " will not find any valid users on this server." + vbCrLf)
                        ParseString = str221 + "Server Disconnecting" + vbCrLf
                    End If
                End If
            Else
                ParseString = str553 + "Domain name required." + vbCrLf
            End If
        Case "RSET"
                ParseString = str250 + "Reset state." + vbCrLf
        Case "QUIT"
            ParseString = str221 + "Server Disconnecting" + vbCrLf
        Case "DATA"
            If FlgAccept Then
                ParseString = str354 + "Enter mail, end with '.' on a line by itself" + vbCrLf
            Else
                ParseString = str221 + "Server Disconnecting" + vbCrLf
            End If
        Case "."
            ParseString = str250 + "Spam Message sent to the bit bucket!" + vbCrLf
            'ParseString = "QUIT"
        Case Else
'            ParseString = str500 + "Command unrecognized: " + ParseArray(0) + vbCrLf
    End Select
End Function

Public Sub SendString(strdata)
    'If for some reason the socket gets closed, a 40006 error is generated that
    'is not trapped. For this reason we check the state before sending.
    On Error GoTo SendStringErr
    If wscFwd.State = sckConnected Then
        Call wscFwd.SendData(Chr$(&H1F) & strdata)
    End If
    Exit Sub
SendStringErr:
    Call LogAccess("Error" + Str$(Err) + " in SendString!")
End Sub

Public Sub StartServer(Optional IPv6 As Boolean)
    With WinsockArray1
        .MaxClients = 25
        .RaiseErrors = True
        .LocalPort = 25
        If (.StartListening(IPv6) = False) Then
            Call LogAccess("Start() failed! ")
        Else
            Call LogAccess("Server() started! ")
            LastTime = Time
            FlgStatus = "Online"
            With wscFwd
                If IPv6 Then
                    .IPv6Flg = 6
                Else
                    .IPv6Flg = 4
                End If
                .LocalPort = 26
                Call .Listen
                If .State <> sckListening Then
                    Call LogAccess("Forward Port 26 not active! ")
                End If
            End With
        End If
    End With
End Sub
Public Sub StopServer()
    On Error GoTo StopServerErr
    If (WinsockArray1.Shutdown) Then
        FlgStatus = "Offline"
        Call LogAccess("Server() stopped! ")
    Else
        Call LogAccess("Shutdown() failed! ")
    End If
    Exit Sub
StopServerErr:
    Call LogAccess("Error" + Str$(Err) + " in Stopping Server!")
End Sub

Private Function GetNotificationWindow() As Long
    Dim h As Long
    h = FindWindowEx(0, 0, "Shell_TrayWnd", vbNullString)
    If h Then h = FindWindowEx(h, 0, "TrayNotifyWnd", vbNullString)
    If h Then h = FindWindowEx(h, 0, "SysPager", vbNullString)
    If h Then
        GetNotificationWindow = FindWindowEx(h, 0, "ToolbarWindow32", vbNullString)
    Else
        GetNotificationWindow = 0
    End If
End Function
Private Function GetProcessNameFromHwnd(ByVal hwnd As Long) As String
    Dim ProcessId As Long, hProcess As Long, hModule As Long, s As String * MAX_PATH
    GetWindowThreadProcessId hwnd, ProcessId
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessId)
    EnumProcessModules hProcess, hModule, 1, ByVal 0
    GetModuleFileNameEx hProcess, hModule, s, MAX_PATH
    GetProcessNameFromHwnd = Left$(s, InStr(s, vbNullChar) - 1)
    CloseHandle hProcess
End Function

Private Sub Form_Load()
    Dim Handle As Long
    TimerInterval = 1000
    If Not QuickConnect Then
        Call LogAccess("Problem with Registry Info!", True)
        Exit Sub
    End If
    Call StartServer(UseIPv6)
End Sub





Private Sub tmrConnect_Timer()
    If Not FwdSocketConnected Then
        SendString ("Not Authorized!")
        Call LogAccess("Timed out waiting for magic string!")
        Call wscFwd_CloseSck
    End If
    'tmrConnect.Interval = 0
End Sub

Private Sub tmrProperties_Timer()
    Dim N%
    Dim tmpStr As String
    With WinsockArray1
        If FlgStatus <> "Offline" Then
            tmpStr = "Online " & .SocketCount & " sockets"
            If tmpStr <> FlgStatus Then 'Update only if changed
                FlgStatus = tmpStr
                If FwdSocketConnected Then 'Send to monitor if enabled
                    SendString (tmpStr)
                End If
            End If
        End If
        'If idle for 30 seconds, close any active sockets
        If Time - LastTime > 0.0003 Then
            If WinsockArray1.ActiveClients > 0 Then
                Debug.Print Time & " Active Sockets = " & CStr(WinsockArray1.ActiveClients)
                For N% = 1 To WinsockArray1.SocketCount
                    Call WinsockArray1.CloseClient(N%)
                Next N%
                Debug.Print Time & " Active Now = " & CStr(WinsockArray1.ActiveClients)
            End If
        End If
    End With
End Sub

Private Sub WinsockArray1_ConnectionClosed(ByVal lngIndex As Long)
    On Error GoTo ConnectionClosedErr
    Call LogAccess(CStr(lngIndex) & ". " & CStr(WinsockArray1.RemoteHostIP(lngIndex)) & " Closed.")
    If FwdSocketConnected Then 'Send to monitor if enabled
        Call SendString(CStr(lngIndex) & ". Closed")
    End If
    strArray(lngIndex) = ""
    Exit Sub
ConnectionClosedErr:
    Call LogAccess("Error" + Str$(Err) + " Encountered Closing Connection on: " + Str$(lngIndex))
End Sub

Private Sub WinsockArray1_DataArrival(ByVal lngIndex As Long, ByVal strdata As String)
    Dim M%, N%
    Dim Response$
    On Error GoTo DataArrivalErr
    N% = InStr(strdata, Chr$(13))
    If N% = 0 Then 'Incomplete line
        strArray(lngIndex) = strArray(lngIndex) + strdata
        Exit Sub
    Else
        M% = 1
    End If
    'Debug.Print strData
    Do Until N% = 0
        strArray(lngIndex) = strArray(lngIndex) + Mid$(strdata, M%, N% - M%)
        If (FlgDataEvents) Then
            Call LogAccess(CStr(lngIndex) & ". " & strArray(lngIndex), True)
            If FwdSocketConnected Then 'Send to monitor if enabled
                Call SendString(CStr(lngIndex) & ". " & strArray(lngIndex))
            End If
        End If
        Response$ = ParseString(lngIndex, strArray(lngIndex), " ")
        'Debug.Print Response$
        If Len(Response$) > 0 Then
            Call WinsockArray1.Send(lngIndex, Response$)
            'Check for QUIT or 501 response
            If Val(Response$) = 221 Then
                Call WinsockArray1.CloseClient(lngIndex)
                Exit Sub
            End If
        End If
        strArray(lngIndex) = ""
        If Mid$(strdata, N% + 1, 1) = Chr(10) Then
            M% = N% + 2
        Else
            M% = N% + 1
        End If
        N% = InStr(M% + 1, strdata, Chr$(13))
    Loop
    strArray(lngIndex) = Mid$(strdata, M%)
    Exit Sub
DataArrivalErr:
    Call LogAccess(strdata)
    Call LogAccess(strArray(lngIndex))
End Sub

Private Sub WinsockArray1_NewConnection(ByVal lngIndex As Long, blnCancel As Boolean)
    Call WinsockArray1.Send(lngIndex, str220 + strGreet + Str$(Now) + " -0700" + vbCrLf)
    Call LogAccess(CStr(lngIndex) & ". " & CStr(WinsockArray1.RemoteHostIP(lngIndex)) + " on port " + CStr(WinsockArray1.RemotePort(lngIndex)))
    If FwdSocketConnected Then 'Send to monitor if enabled
        Call SendString(CStr(lngIndex) & ". Connected " & CStr(WinsockArray1.RemoteHostIP(lngIndex)) + Space$(1) + CStr(WinsockArray1.RemotePort(lngIndex)))
    End If
    LastTime = Time 'Reset timeout
End Sub

Private Sub WinsockArray1_WinsockError(ByVal lngIndex As Long, ByVal lngNumber As Long, ByVal strSource As String, ByVal strDescription As String)
    Call LogAccess(Str$(lngIndex) & ". Error" & CStr(lngNumber) & ": " & strDescription)
    If FwdSocketConnected Then 'Send to monitor if enabled
        Call SendString(Str$(lngIndex) & ". Error " & CStr(lngNumber) & ": " & strDescription)
    End If
End Sub



Private Sub wscFwd_CloseSck()
    On Error GoTo errHandler
    If wscFwd.State <> sckClosed Or sckClosing Then
        wscFwd.CloseSck
    End If
    tmrConnect.Enabled = False
    FwdSocketConnected = False
    Do Until wscFwd.State = sckClosed
        DoEvents
    Loop
    Call LogAccess("Forward Socket Closed! ")
    Call wscFwd.Listen
    If wscFwd.State <> sckListening Then
        Call LogAccess("Forward Port 26 not active! ")
    End If
    Exit Sub
errHandler:
    Call LogAccess("Error " & CStr(Err) & " in wscFwd_CloseSck")
End Sub

Private Sub wscFwd_Connect()
    On Error GoTo SocketError
    Exit Sub
SocketError:
    Call LogAccess("Error " & CStr(Err) & " in wscFwd_Connect")
End Sub

Private Sub wscFwd_ConnectionRequest(ByVal requestID As Long)
    On Error GoTo SocketError
    wscFwd.CloseSck
    wscFwd.Accept requestID
    tmrConnect.Enabled = True
    Exit Sub
SocketError:
    Call LogAccess("Error " & CStr(Err) & " in wscFwd_ConnectionRequest")
End Sub


Private Sub wscFwd_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim strdata As String
    wscFwd.GetData strdata
    Call LogAccess(">>>" & strdata)
    Select Case strdata
        Case "Stop Server"
            Call StopServer
            If FlgStatus = "Offline" Then
                Call SendString(FlgStatus)
            Else
                Call SendString("Stop Server Failed!")
            End If
        Case "Start Server"
            Call StartServer(True)
            If FlgStatus = "Online" Then
                Call SendString(FlgStatus)
            Else
                Call SendString("Start Server Failed!")
            End If
        Case m_strMagic
            Debug.Print "Magic String Received!"
            tmrConnect.Enabled = False
            FwdSocketConnected = True
            SendString ("Online " & WinsockArray1.SocketCount & " sockets")
            Call LogAccess("Forward Socket Opened! ")
    End Select
End Sub




Private Sub wscFwd_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error GoTo SocketError
    Debug.Print "Error" & Number & ":" & Description
    Call LogAccess("Error" & Number & ":" & Description)
    wscFwd_CloseSck
    Exit Sub
SocketError:
    Call LogAccess("Error " & CStr(Err) & " in wscFwd_Error")
End Sub


Private Sub wscFwd_SendComplete()
    On Error GoTo SocketError
    Exit Sub
SocketError:
    Call LogAccess("Error " & CStr(Err) & " in wscFwd_SendComplete")
End Sub


Private Sub wscFwd_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    On Error GoTo SocketError
    Exit Sub
SocketError:
    Call LogAccess("Error " & CStr(Err) & " in wscFwd_SendProgress")
End Sub



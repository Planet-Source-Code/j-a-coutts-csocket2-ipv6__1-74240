VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSocket2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' ---------------------------------------------------------------------------------
' File...........: cSocket2.cls
' Author.........: J.A. Coutts
' Created........: 02/05/11
' Modified.......: 05/05/11
' Version........: 1.0
' Website........: http://www.yellowhead.com
' Contact........: allecnarf@hotmail.com
'
'Copyright (c) 2011 by JAC Computing
'Vernon, BC, Canada
'
'Based on CSocketMaster by Emiliano Scavuzzo
'and CSocket by Oleg Gdalevich
'Subclassing based on WinSubHook2 by Paul Caton
'
' Port of necessary Winsock2 declares, consts, types etc..
' ------------------------------------------------------------------------------
'
'================================
'CONSTANTS
'================================
Private m_enmState              As SockState            'socket state

Public Enum SockState
    sckClosed = 0
    sckOpen
    sckListening
    sckConnectionPending
    sckResolvingHost
    sckHostResolved
    sckConnecting
    sckConnected
    sckClosing
    sckError
End Enum

Public Enum ProtocolConstants
    sckTCPProtocol = 0
    sckUDPProtocol = 1
End Enum

Private Const MSG_PEEK  As Long = &H2

'================================
'EVENTS
'================================
Public Event CloseSck()
Public Event Connect()
Public Event ConnectionRequest(ByVal requestID As Long)
Public Event DataArrival(ByVal bytesTotal As Long)
Public Event Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event SendComplete()
Public Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'================================
'MEMBER VARIABLES
'================================
Private m_lngSocketHandle   As Long   'socket handle
Private m_lngMemoryPointer  As Long   'memory pointer used as buffer when resolving host
Private m_lngMemoryHandle   As Long   'buffer memory handle
Private m_strRemoteHost     As String 'remote host
Private m_lngRemotePort     As Long   'remote port
Private m_strRemotePort     As String 'remote port
Private m_strRemoteHostIP   As String 'remote host ip
Private m_lngLocalPort      As Long   'local port
Private m_strLocalIP        As String 'local IP
Private m_lngLocalPortBind  As Long   'temporary local port
Private m_lngLocalIPIndex   As Long   'Index of local bind port
Private m_IPv               As Long
Private Sa()                As sockaddr
Private Hints               As addrinfo

Private m_colWaitingResolutions As Collection   'hosts waiting to be resolved by the system
Private m_blnAcceptClass        As Boolean      'if True then this is a Accept socket class
Private m_enmProtocol           As ProtocolConstants    'protocol used (TCP / UDP)
Private m_strTag                As String               'tag

Private m_lngSendBufferLen  As Long 'winsock buffer size for sends
Private m_lngRecvBufferLen  As Long 'winsock buffer size for receives

Private m_strSendBuffer As String   'local incoming buffer
Private m_strRecvBuffer As String   'local outgoing buffer

Private Declare Function api_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As sockaddr_in, ByVal namelen As Long) As Long
Private Declare Function api_Connect2 Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As sockaddr, ByVal namelen As Long) As Long
Private Declare Function api_bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Private Declare Function api_bind2 Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef name As sockaddr, ByRef namelen As Long) As Long
Private Declare Function api_listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function api_accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As sockaddr_in, ByRef addrlen As Long) As Long
Private Declare Function api_accept2 Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As sockaddr, ByRef addrlen As Long) As Long

Public Sub WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'  ****  WARNING WARNING WARNING WARNING ******
'This sub MUST be the first on the class. DO NOT attempt
'to change it's location or the code will CRASH.
'This sub receives system messages from our WndProc.
    Select Case uMsg
    Case SOCKET_MESSAGE
        PostSocket LoWord(lParam), HiWord(lParam)
    End Select
End Sub

Public Property Get LocalHostName() As String
    LocalHostName = m_strLocalIP
End Property

Public Property Get LocalIP() As String
    If m_enmState = sckOpen Or m_enmState = sckListening Then
        LocalIP = m_strLocalIP
    Else
'        LocalIP = GetLocalIP
    End If
End Property
Public Property Get LocalPort() As Long
    If m_lngLocalPortBind = 0 Then
        LocalPort = m_lngLocalPort
    Else
        LocalPort = m_lngLocalPortBind
    End If
End Property

Public Property Let LocalPort(ByVal lngPort As Long)
    If m_enmState <> sckClosed Then
        err.Raise sckInvalidOp, "CSocket2.LocalPort", "Invalid operation at current state"
    End If
    If lngPort < 0 Or lngPort > 65535 Then
        err.Raise sckInvalidArg, "CSocket2.LocalPort", "The argument passed to a function was not in the correct format or in the specified range."
    Else
        m_lngLocalPort = lngPort
    End If
End Property

Public Property Get RemotePort() As Long
    RemotePort = m_lngRemotePort
End Property
Public Property Let RemotePort(ByVal lngPort As Long)
    If m_enmProtocol = sckTCPProtocol And m_enmState <> sckClosed Then
        err.Raise sckInvalidOp, "CSocket2.RemotePort", "Invalid operation at current state"
    End If
    If lngPort < 0 Or lngPort > 65535 Then
        err.Raise sckInvalidArg, "CSocket2.RemotePort", "The argument passed to a function was not in the correct format or in the specified range."
    Else
        m_lngRemotePort = lngPort
    End If
End Property

Public Property Get RemoteHostIP() As String
    RemoteHostIP = m_strRemoteHostIP
End Property

Public Property Get RemoteHost() As String
    RemoteHost = m_strRemoteHost
End Property
Public Property Let RemoteHost(ByVal strHost As String)
    If m_enmProtocol = sckTCPProtocol And m_enmState <> sckClosed Then
        err.Raise sckInvalidOp, "CSocket2.RemoteHost", "Invalid operation at current state"
    End If
    'Changed 08/11/2011 to force system to find new IP when name changed
    'Necessary for UDP only
    If m_strRemoteHost <> strHost Then
        m_strRemoteHost = strHost
        m_strRemoteHostIP = ""
    End If
End Property

Private Function BindInternal2(Optional ByVal varLocalPort As Variant, Optional ByVal varLocalIP As Variant) As Boolean
'This function binds a socket to a local port and IP.
'Returns TRUE if it has success.
    If m_enmState = sckOpen Then
        BindInternal2 = True
        Exit Function
    End If
    Dim lngLocalPortInternal As Long
    Dim strLocalHostInternal As String
    Dim strIP As String
    Dim lngResult As Long
    Dim lngErrorCode As Long
    BindInternal2 = False
    'Check if varLocalPort is a number between 0 and 65535
    If Not IsMissing(varLocalPort) Then
        If IsNumeric(varLocalPort) Then
            If varLocalPort < 0 Or varLocalPort > 65535 Then
                BindInternal2 = False
                err.Raise sckInvalidArg, "CSocket2.BindInternal2", "The argument passed to a function was not in the correct format or in the specified range."
            Else
                lngLocalPortInternal = CLng(varLocalPort)
            End If
        Else
            BindInternal2 = False
            err.Raise sckUnsupported, "CSocket2.BindInternal2", "Unsupported variant type."
        End If
    Else
        lngLocalPortInternal = m_lngLocalPort
    End If
    If Not IsMissing(varLocalIP) Then
        If varLocalIP = vbNullString Then
            strLocalHostInternal = GetLocalHostName
        Else
            strLocalHostInternal = CStr(varLocalIP)
        End If
    Else
        strLocalHostInternal = m_strLocalIP
    End If
    Debug.Print "Version =" & Str$(m_IPv)
    strIP = GetIPFromHost(strLocalHostInternal, Str$(lngLocalPortInternal), m_IPv)
    If Not SocketExists Then Exit Function
    'bind the socket
    lngResult = api_bind2(m_lngSocketHandle, Sa(0), LenB(Sa(0)))
    If lngResult = SOCKET_ERROR Then
        lngErrorCode = err.LastDllError
        err.Raise lngErrorCode, "CSocket2.BindInternal2", GetErrorDescription(lngErrorCode)
    Else
        m_strLocalIP = Left$(strIP, InStr(strIP, Chr$(0)) - 1)
        If lngLocalPortInternal <> 0 Then
            Debug.Print "OK Bind HOST: " & strLocalHostInternal & " PORT: " & lngLocalPortInternal
            If DbgFlg Then Call LogError("OK Bind HOST: " & strLocalHostInternal & " PORT:" & Str$(lngLocalPortInternal))
            m_lngLocalPort = lngLocalPortInternal
        Else
            lngResult = GetLocalPort(m_lngSocketHandle)
            If lngResult = SOCKET_ERROR Then
                lngErrorCode = err.LastDllError
                err.Raise lngErrorCode, "CSocket2.BindInternal", GetErrorDescription(lngErrorCode)
            Else
                Debug.Print "OK Bind HOST: " & strLocalHostInternal & " PORT: " & lngResult
                If DbgFlg Then Call LogError("OK Bind HOST: " & strLocalHostInternal & " PORT:" & Str$(lngLocalPortInternal))
                m_lngLocalPortBind = lngResult
            End If
        End If
        BindInternal2 = True
    End If
End Function
Public Sub Bind2(Optional LocalPort As Variant, Optional LocalIP As Variant)
    If m_enmState <> sckClosed Then
        err.Raise sckInvalidOp, "CSocket2.Bind", "Invalid operation at current state"
    End If
    If BindInternal2(LocalPort, LocalIP) Then
        m_enmState = sckOpen: Debug.Print "STATE: sckOpen"
        If DbgFlg Then Call LogError("STATE: sckOpen")
    End If
End Sub

Public Sub Listen()
    If m_enmState <> sckClosed And m_enmState <> sckOpen Then
        err.Raise sckInvalidOp, "CSocket2.Listen", "Invalid operation at current state"
    End If
    If Not SocketExists Then Exit Sub
    If Not BindInternal2 Then Exit Sub
    Dim lngResult As Long
    lngResult = api_listen(m_lngSocketHandle, SOMAXCONN)
    If lngResult = SOCKET_ERROR Then
        Dim lngErrorCode As Long
        lngErrorCode = err.LastDllError
        err.Raise lngErrorCode, "CSocket2.Listen", GetErrorDescription(lngErrorCode)
    Else
        m_enmState = sckListening: Debug.Print "STATE: sckListening"
        If DbgFlg Then Call LogError("STATE: sckListening")
    End If
End Sub

Public Sub Connect2(Optional RemoteHost As Variant, Optional RemotePort As Variant)
    Dim strTemp         As String
    Dim lngResult       As Long
    Dim lngErrorCode    As Long
    Dim blnCancelDisplay As Boolean
    If m_enmState <> sckClosed Then
        err.Raise sckInvalidOp, "CSocket2.Connect", "Invalid operation at current state"
    End If
    If Not IsMissing(RemoteHost) Then
            m_strRemoteHost = CStr(RemoteHost)
    End If
    If m_strRemoteHost = vbNullString Then
        m_strRemoteHost = ""
    End If
    If Not IsMissing(RemotePort) Then
        If IsNumeric(RemotePort) Then
            If CLng(RemotePort) > 65535 Or CLng(RemotePort) < 1 Then
                err.Raise sckInvalidArg, "CSocket2.Connect", "The argument passed to a function was not in the correct format or in the specified range."
            Else
                m_lngRemotePort = CLng(RemotePort)
                m_strRemotePort = CStr(RemotePort)
            End If
        Else
            err.Raise sckUnsupported, "CSocket2.Connect", "Unsupported variant type."
        End If
    End If
    strTemp = GetIPFromHost(m_strRemoteHost, m_strRemotePort, m_IPv)
    If Len(strTemp) Then    'Address found
        If Not SocketExists Then Exit Sub
        'If we are using UDP we just bind the socket and exit
        If m_enmProtocol = sckUDPProtocol Then
            m_strLocalIP = vbNullString
            If BindInternal2 Then
                m_enmState = sckOpen: Debug.Print "STATE: sckOpen"
                If DbgFlg Then Call LogError("STATE: sckOpen")
            End If
            Exit Sub
        Else
            'Using GetAddrInfo, bind is not required for TCP connect
            m_enmState = sckConnecting: Debug.Print "STATE: sckConnecting"
            If DbgFlg Then Call LogError("STATE: sckConnecting")
            lngResult = api_Connect2(m_lngSocketHandle, Sa(0), LenB(Sa(0)))
            'Check and handle errors
            If lngResult = SOCKET_ERROR Then
                lngErrorCode = err.LastDllError
                Debug.Print "Error" & Str$(lngErrorCode)
                If DbgFlg Then Call LogError("Error" & Str$(lngErrorCode))
                If lngErrorCode <> WSAEWOULDBLOCK Then
                    If lngErrorCode = WSAEADDRNOTAVAIL Then
                        err.Raise WSAEADDRNOTAVAIL, "CSocket2.Connect2", GetErrorDescription(WSAEADDRNOTAVAIL)
                    Else
                        m_enmState = sckError: Debug.Print "STATE: sckError"
                        If DbgFlg Then Call LogError("STATE: sckError")
                        blnCancelDisplay = True
                        RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "CSocket2.Connect2", "", 0, blnCancelDisplay)
                        If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "CSocket2.Connect2"
                    End If
                End If
            End If
        End If
    Else
    End If
End Sub
Public Sub CloseSck()
    If m_lngSocketHandle = INVALID_SOCKET Then Exit Sub
    m_enmState = sckClosing: Debug.Print "STATE: sckClosing"
    If DbgFlg Then Call LogError("STATE: sckClosing")
    DestroySocket
    m_lngLocalPortBind = 0
    m_strLocalIP = vbNullString
    m_strRemoteHostIP = ""
    m_strRecvBuffer = ""
    m_strSendBuffer = ""
    m_lngSendBufferLen = 0
    m_lngRecvBufferLen = 0
    m_enmState = sckClosed: Debug.Print "STATE: sckClosed"
    If DbgFlg Then Call LogError("STATE: sckClosed")
End Sub

Public Sub PeekData(ByRef data As Variant, Optional varType As Variant, Optional maxLen As Variant)
    If m_enmProtocol = sckTCPProtocol Then
        If m_enmState <> sckConnected Then
            err.Raise sckBadState, "CSocket2.PeekData", "Wrong protocol or connection state for the requested transaction or request"
            Exit Sub
        End If
    Else
        If m_enmState <> sckOpen Then
            err.Raise sckBadState, "CSocket2.PeekData", "Wrong protocol or connection state for the requested transaction or request"
            Exit Sub
        End If
        If GetBufferLenUDP = 0 Then Exit Sub
    End If
    If Not IsMissing(maxLen) Then
        If IsNumeric(maxLen) Then
            If CLng(maxLen) < 0 Then
                err.Raise sckInvalidArg, "CSocket2.PeekData", "The argument passed to a function was not in the correct format or in the specified range."
            End If
        Else
            If m_enmProtocol = sckTCPProtocol Then
                maxLen = Len(m_strRecvBuffer)
            Else
                maxLen = GetBufferLenUDP
            End If
        End If
    End If
    Dim lngBytesRecibidos  As Long
    lngBytesRecibidos = RecvData(data, True, varType, maxLen)
    Debug.Print "OK Bytes obtained from buffer: " & lngBytesRecibidos
End Sub
Private Sub PostSocket(ByVal lngEventID As Long, ByVal lngErrorCode As Long)
'This procedure is called by the WindowProc callback function
'from the mWinsock2 module. The lngEventID argument is an
'ID of the network event occurred for the socket. The lngErrorCode
'argument contains an error code only if an error was occurred
'during an asynchronous execution.
    'handle any possible error
    If lngErrorCode <> 0 Then
        m_enmState = sckError: Debug.Print "STATE: sckError"
        If DbgFlg Then Call LogError("STATE: sckError")
        Dim blnCancelDisplay As Boolean
        blnCancelDisplay = True
        RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "CSocket2.PostSocket", "", 0, blnCancelDisplay)
        If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "CSocket2.PostSocket"
        Exit Sub
    End If
'    Dim udtSockAddr As sockaddr_in
    Dim newSa As sockaddr
    Dim Sin4 As sockaddr_in
    Dim Sin6 As sockaddr_in6
    Dim lngResult As Long
    Dim lngBytesReceived As Long
    Select Case lngEventID
    '======================================================================
    Case FD_CONNECT
        'Arrival of this message means that the connection initiated by the call
        'of the connect Winsock API function was successfully established.
        Debug.Print "FD_CONNECT " & m_lngSocketHandle
        If DbgFlg Then Call LogError("FD_CONNECT" & Str$(m_lngSocketHandle))
        If m_enmState <> sckConnecting Then
            Debug.Print "WARNING: Omitting FD_CONNECT"
            If DbgFlg Then Call LogError("WARNING: Omitting FD_CONNECT")
            Exit Sub
        End If
        'Get the connection local end-point parameters
        lngResult = getpeername2(m_lngSocketHandle, newSa, LenB(newSa))
        If lngResult = 0 Then
            If newSa.sa_family = AF_INET6 Then
                ReDim bBuffer(0 To INET6_ADDRSTRLEN - 1)    'Resize string buffer
                CopyMemory Sin6, newSa, LenB(Sin6)  'Save to sockaddr_in6
                m_lngRemotePort = mWinsock2.IntegerToUnsigned(ntohs(Sin6.sin6_port))
                lngResult = inet_ntop(AF_INET6, Sin6.sin6_addr, bBuffer(0), INET6_ADDRSTRLEN)
            Else
                ReDim bBuffer(0 To INET_ADDRSTRLEN - 1)    'Resize string buffer
                CopyMemory Sin4, newSa, LenB(Sin4)  'Save to sockaddr_in
                m_lngRemotePort = mWinsock2.IntegerToUnsigned(ntohs(Sin4.sin_port))
                lngResult = inet_ntop(AF_INET, Sin4.sin_addr, bBuffer(0), INET_ADDRSTRLEN)
            End If
            If lngResult Then m_strRemoteHostIP = StringFromPointer(lngResult)
        End If
        m_enmState = sckConnected: Debug.Print "STATE: sckConnected"
        If DbgFlg Then Call LogError("STATE: sckConnected")
        RaiseEvent Connect
    '======================================================================
    Case FD_WRITE
        'This message means that the socket in a write-able
        'state, that is, buffer for outgoing data of the transport
        'service is empty and ready to receive data to send through
        'the network.
        Debug.Print "FD_WRITE " & m_lngSocketHandle
        If m_enmState <> sckConnected Then
            Debug.Print "WARNING: Omitting FD_WRITE"
            If DbgFlg Then Call LogError("WARNING: Omitting FD_WRITE")
           Exit Sub
        End If
        If Len(m_strSendBuffer) > 0 Then
            SendBufferedData
        End If
    '======================================================================
    Case FD_READ
        'Some data has arrived for this socket.
        Debug.Print "FD_READ " & m_lngSocketHandle
        If m_enmProtocol = sckTCPProtocol Then
            If m_enmState <> sckConnected Then
                Debug.Print "WARNING: Omitting FD_READ"
                If DbgFlg Then Call LogError("WARNING: Omitting FD_READ")
                Exit Sub
            End If
            'Call the RecvDataToBuffer function that move arrived data
            'from the Winsock buffer to the local one and returns number
            'of bytes received.
            lngBytesReceived = RecvDataToBuffer
            If lngBytesReceived > 0 Then
                RaiseEvent DataArrival(Len(m_strRecvBuffer))
            End If
        Else 'UDP protocol
            If m_enmState <> sckOpen Then
                Debug.Print "WARNING: Omitting FD_READ"
                If DbgFlg Then Call LogError("WARNING: Omitting FD_READ")
                Exit Sub
            End If
            'If we use UDP we don't remove data from winsock buffer.
            'We just let the user know the amount received so
            'he/she can decide what to do.
            lngBytesReceived = GetBufferLenUDP
            If lngBytesReceived > 0 Then
                RaiseEvent DataArrival(lngBytesReceived)
            End If
            'Now the buffer is emptied no matter what the user
            'dicided to do with the received data
            EmptyBuffer
        End If
    '======================================================================
    Case FD_ACCEPT
        'When the socket is in a listening state, arrival of this message
        'means that a connection request was received. Call the accept
        'Winsock API function in oreder to create a new socket for the
        'requested connection.
        Debug.Print "FD_ACCEPT " & m_lngSocketHandle
        If DbgFlg Then Call LogError("FD_ACCEPT:" & Str$(m_lngSocketHandle))
        If m_enmState <> sckListening Then
            Debug.Print "WARNING: Omitting FD_ACCEPT"
            If DbgFlg Then Call LogError("WARNING: Omitting FD_ACCEPT")
            Exit Sub
        End If
        'lngResult = api_accept(m_lngSocketHandle, udtSockAddr, LenB(udtSockAddr))
        lngResult = api_accept2(m_lngSocketHandle, newSa, LenB(newSa))
        If lngResult = INVALID_SOCKET Then
            lngErrorCode = err.LastDllError
            err.Raise lngErrorCode, "CSocket2.PostSocket", GetErrorDescription(lngErrorCode)
        Else
            'We assign a temporal instance of CSocket2 to
            'handle this new socket until user accepts (or not)
            'the new connection
            mWinsock2.RegisterAccept lngResult
            'We change remote info before firing ConnectionRequest
            'event so the user can see which host is trying to
            'connect.
            Dim lngTempRP As Long
            Dim strTempRHIP As String
            Dim strTempRH As String
            lngTempRP = m_lngRemotePort
            strTempRHIP = m_strRemoteHostIP
            strTempRH = m_strRemoteHost
            GetRemoteInfo lngResult, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
            Debug.Print "OK Accepted socket: " & lngResult
            If DbgFlg Then Call LogError("OK Accepted socket:" & Str$(lngResult))
            RaiseEvent ConnectionRequest(lngResult)
            'we return original info
            If m_enmState = sckListening Then
                 m_lngRemotePort = lngTempRP
                 m_strRemoteHostIP = strTempRHIP
                 m_strRemoteHost = strTempRH
            End If
            'This is very important. If the connection wasn't accepted
            'we must close the socket.
            If IsAcceptRegistered(lngResult) Then
                closesocket lngResult
                mWinsock2.UnregisterSocket lngResult
                mWinsock2.UnregisterAccept lngResult
                Debug.Print "OK Closed accepted socket: " & lngResult
                If DbgFlg Then Call LogError("OK Closed accepted socket:" & Str(lngResult))
            End If
        End If
    '======================================================================
    Case FD_CLOSE
        'This message means that the remote host is closing the conection
        Debug.Print "FD_CLOSE " & m_lngSocketHandle
        If m_enmState <> sckConnected Then
            Debug.Print "WARNING: Omitting FD_CLOSE"
            If DbgFlg Then Call LogError("WARNING: Omitting FD_CLOSE")
            Exit Sub
        End If
        m_enmState = sckClosing: Debug.Print "REMOTE: sckClosing"
        If DbgFlg Then Call LogError("REMOTE: sckClosing")
        RaiseEvent CloseSck
    End Select
End Sub
Public Property Get Protocol() As ProtocolConstants
    Protocol = m_enmProtocol
End Property

Public Property Let Protocol(ByVal enmProtocol As ProtocolConstants)
    If m_enmState <> sckClosed Then
        err.Raise sckInvalidOp, "CSocket2.Protocol", "Invalid operation at current state"
    Else
        m_enmProtocol = enmProtocol
    End If
End Property

Private Function GetRemoteInfo(ByVal lngSocket As Long, ByRef lngRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String) As Boolean
'Retrieves remote info from a connected socket.
'If succeeds returns TRUE and loads the arguments.
'If fails returns FALSE and arguments are not loaded.
    Dim lngResult As Long
    Dim tmpSa As sockaddr
    GetRemoteInfo = False
    lngResult = getpeername2(lngSocket, tmpSa, LenB(tmpSa))
    If lngResult = 0 Then
        GetRemoteInfo = True
        GetRemoteInfoFromSI tmpSa, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
    Else
       lngRemotePort = 0
       strRemoteHostIP = ""
       strRemoteHost = ""
    End If
End Function

Private Function RecvDataToBuffer() As Long
'This function retrieves data from the Winsock buffer
'into the class local buffer. The function returns number
'of bytes retrieved (received).
    Dim arrBuffer() As Byte
    Dim lngBytesReceived As Long
    Dim strBuffTemporal As String
    ReDim arrBuffer(m_lngRecvBufferLen - 1)
    lngBytesReceived = recv(m_lngSocketHandle, arrBuffer(0), m_lngRecvBufferLen, 0&)
    If lngBytesReceived = SOCKET_ERROR Then
        m_enmState = sckError: Debug.Print "STATE: sckError"
        If DbgFlg Then Call LogError("STATE: sckError")
        Dim lngErrorCode As Long
        lngErrorCode = err.LastDllError
        err.Raise lngErrorCode, "CSocket2.RecvDataToBuffer", GetErrorDescription(lngErrorCode)
    ElseIf lngBytesReceived > 0 Then
        strBuffTemporal = StrConv(arrBuffer(), vbUnicode)
        m_strRecvBuffer = m_strRecvBuffer & Left$(strBuffTemporal, lngBytesReceived)
        RecvDataToBuffer = lngBytesReceived
    End If
End Function

Private Sub SendBufferedData()
'Check which protocol we are using to decide which
'function should handle the data sending.
    If m_enmProtocol = sckTCPProtocol Then
        SendBufferedDataTCP
    Else
        SendBufferedDataUDP
    End If
End Sub
Private Sub SendBufferedDataUDP()
'Send buffered data if we are using UDP protocol.
    Dim strIP As String
    Dim arrData() As Byte
    Dim lngBufferLength As Long
    Dim lngResult As Long
    Dim lngErrorCode As Long
    Dim lngTotalSent As Long
    If Len(m_strRemoteHostIP) > 0 Then
        strIP = GetIPFromHost(m_strRemoteHostIP, Str$(m_lngRemotePort), 0)
    ElseIf Len(m_strRemoteHost) > 0 Then
        strIP = GetIPFromHost(m_strRemoteHost, Str$(m_lngRemotePort), m_IPv)
    Else
        MsgBox "No Host Identified!"
        Exit Sub
    End If
    If Len(strIP) = 0 Then
        m_strSendBuffer = ""
        err.Raise sckInvalidArg, "CSocket2.SendBufferedDataUDP", "Invalid argument"
    End If
    lngBufferLength = Len(m_strSendBuffer)
    arrData() = StrConv(m_strSendBuffer, vbFromUnicode)
    m_strSendBuffer = ""
    lngResult = sendto2(m_lngSocketHandle, arrData(0), lngBufferLength, 0&, Sa(0), LenB(Sa(0)))
    If lngResult = SOCKET_ERROR Then
        lngErrorCode = err.LastDllError
        m_enmState = sckError: Debug.Print "STATE: sckError"
        If DbgFlg Then Call LogError("STATE: sckError")
        Dim blnCancelDisplay As Boolean
        blnCancelDisplay = True
        RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "CSocket2.SendBufferedDataUDP", "", 0, blnCancelDisplay)
        If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "CSocket2.SendBufferedDataUDP"
    Else
        Debug.Print "OK Bytes sent: " & lngResult
        If DbgFlg Then Call LogError("OK Bytes sent:" & Str$(lngResult))
        lngTotalSent = lngTotalSent + lngResult
        If Len(m_strSendBuffer) > lngResult Then
            m_strSendBuffer = Mid$(m_strSendBuffer, lngResult + 1)
        Else
            Debug.Print "OK Finished SENDING"
            If DbgFlg Then Call LogError("OK Finished SENDING")
            m_strSendBuffer = ""
            Dim lngTemp As Long
            lngTemp = lngTotalSent
            lngTotalSent = 0
            RaiseEvent SendProgress(lngTemp, 0)
            RaiseEvent SendComplete
        End If
    End If
End Sub


'Send buffered data if we are using TCP protocol.
Private Sub SendBufferedDataTCP()
    Dim arrData()       As Byte
    Dim lngBufferLength As Long
    Dim lngResult    As Long
    Dim lngTotalSent As Long
    Do Until lngResult = SOCKET_ERROR Or Len(m_strSendBuffer) = 0
        lngBufferLength = Len(m_strSendBuffer)
        If lngBufferLength > m_lngSendBufferLen Then
            lngBufferLength = m_lngSendBufferLen
            arrData() = StrConv(Left$(m_strSendBuffer, m_lngSendBufferLen), vbFromUnicode)
        Else
            arrData() = StrConv(m_strSendBuffer, vbFromUnicode)
        End If
        lngResult = send(m_lngSocketHandle, arrData(0), lngBufferLength, 0&)
        If lngResult = SOCKET_ERROR Then
            Dim lngErrorCode As Long
            lngErrorCode = err.LastDllError
            If lngErrorCode = WSAEWOULDBLOCK Then
                Debug.Print "WARNING: Send buffer full, waiting..."
                If DbgFlg Then Call LogError("WARNING: Send buffer full, waiting...")
                If lngTotalSent > 0 Then RaiseEvent SendProgress(lngTotalSent, Len(m_strSendBuffer))
            Else
                m_enmState = sckError: Debug.Print "STATE: sckError"
                If DbgFlg Then Call LogError("STATE: sckError")
                Dim blnCancelDisplay As Boolean
                blnCancelDisplay = True
                RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "CSocket2.SendBufferedData", "", 0, blnCancelDisplay)
                If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "CSocket2.SendBufferedData"
            End If
        Else
            Debug.Print "OK Bytes sent: " & lngResult
            If DbgFlg Then Call LogError("OK Bytes sent:" & Str$(lngResult))
            lngTotalSent = lngTotalSent + lngResult
            If Len(m_strSendBuffer) > lngResult Then
                m_strSendBuffer = Mid$(m_strSendBuffer, lngResult + 1)
            Else
                Debug.Print "OK Finished SENDING"
                If DbgFlg Then Call LogError("OK Finished SENDING")
                m_strSendBuffer = ""
                Dim lngTemp As Long
                lngTemp = lngTotalSent
                lngTotalSent = 0
                RaiseEvent SendProgress(lngTemp, 0)
                RaiseEvent SendComplete
            End If
        End If
    Loop
End Sub


Public Sub Accept2(requestID As Long)
    If m_enmState <> sckClosed Then
        err.Raise sckInvalidOp, "CSocket2.Accept", "Invalid operation at current state"
    End If
    Dim lngResult As Long
    Dim udtSockAddr As sockaddr_in
    Dim lngErrorCode As Long
    m_lngSocketHandle = requestID
    m_enmProtocol = sckTCPProtocol
    ProcessOptions
    If Not mWinsock2.IsAcceptRegistered(requestID) Then
        If IsSocketRegistered(requestID) Then
            err.Raise sckBadState, "CSocket2.Accept", "Wrong protocol or connection state for the requested transaction or request"
        Else
            m_blnAcceptClass = True
            m_enmState = sckConnected: Debug.Print "STATE: sckConnected"
            If DbgFlg Then Call LogError("STATE: sckConnected")
            mWinsock2.RegisterSocket m_lngSocketHandle, ObjPtr(Me), False
            Exit Sub
        End If
    End If
    Dim clsSocket As cSocket2
    Set clsSocket = GetAcceptClass(requestID)
    mWinsock2.UnregisterAccept requestID
    If lngResult = SOCKET_ERROR Then
        lngErrorCode = err.LastDllError
    End If
    GetRemoteInfo m_lngSocketHandle, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
    Debug.Print "Remote Host: " & m_strRemoteHostIP
    If DbgFlg Then Call LogError("Remote Host: " & m_strRemoteHostIP)
    m_enmState = sckConnected: Debug.Print "STATE: sckConnected"
    If DbgFlg Then Call LogError("STATE: sckConnected")
    If clsSocket.BytesReceived > 0 Then
        clsSocket.GetData m_strRecvBuffer
    End If
    mWinsock2.Subclass_ChangeOwner requestID, ObjPtr(Me)
    If Len(m_strRecvBuffer) > 0 Then RaiseEvent DataArrival(Len(m_strRecvBuffer))
    If clsSocket.State = sckClosing Then
        m_enmState = sckClosing: Debug.Print "STATE: sckClosing"
        If DbgFlg Then Call LogError("STATE: sckClosing")
        RaiseEvent CloseSck
    End If
    Set clsSocket = Nothing
End Sub

Public Property Get State() As SockState
    State = m_enmState
End Property

Public Property Get Tag() As String
    Tag = m_strTag
End Property

Public Property Let Tag(ByVal strTag As String)
    m_strTag = strTag
End Property

Public Property Get SocketHandle() As Long
    SocketHandle = m_lngSocketHandle
End Property

Public Property Get Index() As Long
    Index = m_lngLocalIPIndex
End Property

Public Property Get IPv6Flg() As Long
    IPv6Flg = m_IPv
End Property

Public Property Let Index(ByVal IPIndex As Long)
   m_lngLocalIPIndex = IPIndex
End Property

Public Property Let IPv6Flg(ByVal IPvFlg As Long)
    m_IPv = IPvFlg
End Property

Public Property Get BytesReceived() As Long
    If m_enmProtocol = sckTCPProtocol Then
        BytesReceived = Len(m_strRecvBuffer)
    Else
        BytesReceived = GetBufferLenUDP
    End If
End Property


Private Sub Class_Initialize()
    'socket's handle default value
    m_lngSocketHandle = INVALID_SOCKET
    'Set IPv Flag default value
    IPv6Flg = 4 'Defaults to IPv4
    'initiate resolution collection
    Set m_colWaitingResolutions = New Collection
    'initiate processes and winsock service
    mWinsock2.InitiateProcesses
End Sub


Private Function AllocateMemory() As Long
'Allocate some memory for HOSTEN structure and returns
'a pointer to this buffer if no error occurs.
'Returns 0 if it fails.
    m_lngMemoryHandle = GlobalAlloc(GMEM_FIXED, MAXGETHOSTSTRUCT)
    If m_lngMemoryHandle <> 0 Then
        m_lngMemoryPointer = GlobalLock(m_lngMemoryHandle)
        If m_lngMemoryPointer <> 0 Then
            GlobalUnlock (m_lngMemoryHandle)
            AllocateMemory = m_lngMemoryPointer
        Else
            GlobalFree (m_lngMemoryHandle)
            AllocateMemory = m_lngMemoryPointer '0
        End If
    Else
        AllocateMemory = m_lngMemoryHandle '0
    End If
End Function

Public Function GetLocalHostName() As String
    Dim strHostNameBuf As String * LOCAL_HOST_BUFF
    Dim lngResult As Long
    lngResult = gethostname(strHostNameBuf, LOCAL_HOST_BUFF)
    If lngResult = SOCKET_ERROR Then
        GetLocalHostName = vbNullString
        Dim lngErrorCode As Long
        lngErrorCode = err.LastDllError
        err.Raise lngErrorCode, "CSocket2.GetLocalHostName", GetErrorDescription(lngErrorCode)
    Else
        GetLocalHostName = Left(strHostNameBuf, InStr(1, strHostNameBuf, Chr(0)) - 1)
    End If
End Function


Private Function GetBufferLenUDP() As Long
'Returns winsock incoming buffer length from an UDP socket.
    Dim lngResult As Long
    Dim lngBuffer As Long
    lngResult = ioctlsocket(m_lngSocketHandle, FIONREAD, lngBuffer)
    If lngResult = SOCKET_ERROR Then
        GetBufferLenUDP = 0
    Else
        GetBufferLenUDP = lngBuffer
    End If
End Function

Public Sub GetData(ByRef data As Variant, Optional varType As Variant, Optional maxLen As Variant)
    If m_enmProtocol = sckTCPProtocol Then
        If m_enmState <> sckConnected And Not m_blnAcceptClass Then
            err.Raise sckBadState, "CSocket2.GetData", "Wrong protocol or connection state for the requested transaction or request"
            Exit Sub
        End If
    Else
        If m_enmState <> sckOpen Then
            err.Raise sckBadState, "CSocket2.GetData", "Wrong protocol or connection state for the requested transaction or request"
            Exit Sub
        End If
        If GetBufferLenUDP = 0 Then Exit Sub
    End If
    If Not IsMissing(maxLen) Then
        If IsNumeric(maxLen) Then
            If CLng(maxLen) < 0 Then
                err.Raise sckInvalidArg, "CSocket2.GetData", "The argument passed to a function was not in the correct format or in the specified range."
            End If
        Else
            If m_enmProtocol = sckTCPProtocol Then
                maxLen = Len(m_strRecvBuffer)
            Else
                maxLen = GetBufferLenUDP
            End If
        End If
    End If
    Dim lngBytesRecibidos  As Long
    lngBytesRecibidos = RecvData(data, False, varType, maxLen)
    Debug.Print "OK Bytes obtained from buffer: " & lngBytesRecibidos
    If DbgFlg Then Call LogError("OK Bytes obtained from buffer:" & Str$(lngBytesRecibidos))
End Sub
Private Sub EmptyBuffer()
'Empty winsock incoming buffer from an UDP socket.
    Dim B As Byte
    recv m_lngSocketHandle, B, Len(B), 0&
End Sub

Private Function BuildArray(ByVal Size As Long, ByVal blnPeek As Boolean, ByRef lngErrorCode As Long) As Byte()
'Returns a byte array of Size bytes filled with incoming buffer data.
    Dim strdata As String
    If m_enmProtocol = sckTCPProtocol Then
        strdata = Left$(m_strRecvBuffer, CLng(Size))
        BuildArray = StrConv(strdata, vbFromUnicode)
        If Not blnPeek Then
            m_strRecvBuffer = Mid$(m_strRecvBuffer, Size + 1)
        End If
    Else 'UDP protocol
        Dim arrBuffer() As Byte
        Dim lngResult As Long
        Dim tmpSa As sockaddr
        Dim lngFlags As Long
        If blnPeek Then lngFlags = MSG_PEEK
        ReDim arrBuffer(Size - 1)
        lngResult = recvfrom2(m_lngSocketHandle, arrBuffer(0), Size, lngFlags, tmpSa, LenB(tmpSa))
        If lngResult = SOCKET_ERROR Then
            lngErrorCode = err.LastDllError
        End If
        BuildArray = arrBuffer
        GetRemoteInfoFromSI tmpSa, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
    End If
End Function
Private Sub GetRemoteInfoFromSI(ByRef newSa As sockaddr, ByRef lngRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String)
'Gets remote info from a sockaddr_in structure.
    Dim lngResult As Long
    Dim Sin4 As sockaddr_in
    Dim Sin6 As sockaddr_in6
    Dim aLen As Long
    Dim bBuffer() As Byte
    If newSa.sa_family = AF_INET6 Then
        aLen = INET6_ADDRSTRLEN
        CopyMemory Sin6, newSa, LenB(Sin6)  'Save to sockaddr_in6
        lngRemotePort = IntegerToUnsigned(ntohs(Sin6.sin6_port))
        ReDim bBuffer(0 To aLen - 1)    'Resize string buffer
        'Get IPv6 address as string
        lngResult = inet_ntop(AF_INET6, Sin6.sin6_addr, bBuffer(0), aLen)
    Else 'Must be IPv4
        aLen = INET_ADDRSTRLEN
        CopyMemory Sin4, newSa, LenB(Sin4)  'Save to sockaddr_in
        lngRemotePort = IntegerToUnsigned(ntohs(Sin4.sin_port))
        ReDim bBuffer(0 To aLen - 1)  'Resize string buffer
        'Get IPv4 address as string
        lngResult = inet_ntop(AF_INET, Sin4.sin_addr, bBuffer(0), aLen)
    End If
    If lngResult Then strRemoteHostIP = StringFromPointer(lngResult)
    m_strRemoteHost = ""
End Sub

Private Function RecvData(ByRef data As Variant, ByVal blnPeek As Boolean, Optional varClass As Variant, Optional maxLen As Variant) As Long
'This function is to retrieve data from the buffer. If we are using TCP
'then the data is retrieved from a local buffer (m_strRecvBuffer). If we
'are using UDP the data is retrieved from winsock buffer.
'It can be called by two public methods of the class - GetData and PeekData.
'Behavior of the function is defined by the blnPeek argument. If a value of
'that argument is TRUE, the function returns number of bytes in the
'buffer, and copy data from that buffer into the data argument.
'If a value of the blnPeek is FALSE, then this function returns number of
'bytes received, and move data from the buffer into the data
'argument. MOVE means that data will be removed from the buffer.
    Dim blnMaxLenMiss   As Boolean
    Dim blnClassMiss As Boolean
    Dim strRecvData     As String
    Dim lngBufferLen    As Long
    Dim arrBuffer()     As Byte
    Dim lngErrorCode    As Long
    If m_enmProtocol = sckTCPProtocol Then
        lngBufferLen = Len(m_strRecvBuffer)
    Else
        lngBufferLen = GetBufferLenUDP
    End If
    blnMaxLenMiss = IsMissing(maxLen)
    blnClassMiss = IsMissing(varClass)
    'Select type of data
    If varType(data) = vbEmpty Then
        If blnClassMiss Then varClass = vbArray + vbByte
    Else
        varClass = varType(data)
    End If
    'As stated on Winsock control documentation if the
    'data type passed is string or byte array type then
    'we must take into account maxLen argument.
    'If it is another type maxLen is ignored.
    If varClass = vbString Or varClass = vbArray + vbByte Then
        If blnMaxLenMiss Then 'if maxLen argument is missing
            If lngBufferLen = 0 Then
                RecvData = 0
                arrBuffer = StrConv("", vbFromUnicode)
                data = arrBuffer
                Exit Function
            Else
                RecvData = lngBufferLen
                arrBuffer = BuildArray(lngBufferLen, blnPeek, lngErrorCode)
            End If
        Else 'if maxLen argument is not missing
            If maxLen = 0 Or lngBufferLen = 0 Then
                RecvData = 0
                arrBuffer = StrConv("", vbFromUnicode)
                data = arrBuffer
                If m_enmProtocol = sckUDPProtocol Then
                    EmptyBuffer
                    err.Raise WSAEMSGSIZE, "CSocket2.RecvData", GetErrorDescription(WSAEMSGSIZE)
                End If
                Exit Function
            ElseIf maxLen > lngBufferLen Then
                RecvData = lngBufferLen
                arrBuffer = BuildArray(lngBufferLen, blnPeek, lngErrorCode)
            Else
                RecvData = CLng(maxLen)
                arrBuffer() = BuildArray(CLng(maxLen), blnPeek, lngErrorCode)
            End If
        End If
    End If
        Select Case varClass
            Case vbString
                Dim strdata As String
                strdata = StrConv(arrBuffer(), vbUnicode)
                data = strdata
            Case vbArray + vbByte
                data = arrBuffer
            Case vbBoolean
                Dim blnData As Boolean
                If LenB(blnData) > lngBufferLen Then Exit Function
                arrBuffer = BuildArray(LenB(blnData), blnPeek, lngErrorCode)
                RecvData = LenB(blnData)
                CopyMemory blnData, arrBuffer(0), LenB(blnData)
                data = blnData
            Case vbByte
                Dim bytData As Byte
                If LenB(bytData) > lngBufferLen Then Exit Function
                arrBuffer = BuildArray(LenB(bytData), blnPeek, lngErrorCode)
                RecvData = LenB(bytData)
                CopyMemory bytData, arrBuffer(0), LenB(bytData)
                data = bytData
            Case vbCurrency
                Dim curData As Currency
                If LenB(curData) > lngBufferLen Then Exit Function
                arrBuffer = BuildArray(LenB(curData), blnPeek, lngErrorCode)
                RecvData = LenB(curData)
                CopyMemory curData, arrBuffer(0), LenB(curData)
                data = curData
            Case vbDate
                Dim datData As Date
                If LenB(datData) > lngBufferLen Then Exit Function
                arrBuffer = BuildArray(LenB(datData), blnPeek, lngErrorCode)
                RecvData = LenB(datData)
                CopyMemory datData, arrBuffer(0), LenB(datData)
                data = datData
            Case vbDouble
                Dim dblData As Double
                If LenB(dblData) > lngBufferLen Then Exit Function
                arrBuffer = BuildArray(LenB(dblData), blnPeek, lngErrorCode)
                RecvData = LenB(dblData)
                CopyMemory dblData, arrBuffer(0), LenB(dblData)
                data = dblData
            Case vbInteger
                Dim intData As Integer
                If LenB(intData) > lngBufferLen Then Exit Function
                arrBuffer = BuildArray(LenB(intData), blnPeek, lngErrorCode)
                RecvData = LenB(intData)
                CopyMemory intData, arrBuffer(0), LenB(intData)
                data = intData
            Case vbLong
                Dim lngData As Long
                If LenB(lngData) > lngBufferLen Then Exit Function
                arrBuffer = BuildArray(LenB(lngData), blnPeek, lngErrorCode)
                RecvData = LenB(lngData)
                CopyMemory lngData, arrBuffer(0), LenB(lngData)
                data = lngData
            Case vbSingle
                Dim sngData As Single
                If LenB(sngData) > lngBufferLen Then Exit Function
                arrBuffer = BuildArray(LenB(sngData), blnPeek, lngErrorCode)
                RecvData = LenB(sngData)
                CopyMemory sngData, arrBuffer(0), LenB(sngData)
                data = sngData
            Case Else
                err.Raise sckUnsupported, "CSocket2.RecvData", "Unsupported variant type."
        End Select
    'if BuildArray returns an error is handled here
    If lngErrorCode <> 0 Then
        err.Raise lngErrorCode, "CSocket2.RecvData", GetErrorDescription(lngErrorCode)
    End If
End Function
Private Sub FreeMemory()
    'Free memory allocated by AllocateMemory
    If m_lngMemoryHandle <> 0 Then
        m_lngMemoryHandle = 0
        m_lngMemoryPointer = 0
        GlobalFree m_lngMemoryHandle
    End If
End Sub
Private Function GetLocalPort(ByVal lngSocket As Long) As Long
'Returns local port from a connected or bound socket.
'Returns SOCKET_ERROR if fails.
    Dim newSa As sockaddr
    Dim Sin4 As sockaddr_in
    Dim Sin6 As sockaddr_in6
    Dim lngResult As Long
    lngResult = getsockname2(lngSocket, newSa, LenB(newSa))
    If lngResult = SOCKET_ERROR Then
        GetLocalPort = SOCKET_ERROR
    Else
        If newSa.sa_family = AF_INET6 Then
            CopyMemory Sin6, newSa, LenB(Sin6)  'Save to sockaddr_in6
            GetLocalPort = mWinsock2.IntegerToUnsigned(ntohs(Sin6.sin6_port))
        Else
            CopyMemory Sin4, newSa, LenB(Sin4)  'Save to sockaddr_in
            GetLocalPort = mWinsock2.IntegerToUnsigned(ntohs(Sin4.sin_port))
        End If
    End If
End Function

Private Function SocketExists() As Boolean
'Tries to create a socket if there isn't one yet and registers
'it to the control list.
'Returns TRUE if it has success
    SocketExists = True
    Dim lngResult As Long
    Dim lngErrorCode As Long
    'check if there is a socket already
    If m_lngSocketHandle = INVALID_SOCKET Then
        'decide what kind of socket we are creating, TCP or UDP
        If m_enmProtocol = sckTCPProtocol Then
            If m_IPv = 6 Then
                lngResult = Socket(AF_INET6, SOCK_STREAM, IPPROTO_TCP)
            Else
                lngResult = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
            End If
        Else
            If m_IPv = 6 Then
                lngResult = Socket(AF_INET6, SOCK_DGRAM, IPPROTO_UDP)
            Else
                lngResult = Socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
            End If
        End If
        If lngResult = INVALID_SOCKET Then
            m_enmState = sckError: Debug.Print "STATE: sckError"
            Debug.Print "ERROR trying to create socket"
            If DbgFlg Then Call LogError("ERROR trying to create socket")
            SocketExists = False
            lngErrorCode = err.LastDllError
            Dim blnCancelDisplay As Boolean
            blnCancelDisplay = True
            RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "CSocket2.SocketExists", "", 0, blnCancelDisplay)
            If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "CSocket2.SocketExists"
        Else
            Debug.Print "OK Created socket: " & lngResult
            If DbgFlg Then Call LogError("OK Created socket:" & Str$(lngResult))
            m_lngSocketHandle = lngResult
            'set and get some socket options
            ProcessOptions
            SocketExists = mWinsock2.RegisterSocket(m_lngSocketHandle, ObjPtr(Me), True)
        End If
    End If
End Function
Private Sub ProcessOptions()
'Retrieves some socket options.
'If it is an UDP socket also sets SO_BROADCAST option.
    Dim lngResult As Long
    Dim lngBuffer As Long
    Dim lngErrorCode As Long
    If m_enmProtocol = sckTCPProtocol Then
        lngResult = getsockopt(m_lngSocketHandle, SOL_SOCKET, SO_RCVBUF, lngBuffer, LenB(lngBuffer))
        If lngResult = SOCKET_ERROR Then
            lngErrorCode = err.LastDllError
            err.Raise lngErrorCode, "CSocket2.ProcessOptions", GetErrorDescription(lngErrorCode)
        Else
            m_lngRecvBufferLen = lngBuffer
        End If
        lngResult = getsockopt(m_lngSocketHandle, SOL_SOCKET, SO_SNDBUF, lngBuffer, LenB(lngBuffer))
        If lngResult = SOCKET_ERROR Then
            lngErrorCode = err.LastDllError
            err.Raise lngErrorCode, "CSocket2.ProcessOptions", GetErrorDescription(lngErrorCode)
        Else
            m_lngSendBufferLen = lngBuffer
        End If
    Else
        lngBuffer = 1
        lngResult = setsockopt(m_lngSocketHandle, SOL_SOCKET, SO_BROADCAST, lngBuffer, LenB(lngBuffer))
        lngResult = getsockopt(m_lngSocketHandle, SOL_SOCKET, SO_MAX_MSG_SIZE, lngBuffer, LenB(lngBuffer))
        If lngResult = SOCKET_ERROR Then
            lngErrorCode = err.LastDllError
            err.Raise lngErrorCode, "CSocket2.ProcessOptions", GetErrorDescription(lngErrorCode)
        Else
            m_lngRecvBufferLen = lngBuffer
            m_lngSendBufferLen = lngBuffer
        End If
    End If
    Debug.Print "Winsock buffer size for sends: " & m_lngRecvBufferLen
    If DbgFlg Then Call LogError("Winsock buffer size for sends:" & Str$(m_lngRecvBufferLen))
    Debug.Print "Winsock buffer size for receives: " & m_lngSendBufferLen
    If DbgFlg Then Call LogError("Winsock buffer size for receives:" & Str$(m_lngRecvBufferLen))
End Sub


Private Sub DestroySocket()
    'Destroys the socket if it exists and unregisters it
    'from control list.
    If Not m_lngSocketHandle = INVALID_SOCKET Then
        Dim lngResult As Long
        lngResult = closesocket(m_lngSocketHandle)
        If lngResult = SOCKET_ERROR Then
            m_enmState = sckError: Debug.Print "STATE: sckError"
            If DbgFlg Then Call LogError("STATE: sckError")
            Dim lngErrorCode As Long
            lngErrorCode = err.LastDllError
            err.Raise lngErrorCode, "CSocket2.DestroySocket", GetErrorDescription(lngErrorCode)
        Else
            Debug.Print "OK Destroyed socket " & m_lngSocketHandle
            mWinsock2.UnregisterSocket m_lngSocketHandle
            m_lngSocketHandle = INVALID_SOCKET
        End If
    End If
End Sub
Public Function GetIPFromHost(sHost As String, sPort As String, Optional Options As Long) As String
    Dim blnCancelDisplay As Boolean
    Dim Sin4 As sockaddr_in
    Dim Sin6 As sockaddr_in6
    Dim ptrResult As Long
    Dim lRet As Long
    Dim aLen As Long
    Dim bBuffer() As Byte
    Dim N%
    ReDim Sa(10)
    Dim tmpArray(10) As String
    ZeroMemory Hints, Len(Hints)
    Select Case Options
        Case 4
            Hints.ai_family = AF_INET   'Set up for IPv4 Listen
        Case 6
            Hints.ai_family = AF_INET6  'Set up for IPv6 Listen
        Case Else
            Hints.ai_family = AF_UNSPEC 'Unspecified
    End Select
    If Len(sHost) = 0 Then
        Hints.ai_flags = AI_PASSIVE  ' fill in 0 IP for me
    End If
    lRet = getaddrinfo(sHost, sPort, VarPtr(Hints), ptrResult)
    If lRet = 0 Then
        Debug.Print "GetAddrInfo retrieved!"
        If DbgFlg Then Call LogError("GetAddrInfo retrieved!")
    Else
        blnCancelDisplay = True
        RaiseEvent Error(lRet, GetErrorDescription(lRet), 0, "CSocket2.GetIPFromHost", "", 0, blnCancelDisplay)
        If blnCancelDisplay = False Then MsgBox GetErrorDescription(lRet), vbOKOnly, "CSocket2.GetIPFromHost"
        Debug.Print "getaddrinfo() failed with error: " & lRet
        If DbgFlg Then Call LogError("getaddrinfo() failed with error:" & Str$(lRet))
        Exit Function
    End If
    Hints.ai_next = ptrResult   'Point to first structure in linked list
    While Hints.ai_next <> 0 'Enable loop to get all addresses from linked list
        CopyMemory Hints, ByVal Hints.ai_next, LenB(Hints)
            'For N% = 0 To 15: Debug.Print Hex$(PeekB(Hints.ai_addr + N%)) & "|";: Next N%
            '2|0|0|50|60|35|60|32|0|0|0|0|0|0|0|0| = 96.53.96.50:80
        If Hints.ai_family = AF_INET Then
            aLen = INET_ADDRSTRLEN
            CopyMemory Sin4, ByVal Hints.ai_addr, LenB(Sin4)  'Save to sockaddr_in
            CopyMemory Sa(N%), ByVal Hints.ai_addr, LenB(Sin4)    'Save to sockaddr
            ReDim bBuffer(0 To aLen - 1)  'Resize string buffer
            'Get IPv4 address as string
            lRet = inet_ntop(Hints.ai_family, Sin4.sin_addr, bBuffer(0), aLen)
            If lRet Then tmpArray(N%) = StringFromPointer(lRet): N% = N% + 1
        ElseIf Hints.ai_family = AF_INET6 Then
            aLen = INET6_ADDRSTRLEN
            CopyMemory Sin6, ByVal Hints.ai_addr, LenB(Sin6) 'Save to sockaddr_in6
            CopyMemory Sa(N%), ByVal Hints.ai_addr, LenB(Sin6)   'Save to sockaddr
            ReDim bBuffer(0 To aLen - 1)    'Resize string buffer
            'Get IPv6 address as string
            lRet = inet_ntop(Hints.ai_family, Sin6.sin6_addr, bBuffer(0), aLen)
            If lRet Then
                If Sin6.sin6_scope_id > 0 Then
                    tmpArray(N%) = StringFromPointer(lRet) & "%" & Trim(Str$(Sin6.sin6_scope_id))
                Else
                    tmpArray(N%) = StringFromPointer(lRet)
                End If
                N% = N% + 1
            End If
        End If
    Wend
    ReDim Preserve Sa(N% - 1)
    For N% = 0 To UBound(Sa)
        GetIPFromHost = GetIPFromHost + tmpArray(N%) + Chr$(0)
        Debug.Print Str$(N%) & ": " & tmpArray(N%)
        If DbgFlg Then Call LogError(Str$(N%) & ": " & tmpArray(N%))
    Next N%
    lRet = freeaddrinfo(ptrResult) ' free the linked list
End Function

Public Sub SendData(data As Variant)
    Dim arrData() As Byte 'We store the data here before send it
    If m_enmProtocol = sckTCPProtocol Then
        If m_enmState <> sckConnected Then
            err.Raise sckBadState, "CSocket2.SendData", "Wrong protocol or connection state for the requested transaction or request"
            Exit Sub
        End If
    Else 'If we use UDP we create a socket if there isn't one yet
        If Not SocketExists Then Exit Sub
        If Not BindInternal2 Then Exit Sub
        m_enmState = sckOpen: Debug.Print "STATE: sckOpen"
        If DbgFlg Then Call LogError("STATE: sckOpen")
    End If
    'We need to convert data variant into a byte array
    Select Case varType(data)
        Case vbString
            Dim strdata As String
            strdata = CStr(data)
            If Len(strdata) = 0 Then Exit Sub
            ReDim arrData(Len(strdata) - 1)
            arrData() = StrConv(strdata, vbFromUnicode)
        Case vbArray + vbByte
            Dim strArray As String
            strArray = StrConv(data, vbUnicode)
            If Len(strArray) = 0 Then Exit Sub
            arrData() = StrConv(strArray, vbFromUnicode)
        Case vbBoolean
            Dim blnData As Boolean
            blnData = CBool(data)
            ReDim arrData(LenB(blnData) - 1)
            CopyMemory arrData(0), blnData, LenB(blnData)
        Case vbByte
            Dim bytData As Byte
            bytData = CByte(data)
            ReDim arrData(LenB(bytData) - 1)
            CopyMemory arrData(0), bytData, LenB(bytData)
        Case vbCurrency
            Dim curData As Currency
            curData = CCur(data)
            ReDim arrData(LenB(curData) - 1)
            CopyMemory arrData(0), curData, LenB(curData)
        Case vbDate
            Dim datData As Date
            datData = CDate(data)
            ReDim arrData(LenB(datData) - 1)
            CopyMemory arrData(0), datData, LenB(datData)
        Case vbDouble
            Dim dblData As Double
            dblData = CDbl(data)
            ReDim arrData(LenB(dblData) - 1)
            CopyMemory arrData(0), dblData, LenB(dblData)
        Case vbInteger
            Dim intData As Integer
            intData = CInt(data)
            ReDim arrData(LenB(intData) - 1)
            CopyMemory arrData(0), intData, LenB(intData)
        Case vbLong
            Dim lngData As Long
            lngData = CLng(data)
            ReDim arrData(LenB(lngData) - 1)
            CopyMemory arrData(0), lngData, LenB(lngData)
        Case vbSingle
            Dim sngData As Single
            sngData = CSng(data)
            ReDim arrData(LenB(sngData) - 1)
            CopyMemory arrData(0), sngData, LenB(sngData)
        Case Else
            err.Raise sckUnsupported, "CSocket2.SendData", "Unsupported variant type."
    End Select
    'if there's already something in the buffer that means we are
    'already sending data, so we put the new data in the buffer
    'and exit silently
    If Len(m_strSendBuffer) > 0 Then
        m_strSendBuffer = m_strSendBuffer + StrConv(arrData(), vbUnicode)
        Exit Sub
    Else
        m_strSendBuffer = m_strSendBuffer + StrConv(arrData(), vbUnicode)
    End If
    'send the data
    SendBufferedData
End Sub

Private Sub Class_Terminate()
    'clean hostname resolution system
    If Not m_blnAcceptClass Then DestroySocket
    'clean processes and finish winsock service
    mWinsock2.FinalizeProcesses
    'clean resolution collection
    Set m_colWaitingResolutions = Nothing
End Sub



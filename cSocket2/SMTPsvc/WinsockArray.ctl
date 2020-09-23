VERSION 5.00
Object = "{916EA9DB-39C1-4F2C-8377-A1B5F8EF2F4D}#1.0#0"; "CSocket.ocx"
Begin VB.UserControl WinsockArray 
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1245
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   810
   ScaleWidth      =   1245
   ToolboxBitmap   =   "WinsockArray.ctx":0000
   Begin CSocketOCX.Socket Sockets 
      Index           =   0
      Left            =   720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   0
      Picture         =   "WinsockArray.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "WinsockArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' -------------------------------------------------------------------------------
' File...........: WinsockArray.ctl
' Author.........: J.A. Coutts
' Created........: 10/10/2011
' Version........: 1.0
' Website........: http://www.yellowhead.com
' Contact........: allecnarf@hotmail.com
'
' Based upon WinsockArray usercontrol V1.3 written by Will Barden.
' Handles a control array of cSocket objects, and manages events between them
' and the parent.
'
' Unlike MSWinSck, the cSocket control does not return the connecting IP address
' when the CloseSck event is raised. Also, cSocket listens on the default IP
' address, not the zero address. However, the cSocket control does not appear
' to strand the listening socket with a 10035 error (Error10035: Socket is
' non-blocking and the specified operation will block) like MSWinSck.
'
' The LastError property is only used when RaiseErrors is turned off.
'
' All the methods will return a boolean value, and may raise errors if
' RaiseErrors is turned on.
'
' No direct access to the clients is given to the parent, since it could give
' rise to situations that mess up internal usercontrol counters, and won't allow
' raising of proper events to the parent.
'
' MaxClients is used to limit the amount of concurrent clients (and also the
' number of Winsock objects loaded into the array).
'
' If any error is fired during the use of a particular Winsock object, it is
' closed and the client disconnected.
'
' BytesSent and BytesReceived are total values for all the sockets.
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
' Events.
' ------------------------------------------------------------------------------
'
Public Event NewConnection(ByVal lngIndex As Long, ByRef blnCancel As Boolean)
Public Event ConnectionClosed(ByVal lngIndex As Long)
Public Event DataArrival(ByVal lngIndex As Long, ByVal strdata As String)
Public Event WinsockError(ByVal lngIndex As Long, ByVal lngNumber As Long, ByVal strSource As String, ByVal strDescription As String)
Public Event SendProgress(ByVal lngIndex As Long, ByVal lngBytesSent As Long, ByVal lngBytesRemaining As Long)
Public Event SendComplete(ByVal lngIndex As Long)
'
' ------------------------------------------------------------------------------
' Private variables.
' ------------------------------------------------------------------------------
'
Private m_blnRaiseErrors   As Boolean
Private m_strLastError     As String
Private m_lngLocalPort     As Long
Private m_lngMaxClients    As Long
Private m_lngActiveClients As Long
Private m_lngBytesReceived As Long
Private m_lngBytesSent     As Long
Public Property Get RaiseErrors() As Boolean
    RaiseErrors = m_blnRaiseErrors
End Property
Public Property Let RaiseErrors(ByVal Value As Boolean)
    m_blnRaiseErrors = Value
End Property
Public Property Get LastError() As String
    LastError = m_strLastError
End Property
Public Property Let LastError(ByVal Value As String)
    m_strLastError = Value
End Property
Public Property Get MaxClients() As Long
    MaxClients = m_lngMaxClients
End Property
Public Property Let MaxClients(ByVal Value As Long)
    m_lngMaxClients = Value
End Property
Public Property Get ActiveClients() As Long ' Read only.
    ActiveClients = m_lngActiveClients
End Property
Public Property Get LocalPort() As Long
    LocalPort = m_lngLocalPort
End Property
Public Property Let LocalPort(ByVal Value As Long)
    m_lngLocalPort = Value
End Property
Public Property Get BytesReceived() As Long
    BytesReceived = m_lngBytesReceived
End Property
Public Property Let BytesReceived(ByVal Value As Long)
    m_lngBytesReceived = Value
End Property
Public Property Get bytesSent() As Long
    bytesSent = m_lngBytesSent
End Property
Public Property Let bytesSent(ByVal Value As Long)
    m_lngBytesSent = Value
End Property
Public Property Get SocketCount() As Long ' Read only
    SocketCount = Sockets.Count - 1
End Property
Public Property Get RemoteHost(ByVal lngIndex As Long) As String  ' Read only.
    If (lngIndex > 0) And (lngIndex <= Sockets.UBound) Then
        RemoteHost = Sockets(lngIndex).RemoteHost
    End If
End Property
Public Property Get RemoteHostIP(ByVal lngIndex As Long) As String  ' Read only.
    If (lngIndex > 0) And (lngIndex <= Sockets.UBound) Then
        RemoteHostIP = Sockets(lngIndex).RemoteHostIP
    End If
End Property
Public Property Get RemotePort(ByVal lngIndex As Long) As Long  ' Read only.
    If (lngIndex > 0) And (lngIndex <= Sockets.UBound) Then
        RemotePort = Sockets(lngIndex).RemotePort
    End If
End Property
Public Function StartListening(Optional IPv6 As Boolean) As Boolean
    If IPv6 Then
        Sockets(0).IPv6Flg = 6
    Else
        Sockets(0).IPv6Flg = 4
    End If
    On Error GoTo StartError
    'Make sure we've been given a local port to listen on. Without this,
    'the call to Listen() will fail.
    If (m_lngLocalPort) Then
        'If the socket is already listening, and it's listening on the same
        'port, don't bother restarting it.
        If (Sockets(0).State <> sckListening) Or _
          (Sockets(0).LocalPort <> m_lngLocalPort) Then
            With Sockets(0)
                Call .CloseSck
                .LocalPort = m_lngLocalPort
                Call .Listen
            End With
        End If
        'Return true, since the server is now listening for clients.
        StartListening = True
    End If
    Exit Function
StartError:
   'Handle the error - either raise it or save the description.
   If (m_blnRaiseErrors) Then
      Call Err.Raise(Err.Number, Err.Source, Err.Description)
   Else
      m_strLastError = Err.Description
   End If
End Function
Public Function StopListening() As Boolean
    'Stop the listening socket so no more connection requests are received.
    Sockets(0).CloseSck
    StopListening = True
End Function
Public Function Shutdown() As Boolean
    Dim i As Long
    On Error GoTo ShutdownError
    'Close the listening socket first, so no more connection requests.
    Call Sockets(0).CloseSck
    'Now loop through all the clients, close the active ones and
    'unload them all to clear the array from memory.
    For i = 1 To Sockets.UBound
        If (Sockets(i).State <> sckClosed) Then Sockets(i).CloseSck
        Call Unload(Sockets(i))
        RaiseEvent ConnectionClosed(i)
    Next i
    'Return true if all went well.
    Shutdown = True
    Exit Function
ShutdownError:
    'Handle the error - either raise it or save the description.
    If (m_blnRaiseErrors) Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    Else
        m_strLastError = Err.Description
    End If
End Function
Public Function Send(ByVal lngIndex As Long, ByVal strdata As String) As Boolean
    On Error GoTo SendError
    'Send the data on the specified socket.
    Call Sockets(lngIndex).SendData(strdata): DoEvents
    'Return true if it went ok.
    Send = True
    Exit Function
SendError:
    'Handle the error - either raise it or save the description.
    If (m_blnRaiseErrors) Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    Else
        m_strLastError = Err.Description
    End If
End Function
Public Function SendToAll(ByVal strdata As String) As Boolean
    Dim i As Long
    On Error GoTo SendToAllError
    'Loop through the control array of clients and send the data on each one.
    For i = 1 To Sockets.UBound
        If (Sockets(i).State = sckConnected) Then
            Call Sockets(i).SendData(strdata): DoEvents
        End If
    Next i
    'Return true if it all went well.
    SendToAll = True
    Exit Function
SendToAllError:
    'Handle the error - either raise it or save the description.
    If (m_blnRaiseErrors) Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    Else
        m_strLastError = Err.Description
    End If
End Function
Public Function CloseClient(ByVal lngIndex As Long) As Boolean
    On Error GoTo CloseClientError
    'Make sure the index specified is within the range of the control array.
    If (lngIndex > 0) And (lngIndex <= Sockets.UBound) And Sockets(lngIndex).State <> sckClosed Then
        'Close the socket, update the ActiveClients property, and raise a
        'ConnectionClosed event to the parent.
        Call Sockets(lngIndex).CloseSck
        m_lngActiveClients = m_lngActiveClients - 1
        RaiseEvent ConnectionClosed(lngIndex)
    End If
    'Return true if all went well.
    CloseClient = True
    Exit Function
CloseClientError:
    'Handle the error - either raise it or save the description.
    If (m_blnRaiseErrors) Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    Else
        m_strLastError = Err.Description
    End If
End Function

Private Sub Sockets_CloseSck(Index As Integer)
    'Close the socket and raise the event to the parent.
    If (Sockets(Index).State <> sckClosed) Then
        Sockets(Index).CloseSck
        RaiseEvent ConnectionClosed(Index)
        'If this wasn't the listening socket, reduce the ActiveClients property.
        If (Index > 0) Then
            m_lngActiveClients = m_lngActiveClients - 1
        End If
    End If
End Sub

Private Sub Sockets_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i As Long
    Dim j As Long
    Dim blnLoaded As Boolean
    Dim blnCancel As Boolean
    On Error GoTo ConnectionRequestError
    'We shouldn't get ConnectionRequests on any other socket than the listener
    '(index 0), but check anyway. Also check that we're not going to exceed
    'the MaxClients property.
    If (Index = 0) And (m_lngActiveClients < m_lngMaxClients) Then
        'Check to see if we've got any sockets that are free.
        For i = 1 To Sockets.UBound
            If (Sockets(i).State = sckClosed) Then
                j = i
                Exit For
            End If
        Next i
        'If we don't have any free sockets, load another on the array.
        If (j = 0) Then
            blnLoaded = True
            Call Load(Sockets(Sockets.UBound + 1))
            j = Sockets.Count - 1
        End If
        'With the selected socket, reset it and accept the new connection.
        With Sockets(j)
            Call .CloseSck
            Call .Accept(requestID)
        End With
        'Raise the NewConnection event, passing a Cancel boolean ByRef so
        'the parent can cancel the connection if necessary.
        blnCancel = False
        RaiseEvent NewConnection(j, blnCancel)
        If (blnCancel) Then
            Call Sockets(j).CloseSck
            If (blnLoaded) Then Call Unload(Sockets(j))
        Else
            m_lngActiveClients = m_lngActiveClients + 1
        End If
    End If
    Exit Sub
ConnectionRequestError:
    'Close the Winsock that caused the error.
'    Call Winsocks(0).Close
'Possible fix for 10035 Error closing listening socket
    If (Sockets(0).State <> sckListening) Or (Sockets(0).LocalPort <> m_lngLocalPort) Then
        With Sockets(0)
            Call .CloseSck
            .LocalPort = m_lngLocalPort
            Call .Listen
        End With
    End If
    'Handle the error - raise an error or store the description.
    If (m_blnRaiseErrors) Then
        RaiseEvent WinsockError(Index, Err.Number, Err.Source, Err.Description)
    Else
        m_strLastError = Err.Description
    End If
End Sub


Private Sub Sockets_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strdata    As String
    On Error GoTo DataArrivalError
    'Save the amount of data received.
    m_lngBytesReceived = m_lngBytesReceived + bytesTotal
    'Grab the data from the specified Winsock object, and pass it to the parent.
    Call Sockets(Index).GetData(strdata)
    RaiseEvent DataArrival(Index, strdata)
    Exit Sub
DataArrivalError:
    'Close the Winsock and update the ActiveClients property.
    Call Sockets(Index).CloseSck
    If (Index > 0) Then
        m_lngActiveClients = m_lngActiveClients - 1
    End If
    'Raise an error or store the description.
    If (m_blnRaiseErrors) Then
        RaiseEvent WinsockError(Index, Err.Number, Err.Source, Err.Description)
    Else
        m_strLastError = Err.Description
    End If
End Sub

Private Sub Sockets_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Close the Winsock and update the ActiveClients property.
    Call Sockets(Index).CloseSck
    If (Index > 0) Then
        m_lngActiveClients = m_lngActiveClients - 1
    End If
    'Raise an error or store the description.
    If (m_blnRaiseErrors) Then
        RaiseEvent WinsockError(Index, Number, Source, Description)
    Else
        m_strLastError = Description
    End If
End Sub

Private Sub Sockets_SendComplete(Index As Integer)
    'Pass the event on to the parent.
    RaiseEvent SendComplete(Index)
End Sub

Private Sub Sockets_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    'Update the BytesSent property.
    m_lngBytesSent = m_lngBytesSent + bytesSent
    'Pass the event on to the parent.
    RaiseEvent SendProgress(Index, bytesSent, bytesRemaining)
End Sub

Private Sub UserControl_Initialize()
    'Call MsgBox("WinsockArray Usercontrol v2, www.WinsockVB.com", _
      vbOKOnly + vbInformation)
End Sub
Private Sub UserControl_Resize()
    'On Error Resume Next
    'Fit the control around the logo image.
    With UserControl
        .Height = imgIcon.Height
        .Width = imgIcon.Width
    End With
End Sub
Private Sub UserControl_Terminate()
    'Shutdown the server (just in case the parent hasn't called it).
    Call Shutdown
End Sub

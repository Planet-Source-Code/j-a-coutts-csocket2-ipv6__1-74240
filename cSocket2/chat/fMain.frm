VERSION 5.00
Begin VB.Form fMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Chat"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkIPv6 
      Caption         =   "IPv6"
      Height          =   300
      Left            =   3720
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox cmbAddress 
      Height          =   315
      Left            =   960
      TabIndex        =   12
      Top             =   600
      Width           =   5295
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6840
      Top             =   0
   End
   Begin VB.CheckBox chkType 
      Caption         =   "UDP"
      Height          =   300
      Left            =   3720
      TabIndex        =   11
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtLog 
      Height          =   2775
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   8895
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtMessage 
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   3840
      Width           =   7455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetIP 
      Caption         =   "GET ADDRESS"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   7080
      TabIndex        =   3
      Text            =   "225"
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "CONNECT"
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "LISTEN"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   4320
      Width           =   8895
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Port:"
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Domain:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents Server As cSocket2
Attribute Server.VB_VarHelpID = -1

Dim Int_Address() As String
Dim Test_Address As String
Dim Test_Port As Long
'Dim ListenFlg As Boolean
Private Const CB_SHOWDROPDOWN = &H14F

Private Function PeekB(ByVal lpdwData As Long) As Byte
    CopyMemory PeekB, ByVal lpdwData, 1
End Function
Private Function PeekL(ByVal lpdwData As Long) As Long
    CopyMemory PeekL, ByVal lpdwData, 4
End Function
Private Function StateConv() As String
    Select Case Server.State
        Case sckClosed '0
            StateConv = "Socket Closed!"
        Case sckOpen '1
            If Server.Protocol = sckUDPProtocol Then
                StateConv = "Socket Open for " & Server.RemoteHostIP & ":" & CStr(Server.RemotePort)
                cmdConnect.Enabled = False
                cmdListen.Enabled = False
                cmdSend.Enabled = True
                cmdClose.Enabled = True
            Else
                StateConv = "Socket Open!"
            End If
        Case sckListening '2
            StateConv = "Socket Listening!"
        Case sckConnectionPending '3
            StateConv = "Connection Pending!"
        Case sckResolvingHost '4
            StateConv = "Resolving Host!"
        Case sckHostResolved '5
            StateConv = "Host Resolved!"
        Case sckConnecting '6
            StateConv = "Socket Connecting!"
        Case sckConnected '7
            StateConv = Server.RemoteHostIP & " CONNECTED on Port" & Str$(Server.RemotePort)
        Case sckClosing '8
            StateConv = "Socket Closing!"
        Case sckError '9
            StateConv = "Socket Error!"
    End Select
End Function


Private Sub chkIPv6_Click()
    If chkIPv6 Then
        Server.IPv6Flg = 6
    Else
        Server.IPv6Flg = 4
    End If
End Sub

Private Sub chkType_Click()
    If chkType.Value Then 'If UDP then disable Connect button
        Server.Protocol = sckUDPProtocol
    Else
        Server.Protocol = sckTCPProtocol
    End If
    cmdConnect.Enabled = True
    cmdSend.Enabled = False
End Sub


Private Sub cmbAddress_Click()
    Server.Index = CLng(cmbAddress.ListIndex)
    Debug.Print "Index =" & Str$(Server.Index)
End Sub

Private Sub cmbAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdGetIP_Click
    End If
End Sub


Private Sub cmdClose_Click()
    Server.CloseSck
    cmdClose.Enabled = False
    cmdSend.Enabled = False
    cmdConnect.Enabled = True
    cmdListen.Enabled = True
    Timer1.Enabled = True
End Sub

Private Sub cmdConnect_Click()
    Server.Connect2 cmbAddress.Text, txtPort.Text
    txtMessage.SetFocus
End Sub

Private Sub cmdGetIP_Click()
    Dim M%, N%
    Dim tmpAddress As String
    Timer1.Enabled = True
    tmpAddress = cmbAddress.Text
    Test_Address = Server.GetIPFromHost(cmbAddress.Text, txtPort.Text, Server.IPv6Flg)
    Test_Port = txtPort.Text
    cmbAddress.Clear
    M% = 1
    N% = InStr(M%, Test_Address, Chr$(0))
    While N% <> 0
        cmbAddress.AddItem Mid$(Test_Address, M%, N% - M%)
        M% = N% + 1
        N% = InStr(M%, Test_Address, Chr$(0))
    Wend
    If cmbAddress.ListCount > 1 Then
        M% = SendMessage(cmbAddress.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
    ElseIf cmbAddress.ListCount = 0 Then
        cmbAddress = tmpAddress
    Else
        cmbAddress.ListIndex = 0
    End If
End Sub

Private Sub cmdListen_Click()
    Server.Bind2 txtPort.Text
    If Server.State = sckOpen Then
        If Server.Protocol = sckTCPProtocol Then Server.Listen
        cmdClose.Enabled = True
        cmdListen.Enabled = False
        cmdConnect.Enabled = False
        If Server.Protocol = sckUDPProtocol Then
            cmdSend.Enabled = True
        Else
            cmdSend.Enabled = False
        End If
        cmbAddress.Text = Server.LocalHostName
        Timer1.Enabled = True 'Restart staus updates
        txtMessage.SetFocus
    Else
        lblStatus = "Unable to connect to " & cmbAddress.Text
    End If
End Sub

Private Sub cmdSend_Click()
    Debug.Print Server.RemoteHostIP & " :" & CStr(Server.RemotePort)
    Server.SendData txtMessage.Text
End Sub


Private Sub Form_Load()
    Dim M%, N%, I%
    ReDim Int_Address(10)
    If Len(Command$) > 0 Then
        Select Case LCase(Command$)
            Case "debug"
                DbgFlg = True
        End Select
    End If
    Set Server = New cSocket2
    If chkType Then 'Set Protocol
        Server.Protocol = sckUDPProtocol
    Else
        Server.Protocol = sckTCPProtocol
    End If
    If chkIPv6 Then
        Server.IPv6Flg = 6 'Set IP version
    Else
        Server.IPv6Flg = 4
    End If
    'Find and save all internal addresses (not really necessary)
    Test_Address = Server.GetIPFromHost(Server.GetLocalHostName, "")
    M% = 1: I% = 0
    N% = InStr(M%, Test_Address, Chr$(0))
    While N% <> 0
        Int_Address(I%) = Mid$(Test_Address, M%, N% - M%)
        M% = N% + 1
        N% = InStr(M%, Test_Address, Chr$(0))
        I% = I% + 1
    Wend
    ReDim Preserve Int_Address(I% - 1)
End Sub


Private Sub Server_CloseSck()
    Server.CloseSck
    Server.Listen
End Sub

Private Sub Server_Connect()
    cmdConnect.Enabled = False
    cmdListen.Enabled = False
    cmdSend.Enabled = True
    cmdClose.Enabled = True
End Sub

Private Sub Server_ConnectionRequest(ByVal requestID As Long)
    Debug.Print "Status OnConnectRequest ", Server.State
    If DbgFlg Then Call LogError("Status OnConnectRequest:" + Str$(Server.State))
    Server.CloseSck
    Server.Accept2 requestID
    cmdSend.Enabled = True
    cmbAddress.Text = Server.LocalHostName
    Debug.Print "Local Address " & Server.LocalHostName & " on Local Port " & txtPort.Text
    Debug.Print "Remote Address " & Server.RemoteHostIP & " on Remote Port " & Server.RemotePort
End Sub

Private Sub Server_DataArrival(ByVal bytesTotal As Long)
    Dim data As String
    Server.GetData data
    Debug.Print data
    If InStr(data, vbCrLf) Then
        txtLog = txtLog + "--> " + data
    Else
        txtLog = txtLog + "--> " + data + vbCrLf
    End If
End Sub

Private Sub Server_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Timer1.Enabled = False
    lblStatus.Caption = "ERROR " & Number & ": " & ":" & Description
    Server.CloseSck
End Sub


Private Sub Server_SendComplete()
    txtLog = txtLog + "<--" + txtMessage.Text + vbCrLf
    txtMessage = ""
End Sub



Private Sub Timer1_Timer()
    lblStatus.Caption = StateConv()
End Sub


Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdSend.Enabled Then Call cmdSend_Click
End Sub



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{072B0A5C-D6FC-496F-B587-7587679E1D60}#1.0#0"; "CSocket.ocx"
Begin VB.Form SMTPServer 
   Caption         =   "SMTP Pseudo Server"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   Icon            =   "SMTPServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Service"
      Height          =   400
      Left            =   4080
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install Service"
      Height          =   400
      Left            =   3000
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin CSocketOCX.Socket Socket1 
      Left            =   6720
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Setup"
      Height          =   255
      Left            =   6600
      TabIndex        =   0
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdStartServer 
      Caption         =   "Start server"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdStopServer 
      Caption         =   "Stop server"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CheckBox chkDataEvents 
      Caption         =   "Show/Log data events"
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   3360
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkAccept 
      Caption         =   "Accept Mail"
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   3600
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwClients 
      Height          =   2655
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "State"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "RemoteHostIP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "RemotePort"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Error"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Offline"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lblDisplay 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   5415
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLastTime 
      Caption         =   "Last Connect:"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   4320
      Width           =   3135
   End
End
Attribute VB_Name = "SMTPServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const gAppName As String = "SMTPServer"
Const inSep As Byte = &H1F

Dim m_strRemoteHost As String
Dim m_strMagic As String
Dim ServState As SERVICE_STATE
Dim Installed As Boolean

Sub ConnectServer()
    Dim Start As Double
    On Error GoTo ConnectServerErr
    Start = Time
    Debug.Print "Start " & Time
    Do Until Socket1.State = sckConnected
        'Debug.Print Socket1.State
        If Time - Start > 0.0002 Then
            lblStatus.Caption = "Timed Out Connecting"
            Debug.Print "TimedOut " & Time
            Exit Sub
        End If
        DoEvents
    Loop
    Debug.Print "Connected " & Time
    Socket1.SendData (m_strMagic)
    Exit Sub
ConnectServerErr:
    MsgBox ("Error: " & CStr(Err) & ", " & Err.Description)
End Sub

Private Sub DisplayList(strPost As String)
    Dim strArray() As Variant
    Dim objList As ListItem
    Dim I  As Long
    Dim lngIndex As Long
    'Debug.Print strPost
    I = lvwClients.ListItems.Count
    lngIndex = Val(strPost)
    If lngIndex <= I Then
        Set objList = lvwClients.ListItems(lngIndex)
    Else
        For I = I + 1 To lngIndex
            Set objList = lvwClients.ListItems.Add(, , I)
        Next I
    End If
    strArray = ParseString(lngIndex, strPost, " ")
    If Len(strArray(0)) Then
        With objList
           If Len(strArray(1)) Then .SubItems(1) = strArray(1)
           If Len(strArray(2)) Then .SubItems(2) = strArray(2)
           If Len(strArray(3)) Then .SubItems(3) = strArray(3)
           If Len(strArray(4)) Then .SubItems(4) = strArray(4)
        End With
    End If
    Set objList = Nothing
End Sub

Function ParseInput(sBuf As String, Sep As String) As Variant
    'Input should have been checked for leading separater before entry
    Dim M%, N%, N1%
    ReDim ParseArray(20)
    N% = 1
    Do Until N% = 0
        M% = N% + 1
        N% = InStr(M%, sBuf, Sep)
        If N% = 0 Then
            ParseArray(N1%) = Mid$(sBuf, M%)
        Else
            ParseArray(N1%) = Mid$(sBuf, M%, N% - M%)
        End If
        N1% = N1% + 1
    Loop
    ReDim Preserve ParseArray(N1%)
    ParseInput = ParseArray
End Function

Private Function ParseString(lngIndex As Long, A$, Sep$) As Variant
    Dim M%, N%, N1%
    Dim ParseArray(4)
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
        If N1% > 4 Then Exit Do
        Do Until Mid$(A$, N% + 1, 1) <> Sep$
            N% = N% + 1
        Loop
    Loop
    Debug.Print ParseArray(0) & Space$(1) & ParseArray(1) & "|"
    If ParseArray(1) = "Connected" Then
        ParseArray(4) = " "
        lblLastTime.Caption = "Last Connect: " & Time
    ElseIf Left$(ParseArray(1), 6) = "Closed" Then
        ParseArray(2) = vbNullString
        ParseArray(3) = vbNullString
        ParseArray(4) = " "
    ElseIf Left$(ParseArray(1), 5) = "Error" Then
        ParseArray(1) = "Closed"
        ParseArray(2) = vbNullString
        ParseArray(3) = vbNullString
        N% = InStr(1, A$, Sep$)
        ParseArray(4) = Mid$(A$, N% + 1)
    Else
        ParseArray(0) = vbNullString
    End If
    ParseString = ParseArray
End Function

Private Function GetSettings(sKey As String) As String
    'GetSettings = GetSetting(gAppName, "Settings", sKey, "")
    GetSettings = modReg.GetSettingEx(HKEY_LOCAL_MACHINE, "Software\VB Settings\" & gAppName & "\Settings", sKey)
End Function
Private Sub CheckService()
    Dim I%
    If GetServiceConfig() = 0 Then
        Installed = True
        cmdInstall.Caption = "Uninstall Service"
TryAgain:
        ServState = GetServiceStatus()
        Select Case ServState
            Case SERVICE_RUNNING
                cmdInstall.Enabled = False
                cmdStart.Caption = "Stop Service"
                cmdStart.Enabled = True
            Case SERVICE_STOPPED
                cmdInstall.Enabled = True
                cmdStart.Caption = "Start Service"
                cmdStart.Enabled = True
            Case SERVICE_START_PENDING, SERVICE_STOP_PENDING
                DoEvents 'Delay while status changes
                I% = I% + 1
                If I% < 500 Then GoTo TryAgain
            Case Else
                cmdInstall.Enabled = False
                cmdStart.Enabled = False
        End Select
    Else
        Installed = False
        cmdInstall.Caption = "Install Service"
        cmdStart.Enabled = False
        cmdInstall.Enabled = True
    End If
End Sub

Private Sub SaveSettings(sKey As String, sValue As String)
    'SaveSetting gAppName, "Settings", sKey, sValue
    modReg.SaveSettingEx HKEY_LOCAL_MACHINE, "Software\VB Settings\" & gAppName & "\Settings", sKey, sValue
 End Sub

Private Sub cmdInstall_Click()
    Dim sTemp As String
    CheckService
    If Not cmdInstall.Enabled Then Exit Sub
    cmdInstall.Enabled = False
    If Installed Then
        DeleteNTService
    Else
        SetNTService
'        'Add registry key for Adapter
'        sTemp = CreateKey(sDevice)
    End If
    CheckService
End Sub

Private Sub cmdSetup_Click()
    Dim strTemp As String
EnterGreeting:
    strTemp = InputBox("Enter SMTP Connect Message:", , GetSettings("Greeting"))
    If Len(strTemp) > 0 Then
        Call SaveSettings("Greeting", strTemp)
    Else
        MsgBox (strTemp & " is not a valid Greeting." & vbCrLf & "Try Again!")
        GoTo EnterGreeting
    End If
    Call SaveSettings("ShowLog", chkDataEvents)
    Call SaveSettings("AcceptMail", chkAccept)
EnterIPmon:
    strTemp = InputBox("Enter IP to Monitor:", , m_strRemoteHost)
    If Len(strTemp) > 0 And InStr(Socket1.GetIPFromHost(strTemp), strTemp) Then
        Call SaveSettings("Remote", strTemp)
    Else
        MsgBox (strTemp & " is not a valid IP." & vbCrLf & "Try Again!")
        GoTo EnterIPmon
    End If
EnterMagic:
    strTemp = InputBox("Enter Magic Word:", , m_strMagic)
    If Len(strTemp) > 0 Then
        Call SaveSettings("Magic", strTemp)
    Else
        MsgBox (strTemp & " is not a valid Magic Word." & vbCrLf & "Try Again!")
        GoTo EnterMagic
    End If
    MsgBox " Restart SMTPServer for changes to take effect!"
End Sub

Private Sub cmdStart_Click()
    CheckService
    If Not cmdStart.Enabled Then Exit Sub
    cmdStart.Enabled = False
    If ServState = SERVICE_RUNNING Then
        Call Socket1_CloseSck
        StopNTService
    ElseIf ServState = SERVICE_STOPPED Then
        StartNTService
    End If
    CheckService
    If ServState = 1 Then
        Debug.Print "SERVICE_STOPPED"
        lblStatus.Caption = "Offline"
        lblStatus.ForeColor = vbRed
    ElseIf ServState = 4 Then
        Debug.Print "SERVICE_RUNNING"
        MsgBox ("Restart program to see effect!")
    End If
End Sub

Private Sub cmdStartServer_Click()
    On Error GoTo cmdStartErr
    Socket1.SendData "Start Server"
    Exit Sub
cmdStartErr:
    MsgBox ("Error: " & CStr(Err) & ", " & Err.Description)
End Sub


Private Sub cmdStopServer_Click()
    On Error GoTo cmdStopErr
    Socket1.SendData "Stop Server"
    Exit Sub
cmdStopErr:
    MsgBox ("Error: " & CStr(Err) & ", " & Err.Description)
End Sub

Private Sub Form_Activate()
    If InStr(Socket1.GetIPFromHost(""), m_strRemoteHost) _
        Or m_strRemoteHost = "127.0.0.1" _
        Or m_strRemoteHost = "::1" Then
        Debug.Print m_strRemoteHost & " is Local"
        cmdInstall.Visible = True
        cmdStart.Visible = True
    Else
        Debug.Print m_strRemoteHost & " is Remote"
    End If
    Call ConnectServer
End Sub

Private Sub Form_Load()
    Call CheckService
    chkDataEvents = Val(GetSettings("ShowLog"))
    chkAccept = Val(GetSettings("AcceptMail"))
    m_strRemoteHost = GetSettings("Remote")
    m_strMagic = GetSettings("Magic")
    If Len(m_strRemoteHost) Then
        If InStr(m_strRemoteHost, ":") Then
            Socket1.IPv6Flg = 6
        Else
            Socket1.IPv6Flg = 4
        End If
        Socket1.Connect m_strRemoteHost, "26"
        'Socket1.Connect "192.168.1.2", "26"
        'Socket1.Connect "::1", "26"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings("ShowLog", chkDataEvents.Value)
    Call SaveSettings("AcceptMail", chkAccept.Value)
End Sub

Private Sub Form_Resize()
    If Me.ScaleHeight = 0 Then Exit Sub
    With lblStatus
        .Top = Me.ScaleHeight - .Height
        .Width = Me.ScaleWidth / 2
        lblLastTime.Top = .Top
    End With
    With lvwClients
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - 1600
    End With
    With lblDisplay
        .Width = Me.ScaleWidth
        .Top = lvwClients.Height + 25
    End With
    With cmdStartServer
        .Top = lvwClients.Height + lblDisplay.Height + 120
        cmdStopServer.Top = .Top
        cmdInstall.Top = .Top
        cmdStart.Top = .Top
    End With
    With chkDataEvents
        .Top = lvwClients.Height + lblDisplay.Height + 40
        chkAccept.Top = .Top + lblLastTime.Height
        cmdSetup.Top = .Top + lblLastTime.Height
    End With
    If lblStatus.Caption = "Offline" Then
        lblStatus.ForeColor = vbRed
    Else
        lblStatus.ForeColor = vbBlue
    End If
End Sub


Private Sub lblDisplay_Change()
    Dim M%, N%
    Dim CRLFLoc(10) As Integer
    Static lblChngFlg As Boolean
    If Not SMTPServer.Visible Then Exit Sub
    On Error GoTo lblDisplayChangeErr
    If lblChngFlg Then
        lblChngFlg = False
        Exit Sub
    End If
    CRLFLoc(N%) = InStr(lblDisplay.Caption, vbCrLf)
    M% = CRLFLoc(N%)
    Do Until M% = 0
        M% = InStr(M% + 1, lblDisplay.Caption, vbCrLf)
        N% = N% + 1
        CRLFLoc(N%) = M%
    Loop
    If N% > 3 Then
        lblChngFlg = True
        lblDisplay.Caption = Mid$(lblDisplay.Caption, CRLFLoc(N% - 4) + 2)
    End If
    Exit Sub
lblDisplayChangeErr:
    lblDisplay.Caption = ""
End Sub

Private Sub Socket1_CloseSck()
    Socket1.CloseSck
End Sub

Private Sub Socket1_DataArrival(ByVal bytesTotal As Long)
    Dim N%
    Dim Data As String
    Socket1.GetData Data
    Debug.Print Data
    Dim DataArray() As Variant
    If Asc(Data) = inSep Then 'Check for leading separater
        DataArray = ParseInput(Data, Chr$(inSep))
    Else
    End If
    For N% = 0 To UBound(DataArray) - 1
        If Val(DataArray(N%)) Then
            Data = DataArray(N%)
            Call DisplayList(Data)
        Else
            'Data = Mid$(Data, 2)
        End If
        If Left$(DataArray(N%), 6) = "Online" Then
            lblStatus.Caption = DataArray(N%)
            lblStatus.ForeColor = vbBlue
        ElseIf Left$(DataArray(N%), 7) = "Offline" Then
            lblStatus.Caption = DataArray(N%)
            lblStatus.ForeColor = vbRed
            lvwClients.ListItems.Clear
        End If
        lblDisplay.Caption = lblDisplay.Caption & DataArray(N%) & vbCrLf
    Next N%
End Sub



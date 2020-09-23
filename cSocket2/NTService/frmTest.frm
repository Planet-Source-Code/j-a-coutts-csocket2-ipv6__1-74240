VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Object = "{072B0A5C-D6FC-496F-B587-7587679E1D60}#1.0#0"; "CSocket.ocx"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin CSocketOCX.Socket Socket1 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin NTService.NTService NTService1 
      Left            =   960
      Top             =   120
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "Test Service"
      ServiceName     =   "Simple"
      StartMode       =   3
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StopService As Boolean
Dim IPv6 As Boolean
Sub Main()
    On Error GoTo InstallErr
    'if the /install command line switch is set, self-install ourselves as a service!
    If LCase(Trim(Command)) = "/install" Then
        MsgBox NTService1.DisplayName + " Will be Installed as Service"
'        ' enable interaction with desktop
'        NTService.Interactive = True
        If NTService1.Install Then
'            'Save the TimerInterval Parameter in the Registry
'            NTService.SaveSetting "Parameters", "TimerInterval", "300"
            MsgBox NTService1.DisplayName & ": installed successfully"
        Else
            MsgBox NTService1.DisplayName & ": failed to install"
        End If
        End
    ElseIf LCase(Trim(Command)) = "/uninstall" Then
        MsgBox NTService1.DisplayName + " Will be Uninstalled as Service"
        If NTService1.Uninstall Then
            MsgBox NTService1.DisplayName & ": uninstalled successfully"
        Else
            MsgBox NTService1.DisplayName & ": failed to uninstall"
        End If
        End
    'Invalid parameter
    ElseIf Command <> "" Then
        MsgBox "Invalid Parameter"
        End
    End If
'    'Retrive the stored value for the timer interval
'    Timer.Interval = CInt(NTService.GetSetting("Parameters", "TimerInterval", "300"))
    ' enable Pause/Continue. Must be set before StartService
    ' is called or in design mode
    NTService1.ControlsAccepted = svcCtrlPauseContinue
    ' connect service to Windows NT services controller
    NTService1.StartService
    Exit Sub
InstallErr:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
    End
End Sub


Private Sub Form_Load()
    Call Main
    With Socket1
        If IPv6 Then
            .IPv6Flg = 6
        Else
            .IPv6Flg = 4
        End If
        .LocalPort = 226
        Call .Listen
        If .State <> sckListening Then
            MsgBox ("Forward Port 226 not active! ")
        End If
    End With
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
        Case 0  'From Form Control Menu, simply minimize
            Cancel = True
        Case 1  'From Code, do not cancel
        Case 2  'From Windows, log event but do not cancell
            Call NTService1.LogEvent(svcMessageError, svcMessageInfo, "Unload from Windows")
        Case 3  'From Task Manager
        Case 4  'From MDI Form
    End Select
End Sub


Private Sub NTService1_Continue(Success As Boolean)
    'Handle the continue service event
    On Error GoTo ServiceError
    Success = True
    NTService1.LogEvent svcEventInformation, svcMessageInfo, "Service continued"
    Exit Sub
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub


Private Sub NTService1_Control(ByVal mEvent As Long)
    'Take control of the service events
    On Error GoTo ServiceError
'    Label1.Caption = NTService1.DisplayName & " Control signal " & CStr([mEvent])
    Exit Sub
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub NTService1_Pause(Success As Boolean)
    'Pause Event Request
    On Error GoTo ServiceError
    NTService1.LogEvent svcEventError, svcMessageError, "Service paused"
    Success = True
    Exit Sub
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub NTService1_Start(Success As Boolean)
    'Start Event Request
    On Error GoTo ServiceError
    Success = True
    Exit Sub
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub NTService1_Stop()
    'Stop and terminate the Service
    On Error GoTo ServiceError
    StopService = True
    Unload Me
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub


Private Sub Socket1_CloseSck()
    On Error GoTo SocketErr
    Exit Sub
SocketErr:
End Sub

Private Sub Socket1_Connect()
    On Error GoTo SocketErr
    Exit Sub
SocketErr:
End Sub


Private Sub Socket1_ConnectionRequest(ByVal requestID As Long)
    On Error GoTo SocketErr
    Exit Sub
SocketErr:
End Sub


Private Sub Socket1_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo SocketErr
    Exit Sub
SocketErr:
End Sub


Private Sub Socket1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error GoTo SocketErr
    Exit Sub
SocketErr:
End Sub


Private Sub Socket1_SendComplete()
    On Error GoTo SocketErr
    Exit Sub
SocketErr:
End Sub


Private Sub Socket1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    On Error GoTo SocketErr
    Exit Sub
SocketErr:
End Sub



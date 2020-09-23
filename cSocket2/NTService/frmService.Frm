VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Begin VB.Form frmService 
   Caption         =   "NT Service Project"
   ClientHeight    =   1365
   ClientLeft      =   1395
   ClientTop       =   1620
   ClientWidth     =   3570
   Icon            =   "frmService.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1365
   ScaleWidth      =   3570
   Begin VB.Timer Timer 
      Left            =   960
      Top             =   780
   End
   Begin NTService.NTService NTService 
      Left            =   240
      Top             =   780
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "Sample NT Service"
      ServiceName     =   "SampleService"
      StartMode       =   3
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3555
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StopService As Boolean

Private Sub Form_Load()
    On Error GoTo ServiceError
    
    StopService = False
    Label1.Caption = "Loading"
    'Install the service
    If Command = "/i" Then
        ' enable interaction with desktop
        NTService.Interactive = True
        'Install the program as an NT service
        If NTService.Install Then
            'Save the TimerInterval Parameter in the Registry
            NTService.SaveSetting "Parameters", "TimerInterval", "300"
            MsgBox NTService.DisplayName & ": installed successfully"
        Else
            MsgBox NTService.DisplayName & ": failed to install"
        End If
        End
    'Remove the Service Registry Keys and uninstall the service
    ElseIf Command = "/u" Then
        If NTService.Uninstall Then
            MsgBox NTService.DisplayName & ": uninstalled successfully"
        Else
            MsgBox NTService.DisplayName & ": failed to uninstall"
        End If
        End
    'Invalid parameter
    ElseIf Command <> "" Then
        MsgBox "Invalid Parameter"
        End
    End If
    'Retrive the stored value for the timer interval
    Timer.Interval = CInt(NTService.GetSetting("Parameters", "TimerInterval", "300"))
    ' enable Pause/Continue. Must be set before StartService
    ' is called or in design mode
    NTService.ControlsAccepted = svcCtrlPauseContinue
    ' connect service to Windows NT services controller
    NTService.StartService
    Exit Sub
ServiceError:
    Call NTService.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub

'Unload the Service
Private Sub Form_Unload(Cancel As Integer)
    If Not StopService Then
        If MsgBox("Are you sure you want to unload the service?..." & vbCrLf & "the service will be stopped", vbQuestion + vbYesNo, "Stop Service") = vbYes Then
            NTService.StopService
            Label1.Caption = "Stopping"
            Cancel = True
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub NTService_Continue(Success As Boolean)
    'Handle the continue service event
    On Error GoTo ServiceError
    Timer.Enabled = True
    Label1.Caption = "Running"
    Success = True
    NTService.LogEvent svcEventInformation, svcMessageInfo, "Service continued"
    Exit Sub
ServiceError:
    NTService.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub NTService_Control(ByVal mEvent As Long)
    'Take control of the service events
    On Error GoTo ServiceError
    Label1.Caption = NTService.DisplayName & " Control signal " & CStr([mEvent])
    Exit Sub
ServiceError:
    NTService.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub NTService_Pause(Success As Boolean)
    'Pause Event Request
    On Error GoTo ServiceError
    Timer.Enabled = False
    Label1.Caption = "Paused"
    NTService.LogEvent svcEventError, svcMessageError, "Service paused"
    Success = True
    Exit Sub
ServiceError:
    NTService.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub NTService_Start(Success As Boolean)
    'Start Event Request
    On Error GoTo ServiceError
    Label1.Caption = "Running"
    Success = True
    Exit Sub
ServiceError:
    NTService.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub NTService_Stop()
    'Stop and terminate the Service
    On Error GoTo ServiceError
    Label1.Caption = "Stopped"
    StopService = True
    Unload Me
ServiceError:
    NTService.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub Timer_Timer()
    'The program execute this sub when the service is enabled
    Dim nx As Single
    Dim ny As Single
    
    On Error GoTo ServiceError
    nx = Left + Rnd() * 100 - 50
    ny = Top + Rnd() * 100 - 50
    If ny < 0 Then ny = 0
    If nx < 0 Then nx = 0
    If nx > Screen.Width - Width Then nx = Screen.Width - Width
    If ny > Screen.Height - Height Then ny = Screen.Height - Height
    Me.Move nx, ny
    Exit Sub
ServiceError:
    NTService.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

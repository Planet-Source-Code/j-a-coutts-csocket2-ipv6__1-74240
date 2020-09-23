VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Begin VB.Form frmStart 
   Caption         =   "SMTP Service"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   ControlBox      =   0   'False
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin NTService.NTService NTService1 
      Left            =   240
      Top             =   240
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "SMTP svc6"
      ServiceName     =   "SMTPsvc6"
      StartMode       =   3
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function LocalInit() As Long
' Purpose:
'   Starting point for application.
' =====================================================
    Const Title As String = "SMTP Server"
    Dim Handle As Long
    Dim TaskID As Long
    Const sProc As String = "LocalInit"
    On Error GoTo LocalInitErr
    'Check if Program already running
    Handle = frmStart.hwnd
    TaskID = CheckUnique(Title, Handle)
    If TaskID Then
        'Do Nothing
        End             'Window already active form
    End If
    frmStart.Caption = Title
    LocalInit = True
    Load fMain
    Exit Function
LocalInitErr:
    LocalInit = False
End Function




Private Sub Form_Load()
    Load fMain
'    If Len(Trim(Command)) > 0 Then Exit Sub
'    If Not LocalInit Then GoTo fLoadErr
'    On Error GoTo fLoadErr
'    Exit Sub
'fLoadErr:
'    MsgBox "Problem Loading " + AppTitle
'    End
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
        Case 0  'From Form Control Menu, simply minimize
            Cancel = True
        Case 1  'From Code, do not cancel
        Case 2  'From Windows, log event but do not cancell
            NTService1.Log svcEventInformation, svcMessageInfo, "Unload from Windows"
        Case 3  'From Task Manager
        Case 4  'From MDI Form
    End Select
End Sub






Private Sub NTService1_Continue(Success As Boolean)
    On Error GoTo ServiceError
    Success = True
    NTService1.LogEvent svcEventInformation, svcMessageInfo, "Service continued"
    fMain.tmrProperties.Interval = fMain.TimerInterval
    Exit Sub
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub


Private Sub NTService1_Control(ByVal mEvent As Long)
    'Take control of the service events
    On Error GoTo ServiceError
    Exit Sub
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub


Private Sub NTService1_Pause(Success As Boolean)
    'Pause Event Request
    On Error GoTo ServiceError
    NTService1.LogEvent svcEventError, svcMessageError, "Service paused"
    fMain.tmrProperties.Interval = 0
    Success = True
    Exit Sub
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub


Private Sub NTService1_Start(Success As Boolean)
    'Start Event Request
    On Error GoTo ServiceError
    fMain.tmrProperties.Interval = fMain.TimerInterval
    Success = True
    Exit Sub
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub


Private Sub NTService1_Stop()
    'Stop and terminate the Service
    On Error GoTo ServiceError
'    StopService = True
    Unload Me
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub



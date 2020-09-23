Attribute VB_Name = "modStart"
Option Explicit

' Declaration Statements
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef lpcbNeeded As Long) As Long
Public Declare Function GetModuleFileNameEx Lib "psapi" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, Optional lpNumberOfBytesWritten As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

' Shared Constants
Public Const AppTitle As String = "SMTPSvc"
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const PROCESS_QUERY_INFORMATION = (&H400)
Public Const PROCESS_VM_READ = (&H10)
Public Const PROCESS_VM_WRITE = (&H20)
Public Const PROCESS_VM_OPERATION = (&H8)
Public Const MEM_COMMIT = &H1000
Public Const MEM_RESERVE = &H2000
Public Const MEM_RELEASE = &H8000
Public Const PAGE_READWRITE = &H4
Public Const MAX_PATH = 260
Public Const WM_USER = &H400
Public Const TB_BUTTONCOUNT = (WM_USER + 24)
'Public Const TB_GETBUTTON = (WM_USER + 23)
Public Const TB_GETBUTTONTEXTA = (WM_USER + 45)
Public Const WM_COPYDATA = &H4A

'Left-click constants.
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
'Right-click constants.
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up

Public Function GetOwnedWindow(hwnd As Long) As Long
    Dim OwnedHandle As Long
    OwnedHandle = GetDesktopWindow()    'get the desktop handle
    OwnedHandle = GetWindow(OwnedHandle, 5) ' get first top level window
    Do
        If GetParent(OwnedHandle) = hwnd Then
            GetOwnedWindow = OwnedHandle
            Exit Do
        End If
        OwnedHandle = GetWindow(OwnedHandle, 2) 'get next top level window
    Loop Until OwnedHandle = 0
End Function

Public Function CheckUnique(ByVal FormName As String, hIgnore As Long) As Long
' Purpose:
    'FormName is the caption of the desired form, hIgnore is the
    'window handle of the parent form to be ignored if already running.
' =====================================================
    Dim hwnd As Long
    Dim ShowW As Long, SetF As Long, pID As Long
    hwnd = FindWindow(vbNullString, FormName)
    If hwnd = 0 Then Exit Function
    ShowW = ShowWindow(hwnd, 1) 'Restore it in case it is minimized
    SetF = GetForegroundWindow()    'Does not always return current app
    If SetF = frmStart.hwnd Then
        SetF = SetForegroundWindow(hwnd)
        SetF = GetCurrentProcessId()
        SetF = WaitForInputIdle(SetF, 10000)
        hwnd = GetForegroundWindow()
    Else
        hwnd = SetF
    End If
    SetF = GetOwnedWindow(hwnd) 'Get owned top level window
    If SetF = 0 Then
        CheckUnique = hwnd 'No owned Windows found
    Else
        CheckUnique = -SetF  'return neg handle for owned window
        SetF = GetWindowThreadProcessId(SetF, pID)
        SetF = AttachThreadInput(SetF, GetCurrentThreadId(), pID)
    End If
    Exit Function
End Function

Public Sub Main()
    On Error GoTo InstallErr
    'if the /install command line switch is set, self-install ourselves as a service!
    If LCase(Trim(Command)) = "/install" Then
        MsgBox frmStart.NTService1.DisplayName + " Will be Installed as Service"
        frmStart.NTService1.Interactive = True
        If frmStart.NTService1.Install Then
            MsgBox frmStart.NTService1.DisplayName & ": installed successfully"
        Else
            MsgBox frmStart.NTService1.DisplayName & ": failed to install"
        End If
        End
    ElseIf LCase(Trim(Command)) = "/uninstall" Then
        MsgBox frmStart.NTService1.DisplayName + " Will be Uninstalled as Service"
        If frmStart.NTService1.Uninstall Then
            MsgBox frmStart.NTService1.DisplayName & ": uninstalled successfully"
        Else
            MsgBox frmStart.NTService1.DisplayName & ": failed to uninstall"
        End If
        End
    'Invalid parameter
    ElseIf Command <> "" Then
        MsgBox "Invalid Parameter"
        End
    End If
    ' enable Pause/Continue. Must be set before StartService
    ' is called or in design mode
    frmStart.NTService1.ControlsAccepted = svcCtrlPauseContinue
    ' connect service to Windows NT services controller
    frmStart.NTService1.StartService
    Exit Sub
InstallErr:
    Call frmStart.NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
    End
End Sub


Attribute VB_Name = "modReg"
Option Explicit
 
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKEY As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKEY As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
 
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const REG_SZ = 1

Private Const ERROR_SUCCESS As Long = 0
 
Private Function RegQueryStringValue(ByVal HKEY As Long, ByVal strValueName As String)
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
        On Error GoTo 0
    lResult = RegQueryValueEx(HKEY, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = RegQueryValueEx(HKEY, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = StripTerminator(strBuf)
            End If
        End If
    End If
End Function
 
Public Function GetSettingEx(HKEY As Long, sPath As String, sValue As String)
    Dim KeyHand As Long
    Dim datatype As Long
    Call RegOpenKey(HKEY, sPath, KeyHand)
    GetSettingEx = RegQueryStringValue(KeyHand, sValue)
    Call RegCloseKey(KeyHand)
End Function
 
Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
 
Public Sub SaveSettingEx(HKEY As Long, sPath As String, sValue As String, sData As String)
    Dim KeyHand As Long
    Call RegCreateKey(HKEY, sPath, KeyHand)
    Call RegSetValueEx(KeyHand, sValue, 0, REG_SZ, ByVal sData, Len(sData))
    Call RegCloseKey(KeyHand)
End Sub



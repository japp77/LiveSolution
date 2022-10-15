Attribute VB_Name = "ModRegistrazione"
Option Explicit

Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const KEY_ALL_ACCESS = &H3F

Private Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)

Public Function DeleteKey(lPredefinedKey As Long, sKeyName As String)
    Dim lRetVal As Long
    Dim hKey As Long
    lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)
End Function

Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
       Dim lRetVal As Long
       Dim hKey As Long
       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = RegDeleteValue(hKey, sValueName)
       RegCloseKey (hKey)
End Function

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select
End Function
Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5
    Select Case lType
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            lrc = -1
    End Select
QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
QueryValueExError:
   Resume QueryValueExExit
End Function
Public Function CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
    Dim hNewKey As Long
    Dim lRetVal As Long
   
    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Function
Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
       Dim lRetVal As Long
       Dim hKey As Long
       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hKey)
End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
       Dim lRetVal As Long
       Dim hKey As Long
       Dim vValue As Variant
       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       QueryValue = vValue
       RegCloseKey (hKey)
End Function
Public Sub ShowAtStartup(yourexe As String, NomeAvvio As String)
Dim pth As String
    
    pth = MenuOptions.ProgramsPath
    
    If Mid(pth, Len(pth), 1) = "\" Then
        pth = Chr(34) & pth & yourexe & ".exe" & Chr(34)
    Else
        pth = Chr(34) & pth + "\" + yourexe & ".exe" & Chr(34)
    End If
    
    SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\", NomeAvvio, pth, REG_SZ

End Sub
Public Sub DontShowAtStartup(yourexe$)
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\", yourexe$
End Sub






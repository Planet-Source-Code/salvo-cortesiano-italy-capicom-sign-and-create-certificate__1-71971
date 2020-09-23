Attribute VB_Name = "modRegistry"
Option Explicit

Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_USERS As Long = &H80000003
Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
Public Const HKEY_DYN_DATA As Long = &H80000006

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Public Const KEY_READ = &H20019

Public Const MAX_KEY_LENGTH As Long = 255
Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_MULTI_SZ As Long = 7

Public Function GetRegistryValue(ByVal KeyHandle As Long, ByVal KeyName As String, ByVal ValueName As String) As Variant

Dim bytRegValue() As Byte
Dim lngBufferLen As Long
Dim lngDataType As Long
Dim lngRetVal As Long
Dim lngSubKey As Long
Dim lngLongResult As Long
Dim strStringResult As String

If RegOpenKeyEx(KeyHandle, KeyName, 0, KEY_READ, lngSubKey) = 0 Then
    ReDim bytRegValue(0 To 0) As Byte
    lngRetVal = RegQueryValueEx(lngSubKey, ValueName, 0, lngDataType, bytRegValue(0), lngBufferLen)
    If lngRetVal = ERROR_MORE_DATA Then
        ReDim bytRegValue(0 To lngBufferLen - 1) As Byte
        lngRetVal = RegQueryValueEx(lngSubKey, ValueName, 0, lngDataType, bytRegValue(0), lngBufferLen)
    End If
    Select Case lngDataType
        Case REG_DWORD
            CopyMemory lngLongResult, bytRegValue(0), 4
            GetRegistryValue = lngLongResult
        Case REG_SZ, REG_EXPAND_SZ
            strStringResult = Space$(lngBufferLen - 1)
            CopyMemory ByVal strStringResult, bytRegValue(0), lngBufferLen - 1
            GetRegistryValue = strStringResult
        Case REG_BINARY
            If lngBufferLen <> UBound(bytRegValue) + 1 Then
                ReDim Preserve bytRegValue(0 To lngBufferLen - 1) As Byte
            End If
            GetRegistryValue = bytRegValue()
        Case REG_MULTI_SZ
            strStringResult = Space$(lngBufferLen - 2)
            CopyMemory ByVal strStringResult, bytRegValue(0), lngBufferLen - 2
            GetRegistryValue = strStringResult
        Case Else
            RegCloseKey lngSubKey
    End Select
    RegCloseKey lngSubKey
End If

End Function


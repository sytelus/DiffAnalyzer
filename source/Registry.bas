Attribute VB_Name = "Registry"
Option Explicit

Public Enum RegistryRoots
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
End Enum

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&

Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2

Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_SUB_KEY Or KEY_CREATE_LINK Or KEY_SET_VALUE
    'Declares
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Public Function GetRegistryString(ByVal venmRegistryRoot As RegistryRoots, ByVal vsKeyPath As String, ByVal vsKeyName As String, ByVal vsDefaultValue As String) As String
Dim bSuccess As Boolean
Dim bfound As Boolean
Dim lOpenKeyResult As Long
Dim lType As Long
Dim cMyData As String
Dim lLength As Long

    bSuccess = OpenKey(venmRegistryRoot, vsKeyPath, bfound, lOpenKeyResult)
    If bSuccess Then
        If bfound Then
        'Get Setting
            bSuccess = QueryValue(lOpenKeyResult, vsKeyName, lType, cMyData, lLength, bfound)
            If bSuccess Then
                If bfound Then
                    GetRegistryString = Left(cMyData, lLength - 1)
                Else
                    GetRegistryString = vsDefaultValue
                End If
            Else
                GetRegistryString = vsDefaultValue
            End If
        Else
            GetRegistryString = vsDefaultValue
        End If
    Else
        GetRegistryString = vsDefaultValue
    End If
    
    Call CloseKey(lOpenKeyResult)
    
End Function

Private Function OpenKey(ByVal lKey As Long, ByVal cSubKey As String, ByRef bfound As Boolean, ByRef lOpenKey As Long) As Boolean

    On Error GoTo ErrorHandler
    
    Dim lReturn As Long
    Dim bSuccess As Boolean
    
    lReturn = RegOpenKeyEx(lKey, cSubKey, 0&, KEY_QUERY_VALUE, lOpenKey)
    Select Case lReturn
        Case ERROR_SUCCESS
            bfound = True
            bSuccess = True
        Case ERROR_FILE_NOT_FOUND
            bfound = False
            bSuccess = True
        Case Else
            bfound = False
            bSuccess = False
    End Select
    OpenKey = bSuccess

Exit Function

ErrorHandler:
    bfound = False
    bSuccess = False
End Function

Private Function QueryValue(ByVal lKey As Long, ByVal cValueName As String, ByRef lType As Long, ByRef cData As String, ByRef lDataLength As Long, ByRef bfound As Boolean) As Boolean

    On Error GoTo ErrorHandler
    
    Dim lReturn As Long
    Dim bSuccess As Boolean
    
    lDataLength = 255
    cData = String$(lDataLength, 0)
    lReturn = RegQueryValueEx(lKey, cValueName, 0&, lType, cData, lDataLength)
    Select Case lReturn
        Case 0
            bfound = True
            bSuccess = True
        Case ERROR_FILE_NOT_FOUND
            bfound = False
            bSuccess = True
        Case Else
            bfound = False
            bSuccess = False
    End Select
    QueryValue = bSuccess

Exit Function

ErrorHandler:
    bfound = False
    bSuccess = False
End Function

Private Function CloseKey(ByVal lKey As Long) As Boolean

    On Error GoTo ErrorHandler
    
    Dim lReturn As Long
    Dim bSuccess As Boolean
    
    lReturn = RegCloseKey(lKey)
    Select Case lReturn
        Case ERROR_SUCCESS
            bSuccess = True
        Case Else
            bSuccess = False
    End Select
    CloseKey = bSuccess

Exit Function
ErrorHandler:
    bSuccess = False
End Function


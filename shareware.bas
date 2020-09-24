Attribute VB_Name = "modRegistry"
Option Explicit

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_DYN_DATA = &H80000004

Const REG_SZ = 1

' Registry API prototypes
Private Declare Function RegCreateKey Lib "advapi32.dll" _
  Alias "RegCreateKeyA" _
    (ByVal hkey As Long, _
     ByVal lpSubKey As String, _
     phkResult As Long) As Long
     
Private Declare Function RegDeleteKey Lib "advapi32.dll" _
  Alias "RegDeleteKeyA" _
    (ByVal hkey As Long, _
     ByVal lpSubKey As String) As Long
     
Private Declare Function RegDeleteValue Lib "advapi32.dll" _
  Alias "RegDeleteValueA" _
    (ByVal hkey As Long, _
     ByVal lpSubKey As String) As Long
     
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
  Alias "RegQueryValueExA" _
    (ByVal hkey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     lpData As Any, _
     lpcbData As Long) As Long
     
Private Declare Function RegSetValueEx Lib "advapi32.dll" _
  Alias "RegSetValueExA" _
    (ByVal hkey As Long, _
     ByVal lpValueName As String, _
     ByVal Reserved As Long, _
     ByVal dwType As Long, _
     lpData As Any, _
     ByVal cbData As Long) As Long

' Registry error constants
Const API_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_REGISTRY_RECOVERED = 1014&
Const ERROR_REGISTRY_CORRUPT = 1015&
Const ERROR_REGISTRY_IO_FAILED = 1016&
Const ERROR_NOT_REGISTRY_FILE = 1017&
Const ERROR_KEY_DELETED = 1018&
Const ERROR_NO_LOG_SPACE = 1019&
Const ERROR_KEY_HAS_CHILDREN = 1020&
Const ERROR_CHILD_MUST_BE_VOLATILE = 1021&
Const ERROR_RXACT_INVALID_STATE = 1369&

Public Function CreateRegKey(sRegistryKey As String) As Long

Dim lResult         As Long

CreateRegKey = 0           ' Assume success

' Make sure all parameters have values
If Len(sRegistryKey) = 0 Then
    ' The key property is not set, so flag an error
    CreateRegKey = ERROR_BADKEY
    Exit Function
End If

' Make the call to create the key
CreateRegKey = RegCreateKey(HKEY_LOCAL_MACHINE, _
                    sRegistryKey, lResult)
           
End Function

Public Function DeleteRegKey _
            (sRegistryKey As String, _
             sSubKey As String) As Long

Dim lKeyId          As Long
Dim lResult         As Long

DeleteRegKey = 0           ' Assume success

' Make sure all parameters have values
If Len(sRegistryKey) = 0 Then
    ' The key parameter is not set
    DeleteRegKey = ERROR_BADKEY
    Exit Function
End If

If Len(sSubKey) = 0 Then
    ' The sub key parameter is not set
    DeleteRegKey = ERROR_BADKEY
    Exit Function
End If

' Open the key by attempting to create it. If it
' already exists, an ID is returned.
lResult = RegCreateKey(HKEY_LOCAL_MACHINE, sRegistryKey, lKeyId)

If lResult = 0 Then
    ' Got a key ID so delete the entry
    DeleteRegKey = RegDeleteKey(lKeyId, ByVal sSubKey)
End If

End Function

Public Function DeleteRegValue _
            (sRegistryKey As String, _
             sSubKey As String) As Long

Dim lKeyId          As Long
Dim lResult         As Long

DeleteRegValue = 0            ' Assume success

' Make sure all parameters have values
If Len(sRegistryKey) = 0 Then
    ' The key parameter is not set
    DeleteRegValue = ERROR_BADKEY
    Exit Function
End If

If Len(sSubKey) = 0 Then
    ' The sub key parameter is not set
    DeleteRegValue = ERROR_BADKEY
    Exit Function
End If

' Open the key by attempting to create it. If it
' already exists, an ID is returned.
lResult = RegCreateKey(HKEY_LOCAL_MACHINE, sRegistryKey, lKeyId)

If lResult = 0 Then
    ' Got a key ID so delete the value
    DeleteRegValue = RegDeleteValue(lKeyId, ByVal sSubKey)
End If

End Function

Public Function GetRegValue _
            (sRegistryKey As String, _
             sSubKey As String, _
             sKeyValue As String) As Long

Dim lResult             As Long
Dim lKeyId              As Long
Dim lBufferSize         As Long

GetRegValue = 0                ' Assume success

' Clear the return string parameter
sKeyValue = Empty

' Make sure all parameters have values
If Len(sRegistryKey) = 0 Then
    ' The key parameter is not set
    GetRegValue = ERROR_BADKEY
    Exit Function
End If

If Len(sSubKey) = 0 Then
    ' The sub key parameter is not set
    GetRegValue = ERROR_BADKEY
    Exit Function
End If

' Open the key by attempting to create it. If it
' already exists, an ID is returned.
lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
                sRegistryKey, lKeyId)
                
If lResult <> 0 Then
    ' Call failed, can't open the key so exit
    GetRegValue = lResult
    Exit Function
End If

' Determine the size of the data in the registry entry
lResult = RegQueryValueEx(lKeyId, sSubKey, _
                0&, REG_SZ, 0&, lBufferSize)
                
If lBufferSize < 2 Then
    ' No data value available
    Exit Function
End If

' Allocate the needed space for the key data
sKeyValue = String(lBufferSize + 1, " ")

' Get the value of the registry entry
lResult = RegQueryValueEx(lKeyId, sSubKey, _
                0&, REG_SZ, ByVal sKeyValue, lBufferSize)

If lResult <> 0 Then
    ' Unexpected error, return the result
    GetRegValue = lResult

  Else

    ' Trim the null at the end of the returned value
    ' and send it back to the caller
    If InStr(sKeyValue, vbNullChar) > 0 Then
        sKeyValue = Left$(sKeyValue, lBufferSize - 1)
    End If

End If

End Function

Public Function SetRegValue _
            (sRegistryKey As String, _
             sSubKey As String, _
             sKeyValue As String) As Long

Dim lKeyId              As Long
Dim lResult             As Long

SetRegValue = 0                ' Assume success

' Make sure all parameters have values
If Len(sRegistryKey) = 0 Then
    ' The key parameter is not set
    SetRegValue = ERROR_BADKEY
    Exit Function
End If

If Len(sSubKey) = 0 Then
    ' The sub key parameter is not set
    SetRegValue = ERROR_BADKEY
    Exit Function
End If

' Open the key by attempting to create it. If it
' already exists, an ID is returned.
lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
                sRegistryKey, _
                lKeyId)

If lResult <> 0 Then
    ' Call failed, can't open the key so exit
    SetRegValue = lResult
    Exit Function
End If

If Len(sKeyValue) = 0 Then
    ' No key value, so clear any existing entry
    SetRegValue = RegSetValueEx(lKeyId, _
                sSubKey, _
                0&, _
                REG_SZ, _
                0&, _
                0&)

  Else

    ' Set the registry entry to the value
    SetRegValue = RegSetValueEx(lKeyId, _
                sSubKey, _
                0&, _
                REG_SZ, _
                ByVal sKeyValue, _
                Len(sKeyValue) + 1)

End If

End Function

Public Function GetRegErrorText(lStatus As Long) As String

' Evaluate the status and return the error message text.

Select Case lStatus
  Case ERROR_BADDB
    GetRegErrorText = "The configuration registry database " & _
                      "is corrupt."
    
  Case ERROR_BADKEY
    GetRegErrorText = "The configuration registry key is " & _
                      "invalid."
    
  Case ERROR_CANTOPEN
    GetRegErrorText = "The configuration registry key could " & _
                      "not be opened."
    
  Case ERROR_CANTREAD
    GetRegErrorText = "The configuration registry key could " & _
                      "not be read."
    
  Case ERROR_CANTWRITE
    GetRegErrorText = "The configuration registry key could " & _
                      "not be written."
    
  Case ERROR_REGISTRY_RECOVERED
    GetRegErrorText = "One of the files in the Registry " & _
                      "database had to be recovered " & _
                      "by use of a log or alternate copy. " & _
                      "The recovery was successful."
                   
  Case ERROR_REGISTRY_CORRUPT
    GetRegErrorText = "The Registry is corrupt. The structure " & _
                      "of one of the files that contains " & _
                      "Registry data is corrupt, or the " & _
                      "system's image of the file in memory " & _
                      "is corrupt, or the file could not be " & _
                      "recovered because the alternate " & _
                      "copy or log was absent or corrupt."
                   
  Case ERROR_REGISTRY_IO_FAILED
    GetRegErrorText = "An I/O operation initiated by the " & _
                      "Registry failed unrecoverably. " & _
                      "The Registry could not read in, or " & _
                      "write out, or flush, one of the files " & _
                      "that contain the system's image of " & _
                      "the Registry."
                   
  Case ERROR_NOT_REGISTRY_FILE
    GetRegErrorText = "The system has attempted to load or " & _
                      "restore a file into the Registry, but the " & _
                      "specified file is not in a Registry " & _
                      "file format."
                   
  Case ERROR_KEY_DELETED
    GetRegErrorText = "Illegal operation attempted on a " & _
                      "Registry key which has been marked " & _
                      "for deletion."
    
  Case ERROR_NO_LOG_SPACE
    GetRegErrorText = "System could not allocate the required " & _
                      "space in a Registry log."
    
  Case ERROR_KEY_HAS_CHILDREN
    GetRegErrorText = "Cannot create a symbolic link in a " & _
                      "Registry key that already " & _
                      "has subkeys or values."
                   
  Case ERROR_CHILD_MUST_BE_VOLATILE
    GetRegErrorText = "Cannot create a stable subkey under a " & _
                      "volatile parent key."
    
  Case ERROR_RXACT_INVALID_STATE
    GetRegErrorText = "The transaction state of a Registry " & _
                      "subtree is incompatible with the " & _
                      "requested operation."
                   
End Select

End Function


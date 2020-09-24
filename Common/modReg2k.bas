Attribute VB_Name = "modReg2k"
Private Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type

Private Declare Function RegCloseKey _
Lib "advapi32.dll" _
(ByVal lngHKey As Long) _
As Long

Private Declare Function RegCreateKeyEx _
Lib "advapi32.dll" _
Alias "RegCreateKeyExA" _
(ByVal lngHKey As Long, _
ByVal lpSubKey As String, _
ByVal Reserved As Long, _
ByVal lpClass As String, _
ByVal dwOptions As Long, _
ByVal samDesired As Long, _
ByVal lpSecurityAttributes As Long, _
phkResult As Long, _
lpdwDisposition As Long) _
As Long

Private Declare Function RegOpenKeyEx _
Lib "advapi32.dll" _
Alias "RegOpenKeyExA" _
(ByVal lngHKey As Long, _
ByVal lpSubKey As String, _
ByVal ulOptions As Long, _
ByVal samDesired As Long, _
phkResult As Long) _
As Long

Private Declare Function RegQueryValueExString _
Lib "advapi32.dll" _
Alias "RegQueryValueExA" _
(ByVal lngHKey As Long, _
ByVal lpValueName As String, _
ByVal lpReserved As Long, _
lpType As Long, _
ByVal lpData As String, _
lpcbData As Long) _
As Long

Private Declare Function RegQueryValueExLong _
Lib "advapi32.dll" _
Alias "RegQueryValueExA" _
(ByVal lngHKey As Long, _
ByVal lpValueName As String, _
ByVal lpReserved As Long, _
lpType As Long, _
lpData As Long, _
lpcbData As Long) _
As Long

Private Declare Function RegQueryValueExBinary _
Lib "advapi32.dll" _
Alias "RegQueryValueExA" _
(ByVal lngHKey As Long, _
ByVal lpValueName As String, _
ByVal lpReserved As Long, _
lpType As Long, _
ByVal lpData As Long, _
lpcbData As Long) _
As Long
  
Private Declare Function RegQueryValueExNULL _
Lib "advapi32.dll" _
Alias "RegQueryValueExA" _
(ByVal lngHKey As Long, _
ByVal lpValueName As String, _
ByVal lpReserved As Long, _
lpType As Long, _
ByVal lpData As Long, _
lpcbData As Long) _
As Long

Private Declare Function RegSetValueExString _
Lib "advapi32.dll" _
Alias "RegSetValueExA" _
(ByVal lngHKey As Long, _
ByVal lpValueName As String, _
ByVal Reserved As Long, _
ByVal dwType As Long, _
ByVal lpValue As String, _
ByVal cbData As Long) _
As Long

Private Declare Function RegSetValueExLong _
Lib "advapi32.dll" _
Alias "RegSetValueExA" _
(ByVal lngHKey As Long, _
ByVal lpValueName As String, _
ByVal Reserved As Long, _
ByVal dwType As Long, _
lpValue As Long, _
ByVal cbData As Long) _
As Long

Private Declare Function RegSetValueExBinary _
Lib "advapi32.dll" _
Alias "RegSetValueExA" _
(ByVal lngHKey As Long, _
ByVal lpValueName As String, _
ByVal Reserved As Long, _
ByVal dwType As Long, _
ByVal lpValue As Long, _
ByVal cbData As Long) _
As Long
  
Private Declare Function RegEnumKey _
Lib "advapi32.dll" _
Alias "RegEnumKeyA" _
(ByVal lngHKey As Long, _
ByVal dwIndex As Long, _
ByVal lpName As String, _
ByVal cbName As Long) _
As Long

Private Declare Function RegQueryInfoKey _
Lib "advapi32.dll" _
Alias "RegQueryInfoKeyA" _
(ByVal lngHKey As Long, _
ByVal lpClass As String, _
ByVal lpcbClass As Long, _
ByVal lpReserved As Long, _
lpcSubKeys As Long, _
lpcbMaxSubKeyLen As Long, _
ByVal lpcbMaxClassLen As Long, _
lpcValues As Long, _
lpcbMaxValueNameLen As Long, _
ByVal lpcbMaxValueLen As Long, _
ByVal lpcbSecurityDescriptor As Long, _
lpftLastWriteTime As FILETIME) _
As Long

Private Declare Function RegEnumValue _
Lib "advapi32.dll" _
Alias "RegEnumValueA" _
(ByVal lngHKey As Long, _
ByVal dwIndex As Long, _
ByVal lpValueName As String, _
lpcbValueName As Long, _
ByVal lpReserved As Long, _
ByVal lpType As Long, _
ByVal lpData As Byte, _
ByVal lpcbData As Long) _
As Long

Private Declare Function RegDeleteKey _
Lib "advapi32.dll" _
Alias "RegDeleteKeyA" _
(ByVal lngHKey As Long, _
ByVal lpSubKey As String) _
As Long
Private Declare Function RegDeleteValue _
Lib "advapi32.dll" _
Alias "RegDeleteValueA" _
(ByVal lngHKey As Long, _
ByVal lpValueName As String) _
As Long
Public Enum EnumRegistryRootKeys
rrkHKeyClassesRoot = &H80000000
rrkHKeyCurrentUser = &H80000001
rrkHKeyLocalMachine = &H80000002
rrkHKeyUsers = &H80000003
End Enum

Public Enum EnumRegistryValueType
rrkRegSZ = 1
rrkregBinary = 3
rrkRegDWord = 4
End Enum

Private Const mcregOptionNonVolatile = 0
Private Const mcregErrorNone = 0
Private Const mcregErrorBadDB = 1
Private Const mcregErrorBadKey = 2
Private Const mcregErrorCantOpen = 3
Private Const mcregErrorCantRead = 4
Private Const mcregErrorCantWrite = 5
Private Const mcregErrorOutOfMemory = 6
Private Const mcregErrorInvalidParameter = 7
Private Const mcregErrorAccessDenied = 8
Private Const mcregErrorInvalidParameterS = 87
Private Const mcregErrorNoMoreItems = 259

Public Const mcregSynchronize = &H100000

Public Const mcregKeyQueryValue = &H1
Public Const mcregKeySetValue = &H2
Public Const mcregKeyCreateSubKey = &H4
Public Const mcregKeyEnumerateSubKeys = &H8
Public Const mcregKeyCreateLink = &H20
Public Const mcregKeyNotify = &H10
Public Const mcregReadControl = &H20000
Public Const mcregStandardRightsAll = &H1F0000
Public Const mcregStandardRightsRead = (mcregReadControl)
Public Const mcregStandardRightsWrite = (mcregReadControl)

Public Const mcregKeyAllAccess = ((mcregStandardRightsAll Or mcregKeyQueryValue Or mcregKeySetValue Or mcregKeyCreateSubKey Or mcregKeyEnumerateSubKeys Or mcregKeyNotify Or mcregKeyCreateLink) And (Not mcregSynchronize))
Public Const mcregKeyRead = ((mcregStandardRightsRead Or mcregKeyQueryValue Or mcregKeyEnumerateSubKeys Or mcregKeyNotify) And (Not mcregSynchronize))
Public Const mcregKeyWrite = ((mcregStandardRightsWrite Or mcregKeySetValue Or mcregKeyCreateSubKey) And (Not mcregSynchronize))
Public Sub RegistryCreateNewKey( _
eRootKey As EnumRegistryRootKeys, _
strKeyName As String)
On Error Resume Next
Dim lngRetVal As Long
Dim lngHKey As Long
    
On Error GoTo PROC_ERR
lngRetVal = RegCreateKeyEx(eRootKey, strKeyName, 0&, vbNullString, _
mcregOptionNonVolatile, mcregKeyWrite, 0&, lngHKey, 0&)
If lngRetVal = mcregErrorNone Then
RegCloseKey (lngHKey)
End If
PROC_EXIT:
Exit Sub
PROC_ERR:
MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
"RegistryCreateNewKey"
Resume PROC_EXIT
End Sub
Public Sub RegistryDeleteKey( _
eRootKey As EnumRegistryRootKeys, _
strKeyName As String)
On Error Resume Next
Dim lngRetVal As Long
On Error GoTo PROC_ERR
lngRetVal = RegDeleteKey(eRootKey, strKeyName)
    
PROC_EXIT:
Exit Sub
PROC_ERR:
MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
"RegistryDeleteKey"
Resume PROC_EXIT
End Sub
Public Sub RegistryDeleteValue( _
eRootKey As EnumRegistryRootKeys, _
strKeyName As String, _
strValueName As String)
On Error Resume Next

Dim lngRetVal As Long
Dim lngHKey As Long

On Error GoTo PROC_ERR

lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, mcregKeyWrite, _
lngHKey)
If lngRetVal = mcregErrorNone Then
lngRetVal = RegDeleteValue(lngHKey, strValueName)
End If
PROC_EXIT:
Exit Sub
PROC_ERR:
MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
"RegistryDeleteValue"
Resume PROC_EXIT
End Sub
Public Sub RegistryEnumerateSubKeys( _
eRootKey As EnumRegistryRootKeys, _
strKeyName As String, _
astrKeys() As String, _
lngKeyCount As Long)
On Error Resume Next

Dim lngRetVal As Long
Dim lngHKey As Long
Dim lngKeyIndex As Long
Dim strSubKeyName As String
Dim lngSubkeyCount As Long
Dim lngMaxKeyLen As Long
Dim typFT As FILETIME
  
On Error GoTo PROC_ERR

lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, mcregKeyRead, _
lngHKey)
If lngRetVal = mcregErrorNone Then
If mcregErrorNone = lngRetVal Then
If lngSubkeyCount > 0 Then
ReDim astrKeys(lngSubkeyCount - 1) As String
lngKeyIndex = 0
lngMaxKeyLen = lngMaxKeyLen + 1
strSubKeyName = Space$(lngMaxKeyLen)
        
Do While RegEnumKey(lngHKey, lngKeyIndex, strSubKeyName, lngMaxKeyLen + 1) = 0
If InStr(1, strSubKeyName, vbNullChar) > 0 Then
astrKeys(lngKeyIndex) = Left$(strSubKeyName, InStr(1, strSubKeyName, vbNullChar) - 1)
End If
lngKeyIndex = lngKeyIndex + 1
        
Loop
End If
lngKeyCount = lngSubkeyCount
End If
    
RegCloseKey (lngHKey)
End If
PROC_EXIT:
Exit Sub
PROC_ERR:
MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
"RegistryEnumerateSubKeys"
Resume PROC_EXIT
End Sub
Public Sub RegistryEnumerateValues( _
eRootKey As EnumRegistryRootKeys, _
strKeyName As String, _
astrValues() As String, _
lngValueCount As Long)
On Error Resume Next
Dim lngRetVal As Long
Dim lngHKey As Long
Dim lngKeyIndex As Long
Dim strValueName As String
Dim lngTempValueCount As Long
Dim lngMaxValueLen As Long
Dim typFT As FILETIME
  
On Error GoTo PROC_ERR
lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, mcregKeyRead, _
lngHKey)
If lngRetVal = mcregErrorNone Then
lngRetVal = RegQueryInfoKey(lngHKey, vbNullString, 0, 0, 0, _
0, 0, lngTempValueCount, lngMaxValueLen, 0, 0, typFT)
If mcregErrorNone = lngRetVal Then
If lngTempValueCount > 0 Then
ReDim astrValues(lngTempValueCount - 1) As String
lngKeyIndex = 0
lngMaxValueLen = lngMaxValueLen + 1
strValueName = Space$(lngMaxValueLen)
        
Do While RegEnumValue(lngHKey, lngKeyIndex, strValueName, _
lngMaxValueLen + 1, 0, 0, 0, 0) = 0
        
If InStr(1, strValueName, vbNullChar) > 0 Then
astrValues(lngKeyIndex) = Left$(strValueName, InStr(1, strValueName, vbNullChar) - 1)
End If
lngKeyIndex = lngKeyIndex + 1
        
Loop
End If
lngValueCount = lngTempValueCount
End If
    
RegCloseKey (lngHKey)
End If
PROC_EXIT:
Exit Sub
PROC_ERR:
MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
"RegistryEnumerateValues"
Resume PROC_EXIT
End Sub
Public Function RegistryGetKeyValue( _
eRootKey As EnumRegistryRootKeys, _
strKeyName As String, _
strValueName As String) _
As Variant
On Error Resume Next
Dim lngRetVal As Long
Dim lngHKey As Long
Dim varValue As Variant
Dim strValueData As String
Dim abytValueData() As Byte
Dim lngValueData As Long
Dim lngValueType As Long
Dim lngDataSize As Long
  
On Error GoTo PROC_ERR
varValue = Empty
lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0&, mcregKeyRead, _
lngHKey)
If mcregErrorNone = lngRetVal Then
lngRetVal = RegQueryValueExNULL(lngHKey, strValueName, 0&, lngValueType, 0&, lngDataSize)
If lngRetVal = mcregErrorNone Then
Select Case lngValueType
Case rrkRegSZ:
If lngDataSize > 0 Then
strValueData = String(lngDataSize, 0)
lngRetVal = RegQueryValueExString(lngHKey, strValueName, 0&, _
lngValueType, strValueData, lngDataSize)
If InStr(strValueData, vbNullChar) > 0 Then
strValueData = Mid$(strValueData, 1, InStr(strValueData, _
vbNullChar) - 1)
End If
End If
If mcregErrorNone = lngRetVal Then
varValue = Left$(strValueData, lngDataSize)
Else
varValue = Empty
End If
Case rrkRegDWord:
lngRetVal = RegQueryValueExLong(lngHKey, strValueName, 0&, _
lngValueType, lngValueData, lngDataSize)
If mcregErrorNone = lngRetVal Then
varValue = lngValueData
End If
Case rrkregBinary
If lngDataSize > 0 Then
ReDim abytValueData(lngDataSize - 1) As Byte
lngRetVal = RegQueryValueExBinary(lngHKey, strValueName, 0&, _
lngValueType, VarPtr(abytValueData(0)), lngDataSize)
End If
If mcregErrorNone = lngRetVal Then
varValue = abytValueData
Else
varValue = Empty
End If
Case Else
lngRetVal = -1
        
End Select
End If
RegCloseKey (lngHKey)
End If
RegistryGetKeyValue = varValue
  
PROC_EXIT:
Exit Function
PROC_ERR:
MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
"RegistryGetKeyValue"
Resume PROC_EXIT
End Function

Public Sub RegistrySetKeyValue( _
eRootKey As EnumRegistryRootKeys, _
strKeyName As String, _
strValueName As String, _
varData As Variant, _
eDataType As EnumRegistryValueType)
On Error Resume Next
Dim lngRetVal As Long
Dim lngHKey As Long
Dim strData As String
Dim lngData As Long
Dim abytData() As Byte
    
On Error GoTo PROC_ERR
lngRetVal = RegCreateKeyEx(eRootKey, strKeyName, 0&, vbNullString, _
mcregOptionNonVolatile, mcregKeyRead Or mcregKeyWrite, 0&, lngHKey, 0&)
Select Case eDataType
  
Case rrkRegSZ
    strData = varData & vbNullChar
    lngRetVal = RegSetValueExString(lngHKey, strValueName, 0&, eDataType, strData, Len(strData))
Case rrkRegDWord
    lngData = varData
    lngRetVal = RegSetValueExLong(lngHKey, strValueName, 0&, eDataType, lngData, Len(lngData))
Case rrkregBinary
    abytData = varData
    lngRetVal = RegSetValueExBinary(lngHKey, strValueName, 0&, eDataType, VarPtr(abytData(0)), UBound(abytData) + 1)
End Select
RegCloseKey (lngHKey)
PROC_EXIT:
Exit Sub
PROC_ERR:
MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
"RegistrySetKeyValue"
Resume PROC_EXIT
End Sub

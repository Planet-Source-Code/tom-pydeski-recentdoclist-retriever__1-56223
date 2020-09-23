Attribute VB_Name = "Registry"
Option Explicit
Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
End Type
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'
'was Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long
Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
'
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
'
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
'from another -------------------------------------------
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
' Reg Data Types...
Global Const REG_NONE = 0                       ' No value type
Global Const REG_SZ = 1                         ' Unicode nul terminated string
Global Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Global Const REG_BINARY = 3                     ' Free form binary
Global Const REG_DWORD = 4                      ' 32-bit number
Global Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Global Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Global Const REG_LINK = 6                       ' Symbolic Link (unicode)
Global Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Global Const REG_RESOURCE_LIST = 8              ' Resource list in the resource map
Global Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
'Global Const HKEY_CLASSES_ROOT = &H80000000
'Global Const HKEY_CURRENT_USER = &H80000001
'Global Const HKEY_LOCAL_MACHINE = &H80000002
'Global Const HKEY_USERS = &H80000003
' Reg Key ROOT Types
Public Enum RegistryRoots
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum
Private Const ERROR_SUCCESS = 0&
Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259
Global Const KEY_ALL_ACCESS = &H3F
Global Const SYNCHRONIZE = &H100000
Global Const STANDARD_RIGHTS_ALL = &H1F0000
' Reg Key Security Options
Global Const KEY_QUERY_VALUE = &H1
Global Const KEY_SET_VALUE = &H2
Global Const KEY_CREATE_SUB_KEY = &H4
Global Const KEY_ENUMERATE_SUB_KEYS = &H8
Global Const KEY_NOTIFY = &H10
Global Const KEY_CREATE_LINK = &H20
'Global Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Global Const REG_OPTION_NON_VOLATILE = 0
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ
Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte
Const DisplayErrorMsg = False
Public DeskTopPath$
Public StartMenuPath$
Public StartUpPath$
Public ProgramMenuPath$
Public CachePath$
Public Const MAX_PATH = 260
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Dim eTitle$
Dim eMess$
Dim mError As Long

Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
Dim hNewKey As Long 'handle to the new key
Dim lRetVal As Long 'result of the RegCreateKeyEx function
lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
RegCloseKey (hNewKey)
End Sub

Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
Dim Zero As Long, IRetVal As Long, hKey As Long, OrigKeyNam As String
'OrigKeyNam = Left$(sKeyName, InStr(sKeyName + "\", "\") - 1)
'open the specified key
IRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, Zero, KEY_ALL_ACCESS, hKey)
If IRetVal Then MsgBox "RegOpenKey error - " & IRetVal
IRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
If IRetVal Then MsgBox "SetValue error - " & IRetVal
RegCloseKey (hKey)
End Sub

Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
Dim lValue As Long
Dim sValue As String
Select Case lType
    Case REG_SZ
        sValue = vValue & Chr$(0)
        SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
    Case REG_DWORD
        lValue = vValue
        SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
End Select
End Function

Sub QueryValue(sKeyName As String, sValueName As String)
Dim lRetVal As Long 'result of the API functions
Dim hKey As Long    'handle of opened key
Dim vValue As Variant 'setting of queried value
lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, KEY_ALL_ACCESS, hKey)
lRetVal = QueryValueEx(hKey, sValueName, vValue)
MsgBox vValue
RegCloseKey (hKey)
End Sub

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
Dim cch As Long
Dim lrc As Long
Dim lType As Long
Dim lValue As Long
Dim sValue As String
On Error GoTo QueryValueExError
' Determine the size and type of data to be read
lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
If lrc <> ERROR_NONE Then Error 5
Select Case lType
    ' For strings
Case REG_SZ:
    sValue = String(cch, 0)
    lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
    If lrc = ERROR_NONE Then
        vValue = Left$(sValue, cch)
    Else
        vValue = Empty
    End If
    ' For DWORDS
Case REG_DWORD:
    lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
    If lrc = ERROR_NONE Then vValue = lValue
Case Else
    'all other data types not supported
    lrc = -1
End Select
QueryValueExExit:
QueryValueEx = lrc
Exit Function
QueryValueExError:
Resume QueryValueExExit
End Function

Public Sub SaveKey(hKey As Long, strPath As String)
Dim KeyHand&
Dim lResult As Long
lResult = RegCreateKey(hKey, strPath, KeyHand&)
lResult = RegCloseKey(KeyHand&)
End Sub

Public Function GetString(hKey As Long, strPath As String, strValue As String) As String
Dim KeyHand As Long
Dim lResult As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Debug.Print strPath; " "; strValue; " "; hKey
lResult = RegOpenKey(hKey, strPath, KeyHand)
lResult = RegQueryValueEx(KeyHand, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
If lResult = ERROR_SUCCESS Then
    If lValueType = REG_SZ Then
        strBuffer = String(lDataBufferSize, " ")
        lResult = RegQueryValueEx(KeyHand, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
        intZeroPos = InStr(strBuffer, Chr$(0))
        If intZeroPos > 0 Then
            GetString = Left$(strBuffer, intZeroPos - 1)
        Else
            GetString = strBuffer
        End If
    End If
Else
    ' there is a problem
End If
lResult = RegCloseKey(KeyHand)
End Function

Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
Dim KeyHand As Long
Dim R As Long
R = RegCreateKey(hKey, strPath, KeyHand)
R = RegSetValueEx(KeyHand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
R = RegCloseKey(KeyHand)
End Sub

Function GetDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
'text1.text = getdword(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufferSize As Long
Dim R As Long
Dim KeyHand As Long
R = RegOpenKey(hKey, strPath, KeyHand)
' Get length/data type
lDataBufferSize = 4
lResult = RegQueryValueEx(KeyHand, strValueName, 0&, lValueType, lBuf, lDataBufferSize)
If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
        GetDWord = lBuf
    End If
End If
R = RegCloseKey(KeyHand)
End Function

Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
Dim lResult As Long
Dim KeyHand As Long
Dim R As Long
R = RegCreateKey(hKey, strPath, KeyHand)
lResult = RegSetValueEx(KeyHand, strValueName, 0&, REG_DWORD, lData, 4)
R = RegCloseKey(KeyHand)
End Function

Function DeleteKey(KeyName As String) As Long
'usage=DeleteKey HKEY_LOCAL_MACHINE, "Software\OSI\Visualicious\Recovery\" & ListView1.SelectedItem.Text & ".exe"
If MainKeyHandle Then
    rtn = RegOpenKeyEx(MainKeyHandle, KeyName, 0, KEY_WRITE, hKey)
    If rtn = ERROR_SUCCESS Then
        rtn = RegDeleteKey(hKey, KeyName)
        DeleteKey = rtn
        rtn = RegCloseKey(hKey)
    End If
End If
End Function

Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String) As Long
'usage=DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", ListView1.SelectedItem.SubItems(1)
Dim KeyHand As Long
Dim lResult As Long
lResult = RegOpenKey(hKey, strPath, KeyHand)
If lResult <> 0 Then
    MsgBox "Failed to Open " & strPath
    DeleteValue = lResult
    Exit Function
End If
lResult = RegDeleteValue(KeyHand, strValue)
If lResult <> 0 Then
    MsgBox "Failed to Delete " & strValue
    DeleteValue = lResult
    lResult = RegCloseKey(KeyHand)
    Exit Function
End If
lResult = RegCloseKey(KeyHand)
DeleteValue = lResult
End Function

Function SetDWordValue(SubKey As String, Entry As String, Value As Long)
Call ParseKey(SubKey, MainKeyHandle)
If MainKeyHandle Then
    rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey)
    If rtn = ERROR_SUCCESS Then
        rtn = RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4)
        If Not rtn = ERROR_SUCCESS Then
            If DisplayErrorMsg = True Then
                MsgBox ErrorMsg(rtn)
            End If
        End If
        rtn = RegCloseKey(hKey)
    Else
        If DisplayErrorMsg = True Then
            MsgBox ErrorMsg(rtn)
        End If
    End If
End If
End Function

Function GetDWORDValue(SubKey As String, Entry As String)
Call ParseKey(SubKey, MainKeyHandle)
If MainKeyHandle Then
    rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey)
    If rtn = ERROR_SUCCESS Then
        rtn = RegQueryValueExA(hKey, Entry, 0, REG_DWORD, lBuffer, 4)
        If rtn = ERROR_SUCCESS Then
            rtn = RegCloseKey(hKey)
            GetDWORDValue = lBuffer
        Else
            GetDWORDValue = "Error"
            If DisplayErrorMsg = True Then
                MsgBox ErrorMsg(rtn)
            End If
        End If
    Else
        GetDWORDValue = "Error"
        If DisplayErrorMsg = True Then
            MsgBox ErrorMsg(rtn)
        End If
    End If
End If
End Function

Function SetBinaryValue(SubKey As String, Entry As String, Value As String)
Call ParseKey(SubKey, MainKeyHandle)
Dim i As Long
If MainKeyHandle Then
    rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey)
    If rtn = ERROR_SUCCESS Then
        lDataSize = Len(Value)
        ReDim ByteArray(lDataSize)
        For i = 1 To lDataSize
            ByteArray(i) = Asc(Mid$(Value, i, 1))
        Next
        rtn = RegSetValueExB(hKey, Entry, 0, REG_BINARY, ByteArray(1), lDataSize)
        If Not rtn = ERROR_SUCCESS Then
            If DisplayErrorMsg = True Then
                MsgBox ErrorMsg(rtn)
            End If
        End If
        rtn = RegCloseKey(hKey)
    Else
        If DisplayErrorMsg = True Then
            MsgBox ErrorMsg(rtn)
        End If
    End If
End If
End Function

Function GetBinaryValue(SubKey As String, Entry As String)
Call ParseKey(SubKey, MainKeyHandle)
If MainKeyHandle Then
    rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey)
    If rtn = ERROR_SUCCESS Then
        lBufferSize = 1
        rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, 0&, lBufferSize)
        sBuffer = Space(lBufferSize)
        rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize)
        If rtn = ERROR_SUCCESS Then
            rtn = RegCloseKey(hKey)
            GetBinaryValue = sBuffer
        Else
            GetBinaryValue = "Error"
            If DisplayErrorMsg = True Then
                MsgBox ErrorMsg(rtn)
            End If
        End If
    Else
        GetBinaryValue = "Error"
        If DisplayErrorMsg = True Then
            MsgBox ErrorMsg(rtn)
        End If
    End If
End If
End Function

Function GetMainKeyHandle(MainKeyName As String) As Long
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
Select Case MainKeyName
    Case "HKEY_CLASSES_ROOT"
        GetMainKeyHandle = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        GetMainKeyHandle = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
        GetMainKeyHandle = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS"
        GetMainKeyHandle = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA"
        GetMainKeyHandle = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG"
        GetMainKeyHandle = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
        GetMainKeyHandle = HKEY_DYN_DATA
End Select
End Function

Function ErrorMsg(lErrorCode As Long) As String
Dim getErrorMsg As String
Select Case lErrorCode
    Case 1009, 1015
        getErrorMsg = "The Registry Database is corrupt!"
    Case 2, 1010
        getErrorMsg = "Bad Key Name"
    Case 1011
        getErrorMsg = "Can't Open Key"
    Case 4, 1012
        getErrorMsg = "Can't Read Key"
    Case 5
        getErrorMsg = "Access to this key is denied"
    Case 1013
        getErrorMsg = "Can't Write Key"
    Case 8, 14
        getErrorMsg = "Out of memory"
    Case 87
        getErrorMsg = "Invalid Parameter"
    Case 234
        getErrorMsg = "There is more data than the buffer has been allocated to hold."
    Case Else
        getErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
End Select
End Function

Function GetStringValue(SubKey As String, Entry As String)
Call ParseKey(SubKey, MainKeyHandle)
If MainKeyHandle Then
    rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey)
    If rtn = ERROR_SUCCESS Then
        sBuffer = Space(255)
        lBufferSize = Len(sBuffer)
        rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize)
        If rtn = ERROR_SUCCESS Then
            rtn = RegCloseKey(hKey)
            sBuffer = Trim(sBuffer)
            GetStringValue = Left(sBuffer, Len(sBuffer) - 1)
        Else
            GetStringValue = "Error"
            If DisplayErrorMsg = True Then
                MsgBox ErrorMsg(rtn)
            End If
        End If
    Else
        GetStringValue = "Error"
        If DisplayErrorMsg = True Then
            MsgBox ErrorMsg(rtn)
        End If
    End If
End If
End Function

Private Sub ParseKey(KeyName As String, Keyhandle As Long)
rtn = InStr(KeyName, "\")
If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then
    MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + KeyName 'display error to the user
    Exit Sub
ElseIf rtn = 0 Then
    Keyhandle = GetMainKeyHandle(KeyName)
    KeyName = ""
Else
    Keyhandle = GetMainKeyHandle(Left(KeyName, rtn - 1))
    KeyName = Right(KeyName, Len(KeyName) - rtn)
End If
End Sub

Function CreateKey(SubKey As String)
Call ParseKey(SubKey, MainKeyHandle)
If MainKeyHandle Then
    rtn = RegCreateKey(MainKeyHandle, SubKey, hKey)
    If rtn = ERROR_SUCCESS Then
        rtn = RegCloseKey(hKey)
    End If
End If
End Function

Function SetStringValue(SubKey As String, Entry As String, Value As String)
Call ParseKey(SubKey, MainKeyHandle)
If MainKeyHandle Then
    rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey)
    If rtn = ERROR_SUCCESS Then
        rtn = RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value))
        If Not rtn = ERROR_SUCCESS Then
            If DisplayErrorMsg = True Then
                MsgBox ErrorMsg(rtn)
            End If
        End If
        rtn = RegCloseKey(hKey)
    Else
        If DisplayErrorMsg = True Then
            MsgBox ErrorMsg(rtn)
        End If
    End If
End If
End Function

Public Sub GetLocations()
DeskTopPath$ = (GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop"))
StartMenuPath$ = (GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Start Menu"))
StartUpPath$ = (GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup"))
ProgramMenuPath$ = (GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Programs"))
CachePath$ = (GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Cache"))
Dim eM$
eM$ = DeskTopPath$ & vbCrLf
eM$ = eM$ & StartMenuPath$ & vbCrLf
eM$ = eM$ & StartUpPath$ & vbCrLf
eM$ = eM$ & ProgramMenuPath$ & vbCrLf
eM$ = eM$ & CachePath$ & vbCrLf
MsgBox eM$
End Sub

Public Function GetAllSubKeys(ByVal KeyName As String, ByRef SubKeys() As String) As Long
On Error GoTo Oops
'below from kegham
'*****************************************************************************************
'* Function    : GetAllSubKeys
'* Notes       : Retrieves all the subkeys belonging to a registry key.
'*               Returns a long integer value containing the number of retrieved subkeys.
'*****************************************************************************************
Dim Count        As Long
Dim dwReserved   As Long
Dim hKey         As Long
Dim iPos         As Long
Dim lenBuffer    As Long
Dim lIndex       As Long
Dim lRet         As Long
Dim lType        As Long
Dim sCompKey     As String
Dim szBuffer     As String
Erase SubKeys
Count = 0
lIndex = 0
lRet = RegOpenKeyEx(HKEY_CURRENT_USER, KeyName, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
Do While lRet = ERROR_SUCCESS
    szBuffer = String$(MAX_PATH, 0)
    lenBuffer = Len(szBuffer)
    lRet = RegEnumKey(hKey, lIndex, szBuffer, lenBuffer)
    If (lRet = ERROR_SUCCESS) Then
        Count = Count + 1
        ReDim Preserve SubKeys(0 To Count) As String
        iPos = InStr(szBuffer, Chr$(0))
        If (iPos > 0) Then
            SubKeys(Count) = Left$(szBuffer, iPos - 1)
        Else
            SubKeys(Count) = Left$(szBuffer, lenBuffer)
        End If
    End If
    lIndex = lIndex + 1
Loop
If (hKey <> 0) Then RegCloseKey hKey
GetAllSubKeys = Count
GoTo Exit_GetAllSubKeys
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine GetAllSubKeys "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in GetAllSubKeys"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
If (hKey <> 0) Then RegCloseKey hKey
GetAllSubKeys = 0
Exit_GetAllSubKeys:
End Function


Attribute VB_Name = "RecentRegistry"
Option Explicit
'declares and constants from article
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
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Declare Function RegQueryValueExAny Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Global Const REG_BINARY As Long = 3
Global SubKey
Global sKeyName As String
Global RecMax
Global ThisApp
'array for the key names within the root
Global SubKeyName() As String
'array for the maximum number of entries for each key
Global RecentMax() As Long
'2d array for the String values entries for each entry within each subkey
Global RegValue() As String
'maximum number of sub keys
Global RecentSubKeys As Long
Global lRet As Long
Dim i As Integer
Dim RecentDoc$
Public Declare Function NTBeep Lib "kernel32" Alias "Beep" (ByVal FreqHz As Long, ByVal DurationMs As Long) As Long
Global eTitle$
Global eMess$
Global mError As Long
Global AppDir
Global AppName As String

Sub GetRootBinary()
On Error GoTo Oops
Dim lRetVal As Long   'result of API function
Dim hKey As Long      'handle of opened key
Dim vValue As Variant 'setting of value
base:
'recent doc files are in binary form in
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"
'
'open registry key
sKeyName = "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"
lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_ALL_ACCESS, hKey)
If lRetVal <> 0 Then
    MsgBox "Error in GetRootBinary!"
End If
'above opens the right key and
'below goes to settings
'hkey is handle of registry key
'get the upper limit of the values listed in the root key
ReDim Preserve RecentMax(1)
RegQueryInfoKey hKey, "", 0, 0, 0, 0, 0, RecentMax(0), 0, MAX_PATH + 1, 0, 0
'now lets get the number of subkeys in the root key
RegQueryInfoKey hKey, "", 0, 0, RecentSubKeys, 0, 0, 0, 0, 0, 0, 0
'add 2 for the typedurl and runmru
ReDim Preserve RegValue(0 To RecentSubKeys + 2, 0 To RecentMax(0))
'how do we read the subkeys?
ReDim Preserve SubKeyName(0 To RecentSubKeys + 2)
SubKeyName(0) = "" 'the root name has no subkeyname
lRet = GetAllSubKeys(sKeyName, SubKeyName)
'uncomment below to show the keys in the immediate window
'Debug.Print UBound(SubKeyName)
'For i = 1 To RecentSubKeys
'    Debug.Print i, SubKeyName(i)
'Next i
'
For i = 0 To RecentMax(0)
    lRetVal = QueryValueEx(hKey, i, vValue)
    'vValue is the result of query
    If lRetVal = 0 Then
        RegValue(0, i) = vValue
    Else
        'Debug.Print "Missing "; i
    End If
Next i
'close the registry key
RegCloseKey hKey
GoTo Exit_GetRootBinary
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine GetRootBinary "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in GetRootBinary"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_GetRootBinary:
End Sub

Sub GetBinary(SubKeyName As String, SubKeyIndex As Integer)
On Error GoTo Oops
Dim lRetVal As Long   'result of API function
Dim hKey As Long      'handle of opened key
Dim vValue As Variant 'setting of value
base:
'recent doc files are in binary form in
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"
'
'open registry key
lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, SubKeyName, 0, KEY_ALL_ACCESS, hKey)
If lRetVal <> 0 Then
    MsgBox "Error in GetBinary!"
End If
'hkey is handle of registry key
'get the upper limit of the values listed in the root key
ReDim Preserve RecentMax(SubKeyIndex)
RegQueryInfoKey hKey, "", 0, 0, 0, 0, 0, RecentMax(SubKeyIndex), 0, MAX_PATH + 1, 0, 0
For i = 0 To RecentMax(SubKeyIndex)
    lRetVal = QueryValueEx(hKey, i, vValue)
    'vValue is the result of query
    If lRetVal = 0 Then
        RegValue(SubKeyIndex, i) = vValue
    Else
        'Debug.Print "Missing "; i
        RegValue(SubKeyIndex, i) = ""
    End If
Next i
'close the registry key
RegCloseKey hKey
GoTo Exit_GetBinary
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine GetBinary "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in GetBinary"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_GetBinary:
End Sub

Sub GetRegText(SubKeyName As String, KeyRoot As String, Offset As Integer)
On Error GoTo Oops
Dim lRetVal As Long   'result of API function
Dim hKey As Long      'handle of opened key
Dim vValue As Variant 'setting of value
base:
'Run Dialog Recent Menu
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU
'HKEY_USERS\S-1-5-21-127730482-1884467411-3661661970-1007\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU
'typed url list
'HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\TypedURLs
'HKEY_USERS\S-1-5-21-127730482-1884467411-3661661970-1007\Software\Microsoft\Internet Explorer\TypedURLs
'then a listing...url1...url2...etc
'
'open registry key
lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, SubKeyName, 0, KEY_ALL_ACCESS, hKey)
If lRetVal <> 0 Then
    MsgBox "Error in GetRegText!"
End If
'hkey is handle of registry key
'get the upper limit of the values listed in the root key
ReDim Preserve RecentMax(RecentSubKeys + Offset)
RegQueryInfoKey hKey, "", 0, 0, 0, 0, 0, RecentMax(RecentSubKeys + Offset), 0, MAX_PATH + 1, 0, 0
If Offset = 2 Then GoTo AddKeyRoot
For i = 0 To RecentMax(RecentSubKeys + Offset)
    lRetVal = QueryValueEx(hKey, Chr$(65 + i), vValue)
    'vValue is the result of query
    If lRetVal = 0 Then
        RegValue(RecentSubKeys + Offset, i) = vValue
    Else
        'Debug.Print "Missing "; i
        RegValue(RecentSubKeys + Offset, i) = ""
    End If
Next i
GoTo CloseKeys
'
AddKeyRoot:
For i = 0 To RecentMax(RecentSubKeys + Offset)
    lRetVal = QueryValueEx(hKey, KeyRoot & i, vValue)
    'vValue is the result of query
    If lRetVal = 0 Then
        RegValue(RecentSubKeys + Offset, i) = vValue
    Else
        'Debug.Print "Missing "; i
        RegValue(RecentSubKeys + Offset, i) = ""
    End If
Next i
CloseKeys:
'close the registry key
RegCloseKey hKey
GoTo Exit_GetRegText
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine GetRegText "
eMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in GetRegText"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
Alarm
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_GetRegText:
End Sub

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
Dim NumBytes As Long
Dim lrc As Long
Dim lType As Long
Dim lValue As Long
Dim sValue As String
Dim ValIn As Byte
Dim ChIn As String
On Error GoTo QueryValueExError
' Determine the size and type of data to be read
'LTYPE AND NumBytes ARE RETURNED FROM SUB BELOW
lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, NumBytes)
If lrc <> ERROR_NONE Then Error 5
Select Case lType
    ' For strings
Case REG_SZ:
    sValue = String(NumBytes, 0)
    lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, NumBytes)
    If lrc = ERROR_NONE Then
        vValue = Left$(sValue, NumBytes)
    Else
        vValue = Empty
    End If
    ' For DWORDS
Case REG_DWORD:
    lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, NumBytes)
    If lrc = ERROR_NONE Then vValue = lValue
Case REG_BINARY:
    Dim ByteArray() As Byte
    ReDim ByteArray(1 To CInt(NumBytes))
    On Error GoTo Oops
    RecentDoc$ = Space(MAX_PATH + 1)
    'get the data and put it into the byte array
    lrc = RegQueryValueExAny(lhKey, szValueName, 0&, lType, ByteArray(1), NumBytes)
    Dim ByteNum As Integer
    'loop through the byte array and change each byte into a character
    For ByteNum = 1 To UBound(ByteArray)
        ValIn = ByteArray(ByteNum)
        ChIn = Chr(ValIn)
        'add all of the characters to our string
        sValue = sValue & ChIn
    Next ByteNum
    'now convert our string from unicode
    RecentDoc$ = StrConv(sValue, vbFromUnicode)
    vValue = RecentDoc$
    'Clipboard.Clear
    'Clipboard.SetText RecentDoc$
    'Debug.Print RecentDoc$
getfilenames:
    '(end of Case REG_BINARY)
Case Else
    'all other data types not supported
    lrc = -1
End Select
''below is from kegham -  i think it only reads string values
'Dim EnumValue As String
'Dim lBufferLen As Long
'Dim ByteArray() As Byte
'EnumValue = Space(MAX_PATH + 1)
'Erase ByteArray
'lBufferLen = 255
'RegEnumValue hKey, i, EnumValue, MAX_PATH + 1, 0, 0, ByteArray, lBufferLen
'EnumValue = Trim(EnumValue)
QueryValueExExit:
QueryValueEx = lrc
GoTo EndErr
QueryValueExError:
Resume QueryValueExExit
Oops:
'RETRY=4,ABORT=3,IGNORE=5
eMess$ = "Error # " + Str$(Err) + " - " + Error$
Alarm
eTitle$ = AppName + " QueryValueEx"
mError = MsgBox(eMess$, 2, eTitle$)
If mError = 4 Then Resume
If mError = 5 Then Resume Next
EndErr:
End Function

Sub Main()
XPMain
Load RecentDocs
RecentDocs.Show
End Sub

Sub Beeep()
'makes a beep on the pc's internal speaker
NTBeep 500, 50
NTBeep 600, 50
NTBeep 700, 50
NTBeep 800, 50
End Sub

Sub Alarm()
'makes an alert beep on the pc's internal speaker
'usage Call NTBeep(CLng(FreqHz), CLng(LengthMs))
NTBeep 100, 50
NTBeep 200, 50
NTBeep 300, 50
NTBeep 400, 50
NTBeep 500, 50
NTBeep 600, 50
NTBeep 700, 50
NTBeep 800, 50
NTBeep 100, 50
NTBeep 200, 50
NTBeep 300, 50
NTBeep 400, 50
NTBeep 500, 50
NTBeep 600, 50
NTBeep 700, 50
NTBeep 800, 50
DoEvents
End Sub

Function AddZero(StrIn$, intDigits)
'we can actually do this by
'AddZero = Format(StrIn$, String(intDigits, "0"))
'
StrIn$ = Trim(StrIn$)
If Len(StrIn$) >= intDigits Then
    AddZero = StrIn$
    Exit Function
End If
AddZero = String(intDigits - Len(StrIn$), "0") & StrIn$
End Function

Function AddSpace(StrIn$, intDigits)
StrIn$ = Trim(StrIn$)
If Len(StrIn$) >= intDigits Then
    AddSpace = StrIn$
    Exit Function
End If
AddSpace = String(intDigits - Len(StrIn$), " ") & StrIn$
End Function



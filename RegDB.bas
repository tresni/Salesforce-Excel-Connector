Attribute VB_Name = "RegDB"
Option Explicit
 
Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4
 
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
 
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
 
Global Const KEY_QUERY_VALUE = &H1
Global Const KEY_ALL_ACCESS = &H3F
 
Global Const REG_OPTION_NON_VOLATILE = 0

#If VBA7 Then    ' VBA7
Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long

Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As LongPtr, _
    ByVal lpSubKey As String, _
    ByVal Reserved As Long, _
    ByVal lpClass As String, _
    ByVal dwOptions As Long, _
    ByVal samDesired As Long, _
    ByVal lpSecurityAttributes As Long, _
    phkResult As LongPtr, _
    lpdwDisposition As Long) As Long
    
Public Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" _
    Alias "RegOpenKeyExA" ( _
    ByVal hKey As LongPtr, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As LongPtr) As Long
    
Public Declare PtrSafe Function RegQueryValueExString Lib "advapi32.dll" _
    Alias "RegQueryValueExA" ( _
    ByVal hKey As LongPtr, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As String, _
    lpcbData As Long) As Long
    
Public Declare PtrSafe Function RegQueryValueExLong Lib "advapi32.dll" _
    Alias "RegQueryValueExA" ( _
    ByVal hKey As LongPtr, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Long, _
    lpcbData As Long) As Long
    
Public Declare PtrSafe Function RegQueryValueExNULL Lib "advapi32.dll" _
    Alias "RegQueryValueExA" ( _
    ByVal hKey As LongPtr, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Long, _
    lpcbData As Long) As Long
    
Public Declare PtrSafe Function RegSetValueExString Lib "advapi32.dll" _
    Alias "RegSetValueExA" ( _
    ByVal hKey As LongPtr, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    lpData As String, _
    ByVal cbData As Long) As Long
    
 Public Declare PtrSafe Function RegSetValueExLong Lib "advapi32.dll" _
    Alias "RegSetValueExA" ( _
    ByVal hKey As LongPtr, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    lpData As Long, _
    ByVal cbData As Long) As Long

#Else    ' Downlevel when using previous version of VBA7

Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
    "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
    As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
    As Long, phkResult As Long, lpdwDisposition As Long) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
    Long) As Long

Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As String, lpcbData As Long) As Long

Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, lpData As _
    Long, lpcbData As Long) As Long

Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As Long, lpcbData As Long) As Long

Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
    String, ByVal cbData As Long) As Long

Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
    ByVal cbData As Long) As Long
    
#End If


Public Function SetValueEx(ByVal hKey As Long, sValueName As String, _
lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, _
lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, _
lType, lValue, 4)
        End Select
End Function
 
Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
String, vValue As Variant) As Long
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
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
                    sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
                    lValue, cch)
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

Public Function QueryValue(ByVal hKey As Long, sKeyName As String, sValueName As String) As String
    Dim lRetVal As Long         'result of the API functions
    Dim vValue As Variant      'setting of queried value
    Dim hOpenKey As Long
 
    lRetVal = RegOpenKeyEx(hKey, sKeyName, 0, KEY_QUERY_VALUE, hOpenKey)
    lRetVal = QueryValueEx(hOpenKey, sValueName, vValue)
    RegCloseKey (hOpenKey)
    QueryValue = IIf(lRetVal = ERROR_NONE, vValue, "")
    
End Function

Public Sub SetKeyValue(ByVal hKey As Long, sKeyName As String, _
                sValueName As String, vValueSetting As Variant, lValueType As Long)
    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hOpenKey As Long         'handle of open key
 
    'open the specified key
    lRetVal = RegOpenKeyEx(hKey, sKeyName, 0, KEY_ALL_ACCESS, hOpenKey)
    If lRetVal = ERROR_BADKEY Then
      ' create the key if it is missing, added in v5.15
      Dim disp As Long
      lRetVal = RegCreateKeyEx(hKey, sKeyName, 0, "", _
        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, hOpenKey, disp)
    End If
    lRetVal = SetValueEx(hOpenKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hOpenKey)
End Sub
'
' turn an empty string from a missing key into a false
'
Function QueryRegBool(vKey As String) As Boolean
    Dim bb: bb = QueryValue(HKEY_CURRENT_USER, SO_KEY, vKey)
    QueryRegBool = IIf(bb = "", False, bb) ' make sure we pass back a bool
End Function

Function get_locale()
Dim foo:
foo = Application.International
Dim i: For i = 1 To ActiveWorkbook.Styles.Count
    Debug.Print ActiveWorkbook.Styles(i).name
Next
' get_locale = GetLocale()
'Debug.Print "locale is " & get_locale
End Function

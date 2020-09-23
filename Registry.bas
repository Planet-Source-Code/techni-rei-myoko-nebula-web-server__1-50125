Attribute VB_Name = "Registryhandling"
'My Documents : GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal")
'Icon Size    : GetString(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size", 32)
'Desktop      : GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop")
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Const KEY_ALL_ACCESS = &HF003F
    Const HKEY_DYN_DATA = &H80000006
    Const REG_BINARY = 3
    Const REG_DWORD = 4
    Const REG_DWORD_BIG_ENDIAN = 5
    Const REG_DWORD_LITTLE_ENDIAN = 4
    Const REG_EXPAND_SZ = 2
    Const REG_LINK = 6
    Const REG_MULTI_SZ = 7
    Const REG_NONE = 0
    Const REG_RESOURCE_LIST = 8
    Const REG_SZ = 1

Declare Function SystemParametersInfo Lib "USER32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
    Const Spi_seticons As Integer = 88

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String)
On Error Resume Next
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    On Error GoTo 0
    lResult = RegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Or lValueType = REG_EXPAND_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = RegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = StripTerminator(strBuf)
            End If
        End If
    End If
End Function

Public Function GetString(hKey As Long, strpath As String, Optional strvalue As String, Optional default As String = Empty)
On Error Resume Next
    Dim keyhand&, temp As String
    Dim datatype&
    R = RegOpenKey(hKey, strpath, keyhand&)
    temp = RegQueryStringValue(keyhand&, strvalue)
    If temp = Empty Then GetString = default Else GetString = temp
    R = RegCloseKey(keyhand&)
End Function

Function StripTerminator(ByVal strString As String) As String
On Error Resume Next
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Sub SaveString(hKey As Long, strpath As String, strvalue As String, strdata As String)
On Error Resume Next
    Dim keyhand&
    R = RegCreateKey(hKey, strpath, keyhand&)
    R = RegSetValueEx(keyhand&, strvalue, 0, REG_SZ, ByVal strdata, Len(strdata))
    R = RegCloseKey(keyhand&)
End Sub

Public Sub Delstring(hKey As Long, strpath As String, sKey As String)
On Error Resume Next
    Dim keyhand&
    R = RegOpenKey(hKey, strpath, keyhand&)
    R = RegDeleteValue(keyhand&, sKey)
    R = RegCloseKey(keyhand&)
End Sub
Public Function filetype(ByVal Filename As String) As String
On Error Resume Next
    If InStr(Filename, "\") > 0 Then Filename = Right(Filename, Len(Filename) - InStrRev(Filename, "\"))
    Dim stats(0 To 5) As String
    stats(0) = Left(Filename, InStrRev(Filename, "\") - 1)  'Path derivative
    stats(2) = Right(Filename, Len(Filename) - InStrRev(Filename, ".")) 'Extention derivative
    If InStrRev(Filename, ".") = 0 Then stats(2) = UCase(stats(2))
    stats(4) = GetString(HKEY_CLASSES_ROOT, "." & stats(2)) 'File extention name
    stats(5) = GetString(HKEY_CLASSES_ROOT, stats(4), , UCase(stats(2)) & " File") 'File type
    filetype = stats(5)
End Function
Public Sub associate(extention As String, Optional path As String = Empty, Optional fileclassname As String, Optional filetypename As String = Empty, Optional defaulticonfile As String = Empty, Optional defaulticonindex As Long = 0)
    If Left(extention, 1) <> "." Then extention = "." & extention
    Dim tempstr As String
    tempstr = GetString(HKEY_CLASSES_ROOT, extention, , fileclassname)
    
    If filetypename <> Empty Then SaveString HKEY_CLASSES_ROOT, tempstr, Empty, filetypename
    If defaulticonfile <> Empty Then SaveString HKEY_CLASSES_ROOT, tempstr & "\DefaultIcon", Empty, defaulticonfile & ",-" & Abs(defaulticonindex)
    If path <> Empty Then SaveString HKEY_CLASSES_ROOT, tempstr & "\shell\open\command", Empty, path
    
    SystemParametersInfo Spi_seticons, IIf(Active, 1, 0), "0", 0
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Regidit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteTree Lib "advapi32.dll" Alias "RegDeleteTreeA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Byte, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long


 Const REG_SZ As Long = 1
 Const REG_BINARY As Long = 3
 Const REG_DWORD As Long = 4
 Const REG_NONE As Long = 0
 Const REG_QWORD As Long = 11
 Const REG_MULTI_SZ As Long = 7
 Const REG_EXPAND_SZ As Long = 2
 Const HKEY_LOCAL_MACHINE As Long = &H80000002

 Const HKEY_CURRENT_USER As Long = &H80000001
 Const STANDARD_RIGHTS_ALL As Long = &H1F0000
 Const KEY_QUERY_VALUE As Long = &H1
 Const KEY_SET_VALUE As Long = &H2
 Const KEY_CREATE_SUB_KEY As Long = &H4
 Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
 Const KEY_NOTIFY As Long = &H10
 Const KEY_CREATE_LINK As Long = &H20
 Const SYNCHRONIZE As Long = &H100000
 Const KEY_ALL_ACCESS As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Function GetValue(ByVal hkey As Long, ByVal ValueName As String) As Variant
    Dim funResult As Variant
    Dim strBuf As String
    Dim strSize As Long
    strBuf = String(255, Chr(0))
    
    
    RegQueryValueEx hkey, ValueName, 0, REG_SZ, ByVal strBuf, Len(strBuf)
    funResult = Left(strBuf, InStr(strBuf, Chr(0)) - 1)
    GetValue = funResult
'    CloseKey hkey


End Function



'功能: 创建一个键
'返回：TRUE成功  False 失败
'
'hKey 传入一个已经打开的句柄
'lpSubKey 一个要创建的名称


Public Function CreateKey(ByVal KeyPath As String, ByVal lpSubKey As String) As Boolean
On Error GoTo l
Dim funResult As Boolean
Dim ret As Long
Dim mkey As Long
Select Case Left(KeyPath, InStr(KeyPath, "\") - 1)
Case "HKEY_CURRENT_USER"
    mkey = HKEY_CURRENT_USER
Case "HKEY_LOCAL_MACHINE"
    mkey = HKEY_LOCAL_MACHINE
Case Else
    CreateKey = False
    Exit Function
    
End Select

KeyPath = Mid(KeyPath, InStr(KeyPath, "\") + 1)
RegCreateKey mkey, KeyPath & "\" & lpSubKey, ret
If ret <> 0 Then funResult = True
CloseKey ret

CreateKey = funResult
Exit Function
l:
CreateKey = False

End Function

'功能: 打开一个键
'返回: 键句柄

Public Function OpenKey(ByVal hkey As Long, ByVal lpSubKey As String) As Long
    Dim funResult As Long
    RegOpenKeyEx hkey, lpSubKey, 0, KEY_ALL_ACCESS, RetKey
    funResult = RetKey
    OpenKey = funResult
End Function


'功能打开一个键，传入，这个键的完整路径

Public Function OpenKeyPath(ByVal KeyPath As String) As Long
Dim funResult As Long
Dim hkey As Long
Dim keyName As String
Dim ret As Long

Select Case Left(KeyPath, InStr(KeyPath, "\") - 1)
Case "HKEY_CURRENT_USER"
    hkey = HKEY_CURRENT_USER
Case "HKEY_LOCAL_MACHINE"
hkey = HKEY_LOCAL_MACHINE

End Select
KeyPath = Mid(KeyPath, InStr(KeyPath, "\") + 1)
RegOpenKeyEx hkey, KeyPath, 0, KEY_ALL_ACCESS, funResult
OpenKeyPath = funResult
End Function

'功能: 枚举键
'返回: PropertyBag（Count键表示共有多少数据），获取值为索引号，从0开始
'hKey:  传入一个已经打开句柄
Public Function EnumKey(ByVal hkey As Long) As PropertyBag

Dim funResult As New PropertyBag
Dim ret As Long
Dim keyName As String
Dim idx As Long
keyName = String(255, Chr(0))

Do While (ret = RegEnumKey(hkey, idx, keyName, Len(keyName))) <> 0
    funResult.WriteProperty CStr(idx), Left(keyName, InStr(keyName, Chr(0)))
    idx = idx + 1
    keyName = String(255, Chr(0))
Loop

Call funResult.WriteProperty("Count", idx)
Set EnumKey = funResult
Set funResult = Nothing

End Function

'功能：删除键，传入一个完整的路径
Public Function DeleteKey(ByVal KeyPath As String) As Boolean
Dim hkey As Long
Dim keyName As String
Dim ret As Long

Select Case Left(KeyPath, InStr(KeyPath, "\") - 1)
Case "HKEY_CURRENT_USER"
    hkey = HKEY_CURRENT_USER
Case "HKEY_LOCAL_MACHINE"
    hkey = HKEY_LOCAL_MACHINE
    
End Select
KeyPath = Mid(KeyPath, InStr(KeyPath, "\") + 1)

If (InStr(KeyPath, "\")) = 0 Then
    ret = RegDeleteTree(hkey, KeyPath)
    If ret = 0 Then
        DeleteKey = True
    End If
    Exit Function
End If

keyName = Mid(KeyPath, InStrRev(KeyPath, "\") + 1)
KeyPath = Left(KeyPath, Len(KeyPath) - Len(keyName) - 1)
ret = RegDeleteTree(OpenKey(hkey, KeyPath), keyName)
If ret = 0 Then DeleteKey = True

End Function

'功能: 设置REG_SZ键值
'KeyPath 传入一个完整路径
'lpValueName 值名称
'lpData 值
Public Function SetKeyValueREG_SZ(ByVal KeyPath As String, ByVal lpValueName As String, ByVal lpData As String) As Long
Dim ret As Long
Dim hkey As Long

hkey = OpenKeyPath(KeyPath)
ret = RegSetValueEx(hkey, lpValueName, 0, REG_SZ, ByVal lpData, LenB(lpData))
SetKeyValueREG_SZ = ret
CloseKey hkey

End Function

Public Function SetKeyValueREG_SZ2(ByVal hkey As Long, ByVal lpValueName As String, ByVal lpData As String) As Long
ret = RegSetValueEx(hkey, lpValueName, 0, REG_SZ, ByVal lpData, LenB(lpData))
SetKeyValueREG_SZ2 = ret

End Function



'功能: 设置REG_BINARY键值
'KeyPath 传入一个完整路径
'lpValueName 值名称
'lpData 值
Public Function SetKeyValueREG_BINARY(ByVal KeyPath As String, ByVal lpValueName As String, ByVal lpData As Long) As Long
Dim ret As Long
Dim hkey As Long

hkey = OpenKeyPath(KeyPath)
ret = RegSetValueEx(hkey, lpValueName, 0, REG_BINARY, ByVal lpData, 4)
SetKeyValueREG_BINARY = ret
CloseKey hkey

End Function



'功能: 设置REG_DWORD键值
'KeyPath 传入一个完整路径
'lpValueName 值名称
'lpData 值
Public Function SetKeyValueREG_DWORD(ByVal KeyPath As String, ByVal lpValueName As String, ByVal lpData As Long) As Long
Dim ret As Long
Dim hkey As Long

hkey = OpenKeyPath(KeyPath)
ret = RegSetValueEx(hkey, lpValueName, 0, REG_DWORD, lpData, 4)
SetKeyValueREG_DWORD = ret
CloseKey hkey

End Function

'功能: 设置REG_QWORD键值
'KeyPath 传入一个完整路径
'lpValueName 值名称

'lpData 值
Public Function SetKeyValueREG_QWORD(ByVal KeyPath As String, ByVal lpValueName As String, ByVal lpData As Long) As Long
Dim ret As Long
Dim hkey As Long

hkey = OpenKeyPath(KeyPath)
ret = RegSetValueEx(hkey, lpValueName, 0, REG_QWORD, lpData, 8)
SetKeyValueREG_QWORD = ret
CloseKey hkey

End Function

'功能: 设置REG_MULTI_SZ键值
'KeyPath 传入一个完整路径
'lpValueName 值名称

'lpData 值
Public Function SetKeyValueREG_MULTI_SZ(ByVal KeyPath As String, ByVal lpValueName As String, ByVal lpData As String) As Long
Dim ret As Long
Dim hkey As Long

hkey = OpenKeyPath(KeyPath)
ret = RegSetValueEx(hkey, lpValueName, 0, REG_MULTI_SZ, ByVal lpData, Len(lpData))
SetKeyValueREG_MULTI_SZ = ret
CloseKey hkey

End Function

'功能: 设置REG_EXPAND_SZ键值
'KeyPath 传入一个完整路径
'lpValueName 值名称

'lpData 值
Public Function SetKeyValueREG_REG_EXPAND_SZ(ByVal KeyPath As String, ByVal lpValueName As String, ByVal lpData As String) As Long
Dim ret As Long
Dim hkey As Long

hkey = OpenKeyPath(KeyPath)
ret = RegSetValueEx(hkey, lpValueName, 0, REG_EXPAND_SZ, ByVal lpData, Len(lpData))
SetKeyValueREG_REG_EXPAND_SZ = ret
CloseKey hkey

End Function

Public Function DeleteValue(ByVal KeyPath As String, ByVal VlaueName As String) As Boolean
Dim funResult As Boolean
Dim ret As Long
Dim hkey As Long
hkey = OpenKeyPath(KeyPath)

ret = RegDeleteValue(hkey, VlaueName)
If ret = 0 Then DeleteValue = True
CloseKey hkey


End Function


'功能: 枚举值
'返回PropertyBag
'取出  .ReadProperty(索引)
Public Function EnumValue(ByVal KeyPath As String) As PropertyBag
Dim funResult As New PropertyBag
Dim hkey As Long
Dim szBuffer As String, sValue(599) As Byte, dwIndex As Long, ret As Long
    dwIndex = 0
     szBuffer = Space(300)
     hkey = OpenKeyPath(KeyPath)
     
    Do While ret = RegEnumValue(hkey, dwIndex, szBuffer, 255, ByVal 0&, REG_SZ, sValue(0), 600) <> 0
        If InStr(szBuffer, vbNullChar) > 0 Then szBuffer = Left(szBuffer, InStr(szBuffer, vbNullChar) - 1)
        funResult.WriteProperty dwIndex, szBuffer
        dwIndex = dwIndex + 1
        szBuffer = Space(300)
    Loop
    
    funResult.WriteProperty "Count", dwIndex
    Set EnumValue = funResult
    CloseKey hkey
    
End Function
Public Function CloseKey(ByVal hkey As Long)
RegCloseKey hkey
End Function



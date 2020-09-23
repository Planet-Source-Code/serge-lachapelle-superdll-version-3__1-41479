Attribute VB_Name = "modRegistry"
Option Explicit

Public Enum RegKey   ' lPredefinedKey , hMainKey
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

Private Const ERROR_SUCCESS = 0&

Private Const REG_SZ = 1&
Private Const REG_BINARY = 3&
Private Const REG_DWORD = 4&

Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const KEY_CREATE_LINK = &H20&
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = &H20000
Private Const STANDARD_RIGHTS_WRITE = &H20000
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Private Const STANDARD_RIGHTS_ALL = &H1F0000

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lplDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

Private hKey As Long, MainKeyHandle As Long
Private rtn As Long


Public Function SetDWordValue(ByRef sKeyPtr As Long, ByRef sKeyNamePtr As Long, ByVal KeyValue As Long) As Boolean

Dim sKey As String, sKeyName As String
SetDWordValue = False
sKey = CSTOVBS(sKeyPtr)
sKeyName = CSTOVBS(sKeyNamePtr)
ParseKey sKey, MainKeyHandle

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = RegSetValueExA(hKey, sKeyName, 0, REG_DWORD, KeyValue, 4) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
         MsgBox GetErrorMsg(rtn), vbExclamation, "Error"        'display the error
      Else
         SetDWordValue = True
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      MsgBox GetErrorMsg(rtn), vbExclamation, "Error" 'display the error
   End If
End If

End Function


Public Function GetDWordValue(ByRef sKeyPtr As Long, ByRef sKeyNamePtr As Long) As Variant

Dim sKey As String, sKeyName As String, lBuffer As Long
sKey = CSTOVBS(sKeyPtr)
sKeyName = CSTOVBS(sKeyNamePtr)
ParseKey sKey, MainKeyHandle

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = RegQueryValueExA(hKey, sKeyName, 0, REG_DWORD, lBuffer, 4) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         GetDWordValue = lBuffer  'return the value
      Else                        'otherwise, if the value couldnt be retreived
         GetDWordValue = "Error"  'return Error to the user
         MsgBox GetErrorMsg(rtn), vbExclamation, "Error"        'tell the user what was wrong
      End If
   Else 'otherwise, if the key couldnt be opened
      GetDWordValue = "Error"        'return Error to the user
      MsgBox GetErrorMsg(rtn), vbExclamation, "Error"        'tell the user what was wrong
   End If
Else
   GetDWordValue = "Error" 'return Error to the user
End If

End Function


Public Function SetBinaryValue(ByRef sKeyPtr As Long, ByRef sKeyNamePtr As Long, ByRef KeyValuePtr As Long) As Boolean

Dim lDataSize As Long, i As Long, ByteArray() As Byte
Dim sKey As String, sKeyName As String, KeyValue As String
SetBinaryValue = False
sKey = CSTOVBS(sKeyPtr)
sKeyName = CSTOVBS(sKeyNamePtr)
KeyValue = CSTOVBS(KeyValuePtr)
ParseKey sKey, MainKeyHandle

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      lDataSize = Len(KeyValue)
      ReDim ByteArray(lDataSize)
      For i = 1 To lDataSize
      ByteArray(i) = Asc(Mid$(KeyValue, i, 1))
      Next
      rtn = RegSetValueExB(hKey, sKeyName, 0, REG_BINARY, ByteArray(1), lDataSize) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if the was an error writting the value
         MsgBox GetErrorMsg(rtn), vbExclamation, "Error"        'display the error
      Else
         SetBinaryValue = True
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      MsgBox GetErrorMsg(rtn), vbExclamation, "Error" 'display the error
   End If
End If

End Function


Public Function GetBinaryValue(ByRef sKeyPtr As Long, ByRef sKeyNamePtr As Long) As Variant

Dim sBuffer As String, sKey As String, sKeyName As String, lBufferSize As Long
sKey = CSTOVBS(sKeyPtr)
sKeyName = CSTOVBS(sKeyNamePtr)
ParseKey sKey, MainKeyHandle

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened
      lBufferSize = 1
      rtn = RegQueryValueEx(hKey, sKeyName, 0, REG_BINARY, 0, lBufferSize) 'get the value from the registry
      sBuffer = Space(lBufferSize)
      rtn = RegQueryValueEx(hKey, sKeyName, 0, REG_BINARY, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         GetBinaryValue = sBuffer 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         GetBinaryValue = "Error" 'return Error to the user
         MsgBox GetErrorMsg(rtn), vbExclamation, "Error"  'display the error to the user
      End If
   Else 'otherwise, if the key couldnt be opened
      GetBinaryValue = "Error" 'return Error to the user
      MsgBox GetErrorMsg(rtn), vbExclamation, "Error"  'display the error to the user
   End If
Else
   GetBinaryValue = "Error" 'return Error to the user
End If

End Function


Public Function SetStringValue(ByRef sKeyPtr As Long, ByRef sKeyNamePtr As Long, ByRef KeyValuePtr As Long) As Boolean

Dim sKey As String, sKeyName As String, KeyValue As String
SetStringValue = False
sKey = CSTOVBS(sKeyPtr)
sKeyName = CSTOVBS(sKeyNamePtr)
KeyValue = CSTOVBS(KeyValuePtr)
ParseKey sKey, MainKeyHandle

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = RegSetValueEx(hKey, sKeyName, 0, REG_SZ, ByVal KeyValue, Len(KeyValue)) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
         MsgBox GetErrorMsg(rtn), vbExclamation, "Error"        'display the error
      Else
         SetStringValue = True
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      MsgBox GetErrorMsg(rtn), vbExclamation, "Error"        'display the error
   End If
End If

End Function


Public Function GetStringValue(ByRef sKeyPtr As Long, ByRef sKeyNamePtr As Long) As Variant

Dim sBuffer As String, sKey As String, sKeyName As String, lBufferSize As Long
sKey = CSTOVBS(sKeyPtr)
sKeyName = CSTOVBS(sKeyNamePtr)
lBufferSize = 0
sBuffer = ""
ParseKey sKey, MainKeyHandle

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      sBuffer = Space(255)     'make a buffer
      lBufferSize = Len(sBuffer)
      rtn = RegQueryValueEx(hKey, sKeyName, 0, REG_SZ, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         sBuffer = Trim(sBuffer)
         GetStringValue = Left(sBuffer, lBufferSize - 1) 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         GetStringValue = "Error" 'return Error to the user
         MsgBox GetErrorMsg(rtn), vbExclamation, "Error"  'tell the user what was wrong
      End If
   Else 'otherwise, if the key couldnt be opened
      GetStringValue = "Error"       'return Error to the user
      MsgBox GetErrorMsg(rtn), vbExclamation, "Error"        'tell the user what was wrong
   End If
Else
   GetStringValue = "Error"       'return Error to the user
End If

End Function


Public Function CreateKey(ByRef sKeyPtr As Long) As Boolean

    CreateKey = False
    Dim sKey As String
    sKey = CSTOVBS(sKeyPtr)
    
    ParseKey sKey, MainKeyHandle
        
    If MainKeyHandle Then
       rtn = RegCreateKey(MainKeyHandle, sKey, hKey) 'create the key
       If rtn = ERROR_SUCCESS Then 'if the key was created then
          rtn = RegCloseKey(hKey)  'close the key
          CreateKey = True
       Else
          MsgBox GetErrorMsg(rtn), vbExclamation, "Error"
       End If
    End If
    
End Function


Public Function DeleteKey(ByRef KeyNamePtr As Long, Optional ByVal Quiet As Boolean = False) As Boolean

  Dim var1 As String
  Dim KeyName As String
  KeyName = CSTOVBS(KeyNamePtr)
  var1 = KeyName
  DeleteKey = False
  ParseKey KeyName, MainKeyHandle
  
  If MainKeyHandle Then
    If KeyExist2(var1) Then
        rtn = RegDeleteKey(MainKeyHandle, KeyName)
        If (rtn <> ERROR_SUCCESS) Then
             MsgBox GetErrorMsg(rtn), vbExclamation, "Error"    'tell the user what was wrong
        Else
            DeleteKey = True
        End If
    Else
      If Not Quiet Then MsgBox "Key Do Not Exist !", vbExclamation, "Error"
    End If
  End If
  
End Function


Public Function DeleteKeyValue(ByRef sKeyNamePtr As Long, ByRef sValueNamePtr As Long, Optional ByVal Quiet As Boolean = False) As Boolean

  Dim var1 As String, var2 As String
  Dim sKeyName As String, sValueName As String
  sKeyName = CSTOVBS(sKeyNamePtr)
  sValueName = CSTOVBS(sValueNamePtr)
  var1 = sKeyName: var2 = sValueName
  DeleteKeyValue = False
  Dim hKey As Long         'handle of open key
  ParseKey sKeyName, MainKeyHandle
  
  If MainKeyHandle Then
    If KeyValueExist2(var1, var2) Then
        rtn = RegOpenKeyEx(MainKeyHandle, sKeyName, 0, KEY_WRITE, hKey)   'open the specified key
        If (rtn = ERROR_SUCCESS) Then
            rtn = RegDeleteValue(hKey, sValueName)
            If (rtn <> ERROR_SUCCESS) Then
                 MsgBox GetErrorMsg(rtn), vbExclamation, "Error"    'tell the user what was wrong
            Else
                DeleteKeyValue = True
            End If
            rtn = RegCloseKey(hKey)
        End If
    Else
      If Not Quiet Then MsgBox "Key Value Do Not Exist !", vbExclamation, "Error"
    End If
  End If
  
End Function


Public Function KeyExist(ByRef sKeyPtr As Long) As Boolean
    Dim hKey As Long
    Dim sKey As String
    sKey = CSTOVBS(sKeyPtr)
    ParseKey sKey, MainKeyHandle, False

    If MainKeyHandle Then
        rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey) 'open the key
        If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
            KeyExist = True
        Else
            KeyExist = False
        End If
    Else
        KeyExist = False
    End If
    
End Function


Public Function KeyValueExist(ByRef sKeyPtr As Long, ByRef sKeyNamePtr As Long) As Boolean
    Dim hKey As Long
    Dim lActualType As Long
    Dim lSize As Long
    Dim sTmp As String
    Dim sKey As String, sKeyName As String
    sKey = CSTOVBS(sKeyPtr): sKeyName = CSTOVBS(sKeyNamePtr)
    ParseKey sKey, MainKeyHandle, False

    If MainKeyHandle Then
        
        rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey) 'open the key
        If (rtn = ERROR_SUCCESS) Then
            rtn = RegQueryValueEx(hKey, ByVal sKeyName, 0&, lActualType, sTmp, lSize) 'ByVal 0&, lSize)
            If (rtn = ERROR_SUCCESS) Then
                KeyValueExist = True
            Else
                KeyValueExist = False
            End If
        End If
    Else
        KeyValueExist = False
    End If

End Function


Private Function KeyExist2(ByVal sKey As String) As Boolean
    Dim hKey As Long

    ParseKey sKey, MainKeyHandle, False

    If MainKeyHandle Then
        rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey) 'open the key
        If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
            KeyExist2 = True
        Else
            KeyExist2 = False
        End If
    Else
        KeyExist2 = False
    End If
    
End Function


Private Function KeyValueExist2(ByVal sKey As String, ByVal sKeyName As String) As Boolean
    Dim hKey As Long
    Dim lActualType As Long
    Dim lSize As Long
    Dim sTmp As String
    ParseKey sKey, MainKeyHandle, False

    If MainKeyHandle Then
        
        rtn = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey) 'open the key
        If (rtn = ERROR_SUCCESS) Then
            rtn = RegQueryValueEx(hKey, ByVal sKeyName, 0&, lActualType, sTmp, lSize) 'ByVal 0&, lSize)
            If (rtn = ERROR_SUCCESS) Then
                KeyValueExist2 = True
            Else
                KeyValueExist2 = False
            End If
        End If
    Else
        KeyValueExist2 = False
    End If

End Function


Private Sub ParseKey(KeyName As Variant, Keyhandle As Long, Optional ByVal vBox As Boolean = True)
    
Keyhandle = 0
rtn = InStr(KeyName, "\") 'return if "\" is contained in the Keyname

If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   If vBox = True Then MsgBox "Bad Key Name", vbExclamation, "Error"
   Exit Sub 'exit the procedure
ElseIf rtn = 0 Then 'if the Keyname contains no "\"
   Keyhandle = GetMainKeyHandle(KeyName)
   If Keyhandle = 0 Then
     If vBox = True Then MsgBox "Bad Key Name", vbExclamation, "Error"
     Exit Sub
   End If
   KeyName = "" 'leave Keyname blank
Else 'otherwise, Keyname contains "\"
   Keyhandle = GetMainKeyHandle(Left(KeyName, rtn - 1)) 'seperate the Keyname
   If Keyhandle = 0 Then
     If vBox = True Then MsgBox "Bad Key Name", vbExclamation, "Error"
     Exit Sub
   End If
   KeyName = Right(KeyName, Len(KeyName) - rtn)
End If

End Sub


Private Function GetMainKeyHandle(MainKeyName As Variant) As Long
  
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
       Case Else
            GetMainKeyHandle = 0
End Select

End Function

Private Function GetErrorMsg(ByVal lErrorCode As Long) As String
    
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
       Case 1, 1009, 1015
            GetErrorMsg = "The Registry Database is corrupt!"
       Case 2, 6, 1010
            GetErrorMsg = "Bad Key Name"
       Case 3, 1011
            GetErrorMsg = "Can't Open Key"
       Case 4, 1012
            GetErrorMsg = "Can't Read Key"
       Case 1013
            GetErrorMsg = "Can't Write Key"
       Case 5
            GetErrorMsg = "Access to this key is denied"
       Case 8, 14
            GetErrorMsg = "Out of memory"
       Case 7, 87
            GetErrorMsg = "Invalid Parameter"
       Case 234
            GetErrorMsg = "There is more data than the buffer has been allocated to hold."
       Case 259
            GetErrorMsg = "No More Items"
       Case Else
            GetErrorMsg = "Undefined Error Code: " & Str$(lErrorCode)
End Select

End Function


Public Function BinToHexR(ByRef var0 As Long) As Variant
  Dim t As Long
  Dim qaz As String
  Dim qwe As String
  qwe = CSTOVBS(var0)
  
  For t = 1 To Len(qwe)
    qaz = qaz + IIf(Len(Hex(Asc(Mid$(qwe, t, 1)))) > 1, Hex(Asc(Mid$(qwe, t, 1))), "0" & Hex(Asc(Mid$(qwe, t, 1))))
    If t <> Len(qwe) Then qaz = qaz + " "
  Next t
  
  BinToHexR = qaz
  
End Function


Public Function BinToDecR(ByRef var0 As Long) As Variant
  Dim t As Long
  Dim qaz As String
  Dim qwe As String
  qwe = CSTOVBS(var0)
  
  For t = 1 To Len(qwe)
    qaz = qaz + Right$(Str(Asc(Mid$(qwe, t, 1))), Len(Str(Asc(Mid$(qwe, t, 1)))) - 1)
    If t <> Len(qwe) Then qaz = qaz + " "
  Next t
  
  BinToDecR = qaz
  
End Function


Public Function BinToDecA(ByRef var0 As Long) As Variant()
  Dim t As Long
  Dim qaz() As Variant
  Dim qwe As String
  qwe = CSTOVBS(var0)
  ReDim qaz(Len(qwe))
  
  For t = 1 To Len(qwe)
    qaz(t) = Right$(Str(Asc(Mid$(qwe, t, 1))), Len(Str(Asc(Mid$(qwe, t, 1)))) - 1)
  Next t
  
  BinToDecA = qaz
  
End Function


Public Function BinToHexA(ByRef var0 As Long) As Variant()
  Dim t As Long
  Dim qaz() As Variant
  Dim qwe As String
  qwe = CSTOVBS(var0)
  ReDim qaz(Len(qwe))
  
  For t = 1 To Len(qwe)
    qaz(t) = IIf(Len(Hex(Asc(Mid$(qwe, t, 1)))) > 1, Hex(Asc(Mid$(qwe, t, 1))), "0" & Hex(Asc(Mid$(qwe, t, 1))))
  Next t
  
  BinToHexA = qaz
  
End Function


Public Function EnumKey(ByVal hMainKey As RegKey, ByRef sSubKeyPtr As Long, ByVal lIndex As Long, lpStr As Variant) As Boolean
  Dim hKey As Long
  Dim i As Long
  Dim lpStr2 As String
  Dim t As Integer
  Dim sSubKey As String
  sSubKey = CSTOVBS(sSubKeyPtr)
  
  rtn = RegOpenKey(hMainKey, sSubKey, hKey)
  If rtn = ERROR_SUCCESS Then
    lpStr2 = Space(255) + Chr(0)
    rtn = RegEnumKey(hKey, lIndex, lpStr2, Len(lpStr2))
    If rtn = ERROR_SUCCESS Then
      t = 255
      While Mid$(lpStr2, t, 1) = " "
        t = t - 1
      Wend
      lpStr = Left$(lpStr2, t - 1)
      EnumKey = True
    Else
      EnumKey = False
      If rtn <> 259 Then MsgBox GetErrorMsg(rtn), vbExclamation, "Error"
    End If
  Else
    EnumKey = False
    If rtn <> 259 Then MsgBox GetErrorMsg(rtn), vbExclamation, "Error"
  End If
  RegCloseKey hKey
  
End Function


Public Function QueryValue(ByVal lPredefinedKey As RegKey, ByRef sKeyNamePtr As Long, ByRef sValueNamePtr As Long) As Variant
  Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
  Dim lRetVal As Long
  Dim hKey As Long
  Dim vValue As Variant
  Dim sKeyName As String, sValueName As String
  sKeyName = CSTOVBS(sKeyNamePtr)
  sValueName = CSTOVBS(sValueNamePtr)
  lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
  lRetVal = QueryValueEx(hKey, sValueName, vValue)
  QueryValue = vValue
  RegCloseKey (hKey)
End Function


Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
  On Local Error GoTo QueryValueExExit
  Dim cch As Long
  Dim lType As Long
  Dim lValue As Long
  Dim sValue As String
  
  rtn = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
  If rtn <> ERROR_SUCCESS Then MsgBox GetErrorMsg2(rtn), vbExclamation, "Error"
  
  Select Case lType
    Case REG_SZ:
        sValue = String(cch, 0)
        rtn = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
        If rtn = ERROR_SUCCESS Then
          vValue = Left$(sValue, cch)
        Else
          vValue = "ERROR"
        End If
    Case REG_DWORD:
        rtn = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
        If rtn = ERROR_SUCCESS Then
          vValue = lValue
        Else
          vValue = "ERROR"
        End If
    Case Else
        rtn = -1
        vValue = "ERROR"
  End Select

QueryValueExExit:
  QueryValueEx = rtn
End Function


Private Function GetErrorMsg2(ByVal lErrorCode As Long) As String
    
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
       Case 1, 1009, 1015
            GetErrorMsg2 = "The Registry Database is corrupt!"
       Case 2, 6, 1010
            GetErrorMsg2 = "Bad Key Name"
       Case 3, 1011
            GetErrorMsg2 = "Can't Open Key"
       Case 4, 1012
            GetErrorMsg2 = "Can't Read Key"
       Case 5, 1013
            GetErrorMsg2 = "Can't Write Key"
       Case 8
            GetErrorMsg2 = "Access to this key is denied"
       Case 14
            GetErrorMsg2 = "Out of memory"
       Case 7, 87
            GetErrorMsg2 = "Invalid Parameter"
       Case 234
            GetErrorMsg2 = "There is more data than the buffer has been allocated to hold."
       Case 259
            GetErrorMsg2 = "No More Items"
       Case Else
            GetErrorMsg2 = "Undefined Error Code: " & Str$(lErrorCode)
End Select

End Function

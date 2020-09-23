Attribute VB_Name = "modWindows"
Option Explicit

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(0 To 0) As LUID_AND_ATTRIBUTES
End Type

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As PlatformType
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Type WindowsVersionInfo
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As PlatformType
        szCSDVersion As Variant
        dwFullVersion As Variant
        dwTextVersion As Variant
        dwFullTextV As Variant
End Type

Public Enum ShutDownType
  EWX_LOGOFF = &H0
  EWX_SHUTDOWN = &H1
  EWX_REBOOT = &H2
  EWX_POWEROFF = &H8     ' SHUTDOWN is better
End Enum

Public Enum ForceType
  EWX_NORMAL = &H0
  EWX_FORCEIFHUNG = &H10
  EWX_FORCE = &H4        ' better not use !
End Enum

Public Enum PlatformType   ' dwPlatformId
  VER_PLATFORM_WIN32s = 0        ' Unknown Version
  VER_PLATFORM_WIN32_WINDOWS = 1 ' Windows 3.1/95/98/Me
  VER_PLATFORM_WIN32_NT = 2      ' Windows NT/2000/XP/.NET
End Enum

Private Const NoShutDownPrivilege As String = "No ShutDown Privilege !"

Private Declare Function GetVersionEx Lib "Kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function LockWorkStation Lib "user32.dll" () As Long
Private Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "Kernel32.dll" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function GetLastError Lib "Kernel32.dll" () As Long
Private Declare Function FormatMessage Lib "Kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Function isNT2000XP() As Boolean
  Dim lpv As OSVERSIONINFO
  lpv.dwOSVersionInfoSize = Len(lpv)
  GetVersionEx lpv
  If lpv.dwPlatformId = VER_PLATFORM_WIN32_NT Then
    isNT2000XP = True
  Else
    isNT2000XP = False
  End If
End Function

Private Function ShutDownPrivilege() As Boolean
  Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
  Const TOKEN_QUERY As Long = &H8
  Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"
  Const SE_PRIVILEGE_ENABLED As Long = &H2
  Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
  Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
  Const Language_Neutral As Long = &H0
  Const User_Default_Language As Long = &H400
  Const System_Default_Language As Long = &H800
  Dim ErrorNumber As Long
  Dim ErrorMessage As String
  Dim hToken As Long
  Dim tkp As TOKEN_PRIVILEGES
  Dim tkpNULL As TOKEN_PRIVILEGES
  If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then
    ShutDownPrivilege = False
    Exit Function
  End If
  LookupPrivilegeValue vbNullString, SE_SHUTDOWN_NAME, tkp.Privileges(0).pLuid
  tkp.PrivilegeCount = 1
  tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
  AdjustTokenPrivileges hToken, False, tkp, Len(tkp), tkpNULL, Len(tkpNULL)
  ErrorNumber = GetLastError
  If ErrorNumber <> 0 Then
    ErrorMessage = Space$(500)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, ErrorNumber, User_Default_Language, ErrorMessage, Len(ErrorMessage), 0
    MsgBox Trim3(ErrorMessage), vbExclamation, "SuperDLL"
    ShutDownPrivilege = False
    Exit Function
  End If
  ShutDownPrivilege = True
End Function

Public Sub SHUTDOWN(Optional ByVal FT As ForceType = EWX_FORCEIFHUNG, Optional ByVal SDT As ShutDownType = EWX_SHUTDOWN)
  Dim var1 As Long
  If isNT2000XP Then
    If ShutDownPrivilege Then
      var1 = SDT Or FT
      ExitWindowsEx var1, 0
    Else
      MsgBox NoShutDownPrivilege, vbExclamation, "SuperDLL"
    End If
  Else
    var1 = SDT Or FT
    ExitWindowsEx var1, 0
  End If
End Sub

Public Sub LOGOFF(Optional ByVal FT As ForceType = EWX_FORCEIFHUNG)
  Dim var1 As Long
  If isNT2000XP Then
    If ShutDownPrivilege Then
      var1 = EWX_LOGOFF Or FT
      ExitWindowsEx var1, 0
    Else
      MsgBox NoShutDownPrivilege, vbExclamation, "SuperDLL"
    End If
  Else
    var1 = EWX_LOGOFF Or FT
    ExitWindowsEx var1, 0
  End If
End Sub

Public Sub REBOOT(Optional ByVal FT As ForceType = EWX_FORCEIFHUNG)
  Dim var1 As Long
  If isNT2000XP Then
    If ShutDownPrivilege Then
      var1 = EWX_REBOOT Or FT
      ExitWindowsEx var1, 0
    Else
      MsgBox NoShutDownPrivilege, vbExclamation, "SuperDLL"
    End If
  Else
    var1 = EWX_REBOOT Or FT
    ExitWindowsEx var1, 0
  End If
End Sub

Public Sub POWEROFF(Optional ByVal FT As ForceType = EWX_FORCEIFHUNG)
  Dim var1 As Long
  If isNT2000XP Then
    If ShutDownPrivilege Then
      var1 = EWX_POWEROFF Or FT
      ExitWindowsEx var1, 0
    Else
      MsgBox NoShutDownPrivilege, vbExclamation, "SuperDLL"
    End If
  Else
    var1 = EWX_POWEROFF Or FT
    ExitWindowsEx var1, 0
  End If
End Sub

Public Sub LockComputer()
  LockWorkStation
End Sub

Public Function GetWindowsVersion(ByRef wvi As WindowsVersionInfo) As Long
  Dim lpv As OSVERSIONINFO
  Dim qwe As String
  Dim qaz As String
  Dim t As Byte
  lpv.dwOSVersionInfoSize = Len(lpv)
  GetWindowsVersion = GetVersionEx(lpv)
  qwe = ""
  For t = 1 To 128
    qaz = Mid$(lpv.szCSDVersion, t, 1)
    Select Case qaz
      Case Chr$(0), Chr$(32), Chr$(255):
        qwe = qwe & Chr$(32)
      Case Else:
        qwe = qwe & qaz
    End Select
  Next t
  Select Case lpv.dwPlatformId
    Case VER_PLATFORM_WIN32_NT
      Select Case lpv.dwMajorVersion
        Case 3
          wvi.dwTextVersion = "Windows NT 3.51"
        Case 4
          wvi.dwTextVersion = "Windows NT 4.0"
        Case 5
          Select Case lpv.dwMinorVersion
            Case 0
              wvi.dwTextVersion = "Windows 2000"
            Case 1
              wvi.dwTextVersion = "Windows XP"
            Case 2
              wvi.dwTextVersion = "Windows .NET"
            Case Else
              wvi.dwTextVersion = "Windows 2000/XP/.NET"
          End Select
        Case Else
          wvi.dwTextVersion = "Windows NT/2000/XP/.NET"
      End Select
    Case VER_PLATFORM_WIN32_WINDOWS
      Select Case lpv.dwMajorVersion
        Case 3
          wvi.dwTextVersion = "Windows 3.1"
        Case 4
          Select Case lpv.dwMinorVersion
            Case 0
              Select Case Left$(lpv.szCSDVersion, 1)
                Case "C"
                  wvi.dwTextVersion = "Windows 95 C"
                Case "B"
                  wvi.dwTextVersion = "Windows 95 B"
                Case Else
                  wvi.dwTextVersion = "Windows 95"
              End Select
            Case 10
              Select Case Left$(lpv.szCSDVersion, 1)
                Case "A"
                  wvi.dwTextVersion = "Windows 98 SE"
                Case Else
                  wvi.dwTextVersion = "Windows 98"
              End Select
            Case 90
              wvi.dwTextVersion = "Windows Millennium"
            Case Else
              wvi.dwTextVersion = "Windows 95/98/ME"
          End Select
        Case Else
          wvi.dwTextVersion = "Windows 3.1/95/98/ME"
      End Select
    Case Else
      wvi.dwTextVersion = "Unknown Version"
  End Select
  wvi.dwBuildNumber = lpv.dwBuildNumber
  wvi.dwMajorVersion = lpv.dwMajorVersion
  wvi.dwMinorVersion = lpv.dwMinorVersion
  wvi.dwPlatformId = lpv.dwPlatformId
  wvi.szCSDVersion = Trim3(qwe)
  wvi.dwFullVersion = Right$(Str(lpv.dwMajorVersion), Len(Str(lpv.dwMajorVersion)) - 1) & "." & Right$(Str(lpv.dwMinorVersion), Len(Str(lpv.dwMinorVersion)) - 1)
  wvi.dwFullTextV = wvi.dwTextVersion & "   Version " & wvi.dwFullVersion & "   Build " & wvi.dwBuildNumber & "   " & wvi.szCSDVersion
End Function

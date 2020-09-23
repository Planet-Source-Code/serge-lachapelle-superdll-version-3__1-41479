'
' MsgBox QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
'_______________________________________________________________________________________
'
' Dim var1() As Variant
' a$ = GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "AGTSType")
' If a$ <> "Error" Then
'   var1 = BinToHexA(a$)
'   For t = 1 To Len(a$)
'     Debug.Print var1(t);
'   Next t
'   MsgBox BinToHexR(a$)
' End If
'_______________________________________________________________________________________
'
' Dim astr As String
' Dim l As Long
' l = 0
' While EnumKey(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\control\MediaResources", l, astr)
'   Debug.Print astr
'   l = l + 1
' Wend
'_______________________________________________________________________________________
'
' Set Picture1.Picture = LoadResPicture(ID, pType) ' pType : 0 = bitmap, 1 = Icon, 2 = cursor
'_______________________________________________________________________________________
'
' Global bytsound() As Byte                     ' in Module1
'
' bytsound = LoadResData(101, "CUSTOM")         ' in Forms or Modules
'
' PlaySoundM bytsound(0), zWait                 ' if .wav file
'
' file$ = AppPath(App.Path) & "TempFile_Name"   ' if .avi file, use MCI
' Open file$ For Binary As #1 Len = 1
' Put #1, , bytsound
' Close #1
'
' Dim mmFile As mciFile                         ' if .avi file, use MCI
' mmFile.mFile = file$                          ' open , play , pause , resume
' MciCommand "open", mmFile                     ' step , stop , close , getpos
' MciCommand "play wait", mmFile                ' seek , play wait , stepback
' MciCommand "close", mmFile                    ' gettimeformat , setspeed , fullscreen
'_______________________________________________________________________________________
'
' Do
'   DoEvents
'   Debug.Print isKeyDown(VK_CAPSLOCK), isKeyDown(VK_NUMLOCK), isKeyDown(VK_SCROLL)
'   If isKeyDown(VK_RCONTROL) And isKeyDown(VK_1) Then
'     Debug.Print "RIGHT_CONTROL - 1"
'   End If
' Loop Until isKeyDown(VK_ESCAPE)
'_______________________________________________________________________________________

Public Enum VirtualKey
  VK_LBUTTON = &H1
  VK_RBUTTON = &H2
  VK_CTRLBREAK = &H3
  VK_MBUTTON = &H4
  VK_BACKSPACE = &H8
  VK_TAB = &H9
  VK_ENTER = &HD
  VK_SHIFT = &H10
  VK_CONTROL = &H11
  VK_ALT = &H12
  VK_PAUSE = &H13
  VK_CAPSLOCK = &H14
  VK_ESCAPE = &H1B
  VK_SPACE = &H20
  VK_PAGEUP = &H21
  VK_PAGEDOWN = &H22
  VK_END = &H23
  VK_HOME = &H24
  VK_LEFT = &H25
  VK_UP = &H26
  VK_RIGHT = &H27
  VK_DOWN = &H28
  VK_PRINTSCREEN = &H2C
  VK_INSERT = &H2D
  VK_DELETE = &H2E
  VK_0 = &H30
  VK_1 = &H31
  VK_2 = &H32
  VK_3 = &H33
  VK_4 = &H34
  VK_5 = &H35
  VK_6 = &H36
  VK_7 = &H37
  VK_8 = &H38
  VK_9 = &H39
  VK_A = &H41
  VK_B = &H42
  VK_C = &H43
  VK_D = &H44
  VK_E = &H45
  VK_F = &H46
  VK_G = &H47
  VK_H = &H48
  VK_I = &H49
  VK_J = &H4A
  VK_K = &H4B
  VK_L = &H4C
  VK_M = &H4D
  VK_N = &H4E
  VK_O = &H4F
  VK_P = &H50
  VK_Q = &H51
  VK_R = &H52
  VK_S = &H53
  VK_T = &H54
  VK_U = &H55
  VK_V = &H56
  VK_W = &H57
  VK_X = &H58
  VK_Y = &H59
  VK_Z = &H5A
  VK_LWINDOWS = &H5B
  VK_RWINDOWS = &H5C
  VK_APPSPOPUP = &H5D
  VK_NUMPAD0 = &H60
  VK_NUMPAD1 = &H61
  VK_NUMPAD2 = &H62
  VK_NUMPAD3 = &H63
  VK_NUMPAD4 = &H64
  VK_NUMPAD5 = &H65
  VK_NUMPAD6 = &H66
  VK_NUMPAD7 = &H67
  VK_NUMPAD8 = &H68
  VK_NUMPAD9 = &H69
  VK_MULTIPLY = &H6A
  VK_ADD = &H6B
  VK_SUBTRACT = &H6D
  VK_DECIMAL = &H6E
  VK_DIVIDE = &H6F
  VK_F1 = &H70
  VK_F2 = &H71
  VK_F3 = &H72
  VK_F4 = &H73
  VK_F5 = &H74
  VK_F6 = &H75
  VK_F7 = &H76
  VK_F8 = &H77
  VK_F9 = &H78
  VK_F10 = &H79
  VK_F11 = &H7A
  VK_F12 = &H7B
  VK_NUMLOCK = &H90
  VK_SCROLL = &H91
  VK_LSHIFT = &HA0
  VK_RSHIFT = &HA1
  VK_LCONTROL = &HA2
  VK_RCONTROL = &HA3
  VK_LALT = &HA4
  VK_RALT = &HA5
  VK_POINTVIRGULE = &HBA
  VK_ADD_EQUAL = &HBB
  VK_VIRGULE = &HBC
  VK_MINUS_UNDERLINE = &HBD
  VK_POINT = &HBE
  VK_SLASH = &HBF
  VK_TILDE = &HC0
  VK_LEFTBRACKET = &HDB
  VK_BACKSLASH = &HDC
  VK_RIGHTBRACKET = &HDD
  VK_QUOTE = &HDE
  VK_APOSTROPHE = &HDE
End Enum

Public Type mciFile
      IsVideo As Boolean
      mAlias As Variant
      mFile As Variant
      mHeight As Integer
      mLength As Long
      mWidth As Integer
End Type

Public Type WindowsVersionInfo
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

Public Enum TransType
  LWA_OPAQUE = 0
  LWA_COLORKEY = 1
  LWA_ALPHA = 2
End Enum

Public Enum PlatformType   ' dwPlatformId
  VER_PLATFORM_WIN32s = 0        ' Unknown Version
  VER_PLATFORM_WIN32_WINDOWS = 1 ' Windows 3.1/95/98/Me
  VER_PLATFORM_WIN32_NT = 2      ' Windows NT/2000/XP/.NET
End Enum

Public Enum RegKey   ' lPredefinedKey , hMainKey
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

Public Enum DriveTypeVar
  DRIVE_ERROR = 1
  DRIVE_REMOVABLE = 2
  DRIVE_FIXED = 3
  DRIVE_REMOTE = 4
  DRIVE_CDROM = 5
  DRIVE_RAMDISK = 6
End Enum

Public Enum PRIORITY_CLASS
  REALTIME_PRIORITY = &H100
  HIGH_PRIORITY = &H80
  NORMAL_PRIORITY = &H20
  IDLE_PRIORITY = &H40
End Enum

Public Enum FILE_ATTRIBUTE
  FILE_ATTRIBUTE_DIRECTORY = &H10
  FILE_ATTRIBUTE_ARCHIVE = &H20
  FILE_ATTRIBUTE_NORMAL = &H80
  FILE_ATTRIBUTE_READONLY = &H1
  FILE_ATTRIBUTE_HIDDEN = &H2
  FILE_ATTRIBUTE_SYSTEM = &H4
  FILE_ATTRIBUTE_COMPRESSED = &H800
  FILE_ATTRIBUTE_ENCRYPTED = &H40
  FILE_ATTRIBUTE_TEMPORARY = &H100
  FILE_ATTRIBUTE_OFFLINE = &H1000
  FILE_ATTRIBUTE_SPARSE_FILE = &H200
  FILE_ATTRIBUTE_REPARSE_POINT = &H400
  FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
End Enum

Public Type SYSTEMTIME       ' DayOfWeek :
    wYear As Integer         ' ------------
    wMonth As Integer        ' Dimanche = 0
    wDayOfWeek As Integer    ' Lundi    = 1
    wDay As Integer          ' Mardi    = 2
    wHour As Integer         ' Mercredi = 3
    wMinute As Integer       ' Jeudi    = 4
    wSecond As Integer       ' Vendredi = 5
    wMilliseconds As Integer ' Samedi   = 6
End Type

Public Type Search_File_Type
    dwFileAttributes As FILE_ATTRIBUTE
    nFileSize As Currency
    cPath As Variant
    cFileName As Variant
    cPathAndFileName As Variant
    stCreationTime As SYSTEMTIME
    stLastAccessTime As SYSTEMTIME
    stLastWriteTime As SYSTEMTIME
End Type

Public Declare Function isKeyDown Lib "SuperDLL.DLL" (ByVal zkey As VirtualKey) As Boolean
Public Declare Function isAnyKeyDown Lib "SuperDLL.DLL" (Optional ByVal IgnoreLocksKeys As Boolean = False) As Boolean
Public Declare Function Beep2 Lib "SuperDLL.DLL" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function IsCharAlpha Lib "SuperDLL.DLL" (ByVal cChar As Byte) As Boolean
Public Declare Function IsCharAlphaNumeric Lib "SuperDLL.DLL" (ByVal cChar As Byte) As Boolean
Public Declare Function IsCharNumeric Lib "SuperDLL.DLL" (ByVal cChar As Byte) As Boolean
Public Declare Function IsCharLower Lib "SuperDLL.DLL" (ByVal cChar As Byte) As Boolean
Public Declare Function IsCharUpper Lib "SuperDLL.DLL" (ByVal cChar As Byte) As Boolean
Public Declare Function IsStringNumeric Lib "SuperDLL.DLL" (ByVal cString As String) As Boolean
Public Declare Sub StopSound Lib "SuperDLL.DLL" ()
Public Declare Sub PlaySound Lib "SuperDLL.DLL" (ByVal lpszSoundName As String, Optional ByVal zWait As Boolean = False, Optional ByVal LoopSound As Boolean = False)
Public Declare Sub PlaySoundM Lib "SuperDLL.DLL" (lpszSoundName As Byte, Optional ByVal zWait As Boolean = False, Optional ByVal LoopSound As Boolean = False)
Public Declare Sub Sleep Lib "SuperDLL.DLL" (ByVal dwMilliseconds As Long)
Public Declare Sub Sleep2 Lib "SuperDLL.DLL" (ByVal dwMilliseconds As Long)
Public Declare Sub Flash Lib "SuperDLL.DLL" (zForm As Form)

Public Declare Function AppPath Lib "SuperDLL.DLL" (ByVal zPath As String) As Variant
Public Declare Sub Swap Lib "SuperDLL.DLL" (var1 As Variant, var2 As Variant)
Public Declare Function Trim2 Lib "SuperDLL.DLL" (ByVal cString As String) As Variant
Public Declare Function vbExecute Lib "SuperDLL.DLL" (ByVal var1 As String, Optional ByVal ShowError As Boolean = False) As Long
Public Declare Sub End2 Lib "SuperDLL.DLL" (ByVal uExitCode As Long)
Public Declare Function Exec Lib "SuperDLL.DLL" (ByVal CmdLine As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus) As Boolean
Public Declare Function Exec2 Lib "SuperDLL.DLL" (ByVal CmdLine As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal pclass As PRIORITY_CLASS = NORMAL_PRIORITY) As Boolean
Public Declare Function GetExitCode Lib "SuperDLL.DLL" (ByVal CmdLine As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal zWait As Boolean = False, Optional ByVal pclass As PRIORITY_CLASS = NORMAL_PRIORITY) As Variant
Public Declare Function DriveType Lib "SuperDLL.DLL" (ByVal zDrive As String) As DriveTypeVar
Public Declare Function DriveTypeS Lib "SuperDLL.DLL" (ByVal zDrive As String) As Variant
Public Declare Function FileExist Lib "SuperDLL.DLL" (ByVal strPath As String) As Boolean
Public Declare Function Filexist Lib "SuperDLL.DLL" (ByVal strPath As String) As Boolean
Public Declare Function DirExist Lib "SuperDLL.DLL" (ByVal zPath As String) As Boolean
Public Declare Function SetCurDir Lib "SuperDLL.DLL" (ByVal zPath As String) As Boolean
Public Declare Function FreeSpace Lib "SuperDLL.DLL" (ByVal zDrive As String) As Currency
Public Declare Function FileOrDirExist Lib "SuperDLL.DLL" (ByVal zPath As String) As Boolean
Public Declare Function TreeFind Lib "SuperDLL.DLL" (ByVal zPath As String, ByVal zFile As String) As Variant
Public Declare Function SearchFiles Lib "SuperDLL.DLL" (ByVal zPath As String, ByVal zFiles As String, Optional ByVal SubDirs As Boolean = True, Optional ByRef NumberFound As Long = -1) As Search_File_Type()

Public Declare Sub CloseMCI Lib "SuperDLL.DLL" ()
Public Declare Sub MoveMCI Lib "SuperDLL.DLL" (ZmciFile As mciFile, ByVal X As Long, ByVal Y As Long, Optional ByVal X2 As Long = 0, Optional ByVal Y2 As Long = 0)
Public Declare Function MciCommand Lib "SuperDLL.DLL" (ByVal zCommand As String, ZmciFile As mciFile, Optional ByVal zPos As Long = 0, Optional pb1 As PictureBox = Nothing, Optional ByVal UseSuperMCI As Boolean = False) As Variant

Public Declare Function SetDWordValue Lib "SuperDLL.DLL" (ByVal sKey As String, ByVal sKeyName As String, ByVal KeyValue As Long) As Boolean
Public Declare Function GetDWordValue Lib "SuperDLL.DLL" (ByVal sKey As String, ByVal sKeyName As String) As Variant
Public Declare Function SetBinaryValue Lib "SuperDLL.DLL" (ByVal sKey As String, ByVal sKeyName As String, ByVal KeyValue As String) As Boolean
Public Declare Function GetBinaryValue Lib "SuperDLL.DLL" (ByVal sKey As String, ByVal sKeyName As String) As Variant
Public Declare Function SetStringValue Lib "SuperDLL.DLL" (ByVal sKey As String, ByVal sKeyName As String, ByVal KeyValue As String) As Boolean
Public Declare Function GetStringValue Lib "SuperDLL.DLL" (ByVal sKey As String, ByVal sKeyName As String) As Variant
Public Declare Function CreateKey Lib "SuperDLL.DLL" (ByVal sKey As String) As Boolean
Public Declare Function DeleteKey Lib "SuperDLL.DLL" (ByVal Keyname As String, Optional ByVal Quiet As Boolean = False) As Boolean
Public Declare Function DeleteKeyValue Lib "SuperDLL.DLL" (ByVal sKeyName As String, ByVal sValueName As String, Optional ByVal Quiet As Boolean = False) As Boolean
Public Declare Function KeyExist Lib "SuperDLL.DLL" (ByVal sKey As String) As Boolean
Public Declare Function KeyValueExist Lib "SuperDLL.DLL" (ByVal sKey As String, ByVal sKeyName As String) As Boolean
Public Declare Function BinToHexR Lib "SuperDLL.DLL" (ByVal qwe As String) As Variant
Public Declare Function BinToDecR Lib "SuperDLL.DLL" (ByVal qwe As String) As Variant
Public Declare Function BinToDecA Lib "SuperDLL.DLL" (ByVal qwe As String) As Variant()
Public Declare Function BinToHexA Lib "SuperDLL.DLL" (ByVal qwe As String) As Variant()
Public Declare Function EnumKey Lib "SuperDLL.DLL" (ByVal hMainKey As RegKey, ByVal sSubKey As String, ByVal lIndex As Long, lpStr As Variant) As Boolean
Public Declare Function QueryValue Lib "SuperDLL.DLL" (ByVal lPredefinedKey As RegKey, ByVal sKeyName As String, ByVal sValueName As String) As Variant

Public Declare Function isTransparent Lib "SuperDLL.DLL" (zForm As Form) As TransType
Public Declare Function GetTrans Lib "SuperDLL.DLL" (zForm As Form) As Long
Public Declare Function FadeTo Lib "SuperDLL.DLL" (zForm As Form, Optional ByVal Final As Byte = 127, Optional ByVal vStep As Byte = 3) As Boolean
Public Declare Function FadeIn Lib "SuperDLL.DLL" (zForm As Form, Optional ByVal Final As Byte = 255, Optional ByVal vStep As Byte = 3) As Boolean
Public Declare Function FadeOut Lib "SuperDLL.DLL" (zForm As Form, Optional ByVal Final As Byte = 0, Optional ByVal vStep As Byte = 3) As Boolean
Public Declare Function SetTrans Lib "SuperDLL.DLL" (zForm As Form, Optional ByVal vTrans As Byte = 127) As Boolean
Public Declare Function MakeTrans Lib "SuperDLL.DLL" (zForm As Form, Optional ByVal TransColor As Long = &HFF00FF) As Boolean
Public Declare Function MakeOpaque Lib "SuperDLL.DLL" (zForm As Form) As Boolean
Public Declare Sub FormDrag Lib "SuperDLL.DLL" (TheForm As Object)
Public Declare Sub ShapeMe Lib "SuperDLL.DLL" (ByVal Color As Long, ByVal HorizontalScan As Boolean, Optional Name1 As Form = Nothing, Optional Name2 As PictureBox = Nothing)
Public Declare Sub MakeTransparent Lib "SuperDLL.DLL" (TransForm As Form, Optional ByVal zShapeForm As Boolean = True)
Public Declare Sub ChangeMask Lib "SuperDLL.DLL" (zForm As Form, zPict As PictureBox, Optional ByVal lngTransColor As Long = &HFFFFFF)

Public Declare Function DecToHex Lib "SuperDLL.DLL" (ByVal var1 As Long, Optional ByVal AddToNextType As Boolean = True) As Variant
Public Declare Function HexToDec Lib "SuperDLL.DLL" (ByVal var1 As String) As Long
Public Declare Function HexToBin Lib "SuperDLL.DLL" (ByVal var1 As String, Optional ByVal AddToNextType As Boolean = True, Optional ByVal RemoveLeadingZeros As Boolean = False) As Variant
Public Declare Function BinToHex Lib "SuperDLL.DLL" (ByVal var1 As String, Optional ByVal AddToNextType As Boolean = True, Optional ByVal RemoveLeadingZeros As Boolean = False) As Variant
Public Declare Function BinToDec Lib "SuperDLL.DLL" (ByVal var1 As String) As Long
Public Declare Function DecToBin Lib "SuperDLL.DLL" (ByVal var1 As Long, Optional ByVal AddToNextType As Boolean = True, Optional ByVal RemoveLeadingZeros As Boolean = False) As Variant
Public Declare Function HiByte Lib "SuperDLL.DLL" (ByVal Word As Integer) As Byte
Public Declare Function LoByte Lib "SuperDLL.DLL" (ByVal Word As Integer) As Byte
Public Declare Function HiWord Lib "SuperDLL.DLL" (ByVal DWord As Long) As Integer
Public Declare Function LoWord Lib "SuperDLL.DLL" (ByVal DWord As Long) As Integer
Public Declare Function HiByteHiWord Lib "SuperDLL.DLL" (ByVal DWord As Long) As Byte
Public Declare Function LoByteHiWord Lib "SuperDLL.DLL" (ByVal DWord As Long) As Byte
Public Declare Function HiByteLoWord Lib "SuperDLL.DLL" (ByVal DWord As Long) As Byte
Public Declare Function LoByteLoWord Lib "SuperDLL.DLL" (ByVal DWord As Long) As Byte
Public Declare Function MakeWord Lib "SuperDLL.DLL" (ByVal HByte As Byte, ByVal LByte As Byte) As Integer
Public Declare Function MakeDWordB Lib "SuperDLL.DLL" (ByVal HByteHWord As Byte, ByVal LByteHWord As Byte, ByVal HByteLWord As Byte, ByVal LByteLWord As Byte) As Long
Public Declare Function MakeDWordW Lib "SuperDLL.DLL" (ByVal HWord As Integer, LWord As Integer) As Long

Public Declare Function isNT2000XP Lib "SuperDLL.DLL" () As Boolean
Public Declare Sub SHUTDOWN Lib "SuperDLL.DLL" (Optional ByVal FT As ForceType = EWX_FORCEIFHUNG, Optional ByVal SDT As ShutDownType = EWX_SHUTDOWN)
Public Declare Sub LOGOFF Lib "SuperDLL.DLL" (Optional ByVal FT As ForceType = EWX_FORCEIFHUNG)
Public Declare Sub REBOOT Lib "SuperDLL.DLL" (Optional ByVal FT As ForceType = EWX_FORCEIFHUNG)
Public Declare Sub POWEROFF Lib "SuperDLL.DLL" (Optional ByVal FT As ForceType = EWX_FORCEIFHUNG)
Public Declare Sub LockComputer Lib "SuperDLL.DLL" ()
Public Declare Function GetWindowsVersion Lib "SuperDLL.DLL" (ByRef wvi As WindowsVersionInfo) As Long

Public Declare Function GetCurrentX Lib "SuperDLL.DLL" (ByVal hDC As Long) As Variant
Public Declare Function GetCurrentY Lib "SuperDLL.DLL" (ByVal hDC As Long) As Variant
Public Declare Function GetCurrentPosition Lib "SuperDLL.DLL" (ByVal hDC As Long, ByRef X As Long, ByRef Y As Long) As Long
Public Declare Function MoveTo Lib "SuperDLL.DLL" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function LineTo Lib "SuperDLL.DLL" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetPixel Lib "SuperDLL.DLL" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "SuperDLL.DLL" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function DrawLine Lib "SuperDLL.DLL" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DrawTriangle Lib "SuperDLL.DLL" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function DrawAngleCircle Lib "SuperDLL.DLL" (zFormOrPictBox As Object, ByVal X As Single, ByVal Y As Single, ByVal dwRadius As Single, Optional ByVal StartAngle As Single = 0, Optional ByVal EndAngle As Single = 0, Optional ByVal ForColor As Long = -1, Optional ByVal dWidth As Integer = -1, Optional ByVal UpdateForeColorAndDrawWidth As Boolean = False) As Boolean
Public Declare Function DrawAngleEllipse Lib "SuperDLL.DLL" (zFormOrPictBox As Object, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal StartAngle As Single = 0, Optional ByVal EndAngle As Single = 0, Optional ByVal ForColor As Long = -1, Optional ByVal dWidth As Integer = -1, Optional ByVal UpdateForeColorAndDrawWidth As Boolean = False) As Boolean
Public Declare Function DrawCircle Lib "SuperDLL.DLL" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dwRadius As Long) As Long
Public Declare Function DrawEllipse Lib "SuperDLL.DLL" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DrawRectangle Lib "SuperDLL.DLL" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DrawRoundRect Lib "SuperDLL.DLL" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal pcRoundX As Integer, Optional ByVal pcRoundY As Integer = -1) As Long
Public Declare Function SetColor Lib "SuperDLL.DLL" (zFormOrPictBox As Object, Optional ByVal ForColor As Long = -1, Optional ByVal dWidth As Integer = -1, Optional ByVal FilColor As Long = -1, Optional ByVal FilStyle As FillStyleConstants = -1) As Boolean
Public Declare Function FloodFill Lib "SuperDLL.DLL" (zFormOrPictBox As Object, ByVal X As Long, ByVal Y As Long, ByVal BorderColor As Long, Optional ByVal FilColor As Long = -1, Optional ByVal FilStyle As FillStyleConstants = -1, Optional ByVal UpdateFillColorAndFillStyle As Boolean = False) As Long

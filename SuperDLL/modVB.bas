Attribute VB_Name = "modVB"
Option Explicit

Private Const MAX_PATH As Long = 260

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

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

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

Private Type WIN32_FIND_DATA
    dwFileAttributes As FILE_ATTRIBUTE
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
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

Public Enum PRIORITY_CLASS
  REALTIME_PRIORITY = &H100
  HIGH_PRIORITY = &H80
  NORMAL_PRIORITY = &H20
  IDLE_PRIORITY = &H40
End Enum

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As VbAppWinStyle
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Enum DriveTypeVar
  DRIVE_ERROR = 1
  DRIVE_REMOVABLE = 2
  DRIVE_FIXED = 3
  DRIVE_REMOTE = 4
  DRIVE_CDROM = 5
  DRIVE_RAMDISK = 6
End Enum

Private Declare Function EbExecuteLine Lib "vba6.dll" (ByVal StringToExec As Long, ByVal Any1 As Long, ByVal Any2 As Long, ByVal CheckOnly As Long) As Long
Private Declare Sub ExitProcess Lib "Kernel32.dll" (ByVal uExitCode As Long)
Private Declare Function WinExec Lib "Kernel32.dll" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function CreateProcess Lib "Kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "Kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "Kernel32.dll" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetDriveType Lib "Kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function SetCurrentDirectory Lib "Kernel32.dll" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "Kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function FindFirstFile Lib "Kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "Kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "Kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "Kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SearchTreeForFile Lib "imagehlp" (ByVal RootPath As String, ByVal InputPathName As String, ByVal OutputPathBuffer As String) As Long

Private FileArray() As Search_File_Type
Private TotalFilesFound As Long

Public Function AppPath(ByRef zPathPtr As Long) As Variant
  Dim zPath As String
  zPath = Trim3(CSTOVBS(zPathPtr))
  If Right$(zPath, 1) = "\" Then AppPath = zPath Else AppPath = zPath & "\"
End Function

Public Sub Swap(var1 As Variant, var2 As Variant)
  Dim var3 As Variant
  var3 = var1: var1 = var2: var2 = var3
End Sub

Public Function Trim2(ByRef cStringPtr As Long) As Variant
  Dim t As Long
  Dim Z As Long
  Dim cString As String
  cString = CSTOVBS(cStringPtr)
  For t = 1 To Len(cString)
    If Mid$(cString, t, 1) <> " " And Mid$(cString, t, 1) <> Chr$(0) Then Exit For
  Next t
  For Z = Len(cString) To 1 Step -1
    If Mid$(cString, Z, 1) <> " " And Mid$(cString, Z, 1) <> Chr$(0) Then Exit For
  Next Z
  If Z < t Then
    Trim2 = ""
  ElseIf Z = t Then
    Trim2 = Mid$(cString, t, 1)
  Else
    Trim2 = Mid$(cString, t, (Z - t) + 1)
  End If
End Function

Public Function Trim3(ByVal cString As String) As String
  Dim t As Long
  Dim Z As Long
  For t = 1 To Len(cString)
    If Mid$(cString, t, 1) <> " " And Mid$(cString, t, 1) <> Chr$(0) Then Exit For
  Next t
  For Z = Len(cString) To 1 Step -1
    If Mid$(cString, Z, 1) <> " " And Mid$(cString, Z, 1) <> Chr$(0) Then Exit For
  Next Z
  If Z < t Then
    Trim3 = ""
  ElseIf Z = t Then
    Trim3 = Mid$(cString, t, 1)
  Else
    Trim3 = Mid$(cString, t, (Z - t) + 1)
  End If
End Function

Public Function vbExecute(ByRef var1 As Long, Optional ByVal ShowError As Boolean = False) As Long
  On Local Error GoTo ErrHnd
  Dim var2 As String, var3 As Long
  var2 = Trim3(CSTOVBS(var1))
  var3 = EbExecuteLine(StrPtr(var2), 0&, 0&, 1)
  If var3 = 0 Then
    EbExecuteLine StrPtr(var2), 0&, 0&, 0&
  Else
    If ShowError Then Error var3
  End If
  vbExecute = var3
  Exit Function
ErrHnd:
  MsgBox "Error # " & Err.Number & " : " & Err.Description, vbExclamation, "SuperDLL", Err.HelpFile, Err.HelpContext
  vbExecute = var3
End Function

Public Sub End2(ByVal uExitCode As Long)
  ExitProcess uExitCode
End Sub

Public Function Exec(ByRef lpCmdLine As Long, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus) As Boolean
  Dim CmdLine As String, var1 As Long
  CmdLine = Trim3(CSTOVBS(lpCmdLine))
  var1 = WinExec(CmdLine, WindowStyle)
  If var1 > 31 Then
    Exec = True
  Else
    Exec = False
  End If
End Function

Public Function Exec2(ByRef lpCmdLine As Long, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal pclass As PRIORITY_CLASS = NORMAL_PRIORITY) As Boolean
  Dim sinfo As STARTUPINFO, pinfo As PROCESS_INFORMATION, CmdLine As String
  CmdLine = Trim3(CSTOVBS(lpCmdLine))
  sinfo.cb = Len(sinfo)
  sinfo.dwFlags = &H1
  sinfo.wShowWindow = WindowStyle
  If CreateProcess(vbNullString, CmdLine, ByVal 0&, ByVal 0&, 1&, pclass, ByVal 0&, vbNullString, sinfo, pinfo) <> 0 Then
    CloseHandle pinfo.hThread
    CloseHandle pinfo.hProcess
    Exec2 = True
  Else
    Exec2 = False
  End If
End Function

Public Function GetExitCode(ByRef lpCmdLine As Long, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal zWait As Boolean = False, Optional ByVal pclass As PRIORITY_CLASS = NORMAL_PRIORITY) As Variant
  Const Infinite As Long = &HFFFFFFFF
  Const STILL_ACTIVE As Long = &H103
  Dim sinfo As STARTUPINFO, pinfo As PROCESS_INFORMATION
  Dim CmdLine As String, ExitCode As Long
  CmdLine = Trim3(CSTOVBS(lpCmdLine))
  sinfo.cb = Len(sinfo)
  sinfo.dwFlags = &H1
  sinfo.wShowWindow = WindowStyle
  If CreateProcess(vbNullString, CmdLine, ByVal 0&, ByVal 0&, 1&, pclass, ByVal 0&, vbNullString, sinfo, pinfo) <> 0 Then
    If zWait = True Then WaitForSingleObject pinfo.hProcess, Infinite
    Do
      GetExitCodeProcess pinfo.hProcess, ExitCode
      DoEvents
    Loop While ExitCode = STILL_ACTIVE
    CloseHandle pinfo.hThread
    CloseHandle pinfo.hProcess
    GetExitCode = ExitCode
  Else
    GetExitCode = "ERROR"
  End If
End Function

Private Function GetDT(ByVal var1 As String) As String
  Select Case GetDriveType(var1)
    Case DRIVE_ERROR
      GetDT = "ERROR"
    Case DRIVE_REMOVABLE
      GetDT = "REMOVABLE"
    Case DRIVE_FIXED
      GetDT = "FIXED"
    Case DRIVE_REMOTE
      GetDT = "REMOTE"
    Case DRIVE_CDROM
      GetDT = "CDROM"
    Case DRIVE_RAMDISK
      GetDT = "RAMDISK"
    Case Else
      GetDT = "ERROR"
  End Select
End Function

Public Function DriveTypeS(ByRef zDrivePtr As Long) As Variant
  Dim var1 As String, zDrive As String
  zDrive = Trim3(CSTOVBS(zDrivePtr))
  If ((Len(zDrive) = 2) And (Right$(zDrive, 1) = ":")) Or (Len(zDrive) = 1) Then
    var1 = Left$(zDrive, 1)
    Select Case Asc(var1)
      Case 65 To 90
        DriveTypeS = GetDT(var1 & ":")
      Case 97 To 122
        DriveTypeS = GetDT(var1 & ":")
      Case Else
        MsgBox "Use " & Chr$(34) & "DriveType c" & Chr$(34) & " or " & Chr$(34) & "DriveType c:" & Chr$(34), vbExclamation, "INVALID FORMAT !"
        DriveTypeS = "ERROR"
    End Select
  Else
    MsgBox "Use " & Chr$(34) & "DriveType c" & Chr$(34) & " or " & Chr$(34) & "DriveType c:" & Chr$(34), vbExclamation, "INVALID FORMAT !"
    DriveTypeS = "ERROR"
  End If
End Function

Public Function DriveType(ByRef zDrivePtr As Long) As DriveTypeVar
  Dim var1 As String, zDrive As String
  zDrive = Trim3(CSTOVBS(zDrivePtr))
  If ((Len(zDrive) = 2) And (Right$(zDrive, 1) = ":")) Or (Len(zDrive) = 1) Then
    var1 = Left$(zDrive, 1)
    Select Case Asc(var1)
      Case 65 To 90
        DriveType = GetDriveType(var1 & ":")
      Case 97 To 122
        DriveType = GetDriveType(var1 & ":")
      Case Else
        MsgBox "Use " & Chr$(34) & "DriveType c" & Chr$(34) & " or " & Chr$(34) & "DriveType c:" & Chr$(34), vbExclamation, "INVALID FORMAT !"
        DriveType = DRIVE_ERROR
    End Select
  Else
    MsgBox "Use " & Chr$(34) & "DriveType c" & Chr$(34) & " or " & Chr$(34) & "DriveType c:" & Chr$(34), vbExclamation, "INVALID FORMAT !"
    DriveType = DRIVE_ERROR
  End If
End Function

Public Function FileExist(ByRef strPathPtr As Long) As Boolean
  On Local Error GoTo ErrFile
  Dim strPath As String
  strPath = CSTOVBS(strPathPtr)
  Open strPath For Input Access Read As #1
  Close #1
  FileExist = True
  Exit Function
ErrFile:
  FileExist = False
End Function

Public Function Filexist(ByRef strPath As Long) As Boolean
  Filexist = FileExist(strPath)
End Function

Public Function DirExist(ByRef zPathPtr As Long) As Boolean
  On Local Error GoTo ErrDir
  Dim zPath As String
  zPath = CSTOVBS(zPathPtr)
  Dim qwe As String
  qwe = CurDir
  ChDir zPath
  ChDir qwe
  DirExist = True
  Exit Function
ErrDir:
  DirExist = False
End Function

Public Function SetCurDir(ByRef zPathPtr As Long) As Boolean
  Dim zPath As String
  zPath = CSTOVBS(zPathPtr)
  If SetCurrentDirectory(zPath) <> 0 Then
    SetCurDir = True
  Else
    SetCurDir = False
  End If
End Function

Public Function FreeSpace(ByRef zDrivePtr As Long) As Currency
  Dim var1 As Currency, var2 As Currency, var3 As Currency
  Dim var4 As String, zDrive As String
  zDrive = Trim3(CSTOVBS(zDrivePtr))
  If ((Len(zDrive) = 2) And (Right$(zDrive, 1) = ":")) Or (Len(zDrive) = 1) Then
    var4 = Left$(zDrive, 1)
    Select Case Asc(var4)
      Case 65 To 90
        GetDiskFreeSpaceEx var4 & ":", var1, var2, var3
      Case 97 To 122
        GetDiskFreeSpaceEx var4 & ":", var1, var2, var3
      Case Else
        MsgBox "Use " & Chr$(34) & "FreeSpace c" & Chr$(34) & " or " & Chr$(34) & "FreeSpace c:" & Chr$(34), vbExclamation, "INVALID FORMAT !"
        FreeSpace = -1
        Exit Function
    End Select
  Else
    MsgBox "Use " & Chr$(34) & "FreeSpace c" & Chr$(34) & " or " & Chr$(34) & "FreeSpace c:" & Chr$(34), vbExclamation, "INVALID FORMAT !"
    FreeSpace = -1
    Exit Function
  End If
  If var1 = 0 And var2 = 0 And var3 = 0 Then
    FreeSpace = -1
  Else
    FreeSpace = var1 * 10000
  End If
End Function

Public Function FileOrDirExist(ByRef zPathPtr As Long) As Boolean
  Dim zPath As String
  zPath = CSTOVBS(zPathPtr)
  FileOrDirExist = PathFileExists(zPath)
End Function

Public Function TreeFind(ByRef zPathPtr As Long, ByRef zFilePtr As Long) As Variant
  Dim VarTemp As String, zPath As String, zFile As String
  zPath = CSTOVBS(zPathPtr)
  zFile = CSTOVBS(zFilePtr)
  VarTemp = String(MAX_PATH, 0)
  If SearchTreeForFile(zPath, zFile, VarTemp) <> 0 Then
    TreeFind = Trim3(VarTemp)
  Else
    TreeFind = -1
  End If
End Function

Private Function SearchFilesEx(ByVal zPath As String, ByVal zFiles As String, Optional ByVal SubDirs As Boolean = True, Optional ByRef NumberFound As Long = -1, Optional NewSearch As Boolean = True) As Search_File_Type()
  Dim zPathStr As String, DirCount As Long, FileCount As Long, isOK As Boolean
  Dim RetVal As Long, TempSearch() As WIN32_FIND_DATA, DDir() As String, t As Long
  If NewSearch = True Then TotalFilesFound = 0
  If zFiles = vbNullString Or zFiles = "" Then zFiles = "*.*"
  If Right$(zPath, 1) = "\" Then
    zPathStr = zPath
  Else
    zPathStr = zPath & "\"
  End If
  DirCount = 0
  isOK = True
  ReDim TempSearch(1 To 1)
  RetVal = FindFirstFile(zPathStr & "*.*", TempSearch(1))
  If RetVal <> -1 Then
  Do While isOK
    DoEvents
    If (FILE_ATTRIBUTE_DIRECTORY And TempSearch(1).dwFileAttributes) = FILE_ATTRIBUTE_DIRECTORY Then
      If Trim3(TempSearch(1).cFileName) <> "." And Trim3(TempSearch(1).cFileName) <> ".." Then
        DirCount = DirCount + 1
        ReDim Preserve DDir(DirCount)
        DDir(DirCount) = Trim3(TempSearch(1).cFileName)
      End If
    End If
    ReDim TempSearch(1 To 1)
    isOK = FindNextFile(RetVal, TempSearch(1)) <> 0
  Loop
  End If
  FindClose RetVal
  FileCount = 0
  isOK = True
  ReDim TempSearch(1 To 1)
  RetVal = FindFirstFile(zPathStr & zFiles, TempSearch(1))
  If RetVal <> -1 Then
  Do While isOK
    DoEvents
    If Trim3(TempSearch(1).cFileName) <> "." And Trim3(TempSearch(1).cFileName) <> ".." Then
      FileCount = FileCount + 1
      ReDim Preserve FileArray(TotalFilesFound + 1)
      FileArray(TotalFilesFound + 1).cPath = zPathStr
      FileArray(TotalFilesFound + 1).cFileName = Trim3(TempSearch(1).cFileName)
      FileArray(TotalFilesFound + 1).cPathAndFileName = zPathStr & Trim3(TempSearch(1).cFileName)
      FileArray(TotalFilesFound + 1).dwFileAttributes = TempSearch(1).dwFileAttributes
      FileArray(TotalFilesFound + 1).nFileSize = (4294967296@ * TempSearch(1).nFileSizeHigh) + TempSearch(1).nFileSizeLow
      FileTimeToSystemTime TempSearch(1).ftCreationTime, FileArray(TotalFilesFound + 1).stCreationTime
      FileTimeToSystemTime TempSearch(1).ftLastWriteTime, FileArray(TotalFilesFound + 1).stLastWriteTime
      FileTimeToSystemTime TempSearch(1).ftLastAccessTime, FileArray(TotalFilesFound + 1).stLastAccessTime
      TotalFilesFound = TotalFilesFound + 1
      If NumberFound <> -1 Then NumberFound = TotalFilesFound
    End If
    ReDim TempSearch(1 To 1)
    isOK = FindNextFile(RetVal, TempSearch(1)) <> 0
  Loop
  End If
  FindClose RetVal
  If SubDirs = True Then
    If NumberFound = -1 Then
      For t = 1 To DirCount
        DoEvents
        SearchFilesEx zPathStr & DDir(t), zFiles, True, , False
      Next t
    Else
      For t = 1 To DirCount
        DoEvents
        SearchFilesEx zPathStr & DDir(t), zFiles, True, NumberFound, False
      Next t
    End If
  End If
  If NumberFound <> -1 Then NumberFound = TotalFilesFound
  SearchFilesEx = FileArray
End Function

Public Function SearchFiles(ByRef zPathPtr As Long, ByRef zFilesPtr As Long, Optional ByVal SubDirs As Boolean = True, Optional ByRef NumberFound As Long = -1) As Search_File_Type()
  Dim zPath As String, zFiles As String
  zPath = CSTOVBS(zPathPtr)
  zFiles = CSTOVBS(zFilesPtr)
  SearchFiles = SearchFilesEx(zPath, zFiles, SubDirs, NumberFound, True)
End Function

Attribute VB_Name = "modOther"
Option Explicit

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

Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function BeepZ Lib "kernel32.dll" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function IsCharAlphaZ Lib "user32.dll" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharAlphaNumericZ Lib "user32.dll" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharLowerZ Lib "user32.dll" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharUpperZ Lib "user32.dll" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function sndPlaySound2 Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Declare Sub Sleepy Lib "kernel32.dll" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare Function FlashWindow Lib "user32.dll" (ByVal HWND As Long, ByVal bInvert As Long) As Long

Public Function isKeyDown(ByVal zKey As VirtualKey) As Boolean
  Dim var1 As Integer
  If (zKey = VK_CAPSLOCK) Or (zKey = VK_NUMLOCK) Or (zKey = VK_SCROLL) Then
    var1 = &H1
  Else
    var1 = &H80
  End If
  If (GetKeyState(zKey) And var1) = var1 Then
    isKeyDown = True
  Else
    isKeyDown = False
  End If
End Function

Public Function isAnyKeyDown(Optional ByVal IgnoreLocksKeys As Boolean = False) As Boolean
  Dim t As Integer, KD As Boolean
  Dim keystat(0 To 255) As Byte
  GetKeyboardState keystat(0)
  KD = False
  If IgnoreLocksKeys = False Then
    For t = 0 To 255
      If (keystat(t) And &H80) = &H80 Then
        KD = True
        Exit For
      End If
    Next t
  Else
    For t = 0 To 255
      If ((keystat(t) And &H80) = &H80) And (t <> VK_CAPSLOCK) And (t <> VK_NUMLOCK) And (t <> VK_SCROLL) Then
        KD = True
        Exit For
      End If
    Next t
  End If
  If KD = True Then
    isAnyKeyDown = True
  Else
    isAnyKeyDown = False
  End If
End Function

Public Function Beep2(ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
  Beep2 = BeepZ(ByVal dwFreq, ByVal dwDuration)
End Function

Public Function IsCharAlpha(ByVal cChar As Byte) As Boolean
  IsCharAlpha = IsCharAlphaZ(ByVal cChar)
End Function

Public Function IsCharAlphaNumeric(ByVal cChar As Byte) As Boolean
  IsCharAlphaNumeric = IsCharAlphaNumericZ(ByVal cChar)
End Function

Public Function IsCharNumeric(ByVal cChar As Byte) As Boolean
  IsCharNumeric = IsCharAlphaNumericZ(ByVal cChar) And (Not IsCharAlphaZ(ByVal cChar))
End Function

Public Function IsCharLower(ByVal cChar As Byte) As Boolean
  IsCharLower = IsCharLowerZ(ByVal cChar)
End Function

Public Function IsCharUpper(ByVal cChar As Byte) As Boolean
  IsCharUpper = IsCharUpperZ(ByVal cChar)
End Function

Public Function IsStringNumeric(ByRef cString As Long) As Boolean
  Dim t As Long
  Dim q As String
  q = Trim3(CSTOVBS(cString))
  If Len(q) > 0 Then
    For t = 1 To Len(q)
      If Not IsCharNumeric(Asc(Mid$(q, t, 1))) Then
        IsStringNumeric = False
        Exit Function
      End If
    Next t
    IsStringNumeric = True
  Else
    IsStringNumeric = False
  End If
End Function

Public Sub StopSound()
  sndPlaySound vbNullString, 3
End Sub

Public Sub PlaySound(ByRef lpszSoundName As Long, Optional ByVal zWait As Boolean = False, Optional ByVal LoopSound As Boolean = False)
  Dim SoundName As String, sndFlags As Long
  SoundName = Trim3(CSTOVBS(lpszSoundName))
  sndFlags = 2
  If zWait = False Then sndFlags = sndFlags + 1
  If LoopSound = True Then sndFlags = sndFlags + 8
  sndPlaySound SoundName, sndFlags
End Sub

Public Sub PlaySoundM(lpszSoundName As Byte, Optional ByVal zWait As Boolean = False, Optional ByVal LoopSound As Boolean = False)
  Dim sndFlags As Long
  sndFlags = 6
  If zWait = False Then sndFlags = sndFlags + 1
  If LoopSound = True Then sndFlags = sndFlags + 8
  sndPlaySound2 lpszSoundName, sndFlags
End Sub

Public Sub Sleep(ByVal dwMilliseconds As Long)
Dim zz As Single
zz = Timer
Do
  DoEvents
  If Timer >= zz Then
    If (Timer - zz) >= (dwMilliseconds / 1000) Then Exit Do
  Else
    If ((86400 - zz) + Timer) >= (dwMilliseconds / 1000) Then Exit Do
  End If
Loop
End Sub

Public Sub Sleep2(ByVal dwMilliseconds As Long)
  Sleepy ByVal dwMilliseconds
End Sub

Public Sub Flash(zForm As Form)
  FlashWindow zForm.HWND, 1
End Sub

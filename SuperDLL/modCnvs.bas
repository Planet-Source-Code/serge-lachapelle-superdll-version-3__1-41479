Attribute VB_Name = "modCnvs"
Option Explicit

'Private Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lenCString Lib "kernel32.dll" Alias "lstrlenA" (lpString As Long) As Long
Private Declare Function CopyCString Lib "kernel32.dll" Alias "lstrcpynA" (ByVal lpStringDestination As String, lpStringSource As Long, ByVal lngMaxLength As Long) As Long
Private Declare Function HighByte Lib "TLBINF32.DLL" Alias "hibyte" (ByVal Word As Integer) As Byte
Private Declare Function LowByte Lib "TLBINF32.DLL" Alias "lobyte" (ByVal Word As Integer) As Byte
Private Declare Function HighWord Lib "TLBINF32.DLL" Alias "hiword" (ByVal DWord As Long) As Integer
Private Declare Function LowWord Lib "TLBINF32.DLL" Alias "loword" (ByVal DWord As Long) As Integer

'Public Function VBSTOCS(ByVal var1 As String) As String
'  Dim sResult As String
'  Dim qwe As String
'  qwe = var1
'  CopyMemory ByVal VarPtr(sResult), qwe, 4 ' LenB(qwe) ' Len(qwe)
'  VBSTOCS = sResult
'End Function

Public Function CSTOVBS(ByRef lpCString As Long) As String
    Dim lenString As Long, sBuffer As String, lpBuffer As Long
    If lpCString <> 0 Then
        lenString = lenCString(lpCString)
        sBuffer = String$(lenString + 1, 0)
        lpBuffer = CopyCString(sBuffer, lpCString, lenString + 1)
        If Right$(sBuffer, 1) = Chr$(0) Then sBuffer = Left$(sBuffer, Len(sBuffer) - 1)
        CSTOVBS = sBuffer
    Else
      CSTOVBS = "CSTOVBS_ERROR_NO_STRING"
    End If
End Function

Public Function DecToHex(ByVal var1 As Long, Optional ByVal AddToNextType As Boolean = True) As Variant
  Dim var2 As String
  var2 = Hex(var1)
Debug.Print var2
  If AddToNextType Then
    Select Case Len(var2)
      Case 1, 3, 7
        var2 = "0" & var2
      Case 5
        var2 = "000" & var2
      Case 6
        var2 = "00" & var2
    End Select
  End If
  DecToHex = var2
End Function

Public Function HexToDec(ByRef var0 As Long) As Long
  Dim var1 As String
  var1 = Trim3(CSTOVBS(var0))
  On Local Error GoTo ErrCnv
  HexToDec = "&h" & var1
  Exit Function
ErrCnv:
  HexToDec = -1
End Function

Private Function HexBin(ByVal var1 As String, Optional ByVal AddToNextType As Boolean = True, Optional ByVal RemoveLeadingZeros As Boolean = False) As Variant
  Dim t As Long
  Dim qaz As String
  qaz = ""
  If Len(var1) = 0 Then
    HexBin = -1
    Exit Function
  Else
    For t = 1 To Len(var1)
      Select Case UCase$(Mid$(var1, t, 1))
        Case "0"
          qaz = qaz & "0000"
        Case "1"
          qaz = qaz & "0001"
        Case "2"
          qaz = qaz & "0010"
        Case "3"
          qaz = qaz & "0011"
        Case "4"
          qaz = qaz & "0100"
        Case "5"
          qaz = qaz & "0101"
        Case "6"
          qaz = qaz & "0110"
        Case "7"
          qaz = qaz & "0111"
        Case "8"
          qaz = qaz & "1000"
        Case "9"
          qaz = qaz & "1001"
        Case "A"
          qaz = qaz & "1010"
        Case "B"
          qaz = qaz & "1011"
        Case "C"
          qaz = qaz & "1100"
        Case "D"
          qaz = qaz & "1101"
        Case "E"
          qaz = qaz & "1110"
        Case "F"
          qaz = qaz & "1111"
        Case Else
          HexBin = -1
          Exit Function
      End Select
    Next t
  End If
  If RemoveLeadingZeros Then
    For t = 1 To Len(qaz)
      If Mid$(qaz, t, 1) <> "0" Then Exit For
    Next t
    qaz = Mid$(qaz, t)
  ElseIf AddToNextType Then
    Select Case Len(qaz)
      Case 4, 12, 28
        qaz = "0000" & qaz
      Case 20
        qaz = "000000000000" & qaz
      Case 24
        qaz = "00000000" & qaz
    End Select
  End If
  HexBin = qaz
End Function

Public Function HexToBin(ByRef var0 As Long, Optional ByVal AddToNextType As Boolean = True, Optional ByVal RemoveLeadingZeros As Boolean = False) As Variant
  Dim var1 As String
  Dim qwe As Variant
  var1 = Trim3(CSTOVBS(var0))
  qwe = HexBin(var1, AddToNextType, RemoveLeadingZeros)
  If Len(qwe) = 0 Then qwe = "0"
  HexToBin = qwe
End Function

Private Function BinHex(ByVal var0 As String) As Variant
  Select Case UCase$(var0)
    Case "0000", "000", "00", "0"
      BinHex = "0"
    Case "0001", "001", "01", "1"
      BinHex = "1"
    Case "0010", "010", "10"
      BinHex = "2"
    Case "0011", "011", "11"
      BinHex = "3"
    Case "0100", "100"
      BinHex = "4"
    Case "0101", "101"
      BinHex = "5"
    Case "0110", "110"
      BinHex = "6"
    Case "0111", "111"
      BinHex = "7"
    Case "1000"
      BinHex = "8"
    Case "1001"
      BinHex = "9"
    Case "1010"
      BinHex = "A"
    Case "1011"
      BinHex = "B"
    Case "1100"
      BinHex = "C"
    Case "1101"
      BinHex = "D"
    Case "1110"
      BinHex = "E"
    Case "1111"
      BinHex = "F"
    Case Else
      BinHex = -1
  End Select
End Function

Public Function BinToHex(ByRef var0 As Long, Optional ByVal AddToNextType As Boolean = True, Optional ByVal RemoveLeadingZeros As Boolean = False) As Variant
  Dim t As Long
  Dim q As Variant
  Dim qwe As String
  Dim qaz As String
  qwe = Trim3(CSTOVBS(var0))
  qaz = ""
  If Len(qwe) = 0 Then
    BinToHex = -1
    Exit Function
  Else
    Do
      q = BinHex(Right$(qwe, 4))
      If q = -1 Then
        BinToHex = -1
        Exit Function
      End If
      qaz = q & qaz
      If Len(qwe) <= 4 Then
        qwe = ""
      Else
        qwe = Left$(qwe, Len(qwe) - 4)
      End If
    Loop Until Len(qwe) < 1
  End If
  If RemoveLeadingZeros Then
    For t = 1 To Len(qaz)
      If Mid$(qaz, t, 1) <> "0" Then Exit For
    Next t
    qaz = Mid$(qaz, t)
  ElseIf AddToNextType Then
    Select Case Len(qaz)
      Case 1, 3, 7
        qaz = "0" & qaz
      Case 5
        qaz = "000" & qaz
      Case 6
        qaz = "00" & qaz
    End Select
  End If
  If Len(qaz) = 0 Then qaz = "0"
  BinToHex = qaz
End Function

Public Function BinToDec(ByRef var0 As Long) As Long
  Dim qwe As String
  qwe = BinToHex(var0, False, False)
  On Local Error GoTo ErrCnv
  BinToDec = "&h" & qwe
  Exit Function
ErrCnv:
  BinToDec = -1
End Function

Public Function DecToBin(ByVal var1 As Long, Optional ByVal AddToNextType As Boolean = True, Optional ByVal RemoveLeadingZeros As Boolean = False) As Variant
  Dim qwe As String
  Dim qaz As String
  qwe = DecToHex(var1, False)
  qaz = HexBin(qwe, AddToNextType, RemoveLeadingZeros)
  If Len(qaz) = 0 Then qaz = "0"
  DecToBin = qaz
End Function

Public Function HiByte(ByVal Word As Integer) As Byte
  HiByte = HighByte(Word)
End Function

Public Function LoByte(ByVal Word As Integer) As Byte
  LoByte = LowByte(Word)
End Function

Public Function HiWord(ByVal DWord As Long) As Integer
  HiWord = HighWord(DWord)
End Function

Public Function LoWord(ByVal DWord As Long) As Integer
  LoWord = LowWord(DWord)
End Function

Public Function HiByteHiWord(ByVal DWord As Long) As Byte
  HiByteHiWord = HighByte(HighWord(DWord))
End Function

Public Function LoByteHiWord(ByVal DWord As Long) As Byte
  LoByteHiWord = LowByte(HighWord(DWord))
End Function

Public Function HiByteLoWord(ByVal DWord As Long) As Byte
  HiByteLoWord = HighByte(LowWord(DWord))
End Function

Public Function LoByteLoWord(ByVal DWord As Long) As Byte
  LoByteLoWord = LowByte(LowWord(DWord))
End Function

Public Function MakeWord(ByVal HByte As Byte, ByVal LByte As Byte) As Integer
  Dim var1 As Long
  var1 = (256& * HByte) + LByte
  If var1 > 32767 Then var1 = var1 - 65536
  MakeWord = var1
End Function

Public Function MakeDWordB(ByVal HByteHWord As Byte, ByVal LByteHWord As Byte, ByVal HByteLWord As Byte, ByVal LByteLWord As Byte) As Long
  Dim var1 As Currency
  var1 = (16777216@ * HByteHWord) + (65536@ * LByteHWord) + (256@ * HByteLWord) + LByteLWord
  If var1 > 2147483647 Then var1 = var1 - 4294967296@
  MakeDWordB = var1
End Function

Public Function MakeDWordW(ByVal HWord As Integer, LWord As Integer) As Long
  Dim var1 As Currency
  var1 = (16777216@ * HighByte(HWord)) + (65536@ * LowByte(HWord)) + (256@ * HighByte(LWord)) + LowByte(LWord)
  If var1 > 2147483647 Then var1 = var1 - 4294967296@
  MakeDWordW = var1
End Function

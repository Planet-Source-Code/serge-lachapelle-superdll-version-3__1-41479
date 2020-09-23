Attribute VB_Name = "modMCI"
Option Explicit

Private Type mciFile
      IsVideo As Boolean
      mAlias As Variant
      mFile As Variant
      mHeight As Integer
      mLength As Long
      mWidth As Integer
End Type

Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Sub CloseMCI()
  mciExecute "close all"
End Sub

Public Sub MoveMCI(ZmciFile As mciFile, ByVal X As Long, ByVal Y As Long, Optional ByVal X2 As Long = 0, Optional ByVal Y2 As Long = 0)
  If ZmciFile.IsVideo Then mciExecute "put " & ZmciFile.mAlias & " window client at " & Str(X) & " " & Str(Y) & " " & Str(X2) & " " & Str(Y2)
End Sub

Private Function FExist(ByVal strPath As String) As Boolean
  On Local Error GoTo ErrFile
  Open strPath For Input Access Read As #1
  Close #1
  FExist = 1
  Exit Function
ErrFile:
  FExist = 0
End Function

Private Function MciCommand2(zCommand As String, ZmciFile As mciFile, Optional ByVal zPos As Long = 0, Optional pb1 As PictureBox = Nothing, Optional ByVal UseSuperMCI As Boolean = False) As Variant
  Dim rtn As String
  Dim t As Long
  Dim qwe As Single
  If ZmciFile.mAlias = "" Then
    qwe = Timer
    While qwe < 100000
      qwe = qwe * 10
    Wend
    qwe = Int(qwe)
    ZmciFile.mAlias = Right$(Str(qwe), Len(Str(qwe)) - 1)
  End If
  Select Case LCase$(zCommand)
    Case "getpos"
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        rtn = Space$(255)
        mciSendString "status " & ZmciFile.mAlias & " position", rtn, Len(rtn), 0
        MciCommand2 = Val(rtn)
      Else
        MciCommand2 = 0
      End If
    Case "getstatus"
      rtn = Space$(255)
      mciSendString "status " & ZmciFile.mAlias & " mode", rtn, Len(rtn), 0
      For t = 1 To Len(rtn)
        If Mid$(rtn, t, 1) = " " Or Mid$(rtn, t, 1) = Chr$(0) Then Exit For
      Next t
      If t > 1 Then
        MciCommand2 = Left$(rtn, t - 1)
      Else
        MciCommand2 = 0
      End If
    Case "gettimeformat"
      rtn = Space$(255)
      mciSendString "status " & ZmciFile.mAlias & " time format", rtn, Len(rtn), 0
      For t = 1 To Len(rtn)
        If Mid$(rtn, t, 1) = " " Or Mid$(rtn, t, 1) = Chr$(0) Then Exit For
      Next t
      If t > 1 Then
        MciCommand2 = Left$(rtn, t - 1)
      Else
        MciCommand2 = 0
      End If
    Case Else:
      MciCommand2 = mciExecute(zCommand)
      'MsgBox "Unknown MCI Command !", vbExclamation, "Error"
  End Select
End Function

Public Function MciCommand(ByRef zCommandPtr As Long, ZmciFile As mciFile, Optional ByVal zPos As Long = 0, Optional pb1 As PictureBox = Nothing, Optional ByVal UseSuperMCI As Boolean = False) As Variant
  Dim zCommand As String
  Dim rtn As String
  Dim qaz() As String
  Dim t As Long
  Dim qwe As Single
  zCommand = Trim3(CSTOVBS(zCommandPtr))
  If ZmciFile.mAlias = "" Then
    qwe = Timer
    While qwe < 100000
      qwe = qwe * 10
    Wend
    qwe = Int(qwe)
    ZmciFile.mAlias = Right$(Str(qwe), Len(Str(qwe)) - 1)
  End If
  Select Case LCase$(zCommand)
    Case "open":
      If FExist(ZmciFile.mFile) Then
        If MciCommand2("getstatus", ZmciFile) <> 0 Then
          mciExecute "close " & ZmciFile.mAlias
        End If
        Select Case LCase$(Right$(ZmciFile.mFile, 4))
          Case ".avi", ".mpg", "mpeg", ".dat", ".asf", ".wmv", "mpv2", ".mpv", ".mpe", "mp2v", ".m1v"
            ZmciFile.IsVideo = True
          Case Else
            ZmciFile.IsVideo = False
        End Select
        If ZmciFile.IsVideo = True Then
          If pb1 Is Nothing Then
            If UseSuperMCI Then
              MciCommand = mciExecute("open " & Chr$(34) & "SuperMCI!" & ZmciFile.mFile & Chr$(34) & " alias " & ZmciFile.mAlias)
            Else
              MciCommand = mciExecute("open " & Chr$(34) & ZmciFile.mFile & Chr$(34) & " alias " & ZmciFile.mAlias)
            End If
          Else
            If UseSuperMCI Then
              MciCommand = mciExecute("open " & Chr$(34) & "SuperMCI!" & ZmciFile.mFile & Chr$(34) & " alias " & ZmciFile.mAlias & " parent " & pb1.HWND & " style child")
            Else
              MciCommand = mciExecute("open " & Chr$(34) & ZmciFile.mFile & Chr$(34) & " alias " & ZmciFile.mAlias & " parent " & pb1.HWND & " style child")
            End If
          End If
          mciExecute "set " & ZmciFile.mAlias & " seek exactly on"
          rtn = Space$(255)
          mciSendString "status " & ZmciFile.mAlias & " length", rtn, Len(rtn), 0
          ZmciFile.mLength = Val(rtn)
          rtn = Space$(255)
          mciSendString "where " & ZmciFile.mAlias & " destination", rtn, Len(rtn), 0
          qaz = Split(rtn, Chr(32), -1, vbTextCompare)
          ZmciFile.mWidth = Val(qaz(2))
          ZmciFile.mHeight = Val(qaz(3))
        Else
          If UseSuperMCI Then
            MciCommand = mciExecute("open " & Chr$(34) & "SuperMCI!" & ZmciFile.mFile & Chr$(34) & " alias " & ZmciFile.mAlias)
          Else
            MciCommand = mciExecute("open " & Chr$(34) & ZmciFile.mFile & Chr$(34) & " alias " & ZmciFile.mAlias)
          End If
          rtn = Space$(255)
          mciSendString "status " & ZmciFile.mAlias & " length", rtn, Len(rtn), 0
          ZmciFile.mLength = Val(rtn)
        End If
        mciExecute "play " & ZmciFile.mAlias
        mciExecute "stop " & ZmciFile.mAlias
        mciExecute "seek " & ZmciFile.mAlias & " to start wait"
      Else
        MsgBox "File Not Found : " & ZmciFile.mFile, vbExclamation, "Error"
        MciCommand = 0
      End If
    Case "play":
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        If MciCommand2("getstatus", ZmciFile) <> "paused" Then
          mciExecute "stop " & ZmciFile.mAlias
          mciExecute "seek " & ZmciFile.mAlias & " to start wait"
        End If
        MciCommand = mciExecute("play " & ZmciFile.mAlias)
      Else
        MciCommand = 0
      End If
    Case "play wait":
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        If MciCommand2("getstatus", ZmciFile) <> "paused" Then
          mciExecute "stop " & ZmciFile.mAlias
          mciExecute "seek " & ZmciFile.mAlias & " to start wait"
        End If
        MciCommand = mciExecute("play " & ZmciFile.mAlias & " wait")
      Else
        MciCommand = 0
      End If
    Case "fullscreen":
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        MciCommand = mciExecute("play " & ZmciFile.mAlias & " fullscreen")
      Else
        MciCommand = 0
      End If
    Case "resume":
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        MciCommand = mciExecute("play " & ZmciFile.mAlias)
      Else
        MciCommand = 0
      End If
    Case "pause":
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        Select Case LCase$(MciCommand2("getstatus", ZmciFile))
          Case "playing"
            MciCommand = mciExecute("pause " & ZmciFile.mAlias)
            Exit Function
          Case "paused"
            MciCommand = mciExecute("play " & ZmciFile.mAlias)
            Exit Function
          Case "stopped"
            If MciCommand2("getpos", ZmciFile) > 0 Then
              MciCommand = mciExecute("play " & ZmciFile.mAlias)
              Exit Function
            Else
              MciCommand = 1
              Exit Function
            End If
        End Select
      Else
        MciCommand = 0
      End If
    Case "stop":
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        MciCommand = mciExecute("stop " & ZmciFile.mAlias)
        MciCommand = MciCommand And mciExecute("seek " & ZmciFile.mAlias & " to start wait")
      Else
        MciCommand = 0
      End If
    Case "close":
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        MciCommand = mciExecute("close " & ZmciFile.mAlias)
      Else
        MciCommand = 0
      End If
    Case "step":
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        If zPos = 0 Then zPos = 1
        If MciCommand2("getpos", ZmciFile) >= ZmciFile.mLength Then
          mciExecute "stop " & ZmciFile.mAlias
          mciExecute "seek " & ZmciFile.mAlias & " to start wait"
          mciExecute "pause " & ZmciFile.mAlias
        Else
          mciExecute "seek " & ZmciFile.mAlias & " to" & Str(MciCommand2("getpos", ZmciFile) + zPos) & " wait"
          mciExecute "pause " & ZmciFile.mAlias
        End If
        MciCommand = MciCommand2("getpos", ZmciFile)
      Else
        MciCommand = 0
      End If
    Case "stepback":
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        If zPos = 0 Then zPos = 1
        If MciCommand2("getpos", ZmciFile) > 0 Then
          mciExecute "seek " & ZmciFile.mAlias & " to" & Str(MciCommand2("getpos", ZmciFile) - zPos) & " wait"
          mciExecute "pause " & ZmciFile.mAlias
        End If
        MciCommand = MciCommand2("getpos", ZmciFile)
      Else
        MciCommand = 0
      End If
    Case "seek"
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        If MciCommand2("getstatus", ZmciFile) <> "playing" Then
          MciCommand = mciExecute("seek " & ZmciFile.mAlias & " to" & Str(zPos) & " wait")
          mciExecute "pause " & ZmciFile.mAlias
        Else
          MciCommand = mciExecute("seek " & ZmciFile.mAlias & " to" & Str(zPos) & " wait")
          mciExecute "play " & ZmciFile.mAlias
        End If
      Else
        MciCommand = 0
      End If
    Case "getpos"
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        rtn = Space$(255)
        mciSendString "status " & ZmciFile.mAlias & " position", rtn, Len(rtn), 0
        MciCommand = Val(rtn)
      Else
        MciCommand = 0
      End If
    Case "getstatus"
      rtn = Space$(255)
      mciSendString "status " & ZmciFile.mAlias & " mode", rtn, Len(rtn), 0
      For t = 1 To Len(rtn)
        If Mid$(rtn, t, 1) = " " Or Mid$(rtn, t, 1) = Chr$(0) Then Exit For
      Next t
      If t > 1 Then
        MciCommand = Left$(rtn, t - 1)
      Else
        MciCommand = 0
      End If
    Case "gettimeformat"
      rtn = Space$(255)
      mciSendString "status " & ZmciFile.mAlias & " time format", rtn, Len(rtn), 0
      For t = 1 To Len(rtn)
        If Mid$(rtn, t, 1) = " " Or Mid$(rtn, t, 1) = Chr$(0) Then Exit For
      Next t
      If t > 1 Then
        MciCommand = Left$(rtn, t - 1)
      Else
        MciCommand = 0
      End If
    Case "setspeed"
      If MciCommand2("getstatus", ZmciFile) <> 0 Then
        MciCommand = mciExecute("set " & ZmciFile.mAlias & " speed" & Str(zPos * 10))
      Else
        MciCommand = 0
      End If
    Case Else:
      MciCommand = mciExecute(zCommand)
      'MsgBox "Unknown MCI Command !", vbExclamation, "Error"
  End Select
End Function

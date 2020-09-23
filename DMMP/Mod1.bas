Attribute VB_Name = "Mod1"
Option Explicit

Public Type mciFile
      IsVideo As Boolean
      mAlias As Variant
      mfile As Variant
      mHeight As Integer
      mLength As Long
      mWidth As Integer
End Type

Public Declare Sub Sleep Lib "SuperDLL" (ByVal dwMilliseconds As Long)
Public Declare Function FileExist Lib "SuperDLL" (ByVal strPath As String) As Boolean
Public Declare Sub CloseMCI Lib "SuperDLL" ()
Public Declare Sub MoveMCI Lib "SuperDLL" (ZmciFile As mciFile, ByVal X As Long, ByVal Y As Long, Optional ByVal X2 As Long = 0, Optional ByVal Y2 As Long = 0)
Public Declare Function MciCommand Lib "SuperDLL" (ByVal zCommand As String, ZmciFile As mciFile, Optional ByVal zPos As Long = 0, Optional pb1 As PictureBox = Nothing, Optional UseSuperMCI As Boolean = False) As Variant
Public Declare Function SetStringValue Lib "SuperDLL" (ByVal sKey As String, ByVal sKeyName As String, ByVal KeyValue As String)
Public Declare Function KeyExist Lib "SuperDLL" (ByVal sKey As String) As Boolean

Sub main()
  
  Form1.Show

End Sub

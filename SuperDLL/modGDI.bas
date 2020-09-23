Attribute VB_Name = "modGDI"
Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function GetCurrentPositionEx Lib "GDI32.DLL" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
Private Declare Function MoveToEx Lib "GDI32.DLL" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo2 Lib "GDI32.DLL" Alias "LineTo" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetPixel2 Lib "GDI32.DLL" Alias "GetPixel" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel2 Lib "GDI32.DLL" Alias "SetPixel" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function Polygon Lib "GDI32.DLL" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Ellipse Lib "GDI32.DLL" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Rectangle Lib "GDI32.DLL" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "GDI32.DLL" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FloodFill2 Lib "GDI32.DLL" Alias "FloodFill" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
  
Public Function GetCurrentX(ByVal hDC As Long) As Variant
  Dim TMP As POINTAPI
  If GetCurrentPositionEx(hDC, TMP) = 0 Then
    GetCurrentX = "ERROR"
  Else
    GetCurrentX = TMP.X
  End If
End Function

Public Function GetCurrentY(ByVal hDC As Long) As Variant
  Dim TMP As POINTAPI
  If GetCurrentPositionEx(hDC, TMP) = 0 Then
    GetCurrentY = "ERROR"
  Else
    GetCurrentY = TMP.Y
  End If
End Function

Public Function GetCurrentPosition(ByVal hDC As Long, ByRef X As Long, ByRef Y As Long) As Long
  Dim TMP As POINTAPI, var1 As Long
  var1 = GetCurrentPositionEx(hDC, TMP)
  If var1 = 0 Then
    GetCurrentPosition = 0
  Else
    X = TMP.X
    Y = TMP.Y
    GetCurrentPosition = var1
  End If
End Function

Public Function MoveTo(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
  Dim TMP As POINTAPI
  MoveTo = MoveToEx(hDC, X, Y, TMP)
End Function

Public Function LineTo(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
  LineTo = LineTo2(hDC, X, Y)
End Function

Public Function GetPixel(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
  GetPixel = GetPixel2(hDC, X, Y)
End Function

Public Function SetPixel(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
  SetPixel = SetPixel2(hDC, X, Y, crColor)
End Function

Public Function DrawLine(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  Dim TMP2(0 To 1) As POINTAPI
  TMP2(0).X = X1
  TMP2(0).Y = Y1
  TMP2(1).X = X2
  TMP2(1).Y = Y2
  DrawLine = Polygon(hDC, TMP2(0), 2)
End Function

Public Function DrawTriangle(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
  Dim TMP2(0 To 2) As POINTAPI
  TMP2(0).X = X1
  TMP2(0).Y = Y1
  TMP2(1).X = X2
  TMP2(1).Y = Y2
  TMP2(2).X = X3
  TMP2(2).Y = Y3
  DrawTriangle = Polygon(hDC, TMP2(0), 3)
End Function

Public Function DrawAngleCircle(zFormOrPictBox As Object, ByVal X As Single, ByVal Y As Single, ByVal dwRadius As Single, Optional ByVal StartAngle As Single = 0, Optional ByVal EndAngle As Single = 0, Optional ByVal ForColor As Long = -1, Optional ByVal dWidth As Integer = -1, Optional ByVal UpdateForeColorAndDrawWidth As Boolean = False) As Boolean
  On Local Error GoTo ErrHnd
  Const PI As Single = 3.14159265
  Dim SM As Integer, FC As Long, DW As Integer
  If StartAngle < 0 Or EndAngle < 0 Or StartAngle > 360 Or EndAngle > 360 Then
    MsgBox "StartAngle and EndAngle must be between 0 and 360 !", vbExclamation, "SuperDLL"
    DrawAngleCircle = False
    Exit Function
  End If
  If ForColor <> -1 Then
    FC = zFormOrPictBox.ForeColor
    zFormOrPictBox.ForeColor = ForColor
  End If
  If dWidth <> -1 Then
    DW = zFormOrPictBox.DrawWidth
    zFormOrPictBox.DrawWidth = dWidth
  End If
  If StartAngle = 0 And EndAngle = 360 Then EndAngle = 0
  SM = zFormOrPictBox.ScaleMode
  zFormOrPictBox.ScaleMode = 3
  zFormOrPictBox.Circle (X, Y), dwRadius, , (StartAngle * PI) / 180, (EndAngle * PI / 180)
  zFormOrPictBox.ScaleMode = SM
  If UpdateForeColorAndDrawWidth = False Then
    If ForColor <> -1 Then zFormOrPictBox.ForeColor = FC
    If dWidth <> -1 Then zFormOrPictBox.DrawWidth = DW
  End If
  DrawAngleCircle = True
  Exit Function
ErrHnd:
  DrawAngleCircle = False
End Function

Public Function DrawAngleEllipse(zFormOrPictBox As Object, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal StartAngle As Single = 0, Optional ByVal EndAngle As Single = 0, Optional ByVal ForColor As Long = -1, Optional ByVal dWidth As Integer = -1, Optional ByVal UpdateForeColorAndDrawWidth As Boolean = False) As Boolean
  On Local Error GoTo ErrHnd
  Const PI As Single = 3.14159265
  Dim dwRadius As Single, Aspect As Single, SM As Integer, FC As Long, DW As Integer
  If StartAngle < 0 Or EndAngle < 0 Or StartAngle > 360 Or EndAngle > 360 Then
    MsgBox "StartAngle and EndAngle must be between 0 and 360 !", vbExclamation, "SuperDLL"
    DrawAngleEllipse = False
    Exit Function
  End If
  If ForColor <> -1 Then
    FC = zFormOrPictBox.ForeColor
    zFormOrPictBox.ForeColor = ForColor
  End If
  If dWidth <> -1 Then
    DW = zFormOrPictBox.DrawWidth
    zFormOrPictBox.DrawWidth = dWidth
  End If
  If StartAngle = 0 And EndAngle = 360 Then EndAngle = 0
  Aspect = (Y2 - Y1) / (X2 - X1)
  dwRadius = (X2 - X1) / 2
  If (X2 - X1) = (Y2 - Y1) Then
    Aspect = 1
  ElseIf (Y2 - Y1) > (X2 - X1) Then
    dwRadius = (Y2 - Y1) / 2
  End If
  SM = zFormOrPictBox.ScaleMode
  zFormOrPictBox.ScaleMode = 3
  zFormOrPictBox.Circle ((X2 + X1) / 2, (Y2 + Y1) / 2), dwRadius, , (StartAngle * PI) / 180, (EndAngle * PI / 180), Aspect
  zFormOrPictBox.ScaleMode = SM
  If UpdateForeColorAndDrawWidth = False Then
    If ForColor <> -1 Then zFormOrPictBox.ForeColor = FC
    If dWidth <> -1 Then zFormOrPictBox.DrawWidth = DW
  End If
  DrawAngleEllipse = True
  Exit Function
ErrHnd:
  DrawAngleEllipse = False
End Function

Public Function DrawCircle(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dwRadius As Long) As Long
  DrawCircle = Ellipse(hDC, X - dwRadius, Y - dwRadius, X + dwRadius + 1, Y + dwRadius + 1)
End Function

Public Function DrawEllipse(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  DrawEllipse = Ellipse(hDC, X1 + IIf(X2 >= X1, 0, 1), Y1 + IIf(Y2 >= Y1, 0, 1), X2 + IIf(X2 >= X1, 1, 0), Y2 + IIf(Y2 >= Y1, 1, 0))
End Function

Public Function DrawRectangle(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  DrawRectangle = Rectangle(hDC, X1 + IIf(X2 >= X1, 0, 1), Y1 + IIf(Y2 >= Y1, 0, 1), X2 + IIf(X2 >= X1, 1, 0), Y2 + IIf(Y2 >= Y1, 1, 0))
End Function

Public Function DrawRoundRect(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal pcRoundX As Integer, Optional ByVal pcRoundY As Integer = -1) As Long
  Dim X3 As Long, Y3 As Long
  If pcRoundX > 100 Or pcRoundY > 100 Or pcRoundX < 0 Or pcRoundY < -1 Then
    MsgBox "pcRoundX and pcRoundY must be between 0 and 100 !", vbExclamation, "SuperDLL"
    DrawRoundRect = 0
  Else
    X3 = (pcRoundX * (X2 - X1)) / 100
    If pcRoundY = -1 Then
      Y3 = X3
    Else
      Y3 = (pcRoundY * (Y2 - Y1)) / 100
    End If
    DrawRoundRect = RoundRect(hDC, X1 + IIf(X2 >= X1, 0, 1), Y1 + IIf(Y2 >= Y1, 0, 1), X2 + IIf(X2 >= X1, 1, 0), Y2 + IIf(Y2 >= Y1, 1, 0), X3, Y3)
  End If
End Function

Public Function SetColor(zFormOrPictBox As Object, Optional ByVal ForColor As Long = -1, Optional ByVal dWidth As Integer = -1, Optional ByVal FilColor As Long = -1, Optional ByVal FilStyle As FillStyleConstants = -1) As Boolean
  On Local Error GoTo ErrHnd
  If ForColor <> -1 Then zFormOrPictBox.ForeColor = ForColor
  If dWidth <> -1 Then zFormOrPictBox.DrawWidth = dWidth
  If FilColor <> -1 Then zFormOrPictBox.FillColor = FilColor
  If FilStyle <> -1 Then zFormOrPictBox.FillStyle = FilStyle
  SetColor = True
  Exit Function
ErrHnd:
  SetColor = False
End Function

Public Function FloodFill(zFormOrPictBox As Object, ByVal X As Long, ByVal Y As Long, ByVal BorderColor As Long, Optional ByVal FilColor As Long = -1, Optional ByVal FilStyle As FillStyleConstants = -1, Optional ByVal UpdateFillColorAndFillStyle As Boolean = False) As Long
  On Local Error GoTo ErrHnd
  Dim FC As Long, FS As FillStyleConstants
  If FilColor <> -1 Then
    FC = zFormOrPictBox.FillColor
    zFormOrPictBox.FillColor = FilColor
  End If
  If FilStyle <> -1 Then
    FS = zFormOrPictBox.FillStyle
    zFormOrPictBox.FillStyle = FilStyle
  End If
  FloodFill = FloodFill2(zFormOrPictBox.hDC, X, Y, BorderColor)
  If UpdateFillColorAndFillStyle = False Then
    If FilColor <> -1 Then zFormOrPictBox.FillColor = FC
    If FilStyle <> -1 Then zFormOrPictBox.FillStyle = FS
  End If
  Exit Function
ErrHnd:
  FloodFill = 0
End Function

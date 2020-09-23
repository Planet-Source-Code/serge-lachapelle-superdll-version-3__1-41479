Attribute VB_Name = "modTrans"
Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Public Enum TransType
  LWA_OPAQUE = 0
  LWA_COLORKEY = 1
  LWA_ALPHA = 2
End Enum

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000

Private Declare Function GetLayeredWindowAttributes Lib "user32.dll" (ByVal HWND As Long, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal HWND As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "GDI32.DLL" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "GDI32.DLL" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateRectRgn Lib "GDI32.DLL" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "GDI32.DLL" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "GDI32.DLL" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "GDI32.DLL" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal HWND As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CombineRgn Lib "GDI32.DLL" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private lngRegion As Long
Dim CurRgn As Long, TempRgn As Long

Public Function isTransparent(zForm As Form) As TransType
  On Local Error Resume Next
  Dim vTrans As Byte, ALPHA As TransType, cKey As Long
  GetLayeredWindowAttributes zForm.HWND, cKey, vTrans, ALPHA
  If Err Then
    isTransparent = -1
  Else
    isTransparent = ALPHA
  End If
End Function

Public Function GetTrans(zForm As Form) As Long
  On Local Error Resume Next
  Dim vTrans As Byte, ALPHA As TransType, cKey As Long
  GetLayeredWindowAttributes zForm.HWND, cKey, vTrans, ALPHA
  If ALPHA = LWA_ALPHA Then
    GetTrans = vTrans
  ElseIf ALPHA = LWA_COLORKEY Then
    GetTrans = cKey
  Else
    GetTrans = -1
  End If
  If Err Then
    GetTrans = -1
  End If
End Function

Public Function FadeTo(zForm As Form, Optional ByVal Final As Byte = 127, Optional ByVal vStep As Byte = 3) As Boolean
  On Local Error Resume Next
  Dim vTrans As Long
  vTrans = isTransparent(zForm)
  If vTrans = LWA_ALPHA Then
    vTrans = GetTrans(zForm)
  Else
    vTrans = -1
  End If
  If vTrans = -1 Then
    If zForm.Visible Then
      FadeTo = FadeOut(zForm, Final, vStep)
    Else
      FadeTo = FadeIn(zForm, Final, vStep)
    End If
  ElseIf vTrans = Final Then
    FadeTo = True
    Exit Function
  ElseIf vTrans > Final Then
    FadeTo = FadeOut(zForm, Final, vStep)
  ElseIf vTrans < Final Then
    FadeTo = FadeIn(zForm, Final, vStep)
  End If
  If Err Then
    FadeTo = False
  End If
End Function

Public Function FadeIn(zForm As Form, Optional ByVal Final As Byte = 255, Optional ByVal vStep As Byte = 3) As Boolean
  On Local Error Resume Next
  Dim vTrans As Long, ZFE As Boolean
  vTrans = isTransparent(zForm)
  If vTrans <> LWA_ALPHA Then SetTrans zForm, 0
  vTrans = GetTrans(zForm)
  If vTrans = -1 Then
    SetTrans zForm, 0
    vTrans = 0
  End If
  If vTrans > Final Then
    FadeIn = False
    Exit Function
  End If
  If zForm.Visible = False Then zForm.Show
  ZFE = zForm.Enabled
  If ZFE = True Then zForm.Enabled = False
  While vTrans < Final
    DoEvents
    vTrans = vTrans + vStep
    If vTrans > Final Then vTrans = Final
    SetTrans zForm, vTrans
  Wend
  If ZFE = True Then zForm.Enabled = True
  If Err Then
    FadeIn = False
  Else
    FadeIn = True
  End If
End Function

Public Function FadeOut(zForm As Form, Optional ByVal Final As Byte = 0, Optional ByVal vStep As Byte = 3) As Boolean
  On Local Error Resume Next
  Dim vTrans As Long, ZFE As Boolean
  vTrans = isTransparent(zForm)
  If vTrans <> LWA_ALPHA Then SetTrans zForm, 255
  vTrans = GetTrans(zForm)
  If vTrans = -1 Then
    SetTrans zForm, 255
    vTrans = 255
  End If
  If vTrans < Final Then
    FadeOut = False
    Exit Function
  End If
  If zForm.Visible = False Then zForm.Show
  ZFE = zForm.Enabled
  If ZFE = True Then zForm.Enabled = False
  While vTrans > Final
    DoEvents
    vTrans = vTrans - vStep
    If vTrans < Final Then vTrans = Final
    SetTrans zForm, vTrans
  Wend
  If ZFE = True Then zForm.Enabled = True
  If Final = 0 Then zForm.Hide
  If Err Then
    FadeOut = False
  Else
    FadeOut = True
  End If
End Function

Public Function SetTrans(zForm As Form, Optional ByVal vTrans As Byte = 127) As Boolean
  On Local Error Resume Next
  Dim Msg As Long
  Msg = GetWindowLong(zForm.HWND, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong zForm.HWND, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes zForm.HWND, 0, vTrans, LWA_ALPHA
  If Err Then
    SetTrans = False
  Else
    SetTrans = True
  End If
End Function

Public Function MakeTrans(zForm As Form, Optional ByVal TransColor As Long = &HFF00FF) As Boolean
  On Local Error Resume Next
  Dim Msg As Long
  Msg = GetWindowLong(zForm.HWND, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong zForm.HWND, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes zForm.HWND, TransColor, 0, LWA_COLORKEY
  If Err Then
    MakeTrans = False
  Else
    MakeTrans = True
  End If
End Function

Public Function MakeOpaque(zForm As Form) As Boolean
  On Local Error Resume Next
  Dim Msg As Long
  Msg = GetWindowLong(zForm.HWND, GWL_EXSTYLE)
  Msg = Msg And Not WS_EX_LAYERED
  SetWindowLong zForm.HWND, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes zForm.HWND, 0, 0, LWA_ALPHA
  If Err Then
    MakeOpaque = False
  Else
    MakeOpaque = True
  End If
End Function

Public Sub FormDrag(TheForm As Object)
  ReleaseCapture
  SendMessage TheForm.HWND, &HA1, 2, 0&
End Sub

Public Sub ChangeMask(zForm As Form, zPict As PictureBox, Optional ByVal lngTransColor As Long = &HFFFFFF)
  Dim lngRetr As Long
  lngRegion& = RegionFromBitmap(zPict, lngTransColor)
  lngRetr& = SetWindowRgn(zForm.HWND, lngRegion&, True)
End Sub

Private Function RegionFromBitmap(picSource As PictureBox, Optional ByVal lngTransColor As Long = &HFFFFFF) As Long
  Const RGN_OR As Long = 2
  Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
  Dim lngRgnFinal As Long, lngRgnTmp As Long
  Dim lngStart As Long, lngRow As Long
  Dim lngCol As Long
  picSource.ScaleMode = 3
  lngHeight& = picSource.ScaleHeight
  lngWidth& = picSource.ScaleWidth
  lngRgnFinal& = CreateRectRgn(0, 0, 0, 0)
  For lngRow& = 0 To lngHeight& - 1
    lngCol& = 0
    Do While lngCol& < lngWidth&
      Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) = lngTransColor&
        lngCol& = lngCol& + 1
      Loop
      If lngCol& < lngWidth& Then
        lngStart& = lngCol&
        Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) <> lngTransColor&
          lngCol& = lngCol& + 1
        Loop
        If lngCol& > lngWidth& Then lngCol& = lngWidth&
        lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
        lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR)
        DeleteObject (lngRgnTmp&)
      End If
    Loop
  Next
  RegionFromBitmap& = lngRgnFinal&
End Function

Public Sub ShapeMe(ByVal Color As Long, ByVal HorizontalScan As Boolean, Optional Name1 As Form = Nothing, Optional Name2 As PictureBox = Nothing)
'Color = the color to convert to transparent (easiest to use RGB function to pass in this value)
'HorizontalScan = scan for transparent lines horizonally or vertically.  Try both during development and pick the fastest one.
'Name1 = a Form name.  If used, leave Name2 blank.
'Name2 = a PictureBox name.  If used, leave Name1 blank.
Const RGN_DIFF As Long = 4

Dim X As Integer, Y As Integer 'points on form
Dim dblHeight As Double, dblWidth As Double 'height and width of object
Dim lngHDC As Long 'the hDC property of the object
Dim booMiddleOfSet As Boolean 'used during the gathering of transparent points
Dim colPoints As Collection 'this will hold all usrPoints
Set colPoints = New Collection
Dim Z As Variant 'used during iteration through collection
Dim dblTransY As Double 'these 3 variables hold each point that will be made transparent
Dim dblTransStartX As Double
Dim dblTransEndX As Double
Dim Name As Object 'will hold the name of the object.  Late-bound and slower, but allows different types (in this case Form or PictureBox)

'check out the name or names passed into the subroutine
If Name1 Is Nothing Xor Name2 Is Nothing Then 'we know there is a name in one of them
    If Name1 Is Nothing Then 'set the name
        Set Name = Name2
    Else
        Set Name = Name1
    End If
Else 'both or none hold valid names
    MsgBox "Must pass in the name of either a Form OR a PictureBox.  TransForm received NONE or BOTH.  Function failed.", vbOKOnly, "ShapeMe Subroutine"
    Exit Sub
End If

'initialization
With Name
    .AutoRedraw = True 'object must have this setting
    .ScaleMode = 3 'object must have this setting
    lngHDC = .hDC 'faster to use a variable; VB help recommends using the property, but I didn't encounter any problems
    If HorizontalScan = True Then 'look for lines of transparency horizontally
        dblHeight = .ScaleHeight 'faster to use a variable
        dblWidth = .ScaleWidth 'faster to use a variable
    Else 'look vertically (note that the names "dblHeight" and "dblWidth" are non-sensical now, but this was an easy way to do this
        dblHeight = .ScaleWidth 'faster to use a variable
        dblWidth = .ScaleHeight 'faster to use a variable
    End If 'HorizontalScan = True
End With
booMiddleOfSet = False

'gather all points that need to be made transparent
For Y = 0 To dblHeight  ' Go through each column of pixels on form
    dblTransY = Y
    For X = 0 To dblWidth  ' Go through each line of pixels on form
        'note that using GetPixel appears to be faster than using VB's Point
        If TypeOf Name Is Form Then 'check to see if this is a form and use GetPixel function which is a little faster
            If GetPixel(lngHDC, X, Y) = Color Then  ' If the pixel's color is the transparency color, record it
                If booMiddleOfSet = False Then
                    dblTransStartX = X
                    dblTransEndX = X
                    booMiddleOfSet = True
                Else
                    dblTransEndX = X
                End If 'booMiddleOfSet = False
            Else
                If booMiddleOfSet Then
                    colPoints.Add Array(dblTransY, dblTransStartX, dblTransEndX)
                    booMiddleOfSet = False
                End If 'booMiddleOfSet = True
            End If 'GetPixel(lngHDC, X, Y) = Color
         ElseIf TypeOf Name Is PictureBox Then 'if a PictureBox then use Point; a little slower but works when GetPixel doesn't
            If Name.Point(X, Y) = Color Then
                If booMiddleOfSet = False Then
                    dblTransStartX = X
                    dblTransEndX = X
                    booMiddleOfSet = True
                Else
                    dblTransEndX = X
                End If 'booMiddleOfSet = False
            Else
                If booMiddleOfSet Then
                    colPoints.Add Array(dblTransY, dblTransStartX, dblTransEndX)
                    booMiddleOfSet = False
                End If 'booMiddleOfSet = True
            End If 'Name.Point(X, Y) = Color
        End If 'TypeOf Name Is Form
        
    Next X
Next Y

CurRgn = CreateRectRgn(0, 0, dblWidth, dblHeight)  ' Create base region which is the current whole window

For Each Z In colPoints 'now make it transparent
    TempRgn = CreateRectRgn(Z(1), Z(0), Z(2) + 1, Z(0) + 1)  ' Create a temporary pixel region for this pixel
    CombineRgn CurRgn, CurRgn, TempRgn, RGN_DIFF  ' Combine temp pixel region with base region using RGN_DIFF to extract the pixel and make it transparent
    DeleteObject (TempRgn)  ' Delete the temporary region and free resources
Next

SetWindowRgn Name.HWND, CurRgn, True  ' Finally set the windows region to the final product
'I do not use DeleteObject on the CurRgn, going with the advice in Dan Appleman's book:
'once set to a window using SetWindowRgn, do not delete the region.
Set colPoints = Nothing
End Sub

Public Sub MakeTransparent(TransForm As Form, Optional ByVal zShapeForm As Boolean = True)
  Const RGN_XOR As Long = 3
  Dim ErrorTest As Double
    'In case there's an error, ignore it
    On Error Resume Next
    Dim Regn As Long
    Dim TmpRegn As Long
    Dim TmpControl As Control
    Dim LinePoints(4) As POINTAPI
    'Since the apis work with pixels, change the scalemode
    'To pixels
    TransForm.ScaleMode = 3
    'You have to have a borderless form, this just makes
    'sure it's borderless
    If TransForm.BorderStyle <> 0 Then MsgBox "Change the borderstyle to 0!", vbCritical, "ACK!": Exit Sub
    'makes everything invisible
    Regn = CreateRectRgn(0, 0, 0, 0)
    'A loop to check every control in the form
    For Each TmpControl In TransForm
        'If the control is a line...
        If TypeOf TmpControl Is Line Then
          If Not zShapeForm Then
            'Checks the slope
            If Abs((TmpControl.Y1 - TmpControl.Y2) / (TmpControl.X1 - TmpControl.X2)) > 1 Then
                'If it's more verticle than horizontal then
                'Set the points
                LinePoints(0).X = TmpControl.X1 - 1
                LinePoints(0).Y = TmpControl.Y1
                LinePoints(1).X = TmpControl.X2 - 1
                LinePoints(1).Y = TmpControl.Y2
                LinePoints(2).X = TmpControl.X2 + 1
                LinePoints(2).Y = TmpControl.Y2
                LinePoints(3).X = TmpControl.X1 + 1
                LinePoints(3).Y = TmpControl.Y1
            Else
                'If it's more horizontal than verticle then
                'Set the points
                LinePoints(0).X = TmpControl.X1
                LinePoints(0).Y = TmpControl.Y1 - 1
                LinePoints(1).X = TmpControl.X2
                LinePoints(1).Y = TmpControl.Y2 - 1
                LinePoints(2).X = TmpControl.X2
                LinePoints(2).Y = TmpControl.Y2 + 1
                LinePoints(3).X = TmpControl.X1
                LinePoints(3).Y = TmpControl.Y1 + 1
            End If
            'Creates the new polygon with the points
            TmpRegn = CreatePolygonRgn(LinePoints(0), 4, 1)
          End If
        'If the control is a shape...
        ElseIf TypeOf TmpControl Is Shape Then
            'An if that checks the type
            If TmpControl.Shape = 0 Then
            'It's a rectangle
                TmpRegn = CreateRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height)
            ElseIf TmpControl.Shape = 1 Then
            'It's a square
                If TmpControl.Width < TmpControl.Height Then
                    TmpRegn = CreateRectRgn(TmpControl.Left, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2, TmpControl.Left + TmpControl.Width, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width)
                Else
                    TmpRegn = CreateRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2, TmpControl.Top, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height, TmpControl.Top + TmpControl.Height)
                End If
            ElseIf TmpControl.Shape = 2 Then
            'It's an oval
                TmpRegn = CreateEllipticRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width + 0.5, TmpControl.Top + TmpControl.Height + 0.5)
            ElseIf TmpControl.Shape = 3 Then
            'It's a circle
                If TmpControl.Width < TmpControl.Height Then
                    TmpRegn = CreateEllipticRgn(TmpControl.Left, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2, TmpControl.Left + TmpControl.Width + 0.5, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width + 0.5)
                Else
                    TmpRegn = CreateEllipticRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2, TmpControl.Top, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height + 0.5, TmpControl.Top + TmpControl.Height + 0.5)
                End If
            ElseIf TmpControl.Shape = 4 Then
            'It's a rounded rectangle
                If TmpControl.Width > TmpControl.Height Then
                    TmpRegn = CreateRoundRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width + 1, TmpControl.Top + TmpControl.Height + 1, TmpControl.Height / 4, TmpControl.Height / 4)
                Else
                    TmpRegn = CreateRoundRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width + 1, TmpControl.Top + TmpControl.Height + 1, TmpControl.Width / 4, TmpControl.Width / 4)
                End If
            ElseIf TmpControl.Shape = 5 Then
            'It's a rounded square
                If TmpControl.Width > TmpControl.Height Then
                    TmpRegn = CreateRoundRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2, TmpControl.Top, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height + 1, TmpControl.Top + TmpControl.Height + 1, TmpControl.Height / 4, TmpControl.Height / 4)
                Else
                    TmpRegn = CreateRoundRectRgn(TmpControl.Left, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2, TmpControl.Left + TmpControl.Width + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width + 1, TmpControl.Width / 4, TmpControl.Width / 4)
                End If
            End If
            'If the control is a shape with a transparent background
            If TmpControl.BackStyle = 0 Then
                'Combines the regions in memory and makes a new one
                CombineRgn Regn, Regn, TmpRegn, RGN_XOR
                If TmpControl.Shape = 0 Then
                'Rectangle
                    TmpRegn = CreateRectRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width - 1, TmpControl.Top + TmpControl.Height - 1)
                ElseIf TmpControl.Shape = 1 Then
                'Square
                    If TmpControl.Width < TmpControl.Height Then
                        TmpRegn = CreateRectRgn(TmpControl.Left + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + 1, TmpControl.Left + TmpControl.Width - 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width - 1)
                    Else
                        TmpRegn = CreateRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + 1, TmpControl.Top + 1, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height - 1, TmpControl.Top + TmpControl.Height - 1)
                    End If
                ElseIf TmpControl.Shape = 2 Then
                'Oval
                    TmpRegn = CreateEllipticRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width - 0.5, TmpControl.Top + TmpControl.Height - 0.5)
                ElseIf TmpControl.Shape = 3 Then
                'Circle
                    If TmpControl.Width < TmpControl.Height Then
                        TmpRegn = CreateEllipticRgn(TmpControl.Left + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + 1, TmpControl.Left + TmpControl.Width - 0.5, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width - 0.5)
                    Else
                        TmpRegn = CreateEllipticRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + 1, TmpControl.Top + 1, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height - 0.5, TmpControl.Top + TmpControl.Height - 0.5)
                    End If
                ElseIf TmpControl.Shape = 4 Then
                'Rounded rectangle
                    If TmpControl.Width > TmpControl.Height Then
                        TmpRegn = CreateRoundRectRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height, TmpControl.Height / 4, TmpControl.Height / 4)
                    Else
                        TmpRegn = CreateRoundRectRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height, TmpControl.Width / 4, TmpControl.Width / 4)
                    End If
                ElseIf TmpControl.Shape = 5 Then
                'Rounded square
                    If TmpControl.Width > TmpControl.Height Then
                        TmpRegn = CreateRoundRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + 1, TmpControl.Top + 1, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height, TmpControl.Top + TmpControl.Height, TmpControl.Height / 4, TmpControl.Height / 4)
                    Else
                        TmpRegn = CreateRoundRectRgn(TmpControl.Left + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + 1, TmpControl.Left + TmpControl.Width, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width, TmpControl.Width / 4, TmpControl.Width / 4)
                    End If
                End If
            End If
        Else
          'Create a rectangular region with its parameters
          If Not zShapeForm Then TmpRegn = CreateRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height)
        End If
            'Checks to make sure that the control has a width
            'or else you'll get some weird results
            ErrorTest = 0
            ErrorTest = TmpControl.Width
            If ErrorTest <> 0 Or TypeOf TmpControl Is Line Then
                'Combines the regions
                CombineRgn Regn, Regn, TmpRegn, RGN_XOR
            End If
    Next TmpControl
    'Make the regions
    SetWindowRgn TransForm.HWND, Regn, True
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DirectShow MCI Media Player"
   ClientHeight    =   3732
   ClientLeft      =   120
   ClientTop       =   804
   ClientWidth     =   4692
   Icon            =   "DMMP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3732
   ScaleWidth      =   4692
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   120
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   24
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":10DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":18EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":2100
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":2912
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":3124
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":3936
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":4148
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":495A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":54EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":607E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":6890
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":70A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":78B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":80C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DMMP.frx":88D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cm1 
      Left            =   1080
      Top             =   120
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   600
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   0
      ScaleHeight     =   492
      ScaleWidth      =   492
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2412
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   3852
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   252
         Left            =   540
         TabIndex        =   6
         Top             =   600
         Width           =   252
         _ExtentX        =   445
         _ExtentY        =   445
         _Version        =   393216
         Value           =   100
         BuddyControl    =   "Label5"
         BuddyDispid     =   196618
         OrigLeft        =   600
         OrigTop         =   300
         OrigRight       =   852
         OrigBottom      =   588
         Increment       =   5
         Max             =   200
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   288
         Left            =   2640
         Picture         =   "DMMP.frx":90EA
         ScaleHeight     =   288
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Picture5 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   288
         Left            =   2160
         Picture         =   "DMMP.frx":98EC
         ScaleHeight     =   288
         ScaleWidth      =   480
         TabIndex        =   3
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   288
         Left            =   1680
         Picture         =   "DMMP.frx":A0EE
         ScaleHeight     =   288
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   288
         Left            =   1200
         Picture         =   "DMMP.frx":AC70
         ScaleHeight     =   288
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   288
         Left            =   720
         Picture         =   "DMMP.frx":B472
         ScaleHeight     =   288
         ScaleWidth      =   480
         TabIndex        =   0
         Top             =   1200
         Width           =   480
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   612
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   3612
         _ExtentX        =   6371
         _ExtentY        =   1080
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Max             =   100
         TickStyle       =   2
         TickFrequency   =   10
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "SPEED"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   156
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         Height          =   252
         Left            =   180
         TabIndex        =   13
         Top             =   600
         Width           =   372
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "frames"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Width           =   1908
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   1908
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   1908
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   120
         Left            =   960
         TabIndex        =   9
         Top             =   120
         Width           =   1908
      End
   End
   Begin VB.Menu mfile 
      Caption         =   "&File"
      Begin VB.Menu mopen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mclose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mpopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mStart 
         Caption         =   "&Auto Start"
         Checked         =   -1  'True
      End
      Begin VB.Menu mrepeat 
         Caption         =   "&Repeat"
         Checked         =   -1  'True
      End
      Begin VB.Menu mme1 
         Caption         =   "-"
      End
      Begin VB.Menu mhalf 
         Caption         =   "Zoom  50%"
      End
      Begin VB.Menu mone 
         Caption         =   "Zoom 100%"
      End
      Begin VB.Menu m1andhalf 
         Caption         =   "Zoom 150%"
         Checked         =   -1  'True
      End
      Begin VB.Menu mtwo 
         Caption         =   "Zoom 200%"
      End
      Begin VB.Menu mme2 
         Caption         =   "-"
      End
      Begin VB.Menu mfull 
         Caption         =   "FULL&SCREEN"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mmf As mciFile
Dim az As Integer
Dim zxc As Long
Dim div As Integer
Dim MULT As Single
Dim zWidth As Integer
Dim zHeight As Integer
Dim zStatus As Variant

Private Sub OpenFile()
  Picture1.Width = 0
  Picture1.Height = 0
  Frame1.Width = 3852
If MciCommand("open", mmf, , Picture1, True) Then
  Slider1.Enabled = True
  Label4.Caption = MciCommand("gettimeformat", mmf)
  div = 1
  While mmf.mLength / div >= 10000
    div = div * 10
  Wend
  zHeight = mmf.mHeight * MULT
  zWidth = mmf.mWidth * MULT
  MoveMCI mmf, 0, 0, zWidth, zHeight
  Slider1.Max = CLng(mmf.mLength / div)
  Slider1.TickFrequency = CLng((mmf.mLength / 10) / div)
  Slider1.LargeChange = CLng((mmf.mLength / 10) / div)
  Slider1.SmallChange = CLng((mmf.mLength / 20) / div)
  Label1.Caption = CLng(mmf.mLength / div)
  Label2.Caption = 0
  If mmf.IsVideo Then
    Picture1.Visible = True
    Picture1.Width = zWidth * Screen.TwipsPerPixelX
    Picture1.Height = zHeight * Screen.TwipsPerPixelY
    Frame1.Top = (zHeight * Screen.TwipsPerPixelY) + 1
    Form1.Height = (zHeight * Screen.TwipsPerPixelY) + Frame1.Height + 750
    mme1.Visible = True
    mhalf.Visible = True
    mone.Visible = True
    m1andhalf.Visible = True
    mtwo.Visible = True
    mme2.Visible = True
    mfull.Visible = True
  Else
    Picture1.Visible = False
    Picture1.Width = 0
    Picture1.Height = 0
    Frame1.Top = 0
    Form1.Height = Frame1.Height + 750
    mme1.Visible = False
    mhalf.Visible = False
    mone.Visible = False
    m1andhalf.Visible = False
    mtwo.Visible = False
    mme2.Visible = False
    mfull.Visible = False
  End If
  If Picture1.Width > Frame1.Width Then
    Form1.Width = Picture1.Width + 60
  Else
    Form1.Width = Frame1.Width + 60
  End If
  Frame1.Width = Form1.Width - 60
  Slider1.Width = Frame1.Width - 240
  Label1.Left = (Frame1.Width - Label1.Width) / 2
  Label2.Left = (Frame1.Width - Label2.Width) / 2
  Label3.Left = (Frame1.Width - Label3.Width) / 2
  Label4.Left = (Frame1.Width - Label4.Width) / 2
  For az = Len(mmf.mfile) To 1 Step -1
    If Mid$(mmf.mfile, az, 1) = "\" Then Exit For
  Next az
  Form1.Caption = Mid$(mmf.mfile, az + 1)
  Picture4.Left = (Frame1.Width - Picture4.Width) / 2
  Picture3.Left = Picture4.Left - Picture3.Width
  Picture5.Left = Picture4.Left + Picture4.Width
  Picture2.Left = Picture3.Left - Picture2.Width
  Picture6.Left = Picture5.Left + Picture5.Width
  Label5.Left = (Label1.Left - (Label5.Width + UpDown1.Width)) / 2
  UpDown1.Left = Label5.Left + Label5.Width + 1
  Label6.Left = ((Label1.Left - Label6.Width) / 2) - 30
  Form1.Left = (Screen.Width - Form1.Width) / 2.25
  Form1.Top = (Screen.Height - Form1.Height) / 2.5
Else
  End
End If
End Sub

Private Sub Form_Load()
  MULT = 1.5
  If KeyExist("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\MCI32") Then
    SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\MCI32", "SuperMCI", "mciqtz32.dll"
  Else
    SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MCI32", "SuperMCI", "mciqtz32.dll"
  End If
  CloseMCI
  If Command$ = "" Then
    cm1.ShowOpen
    If cm1.FileName = "" Then End
    mmf.mfile = cm1.FileName
  Else
    mmf.mfile = Command$
  End If
  OpenFile
  If mStart.Checked Then Picture3_MouseUp 1, 0, 1, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  CloseMCI
  End
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then Form1.PopupMenu mpopup
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then Form1.PopupMenu mpopup
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then Form1.PopupMenu mpopup
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then Form1.PopupMenu mpopup
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then Form1.PopupMenu mpopup
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then Form1.PopupMenu mpopup
End Sub

Private Sub mclose_Click()
  Timer1.Enabled = False
  MciCommand "close", mmf
  Slider1.Value = 0
  Label2.Caption = 0
  Label1.Caption = 0
  Picture1.Visible = False
  Slider1.Enabled = False
  Frame1.Top = 0
  Frame1.Width = 3852
  Form1.Height = Frame1.Height + 750
  Form1.Width = Frame1.Width + 60
  Label1.Left = (Frame1.Width - Label1.Width) / 2
  Label2.Left = (Frame1.Width - Label2.Width) / 2
  Label3.Left = (Frame1.Width - Label3.Width) / 2
  Label4.Left = (Frame1.Width - Label4.Width) / 2
  Picture4.Left = (Frame1.Width - Picture4.Width) / 2
  Picture3.Left = (Picture4.Left - Picture3.Width) - 1
  Picture5.Left = (Picture4.Left + Picture4.Width) + 1
  Picture2.Left = (Picture3.Left - Picture2.Width) - 1
  Picture6.Left = (Picture5.Left + Picture5.Width) + 1
  Slider1.Width = Frame1.Width - 240
  Form1.Caption = "DirectShow MCI Media Player"
  Label4.Caption = ""
  Label5.Left = (Label1.Left - (Label5.Width + UpDown1.Width)) / 2
  UpDown1.Left = Label5.Left + Label5.Width + 1
  Label6.Left = ((Label1.Left - Label6.Width) / 2) - 30
  Form1.Left = (Screen.Width - Form1.Width) / 2.25
  Form1.Top = (Screen.Height - Form1.Height) / 2.5
End Sub

Private Sub mexit_Click()
  CloseMCI
  End
End Sub

Private Sub mfull_Click()
  If mfull.Checked Then
    mfull.Checked = False
    Exit Sub
  Else
    mfull.Checked = True
    If MciCommand("getstatus", mmf) = "playing" Then MciCommand "fullscreen", mmf
    Exit Sub
  End If
End Sub

Private Sub mopen_Click()
  cm1.FileName = ""
  cm1.ShowOpen
  If cm1.FileName <> "" Then
    mclose_Click
    mmf.mfile = cm1.FileName
    OpenFile
    If mStart.Checked Then
      Picture3_MouseUp 1, 1, 1, 1
      If mmf.IsVideo And mfull.Checked Then
        Sleep 100
        MciCommand "fullscreen", mmf
      End If
    End If
  End If
End Sub

Private Sub mrepeat_Click()
  If mrepeat.Checked = True Then
    mrepeat.Checked = False
    Exit Sub
  Else
    mrepeat.Checked = True
    Exit Sub
  End If
End Sub

Private Sub mStart_Click()
  If mStart.Checked = True Then
    mStart.Checked = False
    Exit Sub
  Else
    mStart.Checked = True
    Exit Sub
  End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then Picture2.Picture = ImageList1.ListImages(16).Picture
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    If X >= 0 And Y >= 0 And X <= Picture2.Width And Y <= Picture2.Height Then
      Timer1.Enabled = False
      MciCommand "stop", mmf
      Slider1.Value = 0
      Label2.Caption = 0
    End If
    Picture2.Picture = ImageList1.ListImages(15).Picture
  End If
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then Picture3.Picture = ImageList1.ListImages(12).Picture
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    If X >= 0 And Y >= 0 And X <= Picture3.Width And Y <= Picture3.Height Then
      If Val(Label5.Caption) <> 100 Then MciCommand "setspeed", mmf, Val(Label5.Caption)
      MciCommand "play", mmf
      If mmf.IsVideo And mfull.Checked And Shift <> 1 Then MciCommand "fullscreen", mmf
      Timer1.Enabled = True
    End If
    Picture3.Picture = ImageList1.ListImages(11).Picture
  End If
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then Picture4.Picture = ImageList1.ListImages(10).Picture
End Sub

Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    If X >= 0 And Y >= 0 And X <= Picture4.Width And Y <= Picture4.Height Then
      MciCommand "pause", mmf
      If mmf.IsVideo And mfull.Checked Then
        If MciCommand("getstatus", mmf) = "playing" Then MciCommand "fullscreen", mmf
      End If
      Timer1.Enabled = True
    End If
    Picture4.Picture = ImageList1.ListImages(9).Picture
  End If
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then Picture5.Picture = ImageList1.ListImages(4).Picture
End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    If X >= 0 And Y >= 0 And X <= Picture5.Width And Y <= Picture5.Height Then
      Timer1.Enabled = False
      zxc = MciCommand("stepback", mmf, div)
      Slider1.Value = CLng(zxc / div)
      Label2.Caption = CLng(zxc / div)
    End If
    Picture5.Picture = ImageList1.ListImages(3).Picture
  End If
End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then Picture6.Picture = ImageList1.ListImages(6).Picture
End Sub

Private Sub Picture6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    If X >= 0 And Y >= 0 And X <= Picture6.Width And Y <= Picture6.Height Then
      Timer1.Enabled = False
      zxc = MciCommand("step", mmf, div)
      Slider1.Value = CLng(zxc / div)
      Label2.Caption = CLng(zxc / div)
    End If
    Picture6.Picture = ImageList1.ListImages(5).Picture
  End If
End Sub

Private Sub Slider1_KeyUp(KeyCode As Integer, Shift As Integer)
  SSS
End Sub

Private Sub Slider1_Scroll()
  SSS
End Sub

Private Sub Timer1_Timer()
  zxc = MciCommand("getpos", mmf)
  zStatus = MciCommand("getstatus", mmf)
  Slider1.Value = CLng(zxc / div)
  Label2.Caption = CLng(zxc / div)
  If zxc > 0 And zStatus = "stopped" Then
    MciCommand "resume", mmf
  End If
  If zxc >= mmf.mLength And zStatus <> "paused" Then
    Timer1.Enabled = False
    MciCommand "close", mmf
    Slider1.Value = 0
    Label2.Caption = 0
    If MciCommand("open", mmf, , Picture1, True) Then
      MoveMCI mmf, 0, 0, zWidth, zHeight
    Else
      mclose_Click
      Exit Sub
    End If
    If mrepeat.Checked Then
      Picture3_MouseUp 1, 1, 1, 1
      If mmf.IsVideo And mfull.Checked Then
        Sleep 100
        MciCommand "fullscreen", mmf
      End If
    End If
  End If
End Sub

Private Sub SSS()
  Timer1.Enabled = False
  Label2.Caption = Slider1.Value
  MciCommand "seek", mmf, Slider1.Value * div
  Timer1.Enabled = True
End Sub

Private Sub mhalf_Click()
  MULT = 0.5
  ChangeSize
  mhalf.Checked = True
  mone.Checked = False
  m1andhalf.Checked = False
  mtwo.Checked = False
End Sub

Private Sub mone_Click()
  MULT = 1
  ChangeSize
  mhalf.Checked = False
  mone.Checked = True
  m1andhalf.Checked = False
  mtwo.Checked = False
End Sub

Private Sub m1andhalf_Click()
  MULT = 1.5
  ChangeSize
  mhalf.Checked = False
  mone.Checked = False
  m1andhalf.Checked = True
  mtwo.Checked = False
End Sub

Private Sub mtwo_Click()
  MULT = 2
  ChangeSize
  mhalf.Checked = False
  mone.Checked = False
  m1andhalf.Checked = False
  mtwo.Checked = True
End Sub

Private Sub ChangeSize()
  Frame1.Top = 0
  Frame1.Width = 3852
  Form1.Height = Frame1.Height + 750
  Form1.Width = Frame1.Width + 60
  Label1.Left = (Frame1.Width - Label1.Width) / 2
  Label2.Left = (Frame1.Width - Label2.Width) / 2
  Label3.Left = (Frame1.Width - Label3.Width) / 2
  Label4.Left = (Frame1.Width - Label4.Width) / 2
  Picture4.Left = (Frame1.Width - Picture4.Width) / 2
  Picture3.Left = (Picture4.Left - Picture3.Width) - 1
  Picture5.Left = (Picture4.Left + Picture4.Width) + 1
  Picture2.Left = (Picture3.Left - Picture2.Width) - 1
  Picture6.Left = (Picture5.Left + Picture5.Width) + 1
  Slider1.Width = Frame1.Width - 240
  Form1.Left = (Screen.Width - Form1.Width) / 2
  Form1.Top = (Screen.Height - Form1.Height) / 2
  zHeight = mmf.mHeight * MULT
  zWidth = mmf.mWidth * MULT
  MoveMCI mmf, 0, 0, zWidth, zHeight
  Picture1.Width = zWidth * Screen.TwipsPerPixelX
  Picture1.Height = zHeight * Screen.TwipsPerPixelY
  Frame1.Top = (zHeight * Screen.TwipsPerPixelY) + 1
  Form1.Height = (zHeight * Screen.TwipsPerPixelY) + Frame1.Height + 750
  If Picture1.Width > Frame1.Width Then
    Form1.Width = Picture1.Width + 60
  Else
    Form1.Width = Frame1.Width + 60
  End If
  Frame1.Width = Form1.Width - 60
  Slider1.Width = Frame1.Width - 240
  Label1.Left = (Frame1.Width - Label1.Width) / 2
  Label2.Left = (Frame1.Width - Label2.Width) / 2
  Label3.Left = (Frame1.Width - Label3.Width) / 2
  Label4.Left = (Frame1.Width - Label4.Width) / 2
  Picture4.Left = (Frame1.Width - Picture4.Width) / 2
  Picture3.Left = Picture4.Left - Picture3.Width
  Picture5.Left = Picture4.Left + Picture4.Width
  Picture2.Left = Picture3.Left - Picture2.Width
  Picture6.Left = Picture5.Left + Picture5.Width
  Label5.Left = (Label1.Left - (Label5.Width + UpDown1.Width)) / 2
  UpDown1.Left = Label5.Left + Label5.Width + 1
  Label6.Left = ((Label1.Left - Label6.Width) / 2) - 30
  Form1.Left = (Screen.Width - Form1.Width) / 2.25
  Form1.Top = (Screen.Height - Form1.Height) / 2.5
End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then MciCommand "setspeed", mmf, Val(Label5.Caption)
End Sub

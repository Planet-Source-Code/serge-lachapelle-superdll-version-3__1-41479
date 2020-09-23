VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Super"
   ClientHeight    =   4416
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4572
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4416
   ScaleWidth      =   4572
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Caption         =   "Flash - Me"
      Height          =   492
      Left            =   1320
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Fade To Trans Value"
      Height          =   492
      Left            =   2280
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Set Transparent Value"
      Height          =   492
      Left            =   360
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Make Transparent"
      Height          =   492
      Left            =   2280
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Make Opaque"
      Height          =   492
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Get Transparent Value"
      Height          =   492
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Transparent ?"
      Height          =   492
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TRANSPARENT FORM"
      Height          =   492
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FORM  &&  BUTTONS FROM PICTURES"
      Height          =   492
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SHAPED FORM"
      Height          =   492
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FORM  WITH  MASK"
      Height          =   492
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1932
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   252
      Left            =   3000
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   252
      _ExtentX        =   445
      _ExtentY        =   445
      _Version        =   393216
      Value           =   255
      BuddyControl    =   "Label1"
      BuddyDispid     =   196620
      OrigLeft        =   1800
      OrigTop         =   2760
      OrigRight       =   2052
      OrigBottom      =   3012
      Increment       =   16
      Max             =   255
      Min             =   31
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65537
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ProgressBar pg1 
      Height          =   300
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pg2 
      Height          =   300
      Left            =   360
      TabIndex        =   15
      Top             =   2280
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pg3 
      Height          =   300
      Left            =   360
      TabIndex        =   16
      Top             =   3000
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pg4 
      Height          =   300
      Left            =   360
      TabIndex        =   17
      Top             =   3720
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      Height          =   252
      Left            =   2640
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   360
      TabIndex        =   18
      Top             =   480
      Width           =   3852
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transparent Value :"
      Height          =   192
      Left            =   1080
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   1404
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
MsgBox "form size was too big to upload !", vbInformation, "Super"
'  frmMask.Show
End Sub

Private Sub Command10_Click()
  MsgBox isTransparent(Me)
End Sub

Private Sub Command11_Click()
  Flash Me
End Sub

Private Sub Command2_Click()
MsgBox "form size was too big to upload !", vbInformation, "Super"
'  frmButtons.Show
End Sub

Private Sub Command3_Click()
  frmShapes.Show
End Sub

Private Sub Command4_Click()
  frmTrans.Show
End Sub

Private Sub Command5_Click()
  Dim qwe As String
  If isTransparent(Me) = LWA_COLORKEY Then
    qwe = Hex(GetTrans(Me))
    While Len(qwe) < 6
      qwe = "0" & qwe
    Wend
    MsgBox "BGR : " & qwe
  Else
    MsgBox GetTrans(Me)
  End If
End Sub

Private Sub Command6_Click()
  Me.BackColor = &H8000000F
  SetTrans Me, UpDown1.Value
End Sub

Private Sub Command7_Click()
  MakeOpaque Me
  Me.BackColor = &H8000000F
End Sub

Private Sub Command8_Click()
  Me.BackColor = &H8000000F
  FadeTo Me, UpDown1.Value
End Sub

Private Sub Command9_Click()
  Me.BackColor = &HFF00FF
  MakeTrans Me
End Sub

Private Sub Form_Load()
  FadeIn Me
  Me.Refresh
  Load frmShapes
  cvs pg1, 1, 100
'  Load frmMask
  cvs pg2, 1, 100
  Load frmTrans
  cvs pg3, 1, 100
'  Load frmButtons
  cvs pg4, 1, 100
  Label3.Visible = False
  pg1.Visible = False
  pg2.Visible = False
  pg3.Visible = False
  pg4.Visible = False
  Label1.Visible = True
  Label2.Visible = True
  UpDown1.Visible = True
  Command1.Visible = True
  Command2.Visible = True
  Command3.Visible = True
  Command4.Visible = True
  Command5.Visible = True
  Command6.Visible = True
  Command7.Visible = True
  Command8.Visible = True
  Command9.Visible = True
  Command10.Visible = True
  Command11.Visible = True
  Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.BackColor = &H8000000F
  FadeOut Me
  End
End Sub

Private Sub cvs(zPG As ProgressBar, zFrom As Integer, zTo As Integer)
  Dim t As Integer
  Form1.Refresh
  For t = zFrom To zTo
    zPG.Value = t
    If t Mod 3 = 0 Then
      Sleep2 1
      DoEvents
    End If
  Next t
  Form1.Refresh
End Sub

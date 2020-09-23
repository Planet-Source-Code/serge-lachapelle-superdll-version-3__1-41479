VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmButtons 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "frmButtons"
   ClientHeight    =   5448
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5508
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmButtons.frx":0000
   ScaleHeight     =   5448
   ScaleWidth      =   5508
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   852
      Index           =   5
      Left            =   1920
      Picture         =   "frmButtons.frx":6C042
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   5
      Top             =   1920
      Width           =   852
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   852
      Index           =   4
      Left            =   3840
      Picture         =   "frmButtons.frx":6F084
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   4
      Top             =   1920
      Width           =   852
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   852
      Index           =   3
      Left            =   2520
      Picture         =   "frmButtons.frx":720C6
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   3
      Top             =   3840
      Width           =   852
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   852
      Index           =   2
      Left            =   2640
      Picture         =   "frmButtons.frx":75108
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   852
      Index           =   1
      Left            =   240
      Picture         =   "frmButtons.frx":7814A
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   1
      Top             =   3000
      Width           =   852
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   4320
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmButtons.frx":7B18C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmButtons.frx":7E1DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   852
      Index           =   0
      Left            =   360
      Picture         =   "frmButtons.frx":81230
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   0
      Top             =   600
      Width           =   852
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Left            =   1920
      TabIndex        =   6
      Top             =   1560
      Width           =   852
   End
End
Attribute VB_Name = "frmButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  FadeIn Me
End Sub

Private Sub Form_Load()
  Dim t As Byte
  ShapeMe RGB(255, 255, 255), True, Me
  For t = 0 To Picture1.Count - 1
    ShapeMe RGB(0, 0, 0), True, , Picture1(t)
  Next t
  SetTrans Me, 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FormDrag Me
End Sub

Private Sub Picture1_Click(Index As Integer)
  If Index = 5 Then
    FadeOut Me
  Else
    MsgBox Index, vbOKOnly, "CLICK"
  End If
End Sub

Private Sub picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Set Picture1(Index).Picture = ImageList1.ListImages(2).Picture
End Sub

Private Sub picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Set Picture1(Index).Picture = ImageList1.ListImages(1).Picture
End Sub

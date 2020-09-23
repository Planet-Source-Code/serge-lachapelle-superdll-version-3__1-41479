VERSION 5.00
Begin VB.Form frmTrans 
   BorderStyle     =   0  'None
   Caption         =   "frmTrans"
   ClientHeight    =   3768
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4428
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   ScaleHeight     =   3768
   ScaleWidth      =   4428
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   2040
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   840
      Width           =   1452
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   252
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   852
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   252
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   492
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1452
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF8080&
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   1812
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3612
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   492
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   492
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   1680
      X2              =   3480
      Y1              =   600
      Y2              =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      Height          =   972
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   972
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "move me"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3012
   End
End
Attribute VB_Name = "frmTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  FadeIn Me
End Sub

Private Sub Form_Load()
  MakeTransparent Me, False
  SetTrans Me, 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FormDrag Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FormDrag Me
End Sub

Private Sub Label2_Click()
  FadeOut Me
End Sub

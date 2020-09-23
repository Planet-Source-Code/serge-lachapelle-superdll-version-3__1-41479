VERSION 5.00
Begin VB.Form frmShapes 
   BorderStyle     =   0  'None
   Caption         =   "frmShapes"
   ClientHeight    =   3084
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   ScaleHeight     =   3084
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   2040
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   720
      Width           =   1332
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   252
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   852
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   252
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   492
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1332
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   8
      Height          =   852
      Left            =   4680
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   852
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   8
      X1              =   1200
      X2              =   3000
      Y1              =   360
      Y2              =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   492
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   492
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
      TabIndex        =   0
      Top             =   240
      Width           =   252
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   1812
      Left            =   120
      Top             =   120
      Width           =   3612
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   492
      Left            =   3737
      Shape           =   1  'Square
      Top             =   720
      Width           =   492
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   8
      Height          =   372
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   3612
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   1812
      Left            =   4229
      Shape           =   1  'Square
      Top             =   120
      Width           =   1812
   End
End
Attribute VB_Name = "frmShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  FadeIn Me
End Sub

Private Sub Form_Load()
  MakeTransparent Me, True
  SetTrans Me, 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FormDrag Me
End Sub

Private Sub Label2_Click()
  FadeOut Me
End Sub

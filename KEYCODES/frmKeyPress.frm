VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CODIOS DE TECLA IF NEEDED"
   ClientHeight    =   3150
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblCharAscii 
      AutoSize        =   -1  'True
      Caption         =   "CHAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   9
      Top             =   1305
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CODE"
      Height          =   195
      Left            =   45
      TabIndex        =   8
      Top             =   585
      Width           =   450
   End
   Begin VB.Label LABEL6 
      AutoSize        =   -1  'True
      Caption         =   "ASCII"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   900
      Width           =   405
   End
   Begin VB.Label lblSift 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   195
      Left            =   1170
      TabIndex        =   6
      Top             =   1980
      Width           =   45
   End
   Begin VB.Label label5 
      AutoSize        =   -1  'True
      Caption         =   "Sift Pressed?"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   1980
      Width           =   930
   End
   Begin VB.Label lblAscii 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "KeyAscii"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   4
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "PRESS A KEY PARA VER KeyCode (VB)"
      Height          =   195
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   2940
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   2475
      Shape           =   3  'Circle
      Top             =   540
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   2970
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Last One:"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   2970
      Width           =   690
   End
   Begin VB.Label lblKeyCode 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "KeyCode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   765
      TabIndex        =   0
      Top             =   495
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
lblKeyCode = KeyCode

Shape1.FillColor = RGB(255, 0, 0)
lblSift.Caption = Shift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
lblAscii = KeyAscii
lblCharAscii.Caption = Chr(KeyAscii)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Label3.Caption = lblKeyCode.Caption
lblKeyCode.Caption = "-"
lblAscii = "-"
Shape1.FillColor = RGB(0, 255, 0)
End Sub


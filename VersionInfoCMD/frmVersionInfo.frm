VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   2130
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "frmVersionInfo.frx":0000
      Top             =   540
      Width           =   4605
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   45
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   45
      Width           =   4605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VER INFO"
      Height          =   375
      Left            =   1395
      TabIndex        =   0
      Top             =   2745
      Width           =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2 = Version1.Version1(Text1.Text)
End Sub

Private Sub Form_Load()
Text1.Text = "NOFILE"
End Sub

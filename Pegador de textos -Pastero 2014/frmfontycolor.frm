VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones"
   ClientHeight    =   3495
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   1320
   Icon            =   "frmfontycolor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   88
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScrollFontSize 
      Height          =   1095
      LargeChange     =   2
      Left            =   120
      Max             =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Negro y Blanco"
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Verde y Negro"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Ejemplo 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Ejemplo 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   6
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Ejemplo 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   5
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdListo 
      Cancel          =   -1  'True
      Caption         =   "Listo"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Comic Sans Ms"
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   2400
      Width           =   375
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdColor_Click(Index As Integer)
With Form1.Text1
.Tag = "NoDo"
.SelStart = 0
.SelLength = Len(.Text)
Select Case Index
Case 1 '"Normal"
.Font.Name = "Ms Sans Serif"
.BackColor = RGB(255, 255, 255)
.SelColor = RGB(0, 0, 0)
Case 0 ' "Hacker"
.Font.Name = "Terminal"
.BackColor = RGB(0, 0, 0)
.SelColor = RGB(110, 255, 110)
'.SelStart = Len(.Text)
Case 2 ' "Comic"
.Font.Name = "Comic Sans Ms"
.SelColor = RGB(0, 0, 0)
.BackColor = RGB(255, 255, 250)
End Select

.Tag = ""
End With
End Sub

Private Sub cmdColorHacker_Click()
Me.Tag = "Hacker"
End Sub

Private Sub cmdColorNormal_Click()
Me.Tag = "Normal"
End Sub

Private Sub cmdComic_Click()
Me.Tag = "Comic"
End Sub

Private Sub cmdListo_Click()
Unload Me
End Sub


Private Sub Form_Load()
    'centrar el formulario
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
ponerArribaOptions
VScrollFontSize.Value = Form1.Text1.Font.Size
Label1.Caption = Form1.Text1.Font.Size

End Sub

Private Sub VScrollFontSize_Change()
Dim n As Integer
Form1.Text1.Tag = "NoDo"
Form1.Text1.Font.Size = VScrollFontSize.Value + 1
Label1.Caption = Form1.Text1.Font.Size
Form1.Text1.Tag = ""
cmdColor_Click (0)


End Sub


Function ponerArribaOptions()
Dim hwnd As Long
Dim x, y, minw, minh
Dim ok As Long
'w = Int(Screen.Width / Screen.TwipsPerPixelX)
x = 0: y = 0
minw = Me.ScaleWidth
minh = Me.ScaleHeight
ok = SetWindowPos(Me.hwnd, -1, x, y, minw, minh, 0)
End Function

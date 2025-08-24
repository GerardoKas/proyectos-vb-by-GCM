VERSION 5.00
Begin VB.Form frmColor 
   Caption         =   "Form2"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmColor.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1635
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   285
      Left            =   45
      ScaleHeight     =   225
      ScaleWidth      =   4545
      TabIndex        =   12
      Top             =   1305
      Width           =   4605
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   150
      Index           =   0
      LargeChange     =   16
      Left            =   600
      Max             =   255
      SmallChange     =   6
      TabIndex        =   8
      Top             =   225
      Width           =   2530
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   150
      Index           =   1
      LargeChange     =   16
      Left            =   600
      Max             =   255
      SmallChange     =   6
      TabIndex        =   7
      Top             =   600
      Width           =   2530
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   150
      Index           =   2
      LargeChange     =   16
      Left            =   600
      Max             =   255
      SmallChange     =   6
      TabIndex        =   6
      Top             =   990
      Width           =   2530
   End
   Begin VB.CommandButton cmdTransp 
      Caption         =   "Transp."
      Height          =   330
      Left            =   3870
      TabIndex        =   5
      Top             =   135
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   330
      Left            =   3870
      TabIndex        =   3
      Top             =   900
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3195
      TabIndex        =   2
      Text            =   "00"
      Top             =   855
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3195
      TabIndex        =   1
      Text            =   "00"
      Top             =   495
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   3195
      TabIndex        =   0
      Text            =   "00"
      Top             =   135
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "AZUL"
      Height          =   240
      Index           =   2
      Left            =   45
      TabIndex        =   11
      Top             =   945
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "VERDE"
      Height          =   240
      Index           =   1
      Left            =   45
      TabIndex        =   10
      Top             =   540
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "ROJO"
      Height          =   240
      Index           =   0
      Left            =   45
      TabIndex        =   9
      Top             =   180
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COLOR"
      Height          =   330
      Left            =   3870
      TabIndex        =   4
      Top             =   495
      Width           =   735
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public hexColor As String
Public Red As String
Public Green As String
Public Blue As String
Public Event Seleccionado(hexColor)
Private colors
Public ObjetoStylo As HTMLStyleSheetRule
Public ReglaStylo As Variant

Public Sub loadObj(ByRef obj As Variant)
    ReglaStylo = obj
End Sub

Private Sub Form_Load()
showColors Picture1
End Sub

Private Sub cmdTransp_Click()
hexColor = "transparent"
Unload Me
End Sub

Private Sub Command1_Click()
Red = Text1(0)
Green = Text1(1)
Blue = Text1(2)
hexColor = Red & Green & Blue
addColorHex hexColor
Unload Me
End Sub

Public Function setColorHTML(HexValue As String)
If HexValue = "transparent" Then
    Exit Function
End If
If Left(HexValue, 1) = "#" Then
    HexValue = Right(HexValue, Len(HexValue) - 1)
End If
rojo$ = Mid(HexValue, 1, 2)
verde$ = Mid(HexValue, 3, 2)
azul$ = Mid(HexValue, 5, 2)

HScroll1(0).value = toDec(rojo)
HScroll1(1).value = toDec(verde)
HScroll1(2).value = toDec(azul)
End Function

Public Function setColorDec(r, g, b)
HScroll1(0).value = (r)
HScroll1(1).value = (g)
HScroll1(2).value = (b)
End Function

Private Sub HScroll1_Change(Index As Integer)
Text1(Index) = Hex(HScroll1(Index).value)
If Len(Text1(Index)) = 1 Then Text1(Index) = "0" & Text1(Index).Text
Label1.BackColor = RGB(HScroll1(0).value, HScroll1(1).value, HScroll1(2).value)
Label1.Caption = Text1(0) & Text1(1) & Text1(2)
Select Case Index
    Case 0
        Label2(0).BackColor = RGB(HScroll1(0), 0, 0)
    Case 1
        Label2(1).BackColor = RGB(0, HScroll1(1), 0)
    Case 2
        Label2(2).BackColor = RGB(0, 0, HScroll1(2))
End Select
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim pos As Long
'total = (UBound(colors) + 1)
este = Int(X / 16)
If este > UBound(modColors.colors) Then Exit Sub
Col = modColors.colors(este)
setColorDec r(Col), g(Col), b(Col)
End Sub

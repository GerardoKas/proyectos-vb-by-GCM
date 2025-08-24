VERSION 5.00
Begin VB.Form frmParams 
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   2550
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRel 
      Caption         =   "Mantener"
      Height          =   240
      Left            =   1260
      TabIndex        =   12
      Top             =   1485
      Value           =   1  'Checked
      Width           =   1005
   End
   Begin VB.TextBox txtRelY 
      Height          =   285
      Left            =   675
      TabIndex        =   10
      Text            =   "3"
      Top             =   1485
      Width           =   510
   End
   Begin VB.TextBox txtRelX 
      Height          =   285
      Left            =   45
      TabIndex        =   9
      Text            =   "4"
      Top             =   1485
      Width           =   510
   End
   Begin VB.CheckBox chkWH 
      Caption         =   "Mantener"
      Enabled         =   0   'False
      Height          =   240
      Left            =   1260
      TabIndex        =   8
      Top             =   855
      Width           =   1005
   End
   Begin VB.CheckBox chkXY 
      Caption         =   "Mantener"
      Enabled         =   0   'False
      Height          =   240
      Left            =   1260
      TabIndex        =   7
      Top             =   315
      Width           =   1005
   End
   Begin VB.TextBox txtHeight 
      Enabled         =   0   'False
      Height          =   285
      Left            =   675
      TabIndex        =   3
      Text            =   "Height"
      Top             =   855
      Width           =   510
   End
   Begin VB.TextBox txtWidth 
      Enabled         =   0   'False
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Text            =   "Width"
      Top             =   855
      Width           =   510
   End
   Begin VB.TextBox txtTop 
      Enabled         =   0   'False
      Height          =   285
      Left            =   675
      TabIndex        =   1
      Text            =   "LEFT"
      Top             =   270
      Width           =   510
   End
   Begin VB.TextBox txtLeft 
      Enabled         =   0   'False
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Text            =   "TOP"
      Top             =   270
      Width           =   510
   End
   Begin VB.Label lblRelActual 
      AutoSize        =   -1  'True
      Caption         =   "REL ACTUAL"
      Height          =   195
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   585
      TabIndex        =   11
      Top             =   1530
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Relacion de Tamaño Ancho/Alto"
      Height          =   195
      Left            =   45
      TabIndex        =   6
      Top             =   1215
      Width           =   2340
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ancho y Altura - W,H"
      Enabled         =   0   'False
      Height          =   195
      Left            =   45
      TabIndex        =   5
      Top             =   630
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Coordenadas Origen X,Y"
      Enabled         =   0   'False
      Height          =   195
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   1755
   End
End
Attribute VB_Name = "frmParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mygcm As GCM_Picture

Public Sub ColocarDatos(ctlGcm As GCM_Picture)
Set mygcm = ctlGcm
If ctlGcm.marcaWidth = 0 Or ctlGcm.marcaHeight = 0 Then Exit Sub
setXY ctlGcm.marcaLeft, ctlGcm.marcaTop
setWH ctlGcm.marcaWidth, ctlGcm.marcaHeight
setRel ctlGcm.marcaWidth, ctlGcm.marcaHeight
End Sub

Sub setXY(left, top)
txtLeft = left
txtTop = top
End Sub

Sub setWH(w, h)
txtWidth = w
txtHeight = h
End Sub

Sub setRel(w, h)
Dim newW As Long, newH As Long
If chkRel.Value = 0 Then Exit Sub
arel = w / h
perfrel = CSng(txtRelX) / CSng(txtRelY)
'si rel >1 then w > h
'si rel <1 then w < h
Debug.Print "DIFERENCIA DE REL: " & Abs(Abs(arel) - Abs(perfrel))

If Abs(Abs(arel) - Abs(perfrel)) < 0.1 Then Exit Sub
If arel > 1 Then
    newW = w
    newH = (w * txtRelY) / txtRelX
ElseIf arel < 1 Or arel = 1 Then
    newH = h
    newW = (h * txtRelX) / txtRelY
End If
newrel = newW / newH
lblRelActual = Round(newrel, 6)
mygcm.ChangeMark mygcm.marcaLeft, mygcm.marcaTop, newW, newH
End Sub
'Sub setRel(w, h)
'cont = 1
'r = w / h
'res = cont * r
'Do While Round(res, 3) <> res
'cont = cont + 1
'res = cont * r
'If cont > 200 Then
'    Exit Do
'End If
'Loop
''MsgBox "rel=" & res & ":" & cont
'txtRelX = res
'txtRelY = cont
'End Sub
Private Sub Form_Load()
calcRel
End Sub

Private Sub txtRelX_Change()
calcRel
End Sub

Private Sub txtRelY_Change()
calcRel
End Sub


Sub calcRel()
setTxtDots txtRelX
setTxtDots txtRelY
On Error Resume Next
lblRelActual = Round(CSng(txtRelX) / CSng(txtRelY), 6)
End Sub

Sub setTxtDots(text As TextBox)
If InStr(1, text, ".") Then
text = Replace(text, ".", ",")
text.SelStart = InStr(1, text, ",")
End If
End Sub

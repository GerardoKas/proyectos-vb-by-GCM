VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pastero Search&Replace"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
   Begin VB.CommandButton Command1 
      Caption         =   "Lista RegExps"
      Height          =   375
      Left            =   45
      TabIndex        =   13
      Top             =   1845
      Width           =   1365
   End
   Begin VB.CommandButton cmdRegExpReplace 
      Caption         =   "RegExpReplace"
      Height          =   330
      Left            =   1485
      TabIndex        =   12
      Top             =   2520
      Width           =   1410
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   0
      TabIndex        =   11
      Top             =   2925
      Width           =   4605
   End
   Begin VB.CommandButton cmdRegExpFind 
      Caption         =   "RegExpFind"
      Height          =   330
      Left            =   45
      TabIndex        =   10
      Top             =   2520
      Width           =   1365
   End
   Begin VB.CommandButton cmdUndoReplace 
      Caption         =   "Undo Replace"
      Height          =   330
      Left            =   3015
      TabIndex        =   9
      Top             =   2520
      Width           =   1545
   End
   Begin VB.CommandButton cmdCountAll 
      Caption         =   "COUNT ALL"
      Height          =   330
      Left            =   3330
      TabIndex        =   8
      Top             =   945
      Width           =   1230
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Reemplazar Todas"
      Height          =   555
      Left            =   3330
      TabIndex        =   7
      Top             =   1800
      Width           =   1230
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Reemplazar"
      Height          =   375
      Left            =   3330
      TabIndex        =   6
      Top             =   1395
      Width           =   1230
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "SIGUIENTE"
      Height          =   330
      Left            =   3330
      TabIndex        =   5
      Top             =   585
      Width           =   1230
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "BUSCAR"
      Height          =   330
      Left            =   3330
      TabIndex        =   4
      Top             =   225
      Width           =   1230
   End
   Begin VB.TextBox Text2 
      Height          =   645
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      ToolTipText     =   "Text Replace"
      Top             =   945
      Width           =   3165
   End
   Begin VB.TextBox TxtBusca 
      Height          =   465
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Text Find"
      Top             =   225
      Width           =   3165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Buscar"
      Height          =   195
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Reemplazo"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   795
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txt As RichTextBox
Dim posActual As Long
Dim backupText As String
Const septor = "-"

Private Sub cmdFind_Click()
posActual = buscar(Text1.Text)
If posActual < 1 Then
MsgBox "Not Found"
End If
End Sub

Private Sub cmdFindNext_Click()
posActual = buscar(Text1.Text, posActual + 2)
If posActual < 1 Then
MsgBox "Not Found"
End If
End Sub


Private Sub cmdRegExpFind_Click()
Dim rex As RegExp, match As match
Set rex = New RegExp
habilitar False
List1.Clear
rex.Global = True
'rex.MultiLine = False
rex.IgnoreCase = True
rex.Pattern = Text1
Set matches = rex.Execute(txt.Text)
For Each match In matches
    List1.AddItem match.FirstIndex & septor & match.Value
Next
habilitar True
End Sub

Private Sub cmdRegExpReplace_Click()
Dim rex As RegExp, match As match
Set rex = New RegExp
habilitar False
List1.Clear
rex.Global = True
'rex.MultiLine = True
rex.IgnoreCase = True
rex.Pattern = Text1.Text
If rex.Test(txt) Then
backupText = txt.Text
txt.Text = rex.Replace(txt.Text, Text2.Text)
End If
habilitar True
End Sub

Private Sub cmdReplace_Click()
'reemplazar
found = buscar(TxtBusca, posActual)
If found > 1 Then
txt.SelText = Text2
Else
    MsgBox "Not Found"
End If
End Sub

Private Sub cmdReplaceAll_Click()
backupText = txt.Text
found = buscar(TxtBusca, posActual)
If found > 1 Then
txt.SelText = Text2
Else
    MsgBox "Not Found"
End If
End Sub

Private Sub cmdCountAll_Click()
Dim count As Integer
count = 0
Do
anterior = buscar(TxtBusca.Text, anterior + 2)
If anterior >= 1 Then
count = count + 1
End If
Loop While anterior >= 1
MsgBox "Total de:" & vbCrLf & Text1.Text & vbCrLf & "--> " & count
End Sub


Private Sub cmdUndoReplace_Click()
txt.Text = backupText
End Sub

Private Sub Command1_Click()
frmListaRegs.Show 0, Me
End Sub

Private Sub Form_Load()
Set txt = Form1.Text1

ok = SetWindowPos(hwnd, -1, 10 + Form2.ScaleLeft + Form1.ScaleWidth, Form2.ScaleTop, Form2.ScaleWidth + 10, Form2.ScaleHeight + 10, 0)

End Sub

Public Function buscar(cadena As String, Optional start As Long = 1) As Long
Dim fnd As Boolean, fndPos As Long, st As Long, i As Long
If Len(txt) > (32000) Then
    MsgBox "Mejor No hacerlo, tardaria mucho. Prueba Con RegeExp Find"
    Exit Function
End If
habilitar False
st = 1
fndPos = start
If start < 1 Then start = 1
For i = start To Len(txt)
    If StrComp(Mid(cadena, st, 1), Mid(txt, i, 1), vbTextCompare) = 0 Then
        If st = 1 Then
           fndPos = i - 1
        End If
       If st = Len(cadena) Then
            'MsgBox "found"
            txt.SelStart = fndPos
            txt.SelLength = Len(cadena)
            txt.SetFocus
            buscar = fndPos
            habilitar True
            Exit Function
        End If
        st = st + 1
        fnd = True
    Else
        fnd = False
        st = 1
    End If
Next
buscar = -1
habilitar True
End Function

Private Sub List1_DblClick()
cad = List1.List(List1.ListIndex)
sep = InStr(1, cad, septor, vbBinaryCompare)
indexpos = Left(cad, sep - 1)
txtcad = Right(cad, Len(cad) - sep)
txt.SelStart = indexpos
txt.SelLength = Len(txtcad)
End Sub

Sub habilitar(si As Boolean)
If si Then
Me.MousePointer = vbDefault
Me.Enabled = True
Else
    Me.MousePointer = vbHourglass
    Me.Enabled = False
End If
End Sub


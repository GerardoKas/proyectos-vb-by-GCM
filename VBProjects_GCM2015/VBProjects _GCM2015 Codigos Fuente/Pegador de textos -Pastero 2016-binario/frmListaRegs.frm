VERSION 5.00
Begin VB.Form frmListaRegs 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Añadir Actual"
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   90
      Width           =   1365
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Text            =   "Nombre Para La Expresion"
      Top             =   90
      Width           =   1995
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   4605
   End
End
Attribute VB_Name = "frmListaRegs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const myList = "RegExpList.dat"
Dim oFind As TextBox
Dim oReplace As TextBox
Const PutTitle = "Titulo de la Expresion"

Private Sub Command1_Click()
If oFind.Text = "" Then MsgBox "No hay ninguna Expresion de Busqueda para añadir": Exit Sub

List1.AddItem txtTitle & "***" & oFind.Text & "***" & oReplace.Text
End Sub

Private Sub Form_Load()
Set oFind = Form2.TxtBusca
Set oReplace = Form2.txtReplace
txtTitle.Text = PutTitle
CargarLista
End Sub

Sub CargarLista()
If Dir$(myList, vbNormal Or vbReadOnly) = "" Then
    MsgBox "no Hay Items"
    Exit Sub
End If
f = FreeFile
Open myList For Input As #f
Do While Not EOF(f)
    Line Input #f, linea
    List1.AddItem linea
Loop
Close #f
End Sub

Private Sub Form_Unload(Cancel As Integer)
f = FreeFile
Open myList For Output As #f
For i = 0 To List1.ListCount - 1
    Print #f, List1.List(i)
Next
Close #f
End Sub

Private Sub List1_DblClick()
pts = Split(List1.List(List1.ListIndex), "***")
oFind.Text = pts(1)
oReplace.Text = pts(2)
End Sub

Private Sub txtTitle_Click()
If txtTitle.Text = PutTitle Then
    txtTitle.Text = ""
End If
End Sub

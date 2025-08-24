VERSION 5.00
Begin VB.Form frmEfectFiles 
   Caption         =   "Form2"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2865
   LinkTopic       =   "Form2"
   ScaleHeight     =   1590
   ScaleWidth      =   2865
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   1275
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Cargar Lista"
      Height          =   285
      Left            =   1395
      TabIndex        =   2
      Top             =   45
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar Nuevo"
      Height          =   285
      Left            =   1395
      TabIndex        =   1
      Top             =   1215
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1215
      Width           =   1230
   End
End
Attribute VB_Name = "frmEfectFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public savefile As Boolean
Public loadfile As Boolean
Public theLoadFile As String
Private Sub cmdLoad_Click()
theLoadFile = File1.filename
cargarEfectos theLoadFile, Form1.lstEfectos
Unload Me
End Sub

Private Sub cmdSave_Click()
grabarEfectos Text1.Text, Form1.lstEfectos
MsgBox "grabado como " & Text1 & ".DATA"
Unload Me
End Sub

Private Sub Form_Load()
File1.Pattern = "*.DATA"
File1.Path = App.Path
If loadfile = True Then
    cmdLoad.Enabled = True
    cmdSave.Enabled = False
Else
    cmdLoad.Enabled = False
    cmdSave.Enabled = True
End If
End Sub


Sub grabarEfectos(nombre, list As ListBox)
f = FreeFile
Open App.Path & "\" & nombre & ".DATA" For Output As #f
For i = 0 To list.ListCount - 1
    Print #f, list.list(i)
Next
Close #f
End Sub

Sub cargarEfectos(file, list As ListBox)
f = FreeFile
Open App.Path & "\" & file For Input As #f
Do While Not EOF(f)
    Line Input #f, linea
    list.AddItem linea
Loop
End Sub

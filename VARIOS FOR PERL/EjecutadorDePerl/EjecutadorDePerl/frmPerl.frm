VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "frmPerl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "New Input Always"
      Height          =   240
      Left            =   4950
      TabIndex        =   10
      Top             =   1950
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   4950
      TabIndex        =   4
      Top             =   0
      Width           =   1575
      Begin VB.CommandButton cmdOutput 
         Caption         =   "Save Output"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1575
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Resetear"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1230
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditor 
         Caption         =   "Editar con ..."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   555
         Width           =   1335
      End
      Begin VB.CommandButton cmdejecutar 
         BackColor       =   &H000080FF&
         Caption         =   "Ejecutar"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   900
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "Cargar Script"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1335
      End
   End
   Begin VB.TextBox txtPerl 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "La Ruta del Perl"
      Top             =   120
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Text            =   "Perl Params"
      ToolTipText     =   "Parametros Previos Para el Perl"
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "Elige un script de perl para ejecutar"
      ToolTipText     =   "El Script que se va a ejecutar"
      Top             =   480
      Width           =   4695
   End
   Begin VB.TextBox txtParams 
      Height          =   1350
      Left            =   120
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      ToolTipText     =   "La lista de parametros o archivos enviados al Script"
      Top             =   840
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Proyecto = "ElPerl"
Const ElPerl = "c:\perl\bin\perl.exe"
Const Seccion = "general"

Private directorio As String
Private nuevoPerl As String
Private Params(5, 1) As String
Private hayFicheros As Boolean
Private Editor As String
Private txtOutput As String

Private Sub Command3_Click()
On Error Resume Next

DeleteSetting Proyecto, Seccion, "path"
DeleteSetting Proyecto, Seccion, "lastfile"
DeleteSetting Proyecto, Seccion, "editor"
Combo1.Text = ""
txtOutput = ""
txtParams = ""
Combo1.Text = ""

End Sub

Private Sub Form_Load()
txtPerl.Text = ElPerl
Params(0, 0) = "-c"
Params(0, 1) = "Check Syntax"
Params(1, 0) = "-w"
Params(1, 1) = "Warnings"
Params(2, 0) = "-e"
Params(2, 1) = "Evaluar Cadena"
Params(3, 0) = "-d"
Params(3, 1) = "Under Debugger"
CargarComboParams
directorio = GetSetting(Proyecto, Seccion, "path")
nuevoPerl = GetSetting(Proyecto, Seccion, "lastFile")
txtParams.Text = GetSetting(Proyecto, Seccion, "params")
Editor = GetSetting(Proyecto, Seccion, "editor")

nuevoItem (nuevoPerl)

End Sub

Private Sub Form_Resize()
txtParams.Height = Me.ScaleHeight - Combo1.Height - Combo1.Top - 100

End Sub

Private Sub Form_Unload(Cancel As Integer)
If directorio <> "" Then
SaveSetting Proyecto, Seccion, "path", directorio
End If
If nuevoPerl <> "" Then
SaveSetting Proyecto, Seccion, "lastFile", nuevoPerl
End If
If Editor <> "" Then
SaveSetting Proyecto, Seccion, "editor", Editor
End If
If txtParams.Text <> "" Then
SaveSetting Proyecto, Seccion, "params", txtParams.Text
End If

End Sub

''''''''''''''''
Private Sub Combo1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.Files.Count >= 1 Then
nuevoItem (Data.Files(1))
End If
End Sub

Private Sub Command1_Click()
cmdlg.Filter = "Programas Perl(*.pl *.pm)|*.pl;*.pm|Todos los archivos *.*|*.*"
cmdlg.DialogTitle = "Abrir Script Perl"
If directorio <> "" Then cmdlg.InitDir = directorio
cmdlg.ShowOpen
If cmdlg.filename = "" Then
Exit Sub
End If
nuevoItem (cmdlg.filename)
End Sub

Private Sub nuevoItem(fichero As String)
nuevoPerl = fichero
Combo1.Text = fichero
directorio = basepath(fichero)
For i = 0 To Combo1.ListCount - 1
If fichero = Combo1.List(i) Then Exit Sub
Next
Combo1.AddItem fichero
End Sub
''''''''''''''
Private Sub cmdEditor_Click()
checkEditor
Shell Editor & " " & nuevoPerl, vbNormalFocus
End Sub

Private Sub cmdOutput_Click()
Dim cmd As String

cmdlg.Filter = "Texto *.txt|*.txt|Todos los archivos *.*|*.*"
cmdlg.DialogTitle = "Guardar La Salida del Programa"
cmdlg.filename = txtOutput

If directorio <> "" Then cmdlg.InitDir = directorio
cmdlg.ShowOpen
If cmdlg.filename = "" Then
Exit Sub
End If
txtOutput = cmdlg.filename
cmd = CrearComando
cmd = cmd & " > " & Chr(34) & Chr(34)
bat = grabarBat(cmd)
Shell bat, vbNormalFocus
DoEvents
checkEditor
Shell Editor & " " & Chr(34) & txtOutput & Chr(34)
End Sub

Sub CargarComboParams()
For i = 0 To UBound(Params, 1)
Combo2.AddItem Params(i, 0) & " (" & Params(i, 1) & ")"
Next
Combo2.ListIndex = 1
End Sub

Function CrearComando()
Dim linea As String
linea = ElPerl & " " & getPerlParam() & _
        " " & getPerlScript() & " " & getFicheros()
CrearComando = linea
End Function

Private Sub cmdejecutar_Click()
Dim cmd As String
If getPerlScript = "" Then Exit Sub

cmd = CrearComando
bat = grabarBat(cmd)
Shell bat, vbMaximizedFocus
End Sub

Function grabarBat(cadena As String)
savefile = nuevoPerl & ".BAT"
f = FreeFile
Open savefile For Output As #f
Print #f, "@echo off"
Print #f, "title " & basename(nuevoPerl)
Print #f, "echo " & nuevoPerl
Print #f, "echo."
Print #f, cadena
Print #f, "echo."
Print #f, "pause"
Close #f
grabarBat = savefile
End Function

Function getPerlScript()
On Error Resume Next
prog = Combo1.Text
If IsError(FileLen(prog)) Then
getPerlScript = ""
nuevoPerl = ""
MsgBox "El Fichero de script no existe " & vbCrLf & Combo1.Text & vbCrLf & "Elige un script que exista", vbCritical
Else
getPerlScript = Chr(34) & prog & Chr(34)
nuevoPerl = prog
End If
End Function

Function getPerlParam()
txt = Combo2.Text
part = Split(Combo2.Text, "(")
If part(0) = "" Or UBound(part) = 0 Then
getPerlParam = ""
Exit Function
End If
getPerlParam = part(0)
End Function

Function getFicheros()
If hayFicheros Then
getFicheros = txtParams.Text
End If
End Function

Private Sub txtParams_Change()
If (txtParams.Text <> "") Then
hayFicheros = True
End If
End Sub

Private Sub txtParams_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
texto = ""
For i = 1 To Data.Files.Count
    texto = texto & Chr(34) & Data.Files.Item(i) & Chr(34) & " "
Next
txtParams.Text = texto
End Sub

Function basepath(filename As String)
pos = InStrRev(filename, "\")
basepath = Left(filename, pos)
End Function


Function basename(filename As String)
Dim pos As Integer
pos = InStrRev(filename, "\")
basename = Right(filename, Len(filename) - pos)
End Function


Public Sub checkEditor()
If Editor = "" Then
If MsgBox("No hay editor definido. Elige Si Para elegir un editor. No para abrirlo con el notepad", vbYesNo, "Elegir un editor") = vbNo Then
Editor = "notepad.exe"
Else
cmdlg.Filter = "Ejecutables |*.exe;*.com"
cmdlg.DialogTitle = "Elegir un editor"
cmdlg.ShowOpen
If cmdlg.filename = "" Then Exit Sub
Editor = cmdlg.filename
End If
If Editor = "" Then Editor = "notepad.exe"

End If
End Sub

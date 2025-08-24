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
   Begin VB.TextBox txtParams 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   855
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      ToolTipText     =   "Drop Files For Input (Any File or Folder)"
      Top             =   2880
      Width           =   3795
   End
   Begin VB.TextBox txtOpcion 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Text            =   "OPCION"
      Top             =   360
      Width           =   600
   End
   Begin VB.TextBox txtFileInput 
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   855
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "NoFile"
      ToolTipText     =   "Drop PerlFile For Exec"
      Top             =   2565
      Width           =   2670
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PerlFile"
      Height          =   195
      Left            =   270
      TabIndex        =   5
      Top             =   2610
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Drop Input"
      Height          =   195
      Left            =   45
      TabIndex        =   4
      Top             =   2925
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MENUS "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim perlfile As String

Private Sub Form_Activate()
txtOpcion.SetFocus
txtOpcion = ""
End Sub

Private Sub Form_Load()
txtFileInput = FileInput(Command$)
If txtFileInput = "" Then Exit Sub

perlfile = txtFileInput
leerMenus
End Sub

Function FileInput(cadena)
entrada = Trim(cadena)
If Left(entrada, 1) = """" Then
    entrada = Mid(entrada, 2, Len(entrada) - 2)
End If
FileInput = entrada
End Function

Function cmdPerlNormal()
proceso "perl " & perlfile
End Function

Sub proceso(cmdline As String)
Set WshShell = CreateObject("WScript.Shell")
modo = 1
retorno = WshShell.Run("cmd  /C " & cmdline & " & pause", modo, True)
Debug.Print "RET " & retorno
End Sub

Function Menus(tipo As String, Optional numMenu As Integer) As Variant
Select Case numMenu
Case 1
t = "Normal"
f = "perl %FILE% %PARAMS%"
Case 2
t = "Warnings"
f = "perl -W %FILE% %PARAMS%"
Case 3
t = "Check Syntax"
f = "perl -c %FILE%"
End Select
Select Case tipo
Case "Total"
Menus = 4
Case "Item"
Menus = t
Case "Funcion"
Menus = f
End Select
End Function

Sub leerMenus()
Dim i As Integer, numM As Integer, c As String
numM = Menus("Total")
c = "ELIGE OPCION" & vbCrLf
For i = 1 To numM
c = c & i & ". " & Menus("Item", i) & " (" & Menus("Funcion", i) & ") " & vbCrLf
Next
Label1.Caption = c
txtOpcion.Top = Label1.Height + Label1.Top + 40
txtOpcion.Left = Label1.Left
End Sub

Private Sub txtParams_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
t = ""
For i = 1 To Data.Files.Count
t = t & " " & """" & FileInput(Data.Files.Item(i)) & """"
Next
txtParams.Text = t
txtOpcion.SetFocus
End Sub

Private Sub txtFileInput_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.Files.Count > 0 Then
t = Data.Files(1)
If InStrRev(t, "pl") > (Len(t) - 3) Then
txtFileInput = FileInput(t)
perlfile = txtFileInput
Else
Me.SetFocus
MsgBox "Eso no es un PerlScript. "
End If
End If
txtOpcion.SetFocus
End Sub

Private Sub txtOpcion_Change()
o = Val(txtOpcion.Text)
If o < 1 Or o > 9 Then
txtOpcion = ""
Exit Sub
End If
cmdline$ = preproceso(Menus("Funcion", CInt(o)))
proceso cmdline
txtOpcion = ""
End Sub

Function preproceso(texto As String)
preproceso = Replace(texto, "%FILE%", """" & perlfile & """")
preproceso = Replace(preproceso, "%PARAMS%", txtParams)

End Function

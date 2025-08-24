VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   6000
   Icon            =   "frmlarga.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Text1 
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      AutoVerbMenu    =   -1  'True
      OLEDropMode     =   1
      TextRTF         =   $"frmlarga.frx":038A
   End
   Begin VB.TextBox txtComunica 
      Height          =   780
      Left            =   4140
      LinkTopic       =   "CopiarRuta|Pasar"
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Contenedor"
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   5940
      Begin VB.ComboBox cmbCOMANDOS 
         Height          =   315
         ItemData        =   "frmlarga.frx":046B
         Left            =   720
         List            =   "frmlarga.frx":0481
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   0
         Width           =   1545
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Copy"
         Height          =   315
         Left            =   2280
         MaskColor       =   &H0000C0C0&
         MousePointer    =   13  'Arrow and Hourglass
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reg"
         Height          =   300
         Left            =   5445
         TabIndex        =   5
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Updir"
         Height          =   315
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   690
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "Clear"
         Height          =   315
         Left            =   4800
         TabIndex        =   3
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "ToFile"
         Height          =   315
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.TextBox Text1ld 
      Height          =   1215
      Left            =   720
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "frmlarga.frx":04B8
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const myCaption = "Drag & Drop Files - GCM2005 (CopiarRuta)"
Private ficheros As String
Private colfiles As Collection
Private using As Collection
Private temp As Collection
Private barra As String


Private Sub cmbCOMANDOS_Click()
Select Case cmbCOMANDOS.List(cmbCOMANDOS.ListIndex)
Case "LINKS"
makelinks
Case "IMG"
makeImg
Case "JSCRIPT"
makeJs
Case "PERL BAT"
makePerl
Case "RESET"
simple
Case "AUDIO LIST"
'remplazar mp3
End Select
End Sub

Private Sub cmdBorrar_Click()
Set colfiles = New Collection
Set using = New Collection
Set temp = New Collection
Me.Caption = myCaption
simple
End Sub

Private Sub cmdSave_Click()
f = FreeFile()
If using.Count = 0 Then Exit Sub
inputtext = InputBox("Nombre de Archivo: " & basepath(using(1)), "Guardar Fichero con nombre", basepath(using(1)) & basename(using(1)) & ".TXT")
If inputtext = "" Then
    MsgBox "No se ha guardado"
    Exit Sub
End If
SaveFile = inputtext & Date$ '"c:\txt-15.txt"   'basename(inputtext)
Path = basepath(colfiles(1))
Open SaveFile For Output As #f
Print #f, Text1.Text
Close #f
End Sub

Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText Text1.Text, vbCFText
DoEvents
Beep
Unload Me
End Sub

Private Sub Command4_Click()
reduxPath
End Sub

Private Sub Form_Load()
barra = "\"
Me.Caption = myCaption
Set colfiles = New Collection
Set using = New Collection
Set temp = New Collection

'If 1 Then
If App.PrevInstance = True Then
    'ahora seremos cliente
    Me.LinkMode = 0
    Me.LinkTopic = ""
    txtComunica.LinkTopic = "CopiaRuta|DDE"
    txtComunica.LinkItem = "txtComunica"
    txtComunica.LinkMode = vbLinkManual
    txtComunica = Command$()
    txtComunica.LinkPoke
    Unload Me
    End
Else
    'seremos servidor
    'MsgBox "servidor"
    Me.LinkMode = 1
    Me.LinkTopic = "DDE"
End If
If Command$() <> "" Then
Set Files = getFiles(Command$())
For Each i In Files
    Poner i
Next
End If
simple
Me.Move 10, 10
cmbCOMANDOS.ListIndex = 0
End Sub

Private Sub Command2_Click()
Dim regfile As String
strreg = "reg add HKCR\*\shell\CopiarRuta\command\ /ve /t REG_SZ /d ""\""-APP-\"" %%1 "" /f"
Debug.Print strreg
strreg = Replace(strreg, "-APP-", App.Path & "\" & App.EXEName & ".exe")
f = FreeFile()

regfile = "c:\regRutero.bat"
Open regfile For Output As #f
Print #f, "@echo off"
Print #f, "echo Registrando Programa"
Print #f, "reg add HKCR\*\shell\CopiarRuta\command\"
Print #f, strreg
Print #f, "echo."
Print #f, "echo Se ha agregado la informacion de registro"
Print #f, "echo."
Print #f, "pause"
Print #f, "del " & regfile
Close #f
Shell regfile, vbNormalFocus
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
Text1.Move 0, 0, Me.ScaleWidth, (Me.ScaleHeight - Frame1.Height)
Frame1.Top = Text1.Height
End Sub

Private Sub RichTextBox1_Change()


End Sub
'Private Sub Text1_OLECompleteDrag(Effect As Long)
'ddddtyn fncnan las tclas d l tatl
'ficheros directorios listas
'End Sub


'Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'For i = 1 To Data.Files.Count
'Poner Data.Files.Item(i)
'Next
'simple
'End Sub

Sub Poner(texto)
colfiles.Add texto
End Sub

Sub simple()
t = ""
Set using = New Collection
CopiarOriginal colfiles, using
For Each i In using
t = t & """" & i & """" & " "
Next
Text1.Text = t
If colfiles.Count > 0 Then
Me.Caption = colfiles.Count & " Items"
Else
Me.Caption = myCaption
End If
End Sub

Sub CopiarOriginal(orig As Collection, ByRef use As Collection)
For Each i In orig
use.Add i
Next
End Sub

Sub reduxPath()
Dim parte As String
Set temp = New Collection
For i = 1 To using.Count
obj = using(i)
p = InStr(1, obj, barra)
parte = Right(obj, Len(obj) - p)
t = t & """" & parte & """" & " "
temp.Add parte
Next
Set using = temp
Text1.Text = t
End Sub

Sub makelinks()
For i = 1 To using.Count
obj = using(i)
obj = Replace(obj, barra, "/")
t = t & "<a href=""" & obj & """>" & basename(obj) & "</a><br>" & vbCrLf
Next
Text1.Text = t
'<a href="$FILE">$NAME</><br>
End Sub

Sub makeImg()
For i = 1 To using.Count
obj = using(i)
obj = Replace(obj, barra, "/")
t = t & "<img src=""" & obj & """ alt=""" & basename(obj) & """><br>" & vbCrLf
Next
Text1.Text = t
'<IMG SRC="$FN" width="$width" height="$HEIGHT" ALT="$NAME" align=left />
End Sub

Sub makePerl()
Dim txt As String
If using.Count = 0 Then Exit Sub
txt = "@echo off & title PERL:" & basename(using(1)) & vbCrLf
txt = txt & "perl.exe -w """ & using(1) & """ %*"
For i = 2 To using.Count
txt = txt & " """ & using(i) & """"
Next
txt = txt & vbCrLf & "echo." & vbCrLf & "pause"
Text1.Text = txt
End Sub

Sub makeJs()
Dim txt As String, medio As String
If using.Count = 0 Then Exit Sub
txt = "var miArray=new Array("
medio = """" & using(1)
For i = 2 To using.Count
medio = medio & """, """ & using(i)
Next
medio = medio & """"
Text1.Text = txt & medio & ");"
End Sub
'////////////////////////////////////////////////


Function basepath(filename)
filename = "c:\"
pos = InStrRev(filename, "\")
basepath = Left(filename, pos)
End Function

Function basename(filename)
Dim pos As Integer
pos = InStrRev(filename, "\")
If pos = 0 Then pos = InStrRev(filename, "/")
basename = Right(filename, Len(filename) - pos)
End Function


Function getFiles(cmd As String) As Collection
Set getFiles = New Collection
Const cmilla = """"
cmd = Trim(cmd)
pos = InStr(1, cmd, cmilla)
If pos = 0 Then
    'fs = Split(cmd, " ")
    'For i = 0 To UBound(fs)
    getFiles.Add cmd 'fs(i)
    'Next
    Exit Function
End If
Do
If pos > 0 Then
    pos2 = InStr(pos + 1, cmd, cmilla)
    getFiles.Add Mid(cmd, pos + 1, pos2 - pos - 1)
End If
pos = InStr(pos2 + 1, cmd, cmilla)
Loop While pos > 0

End Function

Private Sub Text1_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'n t dnndndnd
For i = 1 To Data.Files.Count
Poner Data.Files.Item(i)
Next
simple
End Sub

Private Sub txtComunica_Change()
If txtComunica = "" Then Exit Sub
If Me.LinkMode = 0 Then Exit Sub

Set Files = getFiles(txtComunica.Text)
For Each f In Files
    Poner f
Next
txtComunica = ""
simple
End Sub

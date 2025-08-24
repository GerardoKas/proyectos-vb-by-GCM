VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pastero (for Paste)"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2355
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      OLEDropMode     =   1
      TextRTF         =   $"Form1.frx":038A
   End
   Begin VB.TextBox txtTitulo 
      Height          =   240
      Left            =   1440
      TabIndex        =   3
      Text            =   "Nombre Del Archivo"
      Top             =   0
      Width           =   1800
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1650
      Begin VB.CommandButton cmdMenu 
         BackColor       =   &H80000007&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   0
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdMinMax 
         BackColor       =   &H80000007&
         Caption         =   "Maxi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   315
         TabIndex        =   2
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdPaste 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   "Paste"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   810
         MaskColor       =   &H00004080&
         TabIndex        =   1
         Top             =   0
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   3600
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuNuevo 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Guardar Como..."
      End
      Begin VB.Menu mnuSaveOver 
         Caption         =   "Guardar Encima"
      End
      Begin VB.Menu mnuDirs 
         Caption         =   "Guardar En..."
         Begin VB.Menu mnuDirX 
            Caption         =   "Dir 0"
            Index           =   0
         End
      End
      Begin VB.Menu mnuDir 
         Caption         =   "Cambiar Directorio"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "Abrir en Notepad"
      End
      Begin VB.Menu mnNone 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesordenar 
         Caption         =   "Desordenado"
      End
      Begin VB.Menu mnuSortLines 
         Caption         =   "Ordenar Lineas"
      End
      Begin VB.Menu mnuDelDuplicate 
         Caption         =   "Eliminar Duplicados Juntos"
      End
      Begin VB.Menu mnuFindReplace 
         Caption         =   "Buscar o Reemplazar"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "Añadir Formato"
      End
      Begin VB.Menu mnuColorYFuente 
         Caption         =   "Color y Fuente"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdMenu_Click()
loaddirs
PopupMenu mnuMenu
End Sub

Sub loaddirs()
Dim i As Integer

dd = Dir$(savedir, vbDirectory)
mnuDirX.Item(0).Tag = ""
mnuDirX.Item(0).Caption = savedir
i = 1
Do While dd <> ""
    If GetAttr(savedir) = vbDirectory Then
    If mnuDirX.count <= i Then
    Load mnuDirX.Item(i)
    mnuDirX(i).Visible = True
    End If
    mnuDirX.Item(i).Caption = dd
    mnuDirX.Item(i).Tag = savedir
    i = i + 1
    End If
    dd = Dir$
Loop
For i = i To mnuDirX.count - 1
    mnuDirX.Item(i).Visible = False
Next
End Sub

Sub tellSaved(opt As Integer)
Select Case opt
Case 0 ' unsaved
'cambiando y sin grabar
grabado = False
Me.Caption = "Paste:" & filename
Case 1 'saved
'recien grabado
grabado = True
Me.Caption = "Saved. " & filename
txtTitulo.Text = basename(filename)
Case 2 'new
'sin nada escrito, nuevo
grabado = False
Me.Caption = "New"
Case 3 'opened
'abierto un archivo y no cambiado
grabado = True
Me.Caption = "Open:" & filename
txtTitulo = basename(filename)
End Select
End Sub

Private Sub cmdPaste_Click()
If Clipboard.GetFormat(vbCFText) Then
nt = Clipboard.GetText(vbCFText)
ElseIf Clipboard.GetFormat(vbCFRTF) Then
MsgBox "RTF"
nt = Clipboard.GetText(vbCFRTF)
ElseIf Clipboard.GetFormat(vbCFLink) Then
MsgBox "Link"
nt = Clipboard.GetText(vbCFLink)
ElseIf Clipboard.GetFormat(vbCFFiles) Then
'nt = Clipboard.GetText(vbCFFiles)

MsgBox "Clipboard Files " & vbCrLf & "No puedo capturar los ficheros Copiados, pero si los sueltas encima, se copiara el link"
End If
If nt = "" Then Exit Sub
Text1.Text = Text1.Text & Trim(nt) & bar
Text1.SelStart = Len(Text1.Text)

End Sub

Private Sub cmdMinMax_Click()
'mini maxi
If (cmdMinMax.Caption = "Mini") Then
    sizeMini
    'Me.Width = minW * Screen.TwipsPerPixelX
    'Me.Height = minH * Screen.TwipsPerPixelY
    cmdMinMax.Caption = "Maxi"
Else
    'Me.Left = 0
    'Me.Width = maxW
    'Me.Height = maxH * Screen.TwipsPerPixelY
    sizeMaxi
    cmdMinMax.Caption = "Mini"
End If
End Sub


Private Sub poner()
x = GetSetting(MyApp, MySection, "Left", 300)
y = GetSetting(MyApp, MySection, "Top", 0)
If x > Screen.Width Or x < 0 Then x = 0
If y > Screen.Height Or y < 0 Then y = 0
x = x / Screen.TwipsPerPixelX
y = y / Screen.TwipsPerPixelY
ponerArriba Me.hwnd, x, y
End Sub


Private Sub Form_Load()
Dim linea As String
linea$ = Command$
If linea <> "" Then
linea = Replace(Command$, Chr(34), "")
AbrirYLeer (linea)
Else
    tellSaved 2
End If
poner
If savedir = "" Or FileSystem.Dir(savedir, vbDirectory) = "" Then
savedir = "c:\TEMP"
Else
savedir = GetSetting(MyApp, MySection, "LastDir", App.Path)
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim comamdoIcono
Dim comandoExec

SaveSetting MyApp, MySection, "Left", Me.Left
SaveSetting MyApp, MySection, "Top", Me.Top
SaveSetting MyApp, MySection, "LastDir", savedir
comandoicono = App.Path & ".exe" & ",1"
comandoExec = App.Path & ".exe" & " %%*"
comandoicono = "cmd /C reg add HKCR\.nfo\DefaultIcon /ve /d """ & comandoicono & """ /f" 'icono
comandoprevio = "cmd /C reg add HKCR\.nfo\Shell\Open\Command"
comandoExec = "cmd /C reg add HKCR\.nfo\Shell\Open\Command /ve  /d """ & comandoExec & """ /f" 'programa
Debug.Print (comandoicono)
Debug.Print vbCrLf
Debug.Print comandoExec
MsgBox comandoicono & vbCrLf & comandoExec
If antesDeBorrar = False Then Cancel = 1
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Data.GetFormat(vbCFFiles) = True Then
filename = Data.Files(1)
AbrirYLeer (filename)
ElseIf Data.GetFormat(vbCFLink) = True Then
Data.GetData (vbCFLink)
MsgBox "daa vbcflink"
ElseIf Data.GetFormat(vbCFRTF) = True Then
MsgBox "DTA vbcrtf"
ElseIf Data.GetFormat(vbCFText) = True Then
MsgBox "dta vbCFText"
Else
Me.SetFocus
MsgBox "Nothing Happens"
End If
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
With Text1
.Width = Me.ScaleWidth
.Height = Me.ScaleHeight - .Top
End With
availW = -Me.ScaleWidth - Me.ScaleLeft + Screen.Width
availH = Screen.Height - Me.ScaleWidth - Me.ScaleHeight
If Me.ScaleWidth >= 300 Then
    cmdMinMax.Caption = "Mini"
ElseIf Me.ScaleWidth < 300 Then
    cmdMinMax.Caption = "Maxi"
End If
If Me.ScaleHeight >= 300 Then
    cmdMinMax.Caption = "Mini"
ElseIf Me.ScaleHeight < 300 Then
    cmdMinMax.Caption = "Maxi"
End If
txtTitulo.Width = Me.ScaleWidth - txtTitulo.Left
End Sub

Private Sub mnuColorYFuente_Click()
frmOptions.Show 0, Me

End Sub

Private Sub mnuDelDuplicate_Click()
Dim aLines As Variant, linedel() As Boolean

aLines = Split(Text1.Text, vbCrLf)
ReDim linedel(UBound(aLines))
For i = 0 To UBound(aLines) - 1
    If StrComp(aLines(i), aLines(i + 1), vbTextCompare) = 0 Then
        linedel(i + 1) = True
    End If
Next
Dim t As String
For i = 0 To UBound(aLines)
    'si no hay quer eliminarla
    If Not linedel(i) = True Then
        t = t & aLines(i) & vbCrLf
    End If
Next
Text1.Text = t
End Sub

Private Sub mnuDesordenar_Click()
antesDeBorrar
t = Text1.Text
t = Replace(t, vbCrLf, " ")
ps = Split(t, " ")

For i = 0 To UBound(ps)
    num = Int(Rnd(1) * UBound(ps))
    temp = ps(i)
    ps(i) = ps(num)
    ps(num) = temp
Next
Text1.Text = Join(ps, " ")
End Sub

Private Sub mnuDir_Click()
'save dir
Dim fold As String
fold = BrowseForFolder(savedir)
If fold <> "" Then
savedir = fold
Else
MsgBox "Se guardara en el directorio de la Applicacion : " & vbCrLf & App.Path
savedir = App.Path
End If
End Sub

Private Sub mnuDirX_Click(Index As Integer)
savedir = mnuDirX.Item(Index).Tag & "\" & mnuDirX.Item(Index).Caption
If MsgBox("Deseas guardarlo ahora con el nombre: " & vbCrLf & filename, vbInformation + vbYesNo) = vbYes Then
    grabarFichero (filename)
    tellSaved 1
End If
End Sub

Private Sub mnuFindReplace_Click()
'buscar InputBox("Cadena a buscar?")
Form2.Show 0, Me
End Sub

Private Sub mnuNotepad_Click()
Dim tempName As String
If grabado = True Then
tempName = filename
Else
tempName = "c:\TEMP\Pasteado_" & DateTime.Date$ & ".TXT"
grabarFichero tempName
End If
Shell "Notepad.exe " & tempName, vbNormalFocus
End Sub

Private Sub mnuNuevo_Click()
antesDeBorrar
Text1.Text = ""
tellSaved 2
End Sub

Private Sub mnuSaveAs_Click()
'save
filename = getFilename(cmdlg)
grabarFichero filename
txtTitulo = filename
tellSaved 1
End Sub

Private Sub mnuSaveOver_Click()
If savedir = "" Then
mnuDir_Click
End If
grabarFichero savedir & "\" & txtTitulo
tellSaved 1
End Sub

Private Sub mnuSortLines_Click()
Dim aLines As Variant
aLines = Split(Text1.Text, vbCrLf)
For i = 0 To UBound(aLines)
    For j = i + 1 To UBound(aLines)
        'si i es mayor que j
        If StrComp(aLines(i), aLines(j), vbTextCompare) > 0 Then
            temp = aLines(j)
            aLines(j) = aLines(i)
            aLines(i) = temp
        End If
    Next
Next
Text1.Text = Join(aLines, vbCrLf)
End Sub

Private Sub Text1_Change()
If Text1.Tag = "NoDo" Then Exit Sub
If Text1.Text = "" Then
txtTitulo = ""
tellSaved 2
Else
'txtTitulo = crearTitulo(Text1.Text)
End If
txtTitulo = "Pasted-" & Now() & ".txt"
End Sub

Private Sub Text1_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

If Data.GetFormat(vbCFFiles) Then
    For i = 1 To Data.Files.count
    ficheros = ficheros & """" & Data.Files(i) & """" & vbCrLf
    Next
    Text1.Text = Text1.Text & ficheros & bar
ElseIf Data.GetFormat(vbCFLink) Then
    MsgBox "Eso deberia ser un link aparece como vbcflink pero no se que hacer con ello"
ElseIf Data.GetFormat(vbCFText) Then
    Text1.Text = Text1.Text & Trim(Data.GetData(vbCFText)) & bar
ElseIf Data.GetFormat(vbCFRTF) = True Then
    Text1.Text = Text1.Text & Trim(Data.GetData(vbCFRTF)) & bar
ElseIf Data.GetFormat(vbCFEMetafile) = True Then
DoEvents
End If
Text1.SelStart = Len(Text1.Text)
End Sub


Private Sub txtTitulo_Change()
filename = txtTitulo
End Sub

Public Sub AbrirYLeer(fichero As String)
Me.Enabled = False
Me.MousePointer = vbHourglass
If Dir$(fichero, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
Text1.Text = openFile(fichero)
tellSaved 3
Else
MsgBox fichero & vbCrLf & "No es un fichero"
tellSaved 2
End If
Me.Enabled = True
Me.MousePointer = vbDefault
Text1.SelStart = 0
'Me.SetFocus
End Sub


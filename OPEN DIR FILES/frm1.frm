VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "LISTADOS MSDOS"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3195
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2310
      Left            =   3330
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frm1.frx":0000
      Top             =   0
      Width           =   1950
   End
   Begin VB.ListBox List2 
      Height          =   1035
      ItemData        =   "frm1.frx":000F
      Left            =   45
      List            =   "frm1.frx":0011
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   1755
      Width           =   3165
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "frm1.frx":0013
      Left            =   90
      List            =   "frm1.frx":0015
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   630
      Width           =   2580
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Para Listados de ficheros Estilo MSDOS"
      Height          =   195
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   2835
   End
   Begin VB.Menu mnuNormal 
      Caption         =   "Ver"
      Begin VB.Menu mnuAddThisDir 
         Caption         =   "Añadir Directorio Actual a las Rutas (.\)"
      End
      Begin VB.Menu mnuVerTitulo 
         Caption         =   "Ver Solo Nombre"
      End
      Begin VB.Menu mnuVerTodo 
         Caption         =   "Ver Ruta Completa"
      End
      Begin VB.Menu mnuViewProps 
         Caption         =   "Ver Propiedades de Arcivos"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "MenuFile"
      Visible         =   0   'False
      Begin VB.Menu mnuVario 
         Caption         =   "MenuMultiple"
         Index           =   0
      End
   End
   Begin VB.Menu mnuAcciones 
      Caption         =   "Acciones"
      Begin VB.Menu mnuOpenOnCmd 
         Caption         =   "Abrir en CMD"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuOpenCmdHelp 
         Caption         =   "Abrir En CMD con '/?'"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuNotin 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPropiedades 
         Caption         =   "Propiedades"
         Shortcut        =   ^P
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myFile As String
Dim fDir As String
Dim fFile As String
Dim shlFile As FolderItem
Dim shlfolder As Folder
Dim bShowProps As Boolean
Dim shl As Shell
Dim bitisnot As Boolean
Dim bAppendThisDir As Boolean

Private Sub Form_Load()
If Command$ <> "" Then
    myFile = Replace(Command$(), """", "")
    openIt
End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.Files.Count = 0 Then Exit Sub
myFile = Data.Files(1)
openIt
End Sub

Sub openIt()
Dim l As String
Dim f As Integer
List1.Clear
List2.Clear
f = FreeFile
Open myFile For Input As #f
Line Input #f, l
Do While Not EOF(f)
    List1.AddItem l
    getFD l
    List2.AddItem fFile
    Line Input #f, l
Loop
Close #f

'MsgBox "Listo"
Me.Caption = List1.ListCount & " Ficheros En La Lista"
Set shl = New Shell
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
Me.ScaleMode = vbPixels
Label1.Top = 0
List1.Left = 0
List2.Left = List1.Left
List1.Top = Label1.Height + 3
List2.Top = List1.Top
List1.Height = Me.ScaleHeight - Label1.Height - 3
List2.Height = List1.Height

If bShowProps = False Then
List1.Width = Me.ScaleWidth
List2.Width = Me.ScaleWidth
Text1.Visible = False
Else
'List1.Width = Me.ScaleWidth - Screen.TwipsPerPixelX * 250
'List2.Width = Me.ScaleWidth - Screen.TwipsPerPixelY * 250
Text1.Left = Me.ScaleWidth * 2 / 3
Text1.Width = Me.ScaleWidth - Text1.Left
List1.Width = Text1.Left '- Screen.TwipsPerPixelX * 250
List2.Width = Text1.Left '- Screen.TwipsPerPixelY * 250
Text1.Top = List1.Top
Text1.Height = List1.Height
Text1.Visible = True
End If
End Sub


Private Sub List1_DblClick()
texto = List1.List(List1.ListIndex)
getItem (texto)
shlFile.InvokeVerb (0)
End Sub
 
Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyExecute Then
List1_MouseDown 1, 0, 0, 0
End If
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub


Private Sub List2_DblClick()
texto = List1.List(List2.ListIndex)
getItem (texto)
shlFile.InvokeVerb (0)
End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyExecute Then
List2_MouseDown 1, 0, 0, 0
End If
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

List1.ListIndex = List2.ListIndex
If List1.ListIndex = -1 Then Exit Sub
If Button = 2 Then
    loadMenu (List1.List(List2.ListIndex))
End If
loaddata (List1.List(List2.ListIndex))
End Sub

Private Sub List2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
List2.ListIndex = List1.ListIndex
If Button = 2 Then
    loadMenu (List1.List(List1.ListIndex))
End If
loaddata (List1.List(List1.ListIndex))
End Sub



Sub getItem(texto)
getFD (texto)
'If fFile Then Exit Sub
On Error GoTo itisnot
Set shlfolder = shl.NameSpace(fDir)
Set shlFile = shlfolder.ParseName(fFile)
If shlFile Is Nothing Then GoTo itisnot
bitisnot = False
Exit Sub
itisnot:
MsgBox "Esto no es un fichero!"
bitisnot = True
End Sub

Sub loadMenu(texto As String)
Dim f As String, i As Integer
Dim vrbs As FolderItemVerbs, v As FolderItemVerb

getItem (texto)
If bitisnot = True Then Exit Sub
Set vrbs = shlFile.Verbs
For i = 1 To mnuVario.Count - 1
    Unload mnuVario(i)
Next
i = 1
For Each v In vrbs
    Load mnuVario(i)
    If v.Name = "" Then
        mnuVario(i).Caption = "-"
    Else
        mnuVario(i).Caption = v.Name
    End If
    i = i + 1
Next
mnuVario(0).Caption = "Abrir Carpeta!"

PopupMenu mnuFile
End Sub
Sub loaddata(texto)
getItem (texto)
If bitisnot = True Then Exit Sub
fecha = shlFile.ModifyDate
tipo = shlFile.Type
tam = shlFile.Size
If bShowProps = True Then
t = shlFile.Path & vbCrLf
For i = 1 To 40
colx = shlfolder.GetDetailsOf(shlFile, i)
If colx <> "" Then
t = t & vbCrLf & colx
End If
Next
Text1.Text = t
End If
Label1.Caption = shlFile.Name & ", (" & _
Format(fecha, "dd/mm/yyyy") & "), " & _
Format(tam / 1024, "###,##") & "Kb, " & _
colx 'tipo

End Sub
Sub getFD(texto)
p = InStrRev(texto, "\")
If bAppendThisDir = True Then
myPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
fDir = myPath & Left(texto, p)
Else
fDir = Left(texto, p)
End If
fFile = Right(texto, Len(texto) - p)
End Sub

Private Sub mnuAddThisDir_Click()
If mnuAddThisDir.Tag = "" Then
bAppendThisDir = True
mnuAddThisDir.Tag = "FISTRRO"
mnuAddThisDir.Caption = "Quitar Directorio Actual (.\)"
Else
bAppendThisDir = False
mnuAddThisDir.Tag = ""
mnuAddThisDir.Caption = "Añadir Directorio Actual a las Rutas (.\)"
End If
End Sub

Private Sub mnuOpenCmdHelp_Click()
cmdline = "cmd /K cd """ & fDir & """ & """ & fDir & fFile & """ /?"
'MsgBox cmdline
Shell cmdline, vbNormalFocus
End Sub

Private Sub mnuOpenOnCmd_Click()
cmdline = "cmd /K cd """ & fDir & """ & """ & fDir & fFile & """ "
'MsgBox cmdline
Shell cmdline, vbNormalFocus
End Sub

Private Sub mnuPropiedades_Click()
 shlFile.InvokeVerb "Properties"
End Sub

Private Sub mnuVario_Click(Index As Integer)
If Index = 0 Then
shl.Open (shlfolder.Self.Path)
Else
shlFile.InvokeVerb (mnuVario(Index).Caption)
End If
End Sub

Private Sub mnuVerTitulo_Click()
List2.Visible = True
List1.Visible = False
End Sub

Private Sub mnuVerTodo_Click()
List2.Visible = False
List1.Visible = True
End Sub

Private Sub mnuViewProps_Click()
bShowProps = True
Form_Resize
End Sub

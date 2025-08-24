VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "GCM_Explorer"
   ClientHeight    =   3375
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5490
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   6  'Mask Pen Not
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   45
      ScaleHeight     =   240
      ScaleMode       =   0  'User
      ScaleWidth      =   57.313
      TabIndex        =   4
      Top             =   315
      Width           =   240
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   2715
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   4789
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5145
      Begin VB.ComboBox cmbFolders 
         CausesValidation=   0   'False
         Height          =   315
         ItemData        =   "Form1.frx":038A
         Left            =   585
         List            =   "Form1.frx":038C
         TabIndex        =   5
         Text            =   "cmbFolders"
         Top             =   0
         Width           =   4485
      End
      Begin MSComctlLib.ImageList il1 
         Left            =   4095
         Top             =   270
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "Updir"
         Height          =   315
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   240
         Left            =   315
         TabIndex        =   2
         Top             =   360
         Width           =   4515
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "OPCIONES"
      Height          =   1905
      Left            =   0
      TabIndex        =   6
      Top             =   630
      Visible         =   0   'False
      Width           =   5145
      Begin VB.CommandButton cmdDetener 
         Caption         =   "Detener"
         Height          =   420
         Left            =   2520
         TabIndex        =   8
         Top             =   225
         Width           =   2265
      End
      Begin VB.CommandButton cmdNoFolderSize 
         Caption         =   "No Calcular Tamaño"
         Height          =   420
         Left            =   180
         TabIndex        =   7
         Top             =   225
         Width           =   2265
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   825
         Left            =   180
         TabIndex        =   9
         Top             =   810
         Width           =   4620
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Menu"
      Begin VB.Menu mnuEjecutar 
         Caption         =   "Ejecutar"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuExplorar 
         Caption         =   "Explorar Esta Carpeta"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuCmd 
         Caption         =   "Ejecutar en Cmd"
      End
      Begin VB.Menu mnuVentanas 
         Caption         =   "Ventanas"
      End
   End
   Begin VB.Menu mnuSimpleMenu 
      Caption         =   "File Actions"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Abrir"
      End
      Begin VB.Menu mnuVerTexto 
         Caption         =   "Ver ASCII"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuChangeName 
         Caption         =   "Cambiar Nombre"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFileCopy 
         Caption         =   "Copiar"
      End
      Begin VB.Menu mnuFileCut 
         Caption         =   "Cortar"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "Eliminar"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuThisFile 
      Caption         =   "File Win Menu"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuThisFileMenus 
         Caption         =   "menuDelFichero"
         Index           =   0
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Opciones"
      Begin VB.Menu mnuCalcFolderSize 
         Caption         =   "CalcularFolder Sizes"
      End
      Begin VB.Menu mnuMarcarIso 
         Caption         =   "Marcar Nombres Largos (iso)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const espaciado = 10
Dim myFolderItem As FolderItem
Dim myFolderItems As FolderItems
Dim liSelected As ListItem
Dim startOleDrag As Boolean
Dim draggin As Boolean
Dim shl As New Shell
Dim shFolder As folder

Private Sub cmdDetener_Click()
HALTIT = True
End Sub

'actualizar despues de borrar o cortar ctrl-x
'cambiar nombre y actualizar
'arrastrar para mover o copiar


Private Sub cmdNoFolderSize_Click()
calculateFolderSize = False
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
If KeyCode = 116 Then
mostrarDir
End If
End Sub

'///////////////////////////////////////////////////

'
'Private Sub cmdNew_Click()
'Default = nowDir
'cadena = InputBox("Introduce la carpeta", "Nuevo acceso a Carpeta", Default)
'If cadena = "" Or dir(cadena, vbDirectory) = "" Then
'MsgBox "Debes seleccionar Una Carpeta"
'Exit Sub
'End If
'
'For i = 0 To cmbFolders.ListCount - 1
'    If LCase(cadena) = LCase(cmbFolders.List(i)) Then
'    MsgBox "Esa Carpeta ya existe como " & vbCrLf & cmbFolders.List(i)
'    Exit Sub
'    End If
'Next
'cmbFolders.AddItem cadena, 0
'cmbFolders.ListIndex = 0
'numcarpetas = numcarpetas + 1
'End Sub
'
'Private Sub cmdQuitar_Click()
'If MsgBox("deseas eliminar el elemento : " & cmbFolders.Text & " ?", vbYesNo) = vbNo Then Exit Sub
'If cmbFolders.ListIndex = -1 Then
'MsgBox "Debes seleccionar un item de la lista de combo"
'Exit Sub
'End If
'cmbFolders.RemoveItem (cmbFolders.ListIndex)
'numcarpetas = numcarpetas - 1
'If cmbFolders.ListCount = 0 Then
'MsgBox "No hay ninguna carpeta. Si reinicias el programa se añadiran las carpetas por defecto del sistema"
'Exit Sub
'End If
'cmbFolders.ListIndex = 0
'End Sub
'
Private Sub Form_Unload(Cancel As Integer)
    saveConfig cmbFolders
    clearIcons
End Sub

Private Sub Form_Load()
addColumns lv1
iniciarListas lv1, il1, Picture1
loadFolders cmbFolders
'colorear
End Sub

Private Sub Form_Resize()
'Me.ScaleMode = vbPixels
On Error Resume Next
Frame1.Top = 0
Frame1.Left = 0
Frame1.Width = Me.ScaleWidth - Frame1.Left ' - Frame2.Width
cmbFolders.Width = Frame1.Width - cmdQuitar.Left - cmdQuitar.Width - 180
'Drive1.Width = Frame1.Width - Drive1.Left - 180
lv1.Top = Frame1.Height + Frame1.Top
lv1.Height = Me.ScaleHeight - Frame1.Height - espaciado
lv1.Width = Me.ScaleWidth - espaciado
End Sub


Private Sub lv1_AfterLabelEdit(Cancel As Integer, NewString As String)
'comprobar newstring caracteres validos
myFolderItem.Name = NewString
'mostrarDir
End Sub

Private Sub lv1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim col As columnas
col = ColumnHeader.Index
lv1.Sorted = True
lv1.SortKey = ColumnHeader.Index - 1
If InStr(1, ColumnHeader.Key, "ORDER") > 0 Then
    lv1.SortKey = lv1.ColumnHeaders("NUMBER-Tamano").Index - 1
End If
If lv1.SortOrder = lvwAscending Then
    lv1.SortOrder = lvwDescending
 
Else
    lv1.SortOrder = lvwAscending
End If
End Sub


Private Sub lv1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then
cmdUp_Click
End If
End Sub

Private Sub lv1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim shitem As FolderItem

Set liSelected = lv1.HitTest(X, Y)
If liSelected Is Nothing Then Exit Sub
lv1.SelectedItem.Selected = False
liSelected.Selected = True

Set shFolder = shl.NameSpace(myFolder)
Set myFolderItem = shFolder.ParseName(liSelected)

If myFolderItem Is Nothing Then
Set myFolderItem = shFolder.ParseName(liSelected.SubItems(6))
'Label1.Caption = "no se pudo leer el archivo"
'Set myFolderItem = shl.NameSpace(lv1.SelectedItem.SubItems(6))
Exit Sub
End If
Label1.Caption = myFolderItem.Name & ", " & myFolderItem.size & ", " & myFolderItem.ModifyDate
If Button = 2 Then
    loadFileMenu mnuThisFileMenus, myFolderItem.Verbs
    PopupMenu mnuThisFile
End If
End Sub


Private Sub lv1_DblClick()
    If lv1.SelectedItem Is Nothing Then Exit Sub
    Set liSelected = lv1.SelectedItem
    Set shFolder = shl.NameSpace(myFolder)
    Set myFolderItem = shFolder.ParseName(liSelected.Text)
    If myFolderItem Is Nothing Then
        'para network
'        Set shFolder = shl.NameSpace(myFolder & liSelected.Text)
'        If Not (shFolder Is Nothing) Then
'        If shFolder.Self.IsFolder Then
'        cmbFolders.Text = shFolder.Self.path
'        mostrarDir
'        End If
'        End If
    Set myFolderItem = shFolder.ParseName(liSelected.SubItems(6))
    
    End If
    If myFolderItem Is Nothing Then Exit Sub
    If esFolder(myFolderItem.path) Then
        cmbFolders.Text = myFolderItem.path
        mostrarDir
    ElseIf myFolderItem.IsLink Then
        fn = myFolderItem.GetLink.path
        cmbFolders.Text = fn
        mostrarDir
    Else
        mnuFileOpen_Click
    End If
    
End Sub

Private Sub lv1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(15) = True Then
MsgBox Data.Files.Count
End If
End Sub

'Private Sub lv1_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
'Data.Files.Add myFolderItem.path
'Data.GetData 15
'Debug.Print "DATA PEDIDO: " & DataFormat
'Data.SetData myFolderItem.path, 1
'End Sub

Private Sub lv1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
Data.Clear
'Data.SetData
Data.Files.Add "c:\index.html"
'Data.Files.Add myFolderItem.path
'AllowedEffects = 1
'Label1.Caption = "MOVING: " & Data.Files(1)
'lv1.OLEDrag
Data.SetData "COSA", 1
End Sub

Private Sub cmbFolders_Click()
mostrarDir
End Sub

Private Sub cmdUp_Click()
cmbFolders.Text = updir(myFolder)
mostrarDir
End Sub

Private Sub cmbFolders_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
mostrarDir
End If
End Sub

Sub mostrarDir()
Me.MousePointer = vbHourglass
readDir cmbFolders.Text, lv1
If mnuMarcarIso.Checked = True Then
nombresLargos lv1, 31
End If
Me.MousePointer = 0
End Sub


Private Sub mnuCalcFolderSize_Click()
If mnuCalcFolderSize.Checked = True Then
mnuCalcFolderSize.Checked = False
calculateFolderSize = False
Else
mnuCalcFolderSize.Checked = True
calculateFolderSize = True
End If
End Sub

Private Sub mnuChangeName_Click()
lv1.StartLabelEdit
End Sub

Private Sub mnuCmd_Click()
mnuCmd.Enabled = False
End Sub


Private Sub mnuFileDelete_Click()
Dim Papelera As folder
Set Papelera = shl.NameSpace(ssfBITBUCKET)
Papelera.MoveHere (myFolderItem.path)
mostrarDir
End Sub

Private Sub mnuFileOpen_Click()
myFolderItem.InvokeVerb 0
End Sub

Private Sub mnuThisFileMenus_Click(Index As Integer)
texto = mnuThisFileMenus(Index).Caption
myFolderItem.InvokeVerb (texto)
End Sub

Sub loadFileMenu(mmenu, verbos As FolderItemVerbs)
Dim verbo As FolderItemVerb

For i = mmenu.Count To verbos.Count - 1
    Load mmenu(i)
Next
For i = verbos.Count To mmenu.Count - 1
    Unload mmenu(i)
Next
i = 0
For Each verbo In verbos
    If verbo.Name = "" Then
        mmenu(i).Caption = "-"
    Else
        mmenu(i).Caption = verbo.Name
    End If
    On Error Resume Next
    Debug.Print verbo.Name
    'verbo.DoIt
    i = i + 1
Next
End Sub

Private Sub mnuVerTexto_Click()
Shell "cmd /C title """ & fileItemText() & """ & more """ & fileItemText() & """ & pause", vbNormalFocus
End Sub

Function fileItemText()
fileItemText = myFolderItem.path '& myFolderItem.Name
End Function


'Private Sub mnuCmd_Click()
'Dim r As Double
'comando = InputBox("Introduce comando", "Ejecutar", "")
'r = Shell("cmd -k " & comando, vbNormalFocus)
'MsgBox "Resultado : " & r
'End Sub
'
'Private Sub OpenFolders()
'Dim sh As New Shell32.Shell
'Dim objs As Object
'Dim obj As Object
'Set objs = sh.Windows
'If objs.Count > 0 Then
''Unload mnuUnaVentana(0)
'For i = 1 To objs.Count - 1
'    Load mnuUnaVentana(i)
'Next
'For i = 0 To objs.Count - 1
'    Set obj = objs(i)
'    mnuUnaVentana(i).Caption = obj.LocationName 'obj.Document.Folder.Self.Path
'    'MsgBox obj.Document.Folder.Self.Type
'Next
'End If
'End Sub

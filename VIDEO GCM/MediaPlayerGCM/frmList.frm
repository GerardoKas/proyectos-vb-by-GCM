VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmList 
   Caption         =   "GCM_Player Listado"
   ClientHeight    =   2625
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form2"
   MouseIcon       =   "frmList.frx":030A
   ScaleHeight     =   2625
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   1725
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   3043
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   6600394
      BackColor       =   4470059
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuHacer 
      Caption         =   "Hacer"
      Begin VB.Menu mnuLoadDur 
         Caption         =   "Cargar Duraciones"
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "MenuuArchivo"
      Visible         =   0   'False
      Begin VB.Menu mnuFindFile 
         Caption         =   "Ubicar En Carpeta"
      End
      Begin VB.Menu mnuExternOpen 
         Caption         =   "Abrir Externo"
      End
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private numarks As Integer
Private Sub Form_Load()
    ListView1.ColumnHeaders.Add 1, "Nombre", "Nombre"
    ListView1.ColumnHeaders.Add 2, "Tamaño", "Tamaño"
    ListView1.ColumnHeaders.Add 3, "Directorio", "Directorio"
    ListView1.ColumnHeaders.Add 4, "Llegada", "Llegada"
    ListView1.ColumnHeaders.Add 5, "Duracion", "Duracion"
    ListView1.ColumnHeaders.Add 6, "Segundos", "Segundos"
    'ListView1.ColumnHeaders
    ListView1.View = lvwReport
    AñadirList
End Sub
Public Sub marcar(indice As Integer)
ListView1.ListItems(indice + 1).Selected = True
Me.Caption = indice + 1
End Sub
Public Sub AñadirList()
Dim elitem As ListItem
ListView1.ListItems.Clear
For i = 0 To NumArks - 1
Set elitem = ListView1.ListItems.Add(, i & "_", ListaArk2(i).nomArk)
elitem.SubItems(1) = TheSize(ListaArk2(i).Tamano)
elitem.SubItems(2) = ListaArk2(i).Directorio
elitem.SubItems(3) = i
If ListaArk2(i).segundos > 0 Then
    elitem.SubItems(4) = SecsToTime(ListaArk2(i).segundos)
    elitem.SubItems(5) = ListaArk2(i).segundos
Else
elitem.SubItems(4) = "NOT LOADED"
elitem.SubItems(5) = "NOT LOADED"
End If
Next
End Sub
Public Sub clearList()
ListView1.ListItems.Clear
End Sub
'====================
'sxe ordena listaark2()
Private Sub ordena()
valor = ListaArk2(0).Tamano
For z = 0 To 10000
Next
End Sub
Private Sub Form_Resize()
ListView1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.Sorted = True

ListView1.SortKey = ColumnHeader.Index - 1
If ListView1.SortOrder = lvwAscending Then
ListView1.SortOrder = lvwDescending
Else
ListView1.SortOrder = lvwAscending
End If

End Sub

Private Sub ListView1_DblClick()
posicion = selectedPosition()
Form1.Tocar furFilenames_.getFilename(ListaArk2(posicion))
Form1.ZOrder 0
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuArchivo
End If
End Sub

Private Sub ListView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
'MsgBox ("STARTING DRAG")
End Sub

Private Sub mnuExternOpen_Click()
posicion = selectedPosition
Shell getFilename(ListaArk2(posicion))
End Sub

Private Sub mnuLoadDur_Click()
Dim dur As String, secs As Long
Form1.MP1.Visible = False
For i = 0 To NumArks - 1
secs = Form1.getSecs(getFilename(ListaArk2(i)))
dur = SecsToTime(secs)
ListView1.ListItems(i & "_").SubItems(4) = dur
ListView1.ListItems(i & "_").SubItems(5) = secs
ListaArk2(i).duracion = dur
ListaArk2(i).segundos = secs
'Debug.Print Form1.MP1.GetCodecDescription(1)
Next
End Sub

Private Function selectedPosition() As Integer
Key = ListView1.SelectedItem.Key
selectedPosition = VBA.Left(Key, Len(Key) - 1)
End Function

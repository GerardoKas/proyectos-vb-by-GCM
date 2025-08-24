VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmProceso 
   Caption         =   "Form2"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form2"
   ScaleHeight     =   3585
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   45
      TabIndex        =   2
      Top             =   2970
      Width           =   6180
      Begin VB.CommandButton cmdAddToList 
         Caption         =   "Add To List"
         Height          =   330
         Left            =   5085
         TabIndex        =   7
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox txtStyle 
         Height          =   285
         Left            =   3240
         TabIndex        =   6
         Text            =   "cursor:hand;"
         Top             =   180
         Width           =   1770
      End
      Begin VB.TextBox txtSelector 
         Height          =   285
         Left            =   2115
         TabIndex        =   3
         Text            =   "BODY"
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Add Temporal Style"
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   135
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estilo"
         Height          =   195
         Left            =   2835
         TabIndex        =   5
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SELECTOR"
         Height          =   195
         Left            =   1215
         TabIndex        =   4
         Top             =   225
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "Aplicar a Pagina"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      Top             =   2610
      Width           =   1545
   End
   Begin MSComctlLib.ListView lvMarcas 
      Height          =   2535
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StyleCreado As Boolean
Dim objStyle As IHTMLStyleSheet
Dim nItem As ListItem
Const myObjId$ = "MyEditorStyle"

Sub initLv()
With lvMarcas
    .ColumnHeaders.Add 1, , "Nombre"
    .ColumnHeaders.Add 2, , "CSS"
    .ColumnHeaders.Add 3, , "Enabled"
End With
End Sub


Private Sub cmdAplicar_Click()
Dim i As ListItem, elem As IHTMLElement
    Set elem = myDoc.getElementById(myObjId)

If Not elem Is Nothing Then
    Set objStyle = elem.styleSheet
    objStyle.cssText = ""
Else
    Set objStyle = myDoc.createStyleSheet("VOID.CSS")
     objStyle.owningElement.Id = myObjId
     StyleCreado = True
End If

For Each i In lvMarcas.ListItems
    If i.Checked = True Then
    objStyle.addRule i.Text, i.SubItems(1)
    Else
        On Error Resume Next
        objStyle.addRule i.Text, " "
    End If
Next
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim lis As ListItem
initLv
ReadWrite.setIniFile "marcaduras.ini"
secciones = ReadWrite.ReadKeys("marks")

lista = Split(secciones, Chr(0))
For Each i In lista
    If i <> "" Then
    Set lis = lvMarcas.ListItems.Add(, , i)
    dato = ReadWrite.ReadFromFile("marks", CStr(i))
    lis.SubItems(1) = dato
    lis.SubItems(2) = "Falso"
    End If
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As ListItem
For Each i In lvMarcas.ListItems
ReadWrite.WriteToFile "marks", i.Text, i.SubItems(1)
Next

End Sub

Private Sub lvMarcas_ItemClick(ByVal Item As MSComctlLib.ListItem)
Set nItem = Item
txtSelector = nItem.Text
txtStyle = nItem.SubItems(1)
End Sub

Private Sub cmdAddToList_Click()
If txtSelector = "" Then MsgBox "Debes especificar selector (tag, o class, o id)": Exit Sub
If nItem Is Nothing Then
'si no esta editando
    Set posprev = lvMarcas.FindItem(txtSelector)
    If Not posprev Is Nothing Then
        MsgBox "Ya existe un elemento con ese selector"
        Set nItem = posprev
        txtSelector = nItem.Text
        txtStyle = nItem.SubItems(1)
    Else
        'añadir neuvo
        Set nItem = lvMarcas.ListItems.Add(, , txtSelector)
        nItem.SubItems(1) = txtStyle
    End If
Else
    nItem.Text = txtSelector
    nItem.SubItems(1) = txtStyle
    txtSelector = ""
    txtStyle = ""
    Set nItem = Nothing
End If

End Sub


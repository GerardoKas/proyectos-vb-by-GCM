VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProp 
   Caption         =   "Form2"
   ClientHeight    =   4545
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6780
   LinkTopic       =   "Form2"
   ScaleHeight     =   4545
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSearchTags 
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   0
      Width           =   1365
   End
   Begin VB.CommandButton cmdBuiscarTags 
      Caption         =   "BUSCAR"
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   0
      Width           =   825
   End
   Begin VB.TextBox txtInner 
      Height          =   1995
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmProp.frx":0000
      Top             =   2475
      Width           =   3165
   End
   Begin VB.CommandButton cmdParentNode 
      Caption         =   "ParentNode"
      Height          =   330
      Left            =   3195
      TabIndex        =   4
      Top             =   45
      Width           =   1140
   End
   Begin MSComctlLib.TreeView tree1 
      Height          =   2040
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   3598
      _Version        =   393217
      Indentation     =   0
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   4110
      Left            =   3195
      TabIndex        =   2
      Top             =   405
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   7250
      View            =   3
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "cName"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "cValue"
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "cEspecificado"
         Text            =   "Especificado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtValue 
      BackColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   3285
      TabIndex        =   0
      Text            =   "DATO"
      Top             =   4185
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DATO: "
      Height          =   195
      Left            =   4365
      TabIndex        =   1
      Top             =   135
      Width           =   540
   End
   Begin VB.Menu mnOpcion 
      Caption         =   "Opcion"
      Begin VB.Menu mnuDeleteObject 
         Caption         =   "Eliminar Objeto"
      End
   End
End
Attribute VB_Name = "frmProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objeto As IHTMLElement
Private atributo As String
Private editing As Boolean
Private newEdit As String
Private myLi As ListItem
Private objTree As IHTMLElement
Private listaElements As IHTMLElementCollection

Sub loadObject(obj As IHTMLElement)
Dim evts As New Collection
lv1.ListItems.Clear
Me.Caption = obj.tagName & ", " & obj.uniqueID
Label1.Caption = obj.tagName
If obj.tagName <> "!" Then
For Each i In obj.Attributes
'If i.name = "style" Then MsgBox "Paraaqui"
    If Not Left(i.Name, 2) = "on" Then
    If Not Nothing Is i.Attributes Then MsgBox ("aqui hay attrib")
    Set myLi = lv1.ListItems.Add(, , i.Name)
    myLi.SubItems(1) = IIf(IsNull(obj.getAttribute(i.Name)), "?", obj.getAttribute(i.Name))
    myLi.SubItems(2) = i.specified
    If i.specified = True Then myLi.bold = True
    Else
        'es un evento
        evts.Add i
    End If
Next
Else
DoEvents
End If
txtInner.Text = obj.outerHTML
Set objeto = obj
Exit Sub
For Each i In evts
    Set myLi = lv1.ListItems.Add(, , i.Name)
    myLi.SubItems(1) = IIf(IsNull(obj.getAttribute(i.Name)), "?", obj.getAttribute(i.Name))
    myLi.SubItems(2) = i.specified
Next
End Sub

Private Function loadFullTree(obj As IHTMLElement)
Dim chlds As Object
tree1.Nodes.Clear
Set chlds = obj.children
tree1.Nodes.Add , , obj.uniqueID, obj.tagName
For Each i In chlds
tree1.Nodes.Add obj.uniqueID, tvwChild, i.uniqueID, i.nodeName
'tree1.Nodes.Add i.uniqueID, tvwChild, "X" & i.uniqueID, i.innerHTML
'tree1.Nodes.Add i.uniqueID, tvwChild, "L" & i.uniqueID, "SubNodos"
tree1.Nodes.Item(i.uniqueID).Expanded = True
Next

End Function

Sub loadSubTree(obj As Object, clave As String)
Dim chlds As Object
If tree1.Nodes.Item(clave).children > 0 Then Exit Sub
Set chlds = obj.children
For Each i In chlds
    tree1.Nodes.Add clave, tvwChild, i.uniqueID, i.tagName
Next
tree1.Nodes.Item(clave).Expanded = True
End Sub

Private Sub cmdBuiscarTags_Click()
Dim chlds As Object
Dim i As IHTMLElement
tree1.Nodes.Clear
tree1.Nodes.Add , , "XRAIZ", "BUSCA"

Set listaElements = myDoc.getElementsByTagName(txtSearchTags)
For Each i In listaElements
tree1.Nodes.Add "XRAIZ", tvwChild, i.uniqueID, i.tagName & "(" & Len(i.innerHTML) & " chars)"
Next
tree1.Nodes.Item("XRAIZ").Expanded = True
FIN:
Exit Sub
If tree1.Nodes.Item(clave).children > 0 Then Exit Sub
Set chlds = obj.children
For Each i In chlds
    tree1.Nodes.Add clave, tvwChild, i.uniqueID, i.tagName
Next
tree1.Nodes.Item(clave).Expanded = True

End Sub

Private Sub cmdParentNode_Click()
If objeto.parentElement Is Nothing Then
    MsgBox "DosNot Exist"
    Exit Sub
End If
Set objeto = objeto.parentElement

loadFullTree objeto
loadObject objeto
End Sub


Public Sub DoIt(obj As Object)
If obj Is Nothing Then Me.Caption = "No Object Selected ": Exit Sub
loadFullTree obj
loadObject obj
Set objTree = obj
mnuDeleteObject.Caption = "Eliminar " & obj.tagName
Me.Visible = True
'Me.SetFocus
End Sub

Private Sub lv1_BeforeLabelEdit(Cancel As Integer)

Set myLi = lv1.SelectedItem
atributo = myLi.Text

txtValue.Top = lv1.SelectedItem.Top + lv1.Top
txtValue.Left = lv1.ColumnHeaders.Item(2).Left + lv1.Left
txtValue.ZOrder 0
txtValue.Visible = True
txtValue.SetFocus

v = objeto.getAttribute(atributo)

If Not IsNull(v) Then
    txtValue.Text = v
Else
    txtValue.Text = "Null"
End If

txtValue.SelStart = 0
txtValue.SelLength = Len(txtValue)
Label1.Caption = atributo & " = " & txtValue.Text

Cancel = 1
editing = True
End Sub

Private Sub mnuDeleteObject_Click()
'Set xdoc = objTree.Document
If objTree Is Nothing Then Exit Sub
Key$ = objTree.uniqueID
tree1.Nodes.Remove (Key)
objTree.outerHTML = ""

End Sub

Private Sub tree1_Click()
Dim obj As Object
If tree1.SelectedItem Is Nothing Then Exit Sub
Key$ = tree1.SelectedItem.Key
If Left(Key, 1) = "X" Then Exit Sub
Set obj = myDoc.getElementById(Key)
If obj Is Nothing Then Me.Caption = "CANT LOAD...": Exit Sub
loadSubTree obj, Key
loadObject obj
Set objTree = obj
End Sub

Private Sub tree1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu mnOpcion
End If
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
If editing = True And KeyAscii = 13 Then
    newEdit = txtValue.Text
    objeto.setAttribute atributo, newEdit
    myLi.SubItems(1) = objeto.getAttribute(atributo)
    myLi.bold = True
    txtValue.Visible = False
    'txtValue.SetFocus
End If
End Sub

Private Sub txtValue_LostFocus()
txtValue.Visible = False
End Sub

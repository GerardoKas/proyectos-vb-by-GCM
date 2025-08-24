VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SEARCH 'N REPLACE GCM"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frm replace.frx":0000
      Left            =   900
      List            =   "frm replace.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2475
      Width           =   1050
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SEARCH"
      Height          =   420
      Left            =   2925
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2430
      Width           =   915
   End
   Begin VB.CommandButton cmdRestaurar 
      Caption         =   "UNDO"
      Height          =   420
      Left            =   3870
      TabIndex        =   5
      Top             =   2430
      Width           =   735
   End
   Begin VB.CommandButton cmdFiles 
      Caption         =   "FILES"
      Height          =   420
      Left            =   135
      TabIndex        =   3
      Top             =   2430
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   1140
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1260
      Width           =   4380
   End
   Begin VB.TextBox Text1 
      Height          =   1140
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   45
      Width           =   4380
   End
   Begin VB.CommandButton cmdReplace 
      BackColor       =   &H00FFC0C0&
      Caption         =   "REPLACE"
      Height          =   420
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2430
      Width           =   960
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00C0C0FF&
      Caption         =   "DETENER"
      Height          =   420
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2430
      Width           =   1860
   End
   Begin VB.Label lblFicheros 
      AutoSize        =   -1  'True
      Caption         =   "No has Elegido Ficheros"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   2925
      Width           =   1725
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public listado As ListBox
Dim varstop As Boolean
Public activo As Boolean
Dim hecho As Integer, noHecho As Integer
Dim buscar As Boolean
Private Sub cmdBuscar_Click()
buscar = True
cmdReplace_Click
buscar = False
End Sub

Private Sub cmdReplace_Click()
hecho = 0: noHecho = 0
If listado Is Nothing Then MsgBox "NO HAS ELEGIDO FICHEROS": Exit Sub
If listado.ListCount = 0 Then MsgBox "NO HAY FICHEROS ": Exit Sub
ftmVisionado.Visible = True
If buscar = True Then
    text2.Tag = text2.Text
    text2 = Chr(0) & Text1 & Chr(0)
Else
    text2.Tag = text2.Text
End If
cmdReplace.Visible = False
cmdBuscar.Visible = False
cmdStop.Visible = True
For i = 0 To listado.ListCount - 1
    If listado.list(i) <> "" Then
        lblFicheros = "Reemplazando " & i
        DoEvents
        If varstop = True Then varstop = False: Exit For
        reemplazar listado.list(i), Text1.Text, text2.Text
    End If
Next
cmdReplace.Visible = True
cmdBuscar.Visible = True
cmdStop.Visible = False
text2.Text = text2.Tag
MsgBox "COMPLETE - " & vbCrLf & hecho & " COINCIDENCIAS " & vbCrLf & noHecho & " - SIN HALLAR"
lblFicheros.Caption = hecho & " Reemplazos  " & noHecho & " Sin Coincidencias"
End Sub

Private Sub cmdStop_Click()
varstop = True
End Sub

Private Sub cmdFiles_Click()
Form2.Show 1
'If listado Is Nothing Then MsgBox "NO HAS ELEGIDO FICHEROS": Exit Sub
'If listado.ListCount = 0 Then MsgBox "NO HAY FICHEROS ": Exit Sub
lblFicheros = listado.ListCount & " Ficheros Para Hacer Reemplazo (con backup"
End Sub

Function reemplazar(file As String, cad1 As String, cad2 As String) As Boolean
Dim txtstream As TextStream
Dim textoOr As String, textoFin As String
Dim fs As FileSystemObject
Dim rexp As RegExp
Set rexp = New RegExp
rexp.Pattern = cad1
rexp.MultiLine = True
rexp.Global = True
rexp.IgnoreCase = True
Set fs = New FileSystemObject
Set txtstream = fs.OpenTextFile(file, ForReading)
If Not txtstream.AtEndOfStream Then
textoOr = txtstream.ReadAll
txtstream.Close
End If
If (Form1.Combo1.ListIndex = 0) Then
textoFin = Replace(textoOr, cad1, cad2, , , vbTextCompare)
Else
textoFin = rexp.Replace(textoOr, cad2)
End If
If textoOr <> textoFin Then
    If buscar = False Then
        FileCopy file, (file & ".bak")
        On Error GoTo errX
        Set txtstream = fs.OpenTextFile(file, ForWriting)
        txtstream.Write (textoFin)
        txtstream.Close
    End If
    hecho = hecho + 1
    reemplazar = True
    ftmVisionado.lstVisionado.AddItem file, 0
    ftmVisionado.lstVisionado.Selected(0) = True
Else
    noHecho = noHecho + 1
    reemplaz = False
    ftmVisionado.lstVisionado.AddItem file, 0
    ftmVisionado.lstVisionado.Selected(0) = False
End If
errCont:
Exit Function
errX:
MsgBox "ERROR " & Err.Description & " " & Err.Number & vbCrLf & file
GoTo errCont
End Function


Private Sub cmdRestaurar_Click()
If listado Is Nothing Then MsgBox "No hay ficheros en el baúl": Exit Sub

For i = 0 To listado.ListCount - 1
On Error Resume Next
If listado.list(i) <> "" And Dir(listado.list(i) & ".bak") <> "" Then
lblFicheros.Caption = "Restorin " & i
DoEvents
FileCopy (listado.list(i) & ".bak"), listado.list(i)
Kill listado.list(i) & ".bak"
End If
Next
End Sub

Private Sub Form_Load()
Text1.Text = GetSetting(appMe, appSection, "BUSCA", "")
text2.Text = GetSetting(appMe, appSection, "REEMPLAZA", "")
Combo1.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting appMe, appSection, "BUSCA", Text1.Text
SaveSetting appMe, appSection, "REEMPLAZA", text2.Text
'
Unload Form2
Unload Me
End
End Sub

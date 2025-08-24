VERSION 5.00
Begin VB.Form frmSearch 
   Caption         =   "Form2"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form2"
   ScaleHeight     =   1560
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGO 
      Caption         =   "Search!!"
      Height          =   375
      Left            =   45
      TabIndex        =   4
      Top             =   1080
      Width           =   1275
   End
   Begin VB.CheckBox chkdelete 
      Caption         =   "Borrar Antes de Añadir"
      Height          =   330
      Left            =   1935
      TabIndex        =   3
      Top             =   495
      Width           =   1860
   End
   Begin VB.ComboBox cmbTypes 
      Height          =   315
      Left            =   45
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   495
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Folder"
      Height          =   375
      Left            =   3060
      TabIndex        =   1
      Top             =   0
      Width           =   780
   End
   Begin VB.ComboBox cmbDir 
      Height          =   315
      Left            =   45
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   45
      Width           =   2940
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   465
      Left            =   1485
      TabIndex        =   5
      Top             =   1035
      Width           =   2355
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filetypes As String


Private Sub cmbTypes_Click()
Select Case cmbTypes.ListIndex
Case 0
    filetypes = "WAV MP3 WMA ASF"
Case 1
    filetypes = "MID RMI"
Case 2
    filetypes = "MPEG MPG AVI WMV"
Case 3
    filetypes = "MPEG MPG AVI WMV WAV MP3 WMA MID RMI ASF"
End Select
End Sub

Private Sub cmdGO_Click()
Dim unidad As String, raiz As String
Form1.clearList
frmList.clearList
If chkdelete.Value = 1 Then
NumArks = 0
End If
cmbTypes_Click

unidad = getUnidad(cmbDir)
If checkPath(unidad) = False Then
    MsgBox "Esa unidad No existe o daerror" & vbCrLf & unidad
    Exit Sub
End If
raiz = Mid(cmbDir, Len(unidad) + 1, Len(cmbDir) - Len(unidad))
Set Busca.tmpLeyendo = Label1
Busca.BUSKA3 unidad, raiz, True, filetypes, ListaArk2(), NumArks, 0
Form1.AñadirArkivos 'ListaArk2()
frmList.AñadirList
Label1 = "Terminado"
cmbDir.AddItem cmbDir.text
saveFolders
Unload Me
End Sub

Private Sub Form_Load()
cmbTypes.AddItem "AUDIO (WAV,MP3,WMA,ASF)"
cmbTypes.AddItem "MIDI (MID,RMI)"
cmbTypes.AddItem "VIDEO (MPEG,AVI,WMV)"
cmbTypes.AddItem "TODO (AUDIO,VIDEO,MIDI)"
cmbTypes.ListIndex = 3
cmbDir.text = "c:\"
loadFolders
End Sub


Sub saveFolders()
Dim i As Integer
Dim oldText
i = 0
For n = 0 To cmbDir.ListCount - 1
    If cmbDir.List(i) <> oldText Then
    SaveSetting "MPLAY", "DIRS", "Dir" & i, cmbDir.List(i)
    oldText = cmbDir.List(i)
    i = i + 1
   End If
Next
End Sub

Sub loadFolders()
i = 0
dr = GetSetting("MPLAY", "DIRS", "Dir" & i)
While dr <> ""
i = i + 1
cmbDir.AddItem dr
dr = GetSetting("MPLAY", "DIRS", "Dir" & i)
Wend
On Error Resume Next
cmbDir.ListIndex = 0
End Sub

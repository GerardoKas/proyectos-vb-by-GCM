VERSION 5.00
Begin VB.Form frmDropFiles 
   Caption         =   "ELEGIR ARCHIVO"
   ClientHeight    =   1065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1065
   ScaleWidth      =   1560
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUELTA TU FICHERO AQUI"
      Height          =   915
      Left            =   180
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   45
      Width           =   1005
   End
End
Attribute VB_Name = "frmDropFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ficheros As Collection
Public hayFicheros As Boolean

Private Sub Form_Deactivate()
MsgBox "deactivate"
End Sub

Private Sub Form_Load()
Set ficheros = New Collection
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.GetFormat(15) = True Then
For Each i In Data.Files
    ficheros.Add i
Next
Label1.Caption = "LEIDO!!"
hayFicheros = CBool(ficheros.Count)
Unload Me
Else
MsgBox "deben ser ficheros del explorador!", vbExclamation
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Debug.Print "query unload"
End Sub

Private Sub Form_Terminate()
Debug.Print "terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Debug.Print "unload"
End Sub

Private Sub Label1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

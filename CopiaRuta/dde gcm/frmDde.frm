VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ORIGEN"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1965
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   1965
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRecibe 
      Height          =   1005
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   1770
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_LinkOpen(Cancel As Integer)
MsgBox "linkop"
End Sub

Private Sub Form_Load()
    Me.LinkMode = 1 'servidor
    Me.LinkTopic = "MyDDE" 'nombre del vinculo
End Sub

Private Sub txtRecibe_LinkNotify()
MsgBox "Revibioc"
End Sub

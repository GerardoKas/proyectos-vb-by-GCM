VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CALLER"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2205
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   2205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "RECIBIR"
      Height          =   375
      Left            =   1215
      TabIndex        =   2
      Top             =   585
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENVIAR"
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Text            =   "ENIVADORO"
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.LinkPoke
End Sub

Private Sub Command2_Click()
Text1.LinkRequest
End Sub

Private Sub Form_Load()
'linktopic es: appname|nombreTopicdelFrm
'linkitem es cualquier obj del frm
Text1.LinkTopic = "prDdeOrigen|MyDDE"
Text1.LinkItem = "txtRecibe" 'obj al que enviar o recibir
Text1.LinkMode = vbLinkNotify

'con vblinkmanual debera usarse
'obj.linkpoke|linkrequest (enviar|recibir)

'con vblinkautomatic recibe automaticamente
End Sub

Private Sub Text1_LinkNotify()
MsgBox "oooo"
End Sub

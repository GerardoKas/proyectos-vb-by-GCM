VERSION 5.00
Begin VB.Form form2 
   Caption         =   "Opts"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form2"
   ScaleHeight     =   2475
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2040
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "prueba.frx":0000
      Top             =   150
      Width           =   2790
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Simple"
      Height          =   345
      Left            =   75
      TabIndex        =   2
      Top             =   150
      Width           =   1665
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear Image"
      Height          =   315
      Left            =   75
      TabIndex        =   1
      Top             =   600
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear Links"
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   1050
      Width           =   1665
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dropFiles As Collection


Private Sub Form_Initialize()
For Each Item In dropFiles
t = t & Chr(34) & Item & Chr(34) & " "
Next
Text1.Text = t
End Sub


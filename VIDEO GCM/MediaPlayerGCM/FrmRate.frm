VERSION 5.00
Begin VB.Form FrmRate 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RATE"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2310
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   2310
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdreset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   60
      Left            =   120
      Max             =   600
      Min             =   5
      TabIndex        =   0
      Top             =   240
      Value           =   120
      Width           =   1695
   End
   Begin VB.Label lblrate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   90
   End
End
Attribute VB_Name = "FrmRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Elmp As MediaPlayer

Private Sub cmdreset_Click()
HScroll1.Value = 120

End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False

End Sub

Private Sub HScroll1_Change()
Elmp.Rate = HScroll1.Value / 120
lblrate.Caption = Format(CDec(Elmp.Rate), "#.####")
End Sub

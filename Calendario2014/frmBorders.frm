VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBorders 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmBorders.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRematar 
      Caption         =   "Aplicar"
      Height          =   330
      Left            =   3555
      TabIndex        =   12
      Top             =   2700
      Width           =   825
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "Probar"
      Default         =   -1  'True
      Height          =   330
      Left            =   2430
      TabIndex        =   11
      Top             =   2700
      Width           =   960
   End
   Begin VB.HScrollBar scrollMargin 
      Height          =   195
      LargeChange     =   5
      Left            =   135
      Max             =   60
      TabIndex        =   9
      Top             =   2880
      Width           =   1200
   End
   Begin VB.HScrollBar scrollPadding 
      Height          =   195
      LargeChange     =   5
      Left            =   135
      Max             =   60
      TabIndex        =   7
      Top             =   2295
      Width           =   1200
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Color"
      Height          =   285
      Left            =   180
      TabIndex        =   5
      Top             =   1620
      Width           =   1140
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   315
      Width           =   1320
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   195
      LargeChange     =   5
      Left            =   135
      Max             =   60
      TabIndex        =   1
      Top             =   990
      Width           =   1200
   End
   Begin SHDocVwCtl.WebBrowser webX 
      Height          =   2175
      Left            =   2115
      TabIndex        =   0
      Top             =   315
      Width           =   2400
      ExtentX         =   4233
      ExtentY         =   3836
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Margen (Espacio Externo)"
      Height          =   195
      Left            =   45
      TabIndex        =   10
      Top             =   2655
      Width           =   1830
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Padding (Espacio Interno)"
      Height          =   195
      Left            =   45
      TabIndex        =   8
      Top             =   2070
      Width           =   1830
   End
   Begin VB.Label Label3 
      Caption         =   "Color del Borde"
      Height          =   240
      Left            =   45
      TabIndex        =   6
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Grosor del Borde"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   765
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Borde"
      Height          =   195
      Left            =   45
      TabIndex        =   3
      Top             =   90
      Width           =   1005
   End
End
Attribute VB_Name = "frmBorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private myDoc As HTMLDocument
Private myStyleCss As IHTMLStyleSheet
Private myTipoCssRule As HTMLStyleSheetRule
Public objCssRule2 As HTMLStyleSheetRule 'dat cagsad

Private Sub cargaObjeto(obj As HTMLStyleSheetRule)
Set objCssRule2 = obj
'MsgBox obj, vbCritical, "FAIL"
With objCssRule2.Style
    myTipoCssRule.Style.BorderColor = .BorderColor
    myTipoCssRule.Style.BorderWidth = .BorderWidth
    myTipoCssRule.Style.BorderStyle = .BorderStyle
    myTipoCssRule.Style.padding = .padding
    myTipoCssRule.Style.margin = .margin
End With
End Sub
Private Sub cmdAplicar_Click()
objCssRule2.Style.BorderColor = myTipoCssRule.Style.BorderColor
objCssRule2.Style.BorderStyle = myTipoCssRule.Style.BorderStyle
objCssRule2.Style.BorderWidth = myTipoCssRule.Style.BorderWidth
objCssRule2.Style.padding = myTipoCssRule.Style.padding
objCssRule2.Style.margin = myTipoCssRule.Style.margin
End Sub

Private Sub cmdRematar_Click()
cmdAplicar_Click
Unload Me
End Sub

Private Sub Form_Load()
'Dim webX As HTMLDocument
With Combo1
    .AddItem "solid"
    .AddItem "dotted"
    .AddItem "dashed"
    .AddItem "inset"
    .AddItem "outset"
    .AddItem "ridge"
    .AddItem "groove"
End With

'set webx.Document = "tyng"
webX.Navigate2 ("about:blank")
Do While webX.ReadyState <> tagREADYSTATE.READYSTATE_COMPLETE
    DoEvents
Loop
Set myDoc = webX.Document
 'sgBox web1.LocationNam
On Error Resume Next
For i = 0 To 1000
    DoEvents
    Me.Caption = i & " - Calendario"
Next
'hojaestilosfile = "cascadeCalendario.css"
'On Error Resume Next
Set myStyleCss = myDoc.createStyleSheet(".\calendario.css")
myDoc.body.innerHTML = "<table border=1><td>A</td><td>B</td><tr><td>C</td><td>D</td></tr></table>"
myStyleCss.addRule "TD", "background-color:white"
For i = 0 To 100
    DoEvents
Next
'MsgBox myStyleCss.rules.length
Set myTipoCssRule = myStyleCss.rules(0)
cargaObjeto objCssRule2
cargaControles myTipoCssRule
End Sub

Private Sub HScroll1_Change()
On Error Resume Next
myTipoCssRule.Style.BorderWidth = HScroll1.value
End Sub

Private Sub scrollMargin_Change()
myTipoCssRule.Style.margin = scrollMargin
End Sub

Private Sub scrollPadding_Change()
myTipoCssRule.Style.padding = scrollPadding
End Sub

Private Sub cmdColor_Click()
frmColor.Show 1, Me
myTipoCssRule.Style.BorderColor = frmColor.hexColor
End Sub

Private Sub Combo1_Click()
On Error Resume Next

myTipoCssRule.Style.BorderStyle = Combo1.Text
End Sub

Sub cargaControles(obj As HTMLStyleSheetRule)
Combo1.ListIndex = findIndex(Combo1, obj.Style.BorderStyle)
HScroll1.value = Val(obj.Style.BorderWidth)
scrollMargin.value = Val(obj.Style.margin)
scrollPadding.value = Val(obj.Style.margin)
End Sub


Function findIndex(combo As ComboBox, texto As String)
For i = 0 To combo.ListCount - 1
    If texto = combo.List(i) Then
        findIndex = i
        Exit Function
    End If
Next i
findIndex = -1

End Function


VERSION 5.00
Begin VB.Form frmCalText 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmCalText.frx":0000
      Top             =   0
      Width           =   4650
   End
End
Attribute VB_Name = "frmCalText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cal_Text()
myYear = 2005
Dim myDay As Integer 'recorre los 365 dias
For myDay = 1 To 365
xdate = DateSerial(myYear, 1, myDay)
xmonth = DatePart("m", xdate, vbMonday, vbFirstJan1)
xweek = DatePart("w", xdate, vbMonday, vbFirstJan1)
xday = DatePart("d", xdate, vbMonday, vbFirstJan1)
If xday = 1 Then 'dia del mes.
'si es uno se etiqueta un nuevo mes
    mynmonth = MonthName(xmonth)
    addT "", 2
    addT mynmonth, 1
    If xweek <> 1 Then 'si empieza el mes en dia no lunes
        'movemos el cursor hacia ese dia
        For i = 1 To xweek - 1
            addT "  ", 0
        Next i
    End If
End If
If xweek = 1 Then 'dia de la semana lunes...
'si es el primero se inicia una nueva linea
    addT "", 1
End If
addT Format(xday, "0#"), 0
'mywday = WeekdayName(xweek)
'MsgBox mywday
Next myDay

Text1 = myText

End Sub



Sub addT(cadena, separa As Integer)

Select Case separa
    Case 0
        mysep = " "
    Case 1
        mysep = vbCrLf
    Case 2
        mysep = vbCrLf & Space(4)
End Select
myText = myText & cadena & mysep
End Sub

Private Sub Form_Load()
Cal_Text
End Sub

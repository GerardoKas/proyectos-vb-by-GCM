VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCalHTML 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   Icon            =   "frm1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDebug 
      Height          =   2265
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "frm1.frx":038A
      Top             =   315
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   0
      TabIndex        =   3
      Top             =   4590
      Width           =   8520
      Begin VB.ComboBox cmbStyles 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   180
         Width           =   1635
      End
      Begin VB.CommandButton cmdFuente 
         Caption         =   "Fuente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   180
         Width           =   780
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "Color Texto"
         Height          =   285
         Left            =   2610
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Color Fondo"
         Height          =   285
         Left            =   3735
         TabIndex        =   7
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdImagen 
         Caption         =   "Imagen de Fondo"
         Height          =   285
         Left            =   4860
         TabIndex        =   6
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Guardar Calendario"
         Height          =   510
         Left            =   6210
         TabIndex        =   5
         Top             =   675
         Width           =   2130
      End
      Begin VB.CommandButton cmdBordes 
         Caption         =   "Bordes y Espaciado"
         Height          =   285
         Left            =   6345
         TabIndex        =   4
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label lblCombo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   90
         TabIndex        =   12
         Top             =   540
         Width           =   2055
      End
      Begin VB.Label lblButtons 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion del boton"
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   2205
         TabIndex        =   11
         Top             =   540
         Width           =   3840
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Actualizar"
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
   Begin MSComDlg.CommonDialog cmdl 
      Left            =   2790
      Top             =   3870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   4560
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8475
      ExtentX         =   14949
      ExtentY         =   8043
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox txtCss 
      Height          =   4560
      Left            =   3330
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   45
      Width           =   4650
   End
End
Attribute VB_Name = "frmCalHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myHtml As String
Dim myYear As Integer
Dim myColor As String
Dim myFile As String
Dim myFolder As String
Dim myTempFile As String
'Dim myTempFolder As String
Dim myImages As New Collection
Dim numImages As Integer

Dim objStyleSheet As IHTMLStyleSheet
Dim objCssRule As HTMLStyleSheetRule

Const juntarTrimestres = 3
'Dim WithEvents butones As CommandButton
Const doDebug As Boolean = False

Private Sub Cal_Html()
Dim numDias As Long
If myYear = 0 Then myYear = DatePart("yyyy", Now)


Dim myDay As Integer 'recorre los 365 dias
numDias = 365
If DatePart("m", DateSerial(myYear, 2, 29)) = 2 Then
    numDias = 366
End If
'addH myYear, "y"
For myDay = 1 To numDias
xdate = DateSerial(myYear, 1, myDay)
xmonth = DatePart("m", xdate, vbMonday, vbFirstJan1)
xweek = DatePart("w", xdate, vbMonday, vbFirstJan1)
xday = DatePart("d", xdate, vbMonday, vbFirstJan1)
If xweek = 1 Then
'dia lunes... nueva semana
 addH "", "w"
End If
'nuevo mes
If xday = 1 Then
'si es uno se etiqueta un nuevo mes
    mymonthName = MonthName(xmonth)
    addH mymonthName, "m"
    addH "", "w"
    If xweek <> 1 Then 'si empieza el mes en dia no lunes
        'movemos el cursor hacia ese dia
        For i = 1 To xweek - 1
            addH "", "d"
        Next i
    End If
End If
' dia normal o festivo
If xweek = 7 Then
    addH xday, "df"
Else
    addH xday, "d"
End If
'mywday = WeekdayName(xweek)
'MsgBox mywday
Next myDay
addH "", "e"
Do While web1.Busy = True
    DoEvents
Loop

'web1.Document.write myHtml
End Sub


Sub addH(cadena, tipo)
Static beganM As Boolean
Static beganW As Boolean
Static numTrim As Integer
Select Case tipo
Case "y"
    s = "<table><td id=Anno align=center>" & cadena & "</td></table>"
Case "d"
    s = "<td class=Dia align=right>" & cadena & "</td>"
Case "df"
    s = "<td class=Festivo align=right>" & cadena & "</td>"
Case "m"
    If beganM = True Then 'solo ocurre la segunda vez
            'desde la segunda vez hay qye cerrar las anteriores
        If beganW = True Then
            s = s & "</tr>" 'cerramos la semana anterior
            beganW = False
        End If
        s = s & "</table>" & vbCrLf 'cerramos el mes anterior
    End If
    If numTrim Mod juntarTrimestres = 0 Then
        'fin tabla trimestres
        If numTrim > 0 Then
        s = s & "</table>" & vbCrLf
        End If
        'tabla para los trimestres
        s = s & vbCrLf & "<table style='background-color:transparent' class=Trim><tr><td>"
    End If
    s = s & vbCrLf & "<table align=left class=Mes>"
    s = s & "<tr><td colspan=7 align=center class=MesCabecera>" & UCase(cadena) & "</td></tr>"
    s = s & vbCrLf & "<tr>" & diasSemana() & "</tr>" 'los nombres de los dias
    beganM = True
    numTrim = numTrim + 1
Case "w"
    If beganW = True Then
        s = "</tr>"
    End If
    s = s & "<tr class=Semana>"
    beganW = True
Case "e"
    s = "</tr></table></table>"
End Select
myHtml = myHtml & s & vbCrLf
End Sub

Function diasSemana()
s = ""
For i = 1 To 7
    s = s & "<td class=DiaCabecera>" & UCase(Left(WeekdayName(i, True), 2)) & "</td>"
Next
diasSemana = s
End Function

Private Sub cmbStyles_Click()
'DESCRIPCIONES DE AYUDA
Select Case cmbStyles.Text
    Case "Body"
        t = "Se aplica a toda la pagina. Detras de todos los objetos"
    Case "#Anno"
        t = "La cabecera con el año"
    Case ".Mes"
        t = "Se aplica al recuadro del mes. Por detras de los dias. No aplica color de fondo"
    Case ".MesCabecera"
        t = "Nombre del Mes. Se puede poner color de fondo y fuente"
    Case ".Dia"
        t = "Recuadro de Dias del mes, salvo los festivos. Se pueden cambiar todas las propiedades"
    Case ".Festivo"
        t = "Los domingos, se puede tratar igual que los dias"
    Case ".DiaCabecera"
        t = "Los nombres de los dias. Se pueden cambiar todas las propiedades"
    Case ".Semana"
        t = "Las Semanas. No aplica color de fondo ni fuentes, es el contenedor de los dias"
End Select
lblCombo.Caption = t
If objStyleSheet.rules.length > cmbStyles.ListIndex Then
Set objCssRule = objStyleSheet.rules(cmbStyles.ListIndex)
End If
End Sub


Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblButtons.Caption = "Color de Fondo del Objeto. No se aplica a mes, semana. Se puede aplicar en todos los demas"
End Sub

Private Sub cmdBordes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblButtons.Caption = "Aplica Tipo,Color,Grosor de Borde y espaciado del objeto. Se puede aplicar a cualquier objeto"
End Sub

Private Sub cmdColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblButtons.Caption = "Color del texto. No se aplica a los contenedores (mes,semana...). Solo en Dia,Festivo,DiaCabecera y MesCabecera"
End Sub

Private Sub cmdFuente_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblButtons.Caption = "El tipo de fuente, tamaño y grosor"
End Sub

Private Sub cmdImagen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblButtons.Caption = "Imagen para poner en cualquier objeto. Pero mejor en uno de fondo (body,mes,semana...).Es preferible una imagen muy clarita o muy oscura y poner el texto con color adecuado"
End Sub

Private Sub Form_Load()

'myTempFolder = "c:\windows\temp\"
If my = "" Then
    myYear = Year(Now)
Else
    myYear = my
End If
my = InputBox("De que año quieres realizar el Calendario?", "Elegir año para el Calendario", Year(Now) + 1)
myYear = my

myTempFile = "c:\temp\calendario.html"
myHtml = "<html><!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN'>"
myHtml = myHtml & vbCrLf & "<head><title>Calendario de " & myYear & "</title></head>"
myHtml = myHtml & "<body>" & vbCrLf & "<TABLE id=Toda align=center width=100%><caption id=Anno>" & myYear & "</caption><td align=center valign=center>" & vbCrLf
Cal_Html
myHtml = myHtml & vbCrLf & "</td></table>" & vbCrLf & "</body></html>"
savefile myHtml, myTempFile

'myFile = savefile(myHtml, myFile)
web1.Navigate myTempFile
Do While Not web1.ReadyState = READYSTATE_COMPLETE
    DoEvents
Loop
Set objStyleSheet = web1.Document.createStyleSheet("calendario.css")

With cmbStyles
    .AddItem "Body", 0
    .AddItem "#Anno", 1
    .AddItem ".Mes", 2
    .AddItem ".MesCabecera", 3
    .AddItem ".Semana", 4
    .AddItem ".Dia", 5
    .AddItem ".DiaCabecera", 6
    .AddItem ".Festivo", 7
End With
For i = 0 To cmbStyles.ListCount - 1
    cmbStyles.ListIndex = i
    objStyleSheet.addRule cmbStyles.List(i), "color:black"
    setStyle "background-Color", "transparent"
    myImages.Add "", cmbStyles.List(i)
Next
cmbStyles.ListIndex = 0
'frmBorders.Visible = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
web1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - Frame1.Height
Frame1.Move 0, web1.Height, Me.ScaleWidth
End Sub

Private Sub cmdBordes_Click()
Set frmBorders.objCssRule2 = objStyleSheet.rules(cmbStyles.ListIndex)
frmBorders.Show 1, Me
End Sub

'--------------------------------------------------
Private Sub txtCss_DblClick()
objStyleSheet.cssText = txtCss
End Sub

Private Sub txtCss_GotFocus()
txtCss.ZOrder 0
txtCss.Text = objStyleSheet.cssText
End Sub

Private Sub txtCss_LostFocus()
txtCss.ZOrder 1
End Sub
'--------------------------------------------------------

Private Sub cmdFuente_Click()
cmdl.Flags = &H3
cmdl.ShowFont
'aqui da errores de automatzaion
objCssRule.Style.fontWeight = IIf(cmdl.FontBold, "900", "100")
objCssRule.Style.fontFamily = cmdl.FontName
objCssRule.Style.FontSize = cmdl.FontSize
objCssRule.Style.fontStyle = IIf(cmdl.FontItalic, "", "normal")
End Sub

Private Sub cmdColor_Click()
myColor = objStyleSheet.rules(cmbStyles.ListIndex).Style.Color
myColor = cargaColor()
objCssRule.Style.Color = myColor
End Sub

Private Sub cmdBack_Click()
myColor = objStyleSheet.rules(cmbStyles.ListIndex).Style.backgroundColor
myColor = cargaColor()
objCssRule.Style.backgroundColor = "#" & myColor
End Sub

Private Function cargaColor()
frmColor.setColorHTML myColor
'frmColor.loadObj (objStyleSheet.rules(cmbStyles.ListIndex).Style.backgroundColor)
frmColor.Show 1, Me
cargaColor = frmColor.hexColor
End Function


Private Sub cmdImagen_Click()
myimage = cargarImagen()
'myfimage = copiarImagen(myimage, myTempFolder)
If myimage = "" Then Exit Sub
myImages.Remove (cmbStyles.Text)
myImages.Add myimage, cmbStyles.Text
'aqui da eror de auromtaicacion
objCssRule.Style.backgroundImage = "url(" & myimage & ")"
End Sub

Private Sub cmdImagen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
'quitar imagen
On Error Resume Next
myImages.Remove (cmbStyles.Text)
objCssRule.Style.backgroundImage = ""

End If
End Sub

Private Sub cmdUpdate_Click()
objStyleSheet.cssText = objStyleSheet.cssText
web1.Refresh
End Sub

Private Sub cmdSave_Click()
'como esta en codigo en el aire antes hay que guardarlo al disco duro
''''
myFile = getSaveFileName()
If myFile = "" Then
    MsgBox "No se ha guardado el calendario"
    Exit Sub
End If
myFolder = createHTMLFolder(myFile)
saveImages myImages, myFolder  'modifica el css de url image
saveAllPage myFile
web1.Navigate myFile
MsgBox "Se ha guardado el calendario en " & vbCrLf & myFile, vbInformation + vbOKOnly, "Gracias Por Usar Este Programa"

End Sub

Sub saveAllPage(fichero As String)
Dim css
Dim html

html = web1.Document.body.innerHTML
css = objStyleSheet.cssText
filetext = "<style>" & vbCrLf & css & vbCrLf & "</style>" & vbCrLf & html
savefile filetext, fichero
End Sub

Function getSaveFileName()
On Error GoTo ErrName
cmdl.DefaultExt = "*.html"
cmdl.DialogTitle = "Guardar Calendario"
cmdl.Flags = cdlOFNPathMustExist Or cdlOFNExplorer
cmdl.Filter = "Pagina Web Con Imagenes|*.html"
cmdl.FileName = "Calendario " & myYear
cmdl.CancelError = True
cmdl.ShowSave
'If cmdl.Flags And cdlOFNExtensionDifferent Then
'    cmdl.FileName = cmdl.FileName & ".html"
'End If
If cmdl.FileName = "" Then
    getSaveFileName = ""
Else
    getSaveFileName = cmdl.FileName
End If
Exit Function
ErrName:
If Err.Number = 32755 Then
    getSaveFileName = ""
Else
    MsgBox Err.Number & vbCrLf & "Error Al Seleccionar archivo:" & vbCrLf & Err.Description
End If
End Function

Function savefile(texto, fichero)
Dim f As Integer, fname As String
f = FreeFile()

If fichero <> "" Then
    Open fichero For Output As #f
    Print #f, texto
    Close #f
Else
    MsgBox "No se ha guardado el archivo"
End If
savefile = fichero
End Function

Function saveImages(colImages, Directorio)
dirBase = baseName(Directorio)
For i = 0 To cmbStyles.ListCount - 1
    'cmbStyles.ListIndex = i
    file = colImages.Item(cmbStyles.List(i))
    If file <> "" Then
    imname = baseName(file)
    FileCopy file, Directorio & "\" & imname
    setStyle "background-image", "url(" & dirBase & "\" & imname & ")"
    End If
Next
End Function

Function baseName(file)
pos = InStrRev(file, "\")
baseName = Right(file, Len(file) - pos)
End Function

Function copiarImagen(imagen, Directorio)
pos = InStrRev(imagen, "\")
copiarImagen = Right(imagen, Len(imagen) - pos)
FileCopy imagen, Directorio & copiarImagen
End Function

Function createHTMLFolder(file)
pospunto = InStrRev(file, ".")
base = Left(file, pospunto - 1)
createHTMLFolder = base & "_archivos"
If Dir$(createHTMLFolder, vbDirectory) = "" Then
    MkDir createHTMLFolder
End If
End Function

Private Sub web1_TitleChange(ByVal Text As String)
Me.Caption = Text
End Sub

Function setStyle(atributo, value)
On Error GoTo UnError
If cmbStyles.Text = "" Then
    MsgBox "Debes seleccionar una categoria en el recuadro"
    Exit Function
Else
    obj = cmbStyles.Text
    objindex = cmbStyles.ListIndex
End If
    objStyleSheet.rules(objindex).Style.setAttribute atributo, value
objStyleSheet.cssText = objStyleSheet.cssText
txtCss = objStyleSheet.cssText
If doDebug = True Then
    txtDebug.Visible = True
    txtDebug.Text = txtDebug.Text & vbCrLf & atributo & ": " & value
End If
Exit Function
UnError:
MsgBox "Ha ocurrido un error no controlado : " & vbCrLf & Err.Description, vbOKCance + vbCritical
End Function
Function removeImage()

If cmbStyles.Text = "" Then
    MsgBox "Debes seleccionar una categoria en el recuadro"
    Exit Function
Else
    obj = cmbStyles.Text
    objindex = cmbStyles.ListIndex
End If
    objStyleSheet.rules(objindex).Style.backgroundImage = ""
objStyleSheet.cssText = objStyleSheet.cssText
txtCss = objStyleSheet.cssText
End Function

Function cargarImagen()
cmdl.Filter = "Imagenes jpg gif bmp png|*.jpg;*.bmp;*.gif;*.png"
cmdl.DialogTitle = "Cargar Imagen de Fondo"
cmdl.ShowOpen
If cmdl.FileName <> "" Then
    cargarImagen = cmdl.FileName
Else
    cargarImagen = ""
End If
End Function



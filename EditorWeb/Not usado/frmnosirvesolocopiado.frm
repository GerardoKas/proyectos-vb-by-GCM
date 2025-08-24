VERSION 5.00
Begin VB.Form frmnosirvesolocopiado 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmnosirvesolocopiado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
myTempFile = "c:\windows\temp\cal.html"
'myTempFolder = "c:\windows\temp\"
my = InputBox("De que año quieres realizar el Calendario?", "Elegir año para el Calendario", Year(Now))
If my = "" Then
    myYear = Year(Now)
Else
    myYear = my
End If
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
Set stylex = web1.Document.createStyleSheet("calendario.css")

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
    stylex.addRule cmbStyles.List(i), "color:#000000"
    setStyle "background-Color", "transparent"
    myImages.Add "", cmbStyles.List(i)
Next
cmbStyles.ListIndex = 0
'frmBorders.Visible = True
End Sub

Private Sub Form_Resize()
web1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - Frame1.Height
Frame1.Move 0, web1.Height, Me.ScaleWidth
End Sub

Private Sub cmdBordes_Click()
Set frmBorders.objetoCss = stylex.rules(cmbStyles.ListIndex)
frmBorders.Show 1, Me
End Sub

'--------------------------------------------------
Private Sub txtCss_DblClick()
stylex.cssText = txtCss
End Sub

Private Sub txtCss_GotFocus()
txtCss.ZOrder 0
txtCss.Text = stylex.cssText
End Sub

Private Sub txtCss_LostFocus()
txtCss.ZOrder 1
End Sub
'--------------------------------------------------------

Private Sub cmdFuente_Click()
cmdl.Flags = &H3
cmdl.ShowFont
objCss.Style.fontWeight = IIf(cmdl.FontBold, "900", "100")
objCss.Style.fontFamily = cmdl.FontName
objCss.Style.FontSize = cmdl.FontSize
objCss.Style.fontStyle = IIf(cmdl.FontItalic, "", "normal")
End Sub

Private Sub cmdColor_Click()
myColor = stylex.rules(cmbStyles.ListIndex).Style.Color
myColor = cargaColor()
objCss.Style.Color = myColor
End Sub

Private Sub cmdBack_Click()
myColor = stylex.rules(cmbStyles.ListIndex).Style.backgroundColor
myColor = cargaColor()
objCss.Style.backgroundColor = myColor
End Sub

Private Function cargaColor()
frmColor.setColorHTML myColor
'frmColor.loadObj (stylex.rules(cmbStyles.ListIndex).Style.backgroundColor)
frmColor.Show 1, Me
cargaColor = frmColor.hexColor
End Function


Private Sub cmdImagen_Click()
myimage = cargarImagen()
'myfimage = copiarImagen(myimage, myTempFolder)
If myimage = "" Then Exit Sub
myImages.Remove (cmbStyles.Text)
myImages.Add myimage, cmbStyles.Text
objCss.Style.backgroundImage = "url(" & myimage & ")"
End Sub

Private Sub cmdImagen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
'quitar imagen
On Error Resume Next
myImages.Remove (cmbStyles.Text)
objCss.Style.backgroundImage = ""

End If
End Sub

Private Sub cmdUpdate_Click()
stylex.cssText = stylex.cssText
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
html = web1.Document.body.innerHTML
css = stylex.cssText
filetext = "<style>" & vbCrLf & css & vbCrLf & "</style>" & vbCrLf & html
savefile filetext, fichero
End Sub

Function getSaveFileName()
On Error GoTo ErrName
cmdl.DefaultExt = "*.html"
cmdl.DialogTitle = "Guardar Calendario"
cmdl.Flags = cdlOFNPathMustExist Or cdlOFNExplorer
cmdl.Filter = "Pagina Web Con Imagenes|*.html"
cmdl.filename = ""
cmdl.CancelError = True
cmdl.ShowSave
'If cmdl.Flags And cdlOFNExtensionDifferent Then
'    cmdl.FileName = cmdl.FileName & ".html"
'End If
If cmdl.filename = "" Then
    getSaveFileName = ""
Else
    getSaveFileName = cmdl.filename
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
    stylex.rules(objindex).Style.setAttribute atributo, value
stylex.cssText = stylex.cssText
txtCss = stylex.cssText
If doDebug = True Then
    txtDebug.Visible = True
    txtDebug.Text = txtDebug.Text & vbCrLf & atributo & ": " & value
End If
Exit Function
UnError:
MsgBox "Ha ocurrido un error no controlado : " & vbCrLf & Err.Description, vbOKCance + vbCritical
End Function


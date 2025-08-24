VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Gcm Web Editor"
   ClientHeight    =   4905
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5835
   Icon            =   "editorWeb.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4905
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtContents 
      Height          =   1590
      Left            =   45
      TabIndex        =   7
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2805
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"editorWeb.frx":038A
   End
   Begin VB.Frame frmDisplace 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   90
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   1
      Top             =   1845
      Width           =   5715
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   0
      TabIndex        =   2
      Top             =   4185
      Width           =   5775
      Begin VB.CommandButton cmdDropObjects 
         Appearance      =   0  'Flat
         Caption         =   "INSERTAR IMAGENES"
         Height          =   285
         Left            =   1575
         OLEDropMode     =   1  'Manual
         TabIndex        =   8
         Top             =   360
         Width           =   2670
      End
      Begin VB.TextBox lblTooStat 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   2610
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   90
         Width           =   2400
      End
      Begin VB.TextBox lblStatus 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   0
         TabIndex        =   5
         Text            =   "Stautssss"
         Top             =   90
         Width           =   2580
      End
      Begin VB.ComboBox cmbEditTags 
         Height          =   315
         Left            =   4365
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1365
      End
      Begin VB.CommandButton cmdPropiedades 
         Caption         =   "Object Properties"
         Height          =   285
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   1500
      End
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   5445
      Top             =   1575
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "UNKNOWN.HTML"
      InitDir         =   "c:\"
   End
   Begin MSScriptControlCtl.ScriptControl scr 
      Left            =   5400
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      Language        =   "JScript"
      AllowUI         =   -1  'True
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   2235
      Left            =   0
      TabIndex        =   0
      Top             =   1935
      Width           =   5715
      ExtentX         =   10081
      ExtentY         =   3942
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
      Location        =   "http:///"
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   1005
      Left            =   4590
      Top             =   0
      Width           =   1230
   End
   Begin VB.Menu mfile 
      Caption         =   "Fichero"
      Begin VB.Menu mnLoad 
         Caption         =   "Load"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnLoadURL 
         Caption         =   "Load URL"
      End
      Begin VB.Menu mnuOpenInIE 
         Caption         =   "Editar en el Navegador (yunow)"
      End
      Begin VB.Menu mnSave 
         Caption         =   "Save encima"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnSaveAs 
         Caption         =   "Save As..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnSaveComplete 
         Caption         =   "Save Complete web Page"
      End
   End
   Begin VB.Menu mnEdicion 
      Caption         =   "Edicion"
      Begin VB.Menu mnuReemplazar 
         Caption         =   "Reemplazar"
      End
      Begin VB.Menu mnuEditableBody 
         Caption         =   "Hacer BODY EDITABLE"
      End
      Begin VB.Menu mnuColorRTF 
         Caption         =   "COLOREA"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBuscar 
         Caption         =   "Buscar ..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuBuscarSig 
         Caption         =   "Buscar Siguiente"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnop 
      Caption         =   "DOCUMENTO"
      Begin VB.Menu mnuStopWeb 
         Caption         =   "Stop Page Load"
      End
      Begin VB.Menu mnGetSourceText 
         Caption         =   "Recargar Pagina Web"
      End
      Begin VB.Menu mnGetWebText 
         Caption         =   "Recargar Codigo Fuente"
      End
      Begin VB.Menu mnuBasePath 
         Caption         =   "Ruta Relativa de Base"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Styles"
      Begin VB.Menu mnuTableBorder 
         Caption         =   "Ver Tablas Marcadas"
      End
      Begin VB.Menu mnuVerHojasEstilo 
         Caption         =   "Ver Hojas de Estilo"
      End
   End
   Begin VB.Menu mnuScript 
      Caption         =   "Scripts"
      Begin VB.Menu mnuDisableScripts 
         Caption         =   "Disable All Scripts"
      End
      Begin VB.Menu mnuDeleteScripts 
         Caption         =   "Delete All Scripts"
      End
      Begin VB.Menu mnuNull 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecorrerJs 
         Caption         =   "Recorrer Objeto JS"
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "Insertar"
      Begin VB.Menu mnuImage 
         Caption         =   "Imagen"
      End
      Begin VB.Menu mnuHRule 
         Caption         =   "Hor. Rule"
      End
      Begin VB.Menu mnuIFrame 
         Caption         =   "Contenedor IFrame"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents myBody As HTMLBody
Attribute myBody.VB_VarHelpID = -1
Public WithEvents myWindow As HTMLWindow2
Attribute myWindow.VB_VarHelpID = -1

Dim textActive As Boolean
Dim ObjsDone As Boolean


Private Sub cmbEditTags_Click()
Dim sel As IHTMLTxtRange, newSel As IHTMLTxtRange
Dim obj As IHTMLElement
Dim slc As IHTMLSelectionObject
If cmbEditTags.ListIndex <= 0 Then Exit Sub
    On Error GoTo ICant
    
    Set sel = myDoc.selection.createRange
    If sel.Text = "" Then
    MsgBox "No hay texto en la seleccion!"
    Exit Sub
    End If
    'como hacer para que capte los tags circundantes en su solo uso
    sel.Expand 2
    tagcode = getTagCode(cmbEditTags.Text)
       
    X = sel.getBookmark
    
    If tagcode = "%CLEAR%" Then
     mreplaced$ = Replace(sel.Text, vbCrLf, "<BR>")
     
    Else
    
    mreplaced$ = Replace(tagcode, "%T%", sel.htmlText)
    End If
    sel.pasteHTML mreplaced
    lblTooStat = mreplaced
    
    'con esto lo vuelve a seleccionar
    Set newSel = myBody.createTextRange
    'newSel.moveToElementText obj
    If newSel.moveToBookmark(X) Then
    'MsgBox newSel.Text
    newSel.Expand 1
    newSel.Select
    End If
cmbEditTags.ListIndex = -1
web1.SetFocus
Exit Sub
ICant:
MsgBox Err.description, vbExclamation, Err.Source
End Sub


Private Sub cmdDropObjects_Click()
MsgBox "Suelta una imgen encima del boton y se añadaria la imagen en la posicion actual"

End Sub

Private Sub cmdDropObjects_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim thisChild As IHTMLDOMNode, objChild As IHTMLDOMNode
'MsgBox "se insertara el objeto en la posicion elegida"
'Set thisChild = myDoc.activeElement.children(1)
'Set objChild = myDoc.createTextNode("XX-" & Data.Files.Item(1) & "-XX")
'myDoc.insertBefore objChild, thisChild
myDoc.execCommand "InsertImage", False, Data.Files.Item(1)
lblTooStat.Text = "Imagen Insertada!!!"
End Sub

Private Sub Form_Load()
web1.Offline = False
Set GCM_RTF.rtf = txtContents
modTextEditor.tempFile = App.path & "\" & "TEMPFILE.html"
Set fso = New FileSystemObject
cmdlg.DefaultExt = "*.html"
cmdlg.Filter = "WebPage (*.html;*.htm;*.mht)|*.html;*.htm;*.mht|Javascript / VBS(*.js;*.vbs)|*.js;*.vbs|Css (*.css)|*.css|Tutti (*.*)|*.*"

If fso.FileExists(App.path & "\" & modTextEditor.tempFile) Then
    If MsgBox("EXISTE UN FICHERO TEMPORAL ANTIGUO. PROBABLEMENTE PARECIO BORRARSE DESEAS RECUPERARLO?", vbYesNo + vbExclamation) = vbYes Then
        Shell "explorer.exe """ & App.path & """", vbNormalFocus
    End If
End If
web1.Navigate "about:blank"
OriginalFile = "NONELOADED"
loadTags cmbEditTags, "EditTags.ini"

End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data.Files.Count > 0 Then
OriginalFile = Data.Files(1)
getTempName (OriginalFile)
makeBackup OriginalFile, tempFile
showTextFile tempFile
showWebFile tempFile
End If
End Sub

Private Sub Form_Resize()
Dim spaceInf As Long
esp = 80
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
Frame2.Top = Form1.ScaleHeight - Frame2.Height
spaceInf = Form1.ScaleHeight - Frame2.Top
txtContents.Left = esp
web1.Left = esp
txtContents.Height = frmDisplace.Top - esp
txtContents.Width = Form1.ScaleWidth - esp * 2

web1.Top = frmDisplace.Top + frmDisplace.Height
web1.Width = Form1.ScaleWidth - esp * 2
web1.Height = Form1.ScaleHeight - web1.Top - spaceInf - esp
frmDisplace.Width = Form1.ScaleWidth - esp
If Me.ActiveControl = txtContents Then
    Shape1.Move txtContents.Left - 5, txtContents.Top - 5, txtContents.Width + 10, txtContents.Height + 10
ElseIf Me.ActiveControl = web1 Then
    Shape1.Move web1.Left - 5, web1.Top - 5, web1.Width + 10, web1.Height + 10
End If
Frame2.Width = Me.ScaleWidth
lblStatus.Width = Me.ScaleWidth / 2 - esp
lblTooStat.Left = lblStatus.Width + esp
lblTooStat.Width = Me.ScaleWidth / 2 - esp
End Sub
'=================================================
'BARRITY AMOVIL
Private Sub frmDisplace_DblClick()
frmDisplace.Top = Form1.ScaleHeight / 3
Form_Resize
End Sub
Private Sub frmDisplace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
frmDisplace.Top = Y + frmDisplace.Top
If (frmDisplace.Top < 0) Then frmDisplace.Top = 20
If (frmDisplace.Top > Form1.ScaleHeight) Then frmDisplace.Top = Form1.ScaleHeight - frmDisplace.Height
frmDisplace.Left = 0
End If
End Sub
Private Sub frmDisplace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Form_Resize
End Sub
'======================================

Private Sub mnGetSourceText_Click()
saveTextFile tempFile, txtContents.Text
showWebFile tempFile

End Sub

Private Sub mnGetWebText_Click()
'If tempFile = "" Then getTempName
saveTextFile tempFile, getWebHtml 'replacePaths(getWebHtml)
showTextFile tempFile
End Sub

Private Sub mnLoad_Click()
cmdlg.DialogTitle = "Abrir fichero para editar"
cmdlg.ShowOpen
If cmdlg.CancelError = True Or cmdlg.filename = "" Then Exit Sub
OriginalFile = cmdlg.filename
getTempName (OriginalFile)
makeBackup OriginalFile, tempFile
'abriendo el backup. siempre trabajamos sobtre el backup
'solo al salvar se caambia el original
showTextFile tempFile
showWebFile tempFile
End Sub

Private Sub mnLoadURL_Click()
turl = InputBox("Direccion Web PAra Abrir?", , "about:blank")
showWebFile turl
End Sub

Private Sub mnSave_Click()
perderFoco
If OriginalFile = "" Then
makeSaveAs
Else
saveTextFile OriginalFile, txtContents.Text
lblTooStat = "Guardado sobre :" & OriginalFile
End If
End Sub

Private Sub mnSaveAs_Click()
makeSaveAs
End Sub

Sub makeSaveAs()
'On Error Resume Next
Dim oldName As String
perderFoco
cmdlg.DialogTitle = "Guardar Fichero en..."
cmdlg.filename = OriginalFile
cmdlg.ShowSave

If cmdlg.CancelError = True Or cmdlg.filename = "" Then
Exit Sub
End If
oldName = OriginalFile
OriginalFile = cmdlg.filename

makeBackup oldName, oldName & ".BAK"
saveTextFile OriginalFile, txtContents.Text
lblTooStat = "Guardado en :" & OriginalFile
End Sub

Private Sub mnSaveComplete_Click()
myDoc.execCommand "SaveAs", True
End Sub



Private Sub mnuBasePath_Click()
X = InputBox("Ruta de Base Del Documento" & vbCrLf & "En base a la cual estarán los archivos anexos", "Indicar BasePath", basepath)
basepath = X
End Sub

Private Sub mnuBuscar_Click()
wd = InputBox("Palabra Para Buscar", "Buscar Palabra ", "</BODY>", 0, 0)
GCM_RTF.searchWord (wd)
End Sub

Private Sub mnuBuscarSig_Click()
GCM_RTF.searchNext
End Sub

Private Sub mnuColorRTF_Click()
colorTAGS
End Sub

Private Sub mnuDeleteScripts_Click()
web1.SetFocus
Dim i As IHTMLDOMNode
Set scrs = myDoc.getElementsByTagName("SCRIPT")
For Each i In scrs
i.removeNode True

Next
'Debug.Print myDoc.scripts.length
perderFoco

End Sub


Private Sub mnuDisableScripts_Click()
web1.SetFocus
DoEvents
For Each sc In myDoc.scripts
    sc.disabled = True
Next
End Sub

Private Sub mnuEditableBody_Click()
myDoc.body.setAttribute "contentEditable", "true"
web1.SetFocus
End Sub

Private Sub mnuHRule_Click()
web1.SetFocus
myDoc.execCommand "InsertHorizontalRule", False

End Sub

Private Sub mnuIFrame_Click()
web1.SetFocus
myDoc.execCommand "InsertIFrame", True

End Sub

Private Sub mnuImage_Click()
web1.SetFocus
myDoc.execCommand "InsertImage", True

End Sub

Private Sub mnuOpenInIE_Click()
perderFoco
'Shell "rundll32.exe shdocvw.dll,OpenURL " & tempFile & ""
Shell """C:\Archivos de programa\Internet Explorer\iexplore.exe"" """ & tempFile & """", vbNormalFocus
End Sub

Private Sub mnuRecorrerJs_Click()
On Error GoTo ErrScr
scr.Reset
scr.Language = "JScript"
scr.AddObject "document", web1.Document, True
scr.AddObject "window", web1.Document.parentWindow, True

code = getTextFile("recorrerJs.js")
scr.AddCode code
'scr.Run "", ""
Exit Sub
ErrScr:
msg = "Creo Que el script da Error : " & vbCrLf & "Linea: " & scr.Error.line & vbCrLf & scr.Error.Number & " : " & scr.Error.description & vbCrLf & "Text: " & scr.Error.Text
MsgBox msg
lblTooStat = msg
End Sub

Private Sub mnuStopWeb_Click()
web1.Stop
perderFoco
DoEvents
End Sub

Private Sub mnuTableBorder_Click()
frmProceso.Visible = True
Exit Sub
Dim style As IHTMLStyleSheet
Set style = myDoc.createStyleSheet("NoExiste.css")
style.disabled = True
style.cssText = getTextFile("bordesTablas.css")
style.disabled = False
web1.SetFocus
End Sub

Private Sub mnuVerHojasEstilo_Click()
frmCSS.Visible = True
frmCSS.SetFocus
End Sub


'Private Function myBody_onselectstart() As Boolean
'Dim evt As CEventObj, rng As IHTMLTxtRange
'Set evt = myWindow.event
''MsgBox evt.srcElement.tagName
'myBody_onselectstart = False
''evt.srcElement.replaceAdjacentText "afterBegin", "HOLA"
''Set rng = myBody.createTextRange()
''MsgBox rng.Text
'End Function

Private Sub txtContents_GotFocus()
Shape1.Visible = True
Form_Resize
End Sub

Private Sub txtContents_lostFocus()
'recargar web1 con codigo fuente
'If ObjsDone = False Then Exit Sub
Shape1.Visible = False
mnGetSourceText_Click
End Sub


Private Sub web1_GotFocus()
Shape1.Visible = True
Form_Resize
End Sub

Private Sub web1_lostFocus()
'recargar codigo fuente con wev1
Shape1.Visible = False
mnGetWebText_Click
End Sub

'Private Sub txtContents_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'If Data.GetFormat(vbCFFiles) = True Then
''meterFiles Data.Files
'ElseIf Data.GetFormat(vbCFRTF) = True Then
''insertar (Data.GetData(vbCFText))
'ElseIf Data.GetFormat(vbCFText) = True Then
''insertar (Data.GetData(vbCFText))
'End If
'End Sub

'------------------------------------------------------
'----------FUNCIONES CREADAS---------------------------
Private Sub web1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'ultimo en accionarse debe esperar por el javascript
Dim t As String, tDone As String
'lblTooStat.Text = "NOT LOADED"
Select Case web1.ReadyState
Case READYSTATE_UNINITIALIZED
t = "NO INICIADO"
Case READYSTATE_LOADED
t = "CARGADO"
Case READYSTATE_LOADING
t = "CARGANDO"
Case READYSTATE_INTERACTIVE
t = "INTERACTIVO"
Case READYSTATE_COMPLETE
t = "COMPLETO"
End Select
If web1.ReadyState >= READYSTATE_LOADED Then
    Do While web1.Document Is Nothing
        DoEvents
        n = n + 1
        Me.Caption = "Waiting " & n
    Loop
    Set myDoc = web1.Document
    Set myWindow = myDoc.parentWindow
    Set myBody = myDoc.body
    ObjsDone = True
    basepath = myDoc.URL
    
    If OriginalFile = "" Then
        OriginalFile = myDoc.location.href
    End If
tDone = "CARGADO OBJETOS"
End If
lblTooStat = t & " - " & tDone
Me.Caption = myDoc.Title
End Sub

Private Sub web1_StatusTextChange(ByVal Text As String)
If Text = "Listo" Then Exit Sub
lblStatus = Text
End Sub

Private Sub web1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
If Progress > 0 Then
lblStatus = "Doing : " & Progress & "/" & ProgressMax & " - "
End If
End Sub
'
'Private Sub web1_StatusTextChange(ByVal Text As String)
'lblStatus = Text
'End Sub


Private Function myBody_onclick() As Boolean
Set myelement = myEvent.srcElement
myId = myelement.Id
If myId = "" Then
    myId = myelement.uniqueID
End If
cmdPropiedades.Caption = myelement.tagName & ", " & myId

End Function

Private Function myBody_ondrop() As Boolean
MsgBox myEvent.fromElement
'MsgBox myBody.event

End Function

Private Sub myBody_onselect()
Set sel = myDoc.selection.createRange
lblTooStat = sel.Text
End Sub

Private Sub myWindow_onerror(ByVal description As String, ByVal URL As String, ByVal line As Long)
lblTooStat = "Error Encontrado en la Página : " & vbCrLf & description & vbCrLf & vbCrLf & URL & vbCrLf & "Linea : " & line ', vbExclamation + vbOKOnly, "Error del Programador De la Pagina Web"
End Sub

Private Sub cmdPropiedades_Click()
frmProp.DoIt myelement
End Sub


Sub perderFoco()
lblTooStat.SetFocus
DoEvents
End Sub

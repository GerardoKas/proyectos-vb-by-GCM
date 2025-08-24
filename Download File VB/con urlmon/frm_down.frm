VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Downloader By Gcm 2005"
   ClientHeight    =   3495
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9090
   Icon            =   "frm_down.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   606
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   645
      Left            =   4680
      TabIndex        =   10
      Top             =   2790
      Width           =   4380
      Begin VB.CommandButton cmdREadEmbeds 
         Caption         =   "EMBEDS"
         Height          =   330
         Left            =   1980
         TabIndex        =   14
         Top             =   135
         Width           =   960
      End
      Begin VB.CommandButton cmdReadAll 
         Caption         =   "LINKS"
         Height          =   330
         Left            =   0
         TabIndex        =   13
         Top             =   135
         Width           =   870
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Parar_YA"
         Height          =   330
         Left            =   3195
         TabIndex        =   12
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdReadImages 
         Caption         =   "IMAGES"
         Height          =   330
         Left            =   945
         TabIndex        =   11
         Top             =   135
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   645
      Left            =   45
      TabIndex        =   5
      Top             =   2835
      Width           =   4560
      Begin VB.TextBox txtDest 
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Text            =   "C:\"
         Top             =   0
         Width           =   2640
      End
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H00FF8080&
         Caption         =   "Ejecutar"
         Height          =   315
         Left            =   3495
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1020
      End
      Begin VB.CommandButton cmdDestino 
         Caption         =   "Destino"
         Height          =   315
         Left            =   2625
         TabIndex        =   6
         Top             =   0
         Width           =   810
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   240
         Left            =   0
         TabIndex        =   9
         Top             =   375
         Width           =   4515
      End
   End
   Begin VB.CommandButton cmdLeech 
      Caption         =   "Abrir..."
      Height          =   300
      Left            =   3825
      TabIndex        =   4
      Top             =   75
      Width           =   735
   End
   Begin SHDocVwCtl.WebBrowser wb1 
      Height          =   3030
      Left            =   4680
      TabIndex        =   3
      Top             =   90
      Width           =   4335
      ExtentX         =   7646
      ExtentY         =   5345
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
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   300
      Left            =   3240
      TabIndex        =   2
      Top             =   75
      Width           =   540
   End
   Begin VB.ListBox List1 
      Height          =   2310
      Left            =   75
      OLEDropMode     =   1  'Manual
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   450
      Width           =   4515
   End
   Begin VB.TextBox txturl 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Text            =   "http://localhost/phpinfo.php"
      Top             =   75
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   6
      X2              =   303
      Y1              =   27
      Y2              =   27
   End
   Begin VB.Menu mn_CapLista 
      Caption         =   "Listado"
      Begin VB.Menu mnuAutoClean 
         Caption         =   "AutoLimpiar Completados"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuNull1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLimpiarComp 
         Caption         =   "Limpiar Completados"
      End
      Begin VB.Menu mnuLimpiarDesmarcados 
         Caption         =   "Limpiar Desmarcados"
      End
      Begin VB.Menu mnuLimpiarTODO 
         Caption         =   "Limpiar TODO"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuSaveURLList 
         Caption         =   "Guardar Lista URLs"
      End
      Begin VB.Menu mnuLoadURLs 
         Caption         =   "Cargar Lista URLs"
      End
      Begin VB.Menu mnuNone 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenDest 
         Caption         =   "Abrir Carpeta Destino"
      End
      Begin VB.Menu mnuNull 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveWithTitle 
         Caption         =   "Guardar con Titulo de la Pagina"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim autoclean As Boolean
Dim bNavigated As Boolean
Dim makeLeech As Boolean
Dim bHaciendo As Boolean
Dim doStop As Boolean
Dim doNoop As Boolean
Dim numURLS As Long

Public Function DownloadFile(URL As String, LocalFilename As String) As String
Dim mensaje As String
'MsgBox "Iniciamos"
DoEvents
URLDownloadError = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
DoEvents
If URLDownloadError <> 0 Then
mensaje = GetDownloadErrorDescription
Else
mensaje = msgOk
End If
DownloadFile = mensaje
End Function

Sub EXAMPLE_GET_THE_FILE()
Dim URL As String
Dim LocalFilename As String
URL = "http://localhost/index.html"
LocalFilename = "C:\local.txt"
Debug.Print DownloadFile(URL, LocalFilename)
End Sub



Private Sub cmdGo_Click()
If cmdGo.Caption = "Ejecutar" Then
    cmdGo.Caption = "DETENER"
    cmdGo.BackColor = vbRed
    Me.ScaleMode = vbPixels
    Me.ScaleWidth = meWidth
    Me.ScaleHeight = meHeight
    doStop = False
    DownIt
Else
    cmdGo.Caption = "Ejecutar"
    cmdGo.BackColor = &HFF8080
    doStop = True
End If
End Sub

Public Sub DownIt()

Dim URL As String
Dim nTotal As Integer
Dim nOk As Integer
Dim nError As Integer
Dim r As VbMsgBoxResult
Dim fName As String, ftitulo As String, tempname As String
Dim filesys As Object
On Error Resume Next
'TOTAL ESTA EN LA COL ESTA
nTotal = FileList.Count
'CREAR RUTA CORRECTA
sDestino = txtDest.Text
If Right(sDestino, 1) <> "\" Then sDestino = sDestino & "\"
If Dir$(sDestino & "nul") = "" Then
MsgBox "Esa carpeta no existe. Debes Crearla Antes!", vbExclamation
Exit Sub
End If
MsgBox "Se minimizara la aplicacion mientras trabaja. Al terminar volvera a verse", vbInformation, "Preparando"
Me.WindowState = 1
Me.MousePointer = vbHourglass
ChDir (sDestino)
'RECORRER ARCHIVOS PARA BAJAR
For i = 1 To FileList.Count
    URL = FileList.item(i)
    fName = baseName(URL)
    If fName = "" Then fName = "Noname-" & i & ".xxx"
    'HACER COMPROBACION BASICA DE CORRECCION DE FILENAME
    fName = Replace(fName, "?", "!")
    'poosicionar en el actual
    doNoop = True 'variable para que no haga cosa rara al click
    List1.ListIndex = i - 1
    DoEvents
    'SI ESTA MARCADO LO BAJAMOS ELSE (MAS ABAJO)
    If (List1.Selected(i - 1) = True) Then
        'TITULO DE ESTADO
        nahora = nahora + 1
        Me.Caption = nahora & "/" & FileList.Count & " - " & fName
        lblStatus = "Down: " & fName
        'SI EL FICHERO YA EXISTE PREGUNTAR... o no....
        'preguntaSobreescribir sDestino, fName
        DoEvents
        fName = getFreeNumber(fName)
        If doStop = True Then
            MsgBox "Detenido De Repente .", vbCritical
            Exit Sub
        End If
        'PROCEDEMOS A BAJAR EL FICHERO !ahora !ahora!
        msg = DownloadFile(URL, sDestino & fName)
        'COMPROBAMOS QUE BAJO BIEN Y SUMAMOS UN POROTO
        If msg = msgOk Then
            nOk = nOk + 1
        Else
            nError = nError + 1
        End If
        'PONEMOS ESTADO AL LISTBOX
        List1.List(i - 1) = fName & msgSep & msg
        DoEvents
    Else
    'SI ESTABA DESMARCADO LE PONEMOS OMITIDO
        List1.List(i - 1) = fName & " ! Omitido"
    End If
Next
If mnuAutoClean.Checked = True Then
    mnuLimpiarComp_Click
End If
Me.MousePointer = vbArrow
lblStatus = "Total:" & nTotal & ", Correcto:" & nOk & ", Error:" & nError
Me.Caption = myCaption
Me.WindowState = 0
End Sub

Sub preguntaSobreescribir()
        If (Dir$(sDestino & fName) <> "") Then
            r = MsgBox("El Archivo ya existe : " & vbCrLf & sDestino & fName, vbExclamation + vbYesNoCancel, "Sobreescribir?")
            Select Case r
            'SI NO QUISO SEGUIR
            Case vbCancel
                MsgBox "Detenido en " & fName
                Exit Sub
            Case vbNo
            'SI QUIERE CAMBIAR DE FILENAME
                newname = InputBox("Introduce el nuevo nombre de fichero para que no sobreescriba al anterior", "Cambiar Nombre de fichero")
                If (newname) = "" Then
                    MsgBox "Abortado"
                    lblStatus = "abortado en " & fName
                    Exit Sub
                End If
                fName = newname
            Case vbYes
            'SI QUIERE SOBREESCRIBIR
            End Select
        End If
End Sub
Private Sub cmdAdd_Click()
'AÑADIR EL ITEM ESCRITO
addtoList txturl.Text
lblStatus = "Añadido " & txturl.Text
End Sub

Private Sub cmdDestino_Click()
'CARGAR CARPETA DE GUARDADO
Dim newdest As String
newdest = BrowseForFolder(txtDest.Text)
If newdest <> "" Then
txtDest.Text = newdest
End If
End Sub




Private Sub Form_Load()
'MsgBox Command$

Set FileList = New Collection
meWidth = 320
meHeight = 260
OnTop
Me.Caption = myCaption
txtDest = GetSetting(myApp, mySection, "Dir", "c:\")
txturl = GetSetting(myApp, mySection, "Ruta", "http://google.com/search?q=" & Rnd(9999))
lblStatus.Caption = "Preparado"
mnuAutoClean.Checked = True
wb1.Navigate "about:blank"

If Command$ <> "" Then txturl = Command$
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Exit Sub
Frame1.Top = Me.ScaleHeight - Frame1.Height
List1.Height = Me.ScaleHeight - Frame1.Height - Line1.Y1
Frame2.Top = Me.ScaleHeight - Frame2.Height
wb1.Height = Me.ScaleHeight - Frame2.Height - wb1.Top
wb1.Width = Me.ScaleWidth - wb1.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
'GUARDARR SETTINGS DE APP
SaveSetting myApp, mySection, "Dir", txtDest
SaveSetting myApp, mySection, "Ruta", txturl
Unload Me
End
End Sub

Private Sub List1_Click()
If doNoop = True Then doNoop = False: Exit Sub
'AL CLIKAR UN ITEM SE PONE EL URL EN EL TEXTBOX
If (List1.ListIndex = -1) Then Exit Sub
txturl.Text = FileList.item(List1.ListIndex + 1)
txturl.SelStart = Len(txturl)
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
n = List1.ListIndex
If n = -1 Then Exit Sub

List1.RemoveItem (n)
'LIST1 empieza en 0
'FILELIST empieza en 1
FileList.Remove (n + 1)
End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MENU PARA ELIMINAR COMPLETADOS O DESMARCADOS
If (Button = 2) Then
PopupMenu mn_CapLista
End If
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As String
If Data.GetFormat(vbCFLink) = True Then
    MsgBox "Esto es un vbCFlink, lo normal es que sea vbCFText" & vbCrLf & "Si, tengo que hacer que admita accesos directos a internet, para bajar los favoritos"
End If
If Data.GetFormat(vbCFFiles) = True Then
    MsgBox "Esto son ficheros del sistema?"
End If
If Data.GetFormat(vbCFRTF) = True Then
    MsgBox "Esto es RTF ? Que hace aqui? No problem"
End If
If Data.GetFormat(vbCFBitmap) = True Or Data.GetFormat(vbCFDIB) = True Or Data.GetFormat(vbCFDIB) = True Or Data.GetFormat(vbCFEMetafile) = True Or Data.GetFormat(vbCFMetafile) = True Then
    MsgBox "Esto es bitmap"
    MsgBox " " & Data.GetData(vbCFDIB)
End If
If Data.GetFormat(vbCFText) = True Then
    a = Data.GetData(vbCFText)
    addtoList a
End If
lblStatus = "Cargado Fichero"
End Sub

Private Sub addtoList(item As String)
'SE AÑADEN LOS URLS A LA LISTAS
    List1.AddItem baseName(item)
    numURLS = List1.ListCount + 1
    FileList.Add item, "U" & numURLS
    doNoop = True
    List1.Selected(List1.ListCount - 1) = True
'    Duplicados List1, FileList

End Sub

Public Sub Duplicados(listbox As listbox, listf As Collection)
Dim Search1 As Long
Dim Search2 As Long
'Dim KillDupe As Long
Dim numTot As Long
On Error GoTo errDup
    KillDupe = 0
    numTot = listbox.ListCount
    If numTot <> FileList.Count Then
        MsgBox "PROBLEMA : EL LIST Y EL ARRAY DE FICHEROS TIENEN DIFERENTES CANTIDADES : " & numTot & "<>" & FileList.Count
    End If
    lblStatus = "COmprobando repetidos"
    For sa = 1 To numTot
        sb = sa + 1
        For sb = sb To numTot
            If listf.item(sa) = listf.item(sb) Then
                Debug.Print listf.item(sb)
                Debug.Print listbox.List(sb - 1)
                'collections empiezan en 1
                listf.Remove sb
                'listbox empiezan en -1
                listbox.RemoveItem (sb - 1)
                'sb=sb-1
                'TOTALES
                Debug.Print listf.Count
                Debug.Print listbox.ListCount
                numTot = numTot - 1
                Debug.Print numTot
                xmenosde = xmenosde + 1
            End If
        Next
    Next
    MsgBox "Habia " & xmenosde & " Duplicados"
Exit Sub
errDup:
Y = MsgBox("Hubo error mio. Han Quedado Duplicados." & vbclrf & "Se han eliminado " & xmenosde & vbCrLf & vbCrLf & " Deseas Repetir El Eliminado de Duplicados?", vbYesNo + vbQuestion, "MIni Error")
If Y = vbYes Then
Duplicados List1, FileList
End If
End Sub

Private Sub mnuAutoClean_Click()
autoclean = True
End Sub

Private Sub mnuLimpiarComp_Click()
Dim ntotalborrados As Integer
While i < List1.ListCount
    If Right(List1.List(i), 5) = (msgSep & msgOk) Then
        List1.RemoveItem i
        FileList.Remove (i + 1)
        ntotalborrados = ntotalborrados + 1
        i = i - 1
    End If
    i = i + 1
Wend
lblStatus = "Limpiados " & ntotalborrados
End Sub

Private Sub mnuLimpiarDesmarcados_Click()
Dim ntotalborrados As Integer
Dim i As Integer
While i < List1.ListCount
    If List1.Selected(i) = False Then
        List1.RemoveItem i
        FileList.Remove (i + 1)
        ntotalborrados = ntotalborrados + 1
        i = i - 1
    End If
    i = i + 1
Wend
lblStatus = "Quitados " & ntotalborrados
End Sub

Private Sub mnuLimpiarTODO_Click()
Set FileList = New Collection
List1.Clear
End Sub

Private Sub mnuLoadURLs_Click()
Dim tPath As String
Dim tLoadFile As String
Dim dato As String

frmDropFiles.Show 1, Me
If frmDropFiles.ficheros.Count = 0 Then
 MsgBox "No has elegido fichero para Cargar. Nothin", vbQuestion
 Exit Sub
Else
    tLoadFile = frmDropFiles.ficheros(1)
End If

Dim f As Integer
f = FreeFile
Open tLoadFile For Input As #f
Do While Not EOF(f)
Line Input #f, dato
addtoList dato
Loop
Close #f
End Sub

Private Sub mnuOpenDest_Click()
Shell "explorer.exe " & txtDest, vbNormalFocus
End Sub

Private Sub mnuSaveURLList_Click()
Dim tPath As String
Dim tSaveFile As String
tPath = App.Path
tSaveFile = "LIST_URLS[" & Format(Now, "ddMMMyy-hh-mm") & "].txt"
Dim f As Integer
f = FreeFile
Open tPath & "\" & tSaveFile For Output As #f
For Each i In FileList
    Print #f, i
Next
Close #f
If MsgBox("Lista Guardada En:" & vbCrLf & tPath & "\" & tSaveFile & vbCrLf & vbCrLf & "Deseas Abrirla?", vbInformation + vbYesNo, "Done.Open It?") = vbYes Then
    Shell "notepad.exe """ & tSaveFile & """"
End If
End Sub

Private Sub txturl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAdd_Click
End If
End Sub

Private Sub cmdLeech_Click()
If txturl = "" Then Exit Sub
wb1.Offline = False
Me.MousePointer = 11
If Me.WindowState = vbNormal Then
Me.Width = 9210
End If
makeLeech = True
wb1.Navigate txturl
lblStatus = "El navegador esta cargando la pagina..."
MsgBox "Cuando termine de cargar pulsa el boton ReadLinks", vbInformation + vbOKOnly
Me.MousePointer = 0
cmdReadAll.Enabled = False
cmdReadImages.Enabled = False
End Sub

Private Sub cmdReadAll_Click()
lblStatus.Caption = "Haciendo Leech...cof cof"
leerLinks
End Sub

Function leerLinks()
Dim numLinks As Long
'Dim thisCol As New Collection
Dim lnk As String
On Error GoTo ErrLeech

If makeLeech = False Then GoTo AUNNO
If wb1.ReadyState <= 3 Then GoTo AUNNO
bHaciendo = True
For Each i In wb1.Document.body.All
    If bHaciendo = False Then Exit Function
    If i.tagname = "A" Then
        If i.href <> "" Then
            pnot = InStr(1, i.href, "#")
            If pnot >= 1 Then
                lnk = Left(i.href, pnot - 1)
            Else
                lnk = i.href
            End If
            addtoList lnk
            'thisCol.Add item

            numLinks = numLinks + 1
        End If
    End If
    Debug.Print i.tagname
Next

Duplicados List1, FileList
MsgBox "Se han encontrado " & numLinks & " Enlaces Para Bajar"
lblStatus = numLinks & " Paginas añadidas"
Exit Function
ErrLeech:
    MsgBox Err.Number & " Mi Error Leech - " & Err.Description, vbCritical + vbOKOnly, "Err En Leech"
Exit Function
AUNNO:
    MsgBox "No has pulsado el boton  de abrir pagina o aun no se ha cargado la pagina"
End Function
Private Sub cmdReadImages_Click()
'If MsgBox("Deseas bajar tb las imagenes de cada pagina?", vbQuestion + vbYesNo) = vbNo Then Exit Function
On Error GoTo ErrImag
For Each i In FileList
    numIma = 0
    wb1.Navigate i
    DoEvents
    lblStatus.Caption = "Imaging:" & i
    If bHaciendo = False Then Exit Sub
    Do While wb1.ReadyState < 4
        DoEvents
    Loop
    lblStatus.Caption = "Leeching Images"
    For Each im In wb1.Document.All
    lblStatus.Caption = "Leyendo Doc All"
        If im.tagname = "IMG" Then
            lblStatus.Caption = "Now IMG"
            If InStr(1, im.src, "http:") Then
            addtoList im.src
            numIma = numIma + 1
            End If
        End If
    Next
    lblStatus.Caption = numIma & " Imagenes"
Next

Exit Sub
ErrImag:
    MsgBox Err.Number & " Mi Error Image - " & Err.Description & " " & Err.Source, vbCritical + vbOKCancel, "Error En Imagenes"

End Sub
Private Sub cmdREadEmbeds_Click()

Dim numLinks As Long
'On Error GoTo ErrLeech

If makeLeech = False Then GoTo AUNNO
If wb1.ReadyState <= 3 Then GoTo AUNNO
bHaciendo = True

For Each i In wb1.Document.body.All
    If bHaciendo = False Then Exit Sub
    Debug.Print i.tagname
    If UCase(i.tagname) = "OBJECT" Then
        If i.src <> "" Then
            addtoList i.href
            numLinks = numLinks + 1
        End If
    End If
Next


MsgBox "Se han encontrado " & numLinks & " EMBEDS Para Bajar"
lblStatus = numLinks & " Paginas añadidas"
Exit Sub
ErrLeech:
    MsgBox Err.Number & " Mi Error Leech - " & Err.Description & " " & Err.Source, vbCritical + vbOKOnly, "Err En Leech"
Exit Sub
AUNNO:
    MsgBox "No has pulsado el boton  de abrir pagina o aun no se ha cargado la pagina"
End Sub

Private Sub cmdStop_Click()
bHaciendo = False
wb1.Stop
cmdReadAll.Enabled = True
cmdReadImages.Enabled = True

End Sub

Private Sub wb1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
cmdReadAll.Enabled = True
cmdReadImages.Enabled = True
End Sub

Private Sub wb1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
If wb1.Busy = True Then
lblStatus.Caption = "Progreso ... " & Progress & "/" & ProgressMax
End If
End Sub

